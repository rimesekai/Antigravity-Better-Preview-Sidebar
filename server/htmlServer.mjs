import { randomBytes } from 'node:crypto';
import { createReadStream, realpathSync, statSync } from 'node:fs';
import { createServer } from 'node:http';
import { dirname, extname, relative, resolve, sep } from 'node:path';
import {
  HttpError,
  findDeepestWorkspaceRoot,
  isForbiddenSystemPath,
  isPathWithinRoot,
} from './pathSecurity.mjs';

const MAX_ACTIVE_MOUNTS = 32;

export const MIME_TYPES = Object.freeze({
  '.html': 'text/html; charset=utf-8',
  '.htm': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.mjs': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.ico': 'image/x-icon',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.ttf': 'font/ttf',
  '.wasm': 'application/wasm',
});

/** @type {Map<string, { mountRoot: string, entryRelPath: string, createdAt: number }>} */
const mounts = new Map();

/** @type {import('node:http').Server | null} */
let httpServer = null;
/** @type {Promise<number> | null} */
let listeningPromise = null;
let serverPort = 0;

function sendText(res, status, text) {
  res.writeHead(status, {
    'Content-Type': 'text/plain; charset=utf-8',
    'X-Content-Type-Options': 'nosniff',
    'Cache-Control': 'no-store',
  });
  res.end(text);
}

/**
 * 解码并校验挂载路径中的相对段，阻止 .. 穿越与空字节。
 * @param {string} rawRelPath
 * @returns {string[] | null}
 */
function parseSafeSegments(rawRelPath) {
  if (!rawRelPath) return null;
  const parts = rawRelPath.split('/');
  const decodedParts = [];

  for (const part of parts) {
    if (!part) return null;
    let decoded;
    try {
      decoded = decodeURIComponent(part);
    } catch {
      return 'BAD_ENCODING';
    }
    if (!decoded || decoded.includes('\0') || decoded.includes('\\') || decoded.includes('/')) {
      return null;
    }
    if (decoded === '..' || decoded === '.') {
      return null;
    }
    decodedParts.push(decoded);
  }

  return decodedParts;
}

/**
 * 处理独立端口 HTML 静态预览请求：GET/HEAD /m/<mountId>/<relPath...>
 * @param {import('node:http').IncomingMessage} req
 * @param {import('node:http').ServerResponse} res
 */
function handlePreviewRequest(req, res) {
  if (req.method !== 'GET' && req.method !== 'HEAD') {
    sendText(res, 405, 'Method Not Allowed');
    return;
  }

  const rawPathname = (req.url || '/').split('?')[0].split('#')[0];
  const match = /^\/m\/([0-9a-f]{48})\/(.+)$/.exec(rawPathname);
  if (!match) {
    sendText(res, 404, 'Not Found');
    return;
  }

  const mountId = match[1];
  const rawRelPath = match[2];

  const mount = mounts.get(mountId);
  if (!mount) {
    sendText(res, 404, 'Mount Not Found');
    return;
  }

  // 刷新 LRU 位置
  mounts.delete(mountId);
  mounts.set(mountId, mount);

  const segments = parseSafeSegments(rawRelPath);
  if (segments === 'BAD_ENCODING') {
    sendText(res, 400, 'Bad Request');
    return;
  }
  if (!segments) {
    sendText(res, 403, 'Forbidden');
    return;
  }

  const candidate = resolve(mount.mountRoot, ...segments);
  if (!isPathWithinRoot(candidate, mount.mountRoot) || isForbiddenSystemPath(candidate)) {
    sendText(res, 403, 'Forbidden');
    return;
  }

  let realCandidate;
  try {
    realCandidate = realpathSync(candidate);
  } catch (err) {
    if (err && (err.code === 'ENOENT' || err.code === 'ENOTDIR')) {
      sendText(res, 404, 'Not Found');
      return;
    }
    sendText(res, 403, 'Forbidden');
    return;
  }

  if (!isPathWithinRoot(realCandidate, mount.mountRoot) || isForbiddenSystemPath(realCandidate)) {
    sendText(res, 403, 'Forbidden');
    return;
  }

  let stat;
  try {
    stat = statSync(realCandidate);
  } catch {
    sendText(res, 404, 'Not Found');
    return;
  }

  if (!stat.isFile()) {
    sendText(res, 404, 'Not Found');
    return;
  }

  const ext = extname(realCandidate).toLowerCase();
  const contentType = MIME_TYPES[ext] || 'application/octet-stream';

  res.writeHead(200, {
    'Content-Type': contentType,
    'Content-Length': stat.size,
    'X-Content-Type-Options': 'nosniff',
    'Cache-Control': 'no-store',
  });

  if (req.method === 'HEAD') {
    res.end();
    return;
  }

  const stream = createReadStream(realCandidate);
  stream.on('error', () => {
    if (!res.headersSent) {
      sendText(res, 500, 'Internal Server Error');
    } else {
      res.destroy();
    }
  });
  stream.pipe(res);
}

/**
 * 按需懒启动 127.0.0.1:0 原生静态 HTTP 服务。
 * @returns {Promise<number>}
 */
export async function ensureHtmlServerStarted() {
  if (httpServer && serverPort > 0) {
    return serverPort;
  }
  if (listeningPromise) {
    return listeningPromise;
  }

  listeningPromise = new Promise((resolvePromise, rejectPromise) => {
    const server = createServer(handlePreviewRequest);
    server.once('error', (err) => {
      listeningPromise = null;
      rejectPromise(err);
    });
    server.listen(0, '127.0.0.1', () => {
      const addr = server.address();
      if (!addr || typeof addr === 'string') {
        listeningPromise = null;
        rejectPromise(new Error('无法获取 HTML 预览服务监听端口'));
        return;
      }
      httpServer = server;
      serverPort = addr.port;
      resolvePromise(serverPort);
    });
  });

  return listeningPromise;
}

/**
 * 将合法的本地 .html/.htm 文件注册为独立跨源预览挂载点。
 * 若文件位于工作区根目录内，以该工作区根目录为 mountRoot，否则限定在其所在目录。
 * @param {string} resolvedFilePath 已通过安全校验的真实文件路径
 * @param {string[]} [allowedWorkspaceRoots]
 * @returns {Promise<{ url: string, mountId: string, entryRelPath: string, mountRoot: string }>}
 */
export async function mountHtmlTarget(resolvedFilePath, allowedWorkspaceRoots = []) {
  const realFilePath = realpathSync(resolvedFilePath);
  const ext = extname(realFilePath).toLowerCase();
  if (ext !== '.html' && ext !== '.htm') {
    throw new HttpError(415, '仅支持挂载 .html 或 .htm 文件', 'INVALID_HTML_TARGET');
  }

  const normalizedWorkspaceRoots = [];
  for (const root of allowedWorkspaceRoots) {
    try {
      normalizedWorkspaceRoots.push(realpathSync(root));
    } catch {
      // 忽略不存在的候选工作区路径
    }
  }

  const matchedRoot = findDeepestWorkspaceRoot(realFilePath, normalizedWorkspaceRoots);
  const mountRoot = matchedRoot || dirname(realFilePath);

  const relPathNative = relative(mountRoot, realFilePath);
  if (!relPathNative || relPathNative.startsWith('..')) {
    throw new HttpError(403, '入口文件不在挂载根目录内', 'MOUNT_BOUNDARY_ERROR');
  }

  const entrySegments = relPathNative.split(sep).filter(Boolean);
  const entryRelPath = entrySegments.join('/');
  const encodedEntry = entrySegments.map((seg) => encodeURIComponent(seg)).join('/');

  const port = await ensureHtmlServerStarted();
  const mountId = randomBytes(24).toString('hex');

  mounts.set(mountId, {
    mountRoot,
    entryRelPath,
    createdAt: Date.now(),
  });

  // LRU 上限控制，保留最近 32 个活跃挂载
  while (mounts.size > MAX_ACTIVE_MOUNTS) {
    const oldestKey = mounts.keys().next().value;
    if (!oldestKey) break;
    mounts.delete(oldestKey);
  }

  return {
    url: `http://127.0.0.1:${port}/m/${mountId}/${encodedEntry}`,
    mountId,
    entryRelPath,
    mountRoot,
  };
}

/**
 * 获取当前活跃挂载数量（供测试与监控使用）。
 * @returns {number}
 */
export function getActiveMountCount() {
  return mounts.size;
}

/**
 * 关闭独立 HTML 静态预览服务并清理所有挂载点。
 * @returns {Promise<void>}
 */
export async function closeHtmlServer() {
  mounts.clear();
  listeningPromise = null;
  serverPort = 0;

  if (!httpServer) {
    return;
  }

  const server = httpServer;
  httpServer = null;

  await new Promise((resolveClose) => {
    server.close(() => resolveClose());
  });
}
