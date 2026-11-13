import { existsSync, readFileSync } from 'node:fs';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { SidecarApp, Response } from 'sidecar_sdk';
import { closeHtmlServer, mountHtmlTarget, mountPdfTarget } from './htmlServer.mjs';
import { HttpError, validateFilePath } from './pathSecurity.mjs';
import { scanWorkspaceFiles } from './workspaceScanner.mjs';

const PLUGIN_VERSION = '0.1.0';
const MAX_READ_BYTES = 200 * 1024 * 1024; // 200 MB 安全上限
const MAX_PDF_BYTES = 1024 * 1024 * 1024; // 1 GB PDF 流式预览安全上限

/**
 * 解析插件根目录与 dist/ 产物目录（兼容从 server/index.mjs 与 dist/server.mjs 运行）。
 * @returns {{ pluginRoot: string, distDir: string, assetsDir: string }}
 */
export function resolvePluginPaths() {
  const currentDir = dirname(fileURLToPath(import.meta.url));
  const runningInDist =
    existsSync(join(currentDir, 'index.html')) || basename(currentDir) === 'dist';

  const pluginRoot = resolve(currentDir, '..');
  const distDir = runningInDist ? currentDir : join(pluginRoot, 'dist');
  const assetsDir = join(pluginRoot, 'assets');

  return { pluginRoot, distDir, assetsDir };
}

/**
 * 构造 JSON 响应。
 * @param {unknown} payload
 * @param {number} [status]
 * @param {typeof Response} [ResponseClass]
 * @returns {Response}
 */
function jsonResponse(payload, status = 200, ResponseClass = Response) {
  return new ResponseClass(JSON.stringify(payload), {
    contentType: 'application/json; charset=utf-8',
    status,
    headers: {
      'X-Content-Type-Options': 'nosniff',
      'Cache-Control': 'no-store',
    },
  });
}

/**
 * 包装 API 处理器，统一捕获 HttpError 与未知异常并转换为规范 HTTP 状态码。
 * @template T
 * @param {(data: Record<string, any>) => Promise<T> | T} handler
 * @param {typeof Response} [ResponseClass]
 * @returns {(data: Record<string, any>) => Promise<T | Response>}
 */
export function withErrorHandling(handler, ResponseClass = Response) {
  return async (data = {}) => {
    try {
      return await handler(data || {});
    } catch (err) {
      if (err instanceof HttpError) {
        return jsonResponse({ error: err.message, code: err.code }, err.status, ResponseClass);
      }
      const message = err instanceof Error ? err.message : '服务器内部错误';
      return jsonResponse({ error: message, code: 'INTERNAL_ERROR' }, 500, ResponseClass);
    }
  };
}

/**
 * 读取静态文件内容并构造 Response；若缺失则返回 404。
 * @param {string} filePath
 * @param {string} contentType
 * @param {typeof Response} [ResponseClass]
 * @returns {Response}
 */
function serveStaticFile(filePath, contentType, ResponseClass = Response) {
  if (!existsSync(filePath)) {
    return new ResponseClass('Not Found', {
      contentType: 'text/plain; charset=utf-8',
      status: 404,
      headers: {
        'X-Content-Type-Options': 'nosniff',
        'Cache-Control': 'no-store',
      },
    });
  }

  const content = readFileSync(filePath);
  return new ResponseClass(content, {
    contentType,
    status: 200,
    headers: {
      'X-Content-Type-Options': 'nosniff',
      'Cache-Control': 'no-store',
    },
  });
}

/**
 * 注册所有页面路由与 API 路由到指定的 SidecarApp 实例。
 * @param {SidecarApp} app
 * @param {object} [options]
 * @param {typeof Response} [options.ResponseClass]
 * @param {string} [options.distDir]
 * @param {string} [options.assetsDir]
 * @param {typeof mountHtmlTarget} [options.mountHtml]
 * @param {string[] | null} [options.allowedRootsOverride] 仅用于测试覆盖
 */
export function registerRoutes(app, options = {}) {
  const defaultPaths = resolvePluginPaths();
  const ResponseClass = options.ResponseClass || Response;
  const distDir = options.distDir || defaultPaths.distDir;
  const assetsDir = options.assetsDir || defaultPaths.assetsDir;
  const mountHtml = options.mountHtml || mountHtmlTarget;
  const mountPdf = options.mountPdf || mountPdfTarget;
  const allowedRootsOverride = options.allowedRootsOverride || null;

  // 静态页面与脚本资源路由（SidecarApp 为精确匹配）
  app.page('/', () => serveStaticFile(join(distDir, 'index.html'), 'text/html; charset=utf-8', ResponseClass));
  app.page('/index.html', () =>
    serveStaticFile(join(distDir, 'index.html'), 'text/html; charset=utf-8', ResponseClass),
  );
  app.page('/app.js', () =>
    serveStaticFile(join(distDir, 'app.js'), 'application/javascript; charset=utf-8', ResponseClass),
  );
  app.page('/pdf.worker.js', () =>
    serveStaticFile(join(distDir, 'pdf.worker.js'), 'application/javascript; charset=utf-8', ResponseClass),
  );
  app.page('/styles.css', () =>
    serveStaticFile(join(distDir, 'styles.css'), 'text/css; charset=utf-8', ResponseClass),
  );
  app.page('/assets/logo.svg', () =>
    serveStaticFile(join(assetsDir, 'logo.svg'), 'image/svg+xml', ResponseClass),
  );

  // 健康检查
  app.api(
    '/api/health',
    withErrorHandling(() => ({ status: 'ok', version: PLUGIN_VERSION }), ResponseClass),
    'GET',
  );

  // 文件元信息校验与读取
  app.api(
    '/api/fs/stat',
    withErrorHandling((data) => {
      const info = validateFilePath(data.file, {
        workspaceUris: data.workspaceUris,
        allowedRootsOverride,
      });
      return {
        path: info.path,
        name: info.name,
        ext: info.ext,
        kind: info.kind,
        size: info.size,
        mtimeMs: info.mtimeMs,
      };
    }, ResponseClass),
    'POST',
  );

  // 二进制流读取（用于 .pptx / .docx）
  app.api(
    '/api/fs/read',
    withErrorHandling((data) => {
      const info = validateFilePath(data.file, {
        workspaceUris: data.workspaceUris,
        allowedRootsOverride,
      });

      if (info.size > MAX_READ_BYTES) {
        throw new HttpError(413, '文件体积超过 200MB 预览上限', 'FILE_TOO_LARGE');
      }

      const buffer = readFileSync(info.path);
      return new ResponseClass(buffer, {
        contentType: 'application/octet-stream',
        status: 200,
        headers: {
          'X-Content-Type-Options': 'nosniff',
          'Cache-Control': 'no-store',
        },
      });
    }, ResponseClass),
    'POST',
  );

  // 扫描工作区中的可预览文件
  app.api(
    '/api/fs/list',
    withErrorHandling((data) => {
      const files = scanWorkspaceFiles(data.workspaceUris);
      return { files };
    }, ResponseClass),
    'POST',
  );

  // 挂载 HTML 文件到独立跨源预览子端口
  app.api(
    '/api/html/mount',
    withErrorHandling(async (data) => {
      const info = validateFilePath(data.file, {
        workspaceUris: data.workspaceUris,
        allowedRootsOverride,
      });

      if (info.kind !== 'html') {
        throw new HttpError(415, '仅支持将 .html 或 .htm 文件挂载到预览服务', 'NOT_HTML_FILE');
      }

      const mounted = await mountHtml(info.path, info.workspaceRoots);
      return {
        url: mounted.url,
        entryRelPath: mounted.entryRelPath,
        mountId: mounted.mountId,
        mountRoot: mounted.mountRoot,
      };
    }, ResponseClass),
    'POST',
  );

  // 挂载 PDF 文件到独立端口 HTTP Range 流式预览服务
  app.api(
    '/api/pdf/mount',
    withErrorHandling(async (data) => {
      const info = validateFilePath(data.file, {
        workspaceUris: data.workspaceUris,
        allowedRootsOverride,
      });

      if (info.kind !== 'pdf') {
        throw new HttpError(415, '仅支持将 .pdf 文件挂载到流式预览服务', 'NOT_PDF_FILE');
      }

      if (info.size > MAX_PDF_BYTES) {
        throw new HttpError(413, 'PDF 文件体积超过 1GB 预览上限', 'FILE_TOO_LARGE');
      }

      const mounted = await mountPdf(info.path);
      return {
        url: mounted.url,
        mountId: mounted.mountId,
        fileName: info.name,
        size: info.size,
      };
    }, ResponseClass),
    'POST',
  );

  return app;
}

/**
 * 创建并配置 SidecarApp 实例。
 * @param {object} [options]
 * @returns {SidecarApp}
 */
export function createSidecarApp(options = {}) {
  const SidecarAppClass = options.SidecarAppClass || SidecarApp;
  const app = new SidecarAppClass();
  return registerRoutes(app, options);
}

/**
 * 启动 File Preview Sidecar 主服务。
 * @param {object} [options]
 * @returns {Promise<{ app: SidecarApp, server: import('node:http').Server, close: () => Promise<void> }>}
 */
export async function startServer(options = {}) {
  const app = createSidecarApp(options);
  const server = await app.run();

  const close = async () => {
    await closeHtmlServer();
    if (typeof app.close === 'function') {
      await app.close();
    }
  };

  return { app, server, close };
}

// 当直接作为主模块运行时自动启动
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  startServer().catch((err) => {
    console.error('[file-preview] 启动失败:', err);
    process.exit(1);
  });
}
