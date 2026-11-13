import assert from 'node:assert/strict';
import {
  mkdirSync,
  mkdtempSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from 'node:fs';
import http from 'node:http';
import { registerHooks } from 'node:module';
import { homedir, tmpdir } from 'node:os';
import { join } from 'node:path';
import { after, before, describe, it } from 'node:test';

// 注册轻量 sidecar_sdk mock，模拟 Antigravity 宿主注入的 SidecarApp 与 Response 精确行为
registerHooks({
  resolve(specifier, context, nextResolve) {
    if (specifier === 'sidecar_sdk') {
      return {
        shortCircuit: true,
        url: 'mock://sidecar_sdk',
      };
    }
    return nextResolve(specifier, context);
  },
  load(url, context, nextLoad) {
    if (url === 'mock://sidecar_sdk') {
      return {
        shortCircuit: true,
        format: 'module',
        source: `
          import { createServer } from 'node:http';

          export class Response {
            constructor(body, { contentType = 'text/plain', status = 200, headers } = {}) {
              this.body = body;
              this.contentType = contentType;
              this.status = status;
              this.headers = headers;
            }
          }

          export class SidecarApp {
            #token;
            #port;
            #server;
            #pages = new Map();
            #apis = new Map();

            constructor() {
              const portStr = process.env.ANTIGRAVITY_SIDECAR_WEB_PORT || '0';
              this.#token = process.env.ANTIGRAVITY_SIDECAR_UI_TOKEN;
              this.#port = parseInt(portStr, 10);
            }

            page(path, handler) {
              this.#pages.set(path, handler);
            }

            api(path, handler, method = 'POST') {
              if (!this.#apis.has(path)) {
                this.#apis.set(path, new Map());
              }
              this.#apis.get(path).set(method.toUpperCase(), handler);
            }

            run() {
              const server = createServer(async (req, res) => {
                const url = new URL(req.url, 'http://127.0.0.1');
                if (req.method !== 'GET' && this.#token) {
                  const provided = req.headers['x-sidecar-token'] || url.searchParams.get('token');
                  if (provided !== this.#token) {
                    res.writeHead(401, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ error: 'unauthorized' }));
                    return;
                  }
                }

                if (req.method === 'GET' && this.#pages.has(url.pathname)) {
                  const content = await this.#pages.get(url.pathname)();
                  if (content instanceof Response) {
                    res.writeHead(content.status || 200, {
                      'Content-Type': content.contentType,
                      ...content.headers,
                    });
                    res.end(content.body);
                    return;
                  }
                  const isJS = url.pathname.endsWith('.js');
                  res.writeHead(200, {
                    'Content-Type': isJS ? 'application/javascript' : 'text/html',
                  });
                  res.end(content);
                  return;
                }

                if (this.#apis.has(url.pathname)) {
                  const methodMap = this.#apis.get(url.pathname);
                  if (methodMap.has(req.method)) {
                    const handler = methodMap.get(req.method);
                    let body = '';
                    req.on('data', (chunk) => { body += chunk; });
                    req.on('end', async () => {
                      try {
                        const data = {};
                        for (const [k, v] of url.searchParams.entries()) {
                          if (k.toLowerCase() !== 'token') data[k] = v;
                        }
                        if (body) Object.assign(data, JSON.parse(body));
                        const result = await handler(data);
                        if (
                          result instanceof Response ||
                          (result && typeof result === 'object' && result.contentType && result.body !== undefined)
                        ) {
                          const status = result.status || 200;
                          const respBody = typeof result.body === 'string' ? result.body : Buffer.from(result.body);
                          res.writeHead(status, {
                            'Content-Type': result.contentType,
                            ...result.headers,
                          });
                          res.end(respBody);
                          return;
                        }
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify(result));
                      } catch (err) {
                        res.writeHead(500, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: err.message }));
                      }
                    });
                    return;
                  }
                  res.writeHead(405, { 'Content-Type': 'application/json' });
                  res.end(JSON.stringify({ error: 'Method Not Allowed' }));
                  return;
                }

                res.writeHead(404);
                res.end('Not found');
              });

              return new Promise((resolve) => {
                this.#server = server.listen(this.#port, '127.0.0.1', () => resolve(server));
              });
            }

            close() {
              return new Promise((resolve) => (this.#server ? this.#server.close(() => resolve()) : resolve()));
            }
          }
        `,
      };
    }
    return nextLoad(url, context);
  },
});

const {
  HttpError,
  isForbiddenSystemPath,
  isPathWithinRoot,
  parseInputPath,
  validateFilePath,
} = await import('../server/pathSecurity.mjs');

const {
  closeHtmlServer,
  getActiveMountCount,
  mountHtmlTarget,
} = await import('../server/htmlServer.mjs');

const { scanWorkspaceFiles } = await import('../server/workspaceScanner.mjs');
const { startServer } = await import('../server/index.mjs');

function rawHttpGet(urlString, rawPath) {
  const url = new URL(urlString);
  return new Promise((resolve, reject) => {
    const req = http.request(
      {
        hostname: url.hostname,
        port: Number(url.port),
        path: rawPath,
        method: 'GET',
      },
      (res) => {
        const chunks = [];
        res.on('data', (chunk) => chunks.push(chunk));
        res.on('end', () => {
          resolve({
            status: res.statusCode,
            headers: res.headers,
            body: Buffer.concat(chunks),
          });
        });
      },
    );
    req.on('error', reject);
    req.end();
  });
}

describe('pathSecurity 路径安全与格式分类测试', () => {
  let tempRoot;
  let workspaceDir;
  let outsideDir;

  before(() => {
    tempRoot = mkdtempSync(join(tmpdir(), 'fp-sec-test-'));
    workspaceDir = join(tempRoot, 'workspace');
    outsideDir = join(tempRoot, 'outside');
    mkdirSync(workspaceDir, { recursive: true });
    mkdirSync(outsideDir, { recursive: true });
    mkdirSync(join(workspaceDir, '.system_generated', 'logs'), { recursive: true });
    mkdirSync(join(workspaceDir, '.system_generated', 'messages'), { recursive: true });

    writeFileSync(join(workspaceDir, 'deck.pptx'), Buffer.from('PK\x03\x04pptx'));
    writeFileSync(join(workspaceDir, 'report.DOCX'), Buffer.from('PK\x03\x04docx'));
    writeFileSync(join(workspaceDir, 'page.html'), '<h1>Hello</h1>');
    writeFileSync(join(workspaceDir, 'legacy.doc'), 'legacy doc');
    writeFileSync(join(workspaceDir, 'legacy.ppt'), 'legacy ppt');
    writeFileSync(join(workspaceDir, 'notes.txt'), 'plain');
    writeFileSync(join(workspaceDir, '.system_generated', 'logs', 'run.html'), '<p>log</p>');
    writeFileSync(join(workspaceDir, '.system_generated', 'messages', 'msg.html'), '<p>msg</p>');
    writeFileSync(join(outsideDir, 'secret.html'), '<p>secret</p>');
  });

  after(() => {
    rmSync(tempRoot, { recursive: true, force: true });
  });

  it('正确解析 file:// URI 与 ~/ 路径并拒绝空字节', () => {
    const parsedFromUri = parseInputPath(`file:///${join(workspaceDir, 'deck.pptx').replace(/\\/g, '/')}`);
    assert.equal(parsedFromUri.toLowerCase(), join(workspaceDir, 'deck.pptx').toLowerCase());

    const homeExpanded = parseInputPath('~/demo.docx');
    assert.equal(homeExpanded, join(homedir(), 'demo.docx'));

    assert.throws(
      () => parseInputPath('C:\\foo\\bar\0.pptx'),
      (err) => err instanceof HttpError && err.status === 400,
    );
  });

  it('正确处理 Windows 大小写不敏感与盘符根目录 D:\\ 不产生双分隔符', () => {
    assert.equal(isPathWithinRoot('d:\\Projects\\Demo\\a.pptx', 'D:\\', 'win32'), true);
    assert.equal(isPathWithinRoot('D:\\projects\\demo\\a.pptx', 'd:\\Projects', 'win32'), true);
    assert.equal(isPathWithinRoot('D:\\ProjectsOther\\a.pptx', 'D:\\Projects', 'win32'), false);
    assert.equal(isPathWithinRoot('E:\\a.pptx', 'D:\\', 'win32'), false);
  });

  it('拦截 .system_generated/logs 与 .system_generated/messages', () => {
    assert.equal(isForbiddenSystemPath(join(workspaceDir, '.system_generated', 'logs', 'run.html')), true);
    assert.equal(isForbiddenSystemPath(join(workspaceDir, '.system_generated', 'messages', 'msg.html')), true);

    assert.throws(
      () => validateFilePath(join(workspaceDir, '.system_generated', 'logs', 'run.html'), {
        allowedRootsOverride: [workspaceDir],
      }),
      (err) => err instanceof HttpError && err.status === 403,
    );

    assert.throws(
      () => validateFilePath(join(workspaceDir, '.system_generated', 'messages', 'msg.html'), {
        allowedRootsOverride: [workspaceDir],
      }),
      (err) => err instanceof HttpError && err.status === 403,
    );
  });

  it('拦截白名单之外与路径穿越请求', () => {
    assert.throws(
      () => validateFilePath(join(workspaceDir, '..', 'outside', 'secret.html'), {
        allowedRootsOverride: [workspaceDir],
      }),
      (err) => err instanceof HttpError && err.status === 403,
    );
  });

  it('对旧版 .doc / .ppt 与未知格式返回 415 清晰中文提示', () => {
    assert.throws(
      () => validateFilePath(join(workspaceDir, 'legacy.doc'), { allowedRootsOverride: [workspaceDir] }),
      (err) => err instanceof HttpError && err.status === 415 && err.message.includes('.docx'),
    );

    assert.throws(
      () => validateFilePath(join(workspaceDir, 'legacy.ppt'), { allowedRootsOverride: [workspaceDir] }),
      (err) => err instanceof HttpError && err.status === 415 && err.message.includes('.pptx'),
    );

    assert.throws(
      () => validateFilePath(join(workspaceDir, 'notes.txt'), { allowedRootsOverride: [workspaceDir] }),
      (err) => err instanceof HttpError && err.status === 415,
    );
  });

  it('支持大小写混合扩展名分类与合法文件校验', () => {
    const ppt = validateFilePath(join(workspaceDir, 'deck.pptx'), { allowedRootsOverride: [workspaceDir] });
    assert.equal(ppt.kind, 'ppt');
    assert.equal(ppt.ext, '.pptx');

    const docx = validateFilePath(join(workspaceDir, 'report.DOCX'), { allowedRootsOverride: [workspaceDir] });
    assert.equal(docx.kind, 'word');
    assert.equal(docx.ext, '.docx');

    const html = validateFilePath(join(workspaceDir, 'page.html'), { allowedRootsOverride: [workspaceDir] });
    assert.equal(html.kind, 'html');
    assert.equal(html.ext, '.html');
  });

  it('拦截指向白名单外部的目录 junction / 符号链接逃逸', () => {
    const linkDir = join(workspaceDir, 'escape-dir');
    symlinkSync(outsideDir, linkDir, 'junction');

    assert.throws(
      () => validateFilePath(join(linkDir, 'secret.html'), { allowedRootsOverride: [workspaceDir] }),
      (err) => err instanceof HttpError && err.status === 403 && err.code === 'SYMLINK_ESCAPE',
    );
  });
});

describe('htmlServer 独立静态预览服务测试', () => {
  let tempRoot;
  let workspaceDir;
  let outsideDir;

  before(() => {
    tempRoot = mkdtempSync(join(tmpdir(), 'fp-html-test-'));
    workspaceDir = join(tempRoot, 'workspace');
    outsideDir = join(tempRoot, 'outside');
    mkdirSync(join(workspaceDir, 'slides'), { recursive: true });
    mkdirSync(join(workspaceDir, 'assets'), { recursive: true });
    mkdirSync(outsideDir, { recursive: true });

    writeFileSync(
      join(workspaceDir, 'slides', 'index.html'),
      '<!doctype html><link rel="stylesheet" href="./style.css"><script src="./app.js"></script><img src="../assets/logo.svg">',
    );
    writeFileSync(join(workspaceDir, 'slides', 'style.css'), 'body { color: #38bdf8; }');
    writeFileSync(join(workspaceDir, 'slides', 'app.js'), 'console.log("ok");');
    writeFileSync(join(workspaceDir, 'assets', 'logo.svg'), '<svg xmlns="http://www.w3.org/2000/svg"></svg>');
    writeFileSync(join(outsideDir, 'secret.txt'), 'top-secret');
  });

  after(async () => {
    await closeHtmlServer();
    rmSync(tempRoot, { recursive: true, force: true });
  });

  it('支持挂载工作区并完整加载相对路径 CSS/JS/SVG，且带安全响应头', async () => {
    const mounted = await mountHtmlTarget(join(workspaceDir, 'slides', 'index.html'), [workspaceDir]);
    assert.match(mounted.mountId, /^[0-9a-f]{48}$/);
    assert.equal(mounted.entryRelPath, 'slides/index.html');

    const htmlRes = await fetch(mounted.url);
    assert.equal(htmlRes.status, 200);
    assert.match(htmlRes.headers.get('content-type'), /^text\/html/);
    assert.equal(htmlRes.headers.get('x-content-type-options'), 'nosniff');
    assert.equal(htmlRes.headers.get('cache-control'), 'no-store');

    const cssUrl = new URL('./style.css', mounted.url).href;
    const cssRes = await fetch(cssUrl);
    assert.equal(cssRes.status, 200);
    assert.match(cssRes.headers.get('content-type'), /^text\/css/);
    assert.match(await cssRes.text(), /#38bdf8/);

    const jsUrl = new URL('./app.js', mounted.url).href;
    const jsRes = await fetch(jsUrl);
    assert.equal(jsRes.status, 200);
    assert.match(jsRes.headers.get('content-type'), /^application\/javascript/);

    const svgUrl = new URL('../assets/logo.svg', mounted.url).href;
    const svgRes = await fetch(svgUrl);
    assert.equal(svgRes.status, 200);
    assert.equal(svgRes.headers.get('content-type'), 'image/svg+xml');
  });

  it('拦截原始 HTTP 请求中的 .. 与 %2e%2e 路径穿越并返回 403', async () => {
    const mounted = await mountHtmlTarget(join(workspaceDir, 'slides', 'index.html'), [workspaceDir]);

    const res1 = await rawHttpGet(mounted.url, `/m/${mounted.mountId}/../outside/secret.txt`);
    assert.equal(res1.status, 403);

    const res2 = await rawHttpGet(mounted.url, `/m/${mounted.mountId}/%2e%2e/outside/secret.txt`);
    assert.equal(res2.status, 403);
  });

  it('未处于工作区内时 mountRoot 严格收窄为入口文件所在目录，无法向上访问父目录', async () => {
    const mounted = await mountHtmlTarget(join(workspaceDir, 'slides', 'index.html'), []);
    assert.equal(mounted.entryRelPath, 'index.html');
    const res = await rawHttpGet(mounted.url, `/m/${mounted.mountId}/../assets/logo.svg`);
    assert.equal(res.status, 403);
  });

  it('LRU 挂载上限不超过 32 个', async () => {
    for (let i = 0; i < 36; i += 1) {
      await mountHtmlTarget(join(workspaceDir, 'slides', 'index.html'), [workspaceDir]);
    }
    assert.equal(getActiveMountCount(), 32);
  });
});

describe('workspaceScanner 工作区扫描测试', () => {
  let tempRoot;

  before(() => {
    tempRoot = mkdtempSync(join(tmpdir(), 'fp-scan-test-'));
    mkdirSync(join(tempRoot, 'docs', 'sub'), { recursive: true });
    mkdirSync(join(tempRoot, 'node_modules', 'pkg'), { recursive: true });
    mkdirSync(join(tempRoot, 'dist'), { recursive: true });
    mkdirSync(join(tempRoot, '.git'), { recursive: true });
    mkdirSync(join(tempRoot, 'd1', 'd2', 'd3', 'd4', 'd5', 'd6'), { recursive: true });

    writeFileSync(join(tempRoot, 'docs', 'a.docx'), 'docx');
    writeFileSync(join(tempRoot, 'docs', 'sub', 'b.pptx'), 'pptx');
    writeFileSync(join(tempRoot, 'index.htm'), '<h1>hi</h1>');
    writeFileSync(join(tempRoot, 'node_modules', 'pkg', 'ignored.html'), '<p>ignore</p>');
    writeFileSync(join(tempRoot, 'dist', 'ignored.html'), '<p>ignore</p>');
    writeFileSync(join(tempRoot, '.git', 'ignored.html'), '<p>ignore</p>');
    writeFileSync(join(tempRoot, 'd1', 'd2', 'd3', 'd4', 'd5', 'depth5.pptx'), 'ok');
    writeFileSync(join(tempRoot, 'd1', 'd2', 'd3', 'd4', 'd5', 'd6', 'depth6.pptx'), 'too deep');
  });

  after(() => {
    rmSync(tempRoot, { recursive: true, force: true });
  });

  it('跳过 node_modules / dist / .git 并限制最大递归深度为 5', () => {
    const files = scanWorkspaceFiles([`file:///${tempRoot.replace(/\\/g, '/')}`]);
    const relPaths = files.map((f) => f.relPath).sort();

    assert.deepEqual(relPaths, [
      'd1/d2/d3/d4/d5/depth5.pptx',
      'docs/a.docx',
      'docs/sub/b.pptx',
      'index.htm',
    ]);
  });
});

describe('SidecarApp HTTP 路由集成测试', () => {
  let tempRoot;
  let workspaceDir;
  let distDir;
  let assetsDir;
  let running;
  let baseUrl;
  const token = 'test-sidecar-token-123';

  before(async () => {
    tempRoot = mkdtempSync(join(tmpdir(), 'fp-api-test-'));
    workspaceDir = join(tempRoot, 'ws');
    distDir = join(tempRoot, 'dist');
    assetsDir = join(tempRoot, 'assets');
    mkdirSync(workspaceDir, { recursive: true });
    mkdirSync(distDir, { recursive: true });
    mkdirSync(assetsDir, { recursive: true });

    writeFileSync(join(distDir, 'index.html'), '<html><body>Mock UI</body></html>');
    writeFileSync(join(distDir, 'app.js'), 'console.log("ui");');
    writeFileSync(join(distDir, 'styles.css'), 'body{}');
    writeFileSync(join(assetsDir, 'logo.svg'), '<svg></svg>');

    writeFileSync(join(workspaceDir, 'demo.pptx'), Buffer.from([1, 2, 3, 4]));
    writeFileSync(join(workspaceDir, 'demo.html'), '<html><body>Demo</body></html>');
    writeFileSync(join(workspaceDir, 'old.doc'), 'old');

    process.env.ANTIGRAVITY_SIDECAR_WEB_PORT = '0';
    process.env.ANTIGRAVITY_SIDECAR_UI_TOKEN = token;

    running = await startServer({
      distDir,
      assetsDir,
      allowedRootsOverride: [workspaceDir],
    });
    const addr = running.server.address();
    baseUrl = `http://127.0.0.1:${addr.port}`;
  });

  after(async () => {
    await running.close();
    delete process.env.ANTIGRAVITY_SIDECAR_WEB_PORT;
    delete process.env.ANTIGRAVITY_SIDECAR_UI_TOKEN;
    rmSync(tempRoot, { recursive: true, force: true });
  });

  it('GET /api/health 与静态页面路由返回正确内容', async () => {
    const healthRes = await fetch(`${baseUrl}/api/health`);
    assert.equal(healthRes.status, 200);
    assert.deepEqual(await healthRes.json(), { status: 'ok', version: '0.1.0' });

    const pageRes = await fetch(`${baseUrl}/`);
    assert.equal(pageRes.status, 200);
    assert.match(await pageRes.text(), /Mock UI/);

    const logoRes = await fetch(`${baseUrl}/assets/logo.svg`);
    assert.equal(logoRes.status, 200);
    assert.equal(logoRes.headers.get('content-type'), 'image/svg+xml');
  });

  it('POST 接口强制校验 X-Sidecar-Token，未携带返回 401', async () => {
    const res = await fetch(`${baseUrl}/api/fs/stat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file: join(workspaceDir, 'demo.pptx'), workspaceUris: [workspaceDir] }),
    });
    assert.equal(res.status, 401);
  });

  it('POST /api/fs/stat、/api/fs/read、/api/fs/list、/api/html/mount 正常响应', async () => {
    const headers = {
      'Content-Type': 'application/json',
      'X-Sidecar-Token': token,
    };

    const statRes = await fetch(`${baseUrl}/api/fs/stat`, {
      method: 'POST',
      headers,
      body: JSON.stringify({ file: join(workspaceDir, 'demo.pptx'), workspaceUris: [workspaceDir] }),
    });
    assert.equal(statRes.status, 200);
    const statJson = await statRes.json();
    assert.equal(statJson.kind, 'ppt');
    assert.equal(statJson.ext, '.pptx');
    assert.equal(statJson.size, 4);

    const readRes = await fetch(`${baseUrl}/api/fs/read`, {
      method: 'POST',
      headers,
      body: JSON.stringify({ file: join(workspaceDir, 'demo.pptx'), workspaceUris: [workspaceDir] }),
    });
    assert.equal(readRes.status, 200);
    assert.equal(readRes.headers.get('content-type'), 'application/octet-stream');
    assert.equal(readRes.headers.get('x-content-type-options'), 'nosniff');
    const buf = Buffer.from(await readRes.arrayBuffer());
    assert.deepEqual(buf, Buffer.from([1, 2, 3, 4]));

    const listRes = await fetch(`${baseUrl}/api/fs/list`, {
      method: 'POST',
      headers,
      body: JSON.stringify({ workspaceUris: [workspaceDir] }),
    });
    assert.equal(listRes.status, 200);
    const listJson = await listRes.json();
    assert.equal(listJson.files.length, 2);

    const mountRes = await fetch(`${baseUrl}/api/html/mount`, {
      method: 'POST',
      headers,
      body: JSON.stringify({ file: join(workspaceDir, 'demo.html'), workspaceUris: [workspaceDir] }),
    });
    assert.equal(mountRes.status, 200);
    const mountJson = await mountRes.json();
    assert.equal(mountJson.entryRelPath, 'demo.html');

    const htmlFetch = await fetch(mountJson.url);
    assert.equal(htmlFetch.status, 200);
    assert.match(await htmlFetch.text(), /Demo/);

    // 旧版 .doc 返回 415
    const docStat = await fetch(`${baseUrl}/api/fs/stat`, {
      method: 'POST',
      headers,
      body: JSON.stringify({ file: join(workspaceDir, 'old.doc'), workspaceUris: [workspaceDir] }),
    });
    assert.equal(docStat.status, 415);
    const docErr = await docStat.json();
    assert.match(docErr.error, /\.docx/);
  });
});
