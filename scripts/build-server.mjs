/**
 * 使用 esbuild 将服务端入口打包为零依赖的 dist/server.mjs，
 * 并复制 pdfjs-dist 的 Worker 文件到 dist/pdf.worker.js。
 */
import { copyFileSync, existsSync, mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';

const __dirname = dirname(fileURLToPath(import.meta.url));
const rootDir = resolve(__dirname, '..');
const outFile = resolve(rootDir, 'dist', 'server.mjs');

mkdirSync(dirname(outFile), { recursive: true });

await build({
  entryPoints: [resolve(rootDir, 'server', 'index.mjs')],
  outfile: outFile,
  bundle: true,
  minify: true,
  platform: 'node',
  target: 'node20',
  format: 'esm',
  external: ['sidecar_sdk'],
  banner: {
    js: "import { createRequire as __bundleRequire } from 'node:module'; const require = __bundleRequire(import.meta.url);",
  },
});

const pdfWorkerSrc = resolve(rootDir, 'node_modules', 'pdfjs-dist', 'build', 'pdf.worker.min.js');
const pdfWorkerDest = resolve(rootDir, 'dist', 'pdf.worker.js');
if (!existsSync(pdfWorkerSrc)) {
  throw new Error(`[build-server] 缺少 pdf.worker.min.js: ${pdfWorkerSrc}`);
}
copyFileSync(pdfWorkerSrc, pdfWorkerDest);

console.log(`[build-server] 已生成 ${outFile}`);
console.log(`[build-server] 已复制 ${pdfWorkerDest}`);
