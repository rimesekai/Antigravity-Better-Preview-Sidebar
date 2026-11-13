/**
 * 使用 esbuild 将服务端入口打包为零依赖的 dist/server.mjs。
 */
import { mkdirSync } from 'node:fs';
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

console.log(`[build-server] 已生成 ${outFile}`);
