/**
 * 将插件运行时所需的精简文件安装到全局 Antigravity 插件目录：
 * ~/.gemini/config/plugins/file-preview/
 * 不复制 node_modules、src/、server/ 与测试文件。
 */
import { cpSync, existsSync, mkdirSync, rmSync } from 'node:fs';
import { homedir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const rootDir = resolve(__dirname, '..');
const targetDir = join(homedir(), '.gemini', 'config', 'plugins', 'file-preview');

const requiredEntries = [
  'plugin.json',
  'assets',
  'rules',
  'skills',
  'sidecars',
  'dist',
  'README.md',
];

for (const entry of requiredEntries) {
  const srcPath = join(rootDir, entry);
  if (!existsSync(srcPath)) {
    console.error(`[install-local] 缺少必需文件或目录: ${srcPath}，请先执行构建。`);
    process.exit(1);
  }
}

rmSync(targetDir, { recursive: true, force: true });
mkdirSync(targetDir, { recursive: true });

for (const entry of requiredEntries) {
  const srcPath = join(rootDir, entry);
  const destPath = join(targetDir, entry);
  cpSync(srcPath, destPath, { recursive: true });
}

console.log(`[install-local] 已将插件安装至: ${targetDir}`);
