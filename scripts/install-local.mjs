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
  'LICENSE',
];

for (const entry of requiredEntries) {
  const srcPath = join(rootDir, entry);
  if (!existsSync(srcPath)) {
    console.error(`[install-local] 缺少必需文件或目录: ${srcPath}，请先执行构建。`);
    process.exit(1);
  }
}

mkdirSync(targetDir, { recursive: true });

for (const entry of requiredEntries) {
  const srcPath = join(rootDir, entry);
  const destPath = join(targetDir, entry);
  try {
    rmSync(destPath, { recursive: true, force: true });
  } catch {
    // Windows 下若目录句柄被宿主监听占用，跳过整目录删除，交由下方 cpSync force 覆盖
  }
  cpSync(srcPath, destPath, { recursive: true, force: true });
}

console.log(`[install-local] 已将插件安装至: ${targetDir}`);
