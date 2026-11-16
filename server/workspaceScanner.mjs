import { readdirSync, realpathSync, statSync } from 'node:fs';
import { extname, join, relative, sep } from 'node:path';
import { isForbiddenSystemPath, parseWorkspaceRoot } from './pathSecurity.mjs';

const SKIP_DIRS = new Set([
  'node_modules',
  '.git',
  '.svn',
  '.venv',
  'dist',
  'build',
  '.next',
  '.system_generated',
]);

const SUPPORTED_EXTS = new Map([
  ['.pptx', 'ppt'],
  ['.docx', 'word'],
  ['.xlsx', 'excel'],
  ['.xls', 'excel'],
  ['.csv', 'excel'],
  ['.tsv', 'excel'],
  ['.pdf', 'pdf'],
  ['.mp4', 'mp4'],
  ['.html', 'html'],
  ['.htm', 'html'],
]);

const DEFAULT_MAX_DEPTH = 5;
const DEFAULT_LIMIT = 200;
const HARD_COLLECT_CAP = 2000;

/**
 * 判断文件名是否为 Office 临时锁文件或系统隐藏元数据文件（如 ~$*.docx、.~lock.*、._*）。
 * @param {string} name
 * @returns {boolean}
 */
function isIgnoredTempFile(name) {
  return name.startsWith('~$') || name.startsWith('.~') || name.startsWith('._');
}

/**
 * 递归扫描工作区中的可预览文件（.pptx / .docx / .xlsx / .xls / .csv / .tsv / .pdf / .mp4 / .html / .htm）。
 * @param {string[] | string} [workspaceUris]
 * @param {object} [options]
 * @param {number} [options.maxDepth] 最大递归深度（默认 5）
 * @param {number} [options.limit] 返回条数上限（默认 200）
 * @returns {Array<{
 *   path: string,
 *   name: string,
 *   ext: string,
 *   kind: 'ppt' | 'word' | 'excel' | 'html' | 'pdf' | 'mp4',
 *   size: number,
 *   mtimeMs: number,
 *   relPath: string,
 *   workspaceRoot: string,
 * }>}
 */
export function scanWorkspaceFiles(workspaceUris = [], options = {}) {
  const maxDepth = typeof options.maxDepth === 'number' ? options.maxDepth : DEFAULT_MAX_DEPTH;
  const limit = typeof options.limit === 'number' ? options.limit : DEFAULT_LIMIT;

  const rawList = Array.isArray(workspaceUris)
    ? workspaceUris
    : typeof workspaceUris === 'string'
      ? [workspaceUris]
      : [];

  const uniqueRoots = [];
  const seenRootKeys = new Set();

  for (const item of rawList) {
    const parsed = parseWorkspaceRoot(item);
    if (!parsed) continue;

    let realRoot;
    try {
      realRoot = realpathSync(parsed);
      const stat = statSync(realRoot);
      if (!stat.isDirectory()) continue;
    } catch {
      continue;
    }

    const key = process.platform === 'win32' ? realRoot.toLowerCase() : realRoot;
    if (seenRootKeys.has(key)) continue;
    seenRootKeys.add(key);
    uniqueRoots.push(realRoot);
  }

  const results = [];
  const seenFiles = new Set();
  const visitedDirs = new Set();

  for (const rootDir of uniqueRoots) {
    walkDirectory(rootDir, rootDir, 0);
    if (results.length >= HARD_COLLECT_CAP) break;
  }

  results.sort((a, b) => b.mtimeMs - a.mtimeMs);
  return results.slice(0, limit);

  /**
   * @param {string} currentDir
   * @param {string} rootDir
   * @param {number} depth
   */
  function walkDirectory(currentDir, rootDir, depth) {
    if (results.length >= HARD_COLLECT_CAP) return;
    if (isForbiddenSystemPath(currentDir)) return;

    const dirKey = process.platform === 'win32' ? currentDir.toLowerCase() : currentDir;
    if (visitedDirs.has(dirKey)) return;
    visitedDirs.add(dirKey);

    let entries;
    try {
      entries = readdirSync(currentDir, { withFileTypes: true });
    } catch {
      return;
    }

    for (const entry of entries) {
      if (results.length >= HARD_COLLECT_CAP) break;

      const fullPath = join(currentDir, entry.name);

      if (entry.isDirectory()) {
        if (depth >= maxDepth) continue;
        if (entry.name.startsWith('.') || SKIP_DIRS.has(entry.name.toLowerCase())) continue;
        walkDirectory(fullPath, rootDir, depth + 1);
        continue;
      }

      if (!entry.isFile()) continue;
      if (isIgnoredTempFile(entry.name)) continue;

      const ext = extname(entry.name).toLowerCase();
      const kind = SUPPORTED_EXTS.get(ext);
      if (!kind) continue;

      let realFile;
      let stat;
      try {
        realFile = realpathSync(fullPath);
        stat = statSync(realFile);
        if (!stat.isFile()) continue;
      } catch {
        continue;
      }

      const fileKey = process.platform === 'win32' ? realFile.toLowerCase() : realFile;
      if (seenFiles.has(fileKey)) continue;
      seenFiles.add(fileKey);

      results.push({
        path: realFile,
        name: entry.name,
        ext,
        kind,
        size: stat.size,
        mtimeMs: stat.mtimeMs,
        relPath: relative(rootDir, fullPath).split(sep).join('/'),
        workspaceRoot: rootDir,
      });
    }
  }
}
