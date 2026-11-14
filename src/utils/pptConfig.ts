import {
  DEFAULT_EMBEDDED_FONT_LIMITS,
  RECOMMENDED_ZIP_LIMITS,
} from '@aiden0z/pptx-renderer/browser';
import type {
  EmbeddedFontLimits,
  ZipParseLimits,
} from '@aiden0z/pptx-renderer/browser';
import { tr } from './i18n';

const KB = 1024;
const MB = 1024 * KB;
const GB = 1024 * MB;

/**
 * 面向大体积/多条目 PPTX 课件的 ZIP 解析限额。
 * 保持与桌面端可用内存匹配的安全边界，同时放宽默认 4000 条目与媒体上限。
 */
export const PPTX_ZIP_LIMITS: Readonly<Required<ZipParseLimits>> = Object.freeze({
  ...RECOMMENDED_ZIP_LIMITS,
  maxEntries: 50_000,
  maxEntryUncompressedBytes: 256 * MB,
  maxTotalUncompressedBytes: 2 * GB,
  maxMediaBytes: 1536 * MB,
  maxConcurrency: RECOMMENDED_ZIP_LIMITS.maxConcurrency,
});

/**
 * 面向含多款内嵌中文字体 PPTX 课件的字体解析限额。
 */
export const PPTX_EMBEDDED_FONT_LIMITS: Readonly<Required<EmbeddedFontLimits>> = Object.freeze({
  ...DEFAULT_EMBEDDED_FONT_LIMITS,
  maxFaces: 32,
  maxInputBytesPerFace: 32 * MB,
  maxDecompressedBytesPerFace: 64 * MB,
  maxTotalDecompressedBytes: 128 * MB,
  maxProcessingMs: 1000,
});

function formatByteSize(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes < 0) {
    return String(bytes);
  }
  const formatUnit = (value: number, unit: string): string => {
    const text = value.toFixed(value >= 100 ? 0 : 1).replace(/\.0$/, '');
    return `${text} ${unit}`;
  };
  if (bytes >= GB) return formatUnit(bytes / GB, 'GB');
  if (bytes >= MB) return formatUnit(bytes / MB, 'MB');
  if (bytes >= KB) return formatUnit(bytes / KB, 'KB');
  return `${bytes} B`;
}

function formatZipLimitDetail(detail: string): string {
  const entriesMatch = detail.match(/^entries\s+(\d+)\s*>\s*maxEntries\s+(\d+)$/i);
  if (entriesMatch) {
    const [, actual, limit] = entriesMatch;
    return tr(
      `PPTX 压缩包内文件条目数过多（${actual} > 上限 ${limit}），超出当前安全解析限制；请精简拆分课件或在桌面端 PowerPoint / WPS 中打开。`,
      `Too many entries inside PPTX archive (${actual} > limit ${limit}). Exceeds safe parsing limit; please reduce slide assets or open in desktop PowerPoint / WPS.`,
    );
  }

  const totalBytesMatch = detail.match(
    /^total uncompressed bytes\s+(\d+)\s*>\s*maxTotalUncompressedBytes\s+(\d+)$/i,
  );
  if (totalBytesMatch) {
    const actual = Number(totalBytesMatch[1]);
    const limit = Number(totalBytesMatch[2]);
    return tr(
      `PPTX 解压后总体积过大（${formatByteSize(actual)}，${actual} B > 上限 ${formatByteSize(limit)}，${limit} B），超出当前安全解析限制；请压缩内容后重试或在桌面端 PowerPoint / WPS 中打开。`,
      `PPTX uncompressed size is too large (${formatByteSize(actual)}, ${actual} B > limit ${formatByteSize(limit)}, ${limit} B). Please compress assets or open in desktop PowerPoint / WPS.`,
    );
  }

  const mediaBytesMatch = detail.match(/^media bytes\s+(\d+)\s*>\s*maxMediaBytes\s+(\d+)$/i);
  if (mediaBytesMatch) {
    const actual = Number(mediaBytesMatch[1]);
    const limit = Number(mediaBytesMatch[2]);
    return tr(
      `PPTX 内嵌媒体资源体积过大（${formatByteSize(actual)}，${actual} B > 上限 ${formatByteSize(limit)}，${limit} B），超出当前安全解析限制；请压缩图片或音视频后重试，或在桌面端 PowerPoint / WPS 中打开。`,
      `Embedded media inside PPTX is too large (${formatByteSize(actual)}, ${actual} B > limit ${formatByteSize(limit)}, ${limit} B). Please compress media or open in desktop PowerPoint / WPS.`,
    );
  }

  const entryBytesMatch = detail.match(
    /^(.+?)\s+is\s+(\d+)\s+bytes\s*>\s*maxEntryUncompressedBytes\s+(\d+)$/i,
  );
  if (entryBytesMatch) {
    const entryName = entryBytesMatch[1];
    const actual = Number(entryBytesMatch[2]);
    const limit = Number(entryBytesMatch[3]);
    return tr(
      `PPTX 内部资源（${entryName}）解压体积过大（${formatByteSize(actual)}，${actual} B > 上限 ${formatByteSize(limit)}，${limit} B），超出当前安全解析限制；请压缩该资源后重试或在桌面端 PowerPoint / WPS 中打开。`,
      `Internal PPTX entry (${entryName}) uncompressed size is too large (${formatByteSize(actual)}, ${actual} B > limit ${formatByteSize(limit)}, ${limit} B). Please compress it or open in desktop PowerPoint / WPS.`,
    );
  }

  return tr(
    `PPTX 文件体积或对象数量超出安全解析限制（${detail}），请精简文件后重试或在桌面端 PowerPoint / WPS 中打开。`,
    `PPTX file size or entry count exceeds safe parsing limits (${detail}). Please simplify the deck or open in desktop PowerPoint / WPS.`,
  );
}

/**
 * 将 PPTX 解析/渲染阶段的底层异常转换为清晰的中文或英文错误提示。
 */
export function formatPptLoadError(err: unknown): Error {
  const rawMessage =
    err instanceof Error ? err.message : typeof err === 'string' ? err : String(err ?? '');
  const message = rawMessage.trim();

  if (!message) {
    return new Error(
      tr(
        'PPTX 解析失败，请确认文件完整且未损坏。',
        'Failed to parse PPTX. Please verify the file is complete and not corrupted.',
      ),
    );
  }

  const zipLimitMatch = message.match(/^PPTX zip limit exceeded:\s*(.+)$/i);
  if (zipLimitMatch) {
    return new Error(formatZipLimitDetail(zipLimitMatch[1].trim()));
  }

  if (/文件内容为空|\bempty\b|\b0\s*bytes?\b/i.test(message)) {
    return err instanceof Error && /文件内容为空|is empty/i.test(message)
      ? err
      : new Error(tr('PPTX 文件内容为空，无法渲染。', 'PPTX file is empty and cannot be rendered.'));
  }

  if (/encrypt|password|已加密|受密码保护/i.test(message)) {
    return new Error(
      tr(
        'PPTX 文件已加密或受密码保护，暂不支持直接预览，请解除密码后重试。',
        'This PPTX file is encrypted or password-protected and cannot be previewed directly.',
      ),
    );
  }

  if (
    /end of central directory|corrupted zip|invalid zip|not a valid zip|end of data reached|unexpected signature|multi-volumes zip|compression .* unknown|central dir|missing\s+(?:\[Content_Types\]|ppt\/presentation\.xml)/i.test(
      message,
    )
  ) {
    return new Error(
      tr(
        `PPTX 文件格式无效或压缩包已损坏，无法解析（${message}）。`,
        `Invalid PPTX format or corrupted ZIP archive (${message}).`,
      ),
    );
  }

  if (err instanceof Error && /^PPTX[\s\u4e00-\u9fa5]/.test(message) && /[\u4e00-\u9fa5]/.test(message)) {
    return err;
  }

  return new Error(tr(`PPTX 加载或解析失败：${message}`, `Failed to load or parse PPTX: ${message}`));
}
