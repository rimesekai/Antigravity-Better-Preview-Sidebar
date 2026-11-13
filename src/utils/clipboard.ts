/**
 * 将文本写入剪贴板；优先使用 Clipboard API，在无权限或受限 WebView 沙箱中自动降级到 execCommand('copy')。
 */
export async function copyTextToClipboard(text: string): Promise<boolean> {
  if (typeof navigator !== 'undefined' && navigator.clipboard?.writeText) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch {
      // 沙箱或非安全上下文可能拒绝 Clipboard API，继续走隐藏 textarea 降级方案
    }
  }

  if (typeof document === 'undefined' || !document.body) {
    return false;
  }

  const textarea = document.createElement('textarea');
  textarea.value = text;
  textarea.setAttribute('readonly', 'true');
  textarea.style.position = 'fixed';
  textarea.style.top = '-9999px';
  textarea.style.left = '-9999px';
  textarea.style.opacity = '0';
  document.body.appendChild(textarea);

  const prevSelection = document.getSelection();
  const prevRanges: Range[] = [];
  if (prevSelection) {
    for (let i = 0; i < prevSelection.rangeCount; i++) {
      prevRanges.push(prevSelection.getRangeAt(i));
    }
  }

  textarea.focus();
  textarea.select();
  textarea.setSelectionRange(0, textarea.value.length);

  let ok = false;
  try {
    ok = document.execCommand('copy');
  } catch {
    ok = false;
  }

  textarea.remove();

  // 恢复页面原有选区，避免触发复制时破坏用户当前选区
  if (prevSelection) {
    try {
      prevSelection.removeAllRanges();
      for (const range of prevRanges) {
        prevSelection.addRange(range);
      }
    } catch {
      // 忽略选区恢复失败
    }
  }

  return ok;
}
