import type {
  HtmlMountInfo,
  HtmlPreviewMode,
  PreviewRenderer,
  PreviewStateInfo,
} from '../types/preview';

const UNMOUNTED_INTERACTIVE_WARNING =
  '拖拽无路径文件仅支持静态沙箱与源码视图，若需加载同目录资源请从工作区或路径打开';

function buildBustUrl(rawUrl: string): string {
  try {
    const parsed = new URL(rawUrl);
    parsed.searchParams.set('_ts', String(Date.now()));
    return parsed.toString();
  } catch {
    const sep = rawUrl.includes('?') ? '&' : '?';
    return `${rawUrl}${sep}_ts=${Date.now()}`;
  }
}

async function copyTextToClipboard(text: string): Promise<boolean> {
  if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch {
      // fallback below
    }
  }
  const textarea = document.createElement('textarea');
  textarea.value = text;
  textarea.setAttribute('readonly', 'true');
  textarea.style.position = 'fixed';
  textarea.style.top = '-9999px';
  textarea.style.left = '-9999px';
  document.body.appendChild(textarea);
  textarea.select();
  let ok = false;
  try {
    ok = document.execCommand('copy');
  } catch {
    ok = false;
  }
  textarea.remove();
  return ok;
}

export function createHtmlRenderer(): PreviewRenderer {
  let rootEl: HTMLElement | null = null;
  let warningEl: HTMLElement | null = null;
  let stageEl: HTMLElement | null = null;

  let currentMountInfo: HtmlMountInfo | null = null;
  let currentSourceText = '';
  let currentMeta: { name: string; path: string; ext: string } | null = null;
  let htmlMode: HtmlPreviewMode = 'sandbox';
  let activeWarning: string | undefined;
  let destroyed = false;
  let copyTimer: number | null = null;

  const stateListeners = new Set<(info: PreviewStateInfo) => void>();

  function emitState(): void {
    const info: PreviewStateInfo = {
      kind: 'html',
      htmlMode,
      warning: activeWarning,
    };
    for (const cb of stateListeners) {
      try {
        cb(info);
      } catch (err) {
        console.error('[htmlRenderer] state listener error:', err);
      }
    }
  }

  function updateWarningBanner(): void {
    if (!warningEl) return;
    if (activeWarning) {
      warningEl.textContent = activeWarning;
      warningEl.hidden = false;
    } else {
      warningEl.textContent = '';
      warningEl.hidden = true;
    }
  }

  function renderIframeView(interactive: boolean): void {
    if (!stageEl) return;
    stageEl.innerHTML = '';

    const frameWrap = document.createElement('div');
    frameWrap.className = 'fp-html-frame-wrap';

    const iframe = document.createElement('iframe');
    iframe.className = 'fp-html-iframe';
    iframe.title = currentMeta ? `HTML 预览: ${currentMeta.name}` : 'HTML 预览';
    iframe.setAttribute('referrerpolicy', 'no-referrer');

    if (currentMountInfo?.url) {
      const sandboxValue = interactive
        ? 'allow-scripts allow-same-origin allow-forms allow-popups'
        : '';
      iframe.setAttribute('sandbox', sandboxValue);
      iframe.src = buildBustUrl(currentMountInfo.url);
    } else {
      iframe.setAttribute('sandbox', '');
      iframe.srcdoc = currentSourceText;
    }

    frameWrap.appendChild(iframe);
    stageEl.appendChild(frameWrap);
  }

  function renderSourceView(): void {
    if (!stageEl) return;
    stageEl.innerHTML = '';

    const sourceWrap = document.createElement('div');
    sourceWrap.className = 'fp-html-source';

    const headerEl = document.createElement('div');
    headerEl.className = 'fp-html-source-header';

    const normalized = currentSourceText.replace(/\r\n/g, '\n');
    const lineCount = normalized.length === 0 ? 1 : normalized.split('\n').length;

    const metaLabel = document.createElement('span');
    metaLabel.className = 'fp-html-source-meta';
    metaLabel.textContent = `${currentMeta?.name ?? 'document.html'} · ${lineCount} 行`;

    const copyBtn = document.createElement('button');
    copyBtn.type = 'button';
    copyBtn.className = 'fp-html-copy-btn';
    copyBtn.textContent = '复制源码';
    copyBtn.addEventListener('click', () => {
      void copyTextToClipboard(currentSourceText).then((ok) => {
        if (destroyed) return;
        if (copyTimer !== null) {
          window.clearTimeout(copyTimer);
        }
        copyBtn.textContent = ok ? '已复制' : '复制失败';
        copyBtn.classList.toggle('is-copied', ok);
        copyTimer = window.setTimeout(() => {
          copyTimer = null;
          if (!destroyed) {
            copyBtn.textContent = '复制源码';
            copyBtn.classList.remove('is-copied');
          }
        }, 1800);
      });
    });

    headerEl.appendChild(metaLabel);
    headerEl.appendChild(copyBtn);

    const bodyEl = document.createElement('div');
    bodyEl.className = 'fp-html-source-body';

    const gutterEl = document.createElement('pre');
    gutterEl.className = 'fp-html-gutter';
    gutterEl.setAttribute('aria-hidden', 'true');
    const lineNumbers: string[] = [];
    for (let i = 1; i <= lineCount; i++) {
      lineNumbers.push(String(i));
    }
    gutterEl.textContent = lineNumbers.join('\n');

    const codePre = document.createElement('pre');
    codePre.className = 'fp-html-code';
    const codeEl = document.createElement('code');
    codeEl.textContent = normalized;
    codePre.appendChild(codeEl);

    bodyEl.appendChild(gutterEl);
    bodyEl.appendChild(codePre);
    sourceWrap.appendChild(headerEl);
    sourceWrap.appendChild(bodyEl);
    stageEl.appendChild(sourceWrap);
  }

  function renderCurrentMode(): void {
    if (destroyed || !stageEl) return;
    updateWarningBanner();
    if (htmlMode === 'source') {
      renderSourceView();
    } else {
      renderIframeView(htmlMode === 'interactive');
    }
  }

  return {
    mount(container: HTMLElement): void {
      if (destroyed) {
        destroyed = false;
      }
      container.innerHTML = '';

      rootEl = document.createElement('div');
      rootEl.className = 'fp-html-root';

      warningEl = document.createElement('div');
      warningEl.className = 'fp-html-warning';
      warningEl.setAttribute('role', 'alert');
      warningEl.hidden = true;

      stageEl = document.createElement('div');
      stageEl.className = 'fp-html-stage';

      rootEl.appendChild(warningEl);
      rootEl.appendChild(stageEl);
      container.appendChild(rootEl);
    },

    async loadHtml(
      mountInfo: HtmlMountInfo | null,
      sourceText: string,
      meta: { name: string; path: string; ext: string }
    ): Promise<void> {
      currentMountInfo = mountInfo;
      currentSourceText = sourceText ?? '';
      currentMeta = meta;

      if (!currentMountInfo?.url && htmlMode === 'interactive') {
        htmlMode = 'sandbox';
        activeWarning = UNMOUNTED_INTERACTIVE_WARNING;
      } else {
        activeWarning = undefined;
      }

      renderCurrentMode();
      emitState();
    },

    async refresh(): Promise<void> {
      if (destroyed || !currentMeta) return;
      renderCurrentMode();
      emitState();
    },

    setHtmlMode(mode: HtmlPreviewMode): void {
      if (mode !== 'sandbox' && mode !== 'interactive' && mode !== 'source') return;

      if (mode === 'interactive' && !currentMountInfo?.url) {
        htmlMode = 'sandbox';
        activeWarning = UNMOUNTED_INTERACTIVE_WARNING;
        renderCurrentMode();
        emitState();
        return;
      }

      htmlMode = mode;
      activeWarning = undefined;
      renderCurrentMode();
      emitState();
    },

    onStateChange(cb: (info: PreviewStateInfo) => void): void {
      stateListeners.add(cb);
    },

    destroy(): void {
      destroyed = true;

      if (copyTimer !== null) {
        window.clearTimeout(copyTimer);
        copyTimer = null;
      }

      if (stageEl) {
        const iframe = stageEl.querySelector('iframe');
        if (iframe) {
          iframe.src = 'about:blank';
        }
        stageEl.innerHTML = '';
      }

      if (rootEl) {
        rootEl.remove();
      }

      rootEl = null;
      warningEl = null;
      stageEl = null;
      currentMountInfo = null;
      currentSourceText = '';
      currentMeta = null;
      activeWarning = undefined;
      stateListeners.clear();
    },
  };
}
