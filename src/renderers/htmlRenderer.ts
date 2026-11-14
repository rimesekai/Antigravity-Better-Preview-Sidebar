import type {
  HtmlMountInfo,
  HtmlPreviewMode,
  PreviewRenderer,
  PreviewStateInfo,
} from '../types/preview';
import { copyTextToClipboard } from '../utils/clipboard';
import { onLocaleChange, tr } from '../utils/i18n';

function getUnmountedInteractiveWarning(): string {
  return tr(
    '拖拽无路径文件仅支持静态沙箱与源码视图，若需加载同目录资源请从工作区或路径打开',
    'Dropped files without a disk path support Sandbox and Source views only; open from workspace or path to load local assets.',
  );
}

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

export function createHtmlRenderer(): PreviewRenderer {
  let rootEl: HTMLElement | null = null;
  let warningEl: HTMLElement | null = null;
  let stageEl: HTMLElement | null = null;

  let currentMountInfo: HtmlMountInfo | null = null;
  let currentSourceText = '';
  let currentMeta: { name: string; path: string; ext: string } | null = null;
  let htmlMode: HtmlPreviewMode = 'sandbox';
  let hasUnmountedInteractiveWarning = false;
  let destroyed = false;
  let copyTimer: number | null = null;
  let stopLocaleSync: (() => void) | null = null;

  const stateListeners = new Set<(info: PreviewStateInfo) => void>();

  function getActiveWarning(): string | undefined {
    return hasUnmountedInteractiveWarning ? getUnmountedInteractiveWarning() : undefined;
  }

  function emitState(): void {
    const info: PreviewStateInfo = {
      kind: 'html',
      htmlMode,
      warning: getActiveWarning(),
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
    const warning = getActiveWarning();
    if (warning) {
      warningEl.textContent = warning;
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
    iframe.title = currentMeta
      ? tr(`HTML 预览: ${currentMeta.name}`, `HTML Preview: ${currentMeta.name}`)
      : tr('HTML 预览', 'HTML Preview');
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
    metaLabel.textContent = tr(
      `${currentMeta?.name ?? 'document.html'} · ${lineCount} 行`,
      `${currentMeta?.name ?? 'document.html'} · ${lineCount} lines`,
    );

    const copyBtn = document.createElement('button');
    copyBtn.type = 'button';
    copyBtn.className = 'fp-html-copy-btn';
    copyBtn.textContent = tr('复制源码', 'Copy Source');
    copyBtn.addEventListener('click', () => {
      void copyTextToClipboard(currentSourceText).then((ok) => {
        if (destroyed) return;
        if (copyTimer !== null) {
          window.clearTimeout(copyTimer);
        }
        copyBtn.textContent = ok ? tr('已复制', 'Copied') : tr('复制失败', 'Copy failed');
        copyBtn.classList.toggle('is-copied', ok);
        copyTimer = window.setTimeout(() => {
          copyTimer = null;
          if (!destroyed) {
            copyBtn.textContent = tr('复制源码', 'Copy Source');
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

      stopLocaleSync?.();
      stopLocaleSync = onLocaleChange(() => {
        updateWarningBanner();
        if (currentMeta) {
          if (htmlMode === 'source') {
            renderSourceView();
          } else {
            const iframe = stageEl?.querySelector('iframe');
            if (iframe) {
              iframe.title = tr(`HTML 预览: ${currentMeta.name}`, `HTML Preview: ${currentMeta.name}`);
            }
          }
          emitState();
        }
      });
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
        hasUnmountedInteractiveWarning = true;
      } else {
        hasUnmountedInteractiveWarning = false;
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
        hasUnmountedInteractiveWarning = true;
        renderCurrentMode();
        emitState();
        return;
      }

      htmlMode = mode;
      hasUnmountedInteractiveWarning = false;
      renderCurrentMode();
      emitState();
    },

    onStateChange(cb: (info: PreviewStateInfo) => void): void {
      stateListeners.add(cb);
    },

    destroy(): void {
      destroyed = true;
      stopLocaleSync?.();
      stopLocaleSync = null;

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
      hasUnmountedInteractiveWarning = false;
      stateListeners.clear();
    },
  };
}
