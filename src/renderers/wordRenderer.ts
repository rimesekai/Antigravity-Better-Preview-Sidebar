import { renderAsync } from 'docx-preview';
import type { PreviewRenderer, PreviewStateInfo } from '../types/preview';

const MIN_ZOOM = 0.5;
const MAX_ZOOM = 2.0;

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function roundZoom(value: number): number {
  return Math.round(value * 100) / 100;
}

export function createWordRenderer(): PreviewRenderer {
  let rootEl: HTMLElement | null = null;
  let viewportEl: HTMLElement | null = null;
  let sizerEl: HTMLElement | null = null;
  let scaleLayerEl: HTMLElement | null = null;
  let stageEl: HTMLElement | null = null;

  let resizeObserver: ResizeObserver | null = null;
  let rafLayoutId: number | null = null;
  let rafScrollId: number | null = null;
  let loadGeneration = 0;
  let destroyed = false;

  let fitMode = true;
  let manualZoom = 1.0;
  let effectiveZoom = 1.0;
  let currentPage = 1;
  let totalPages = 0;
  let lastBuffer: ArrayBuffer | null = null;
  let lastMeta: { name: string; path: string; ext: string } | null = null;

  const stateListeners = new Set<(info: PreviewStateInfo) => void>();

  function emitState(warning?: string): void {
    const info: PreviewStateInfo = {
      kind: 'word',
      page: totalPages > 0 ? currentPage : undefined,
      totalPages: totalPages > 0 ? totalPages : undefined,
      zoom: roundZoom(effectiveZoom),
      fitMode,
      warning,
    };
    for (const cb of stateListeners) {
      try {
        cb(info);
      } catch (err) {
        console.error('[wordRenderer] state listener error:', err);
      }
    }
  }

  function getPageSections(): HTMLElement[] {
    if (!stageEl) return [];
    return Array.from(stageEl.querySelectorAll<HTMLElement>('.docx-wrapper > section.docx'));
  }

  function updateLayout(): void {
    if (destroyed || !viewportEl || !sizerEl || !scaleLayerEl || !stageEl) return;

    const sections = getPageSections();
    const nextTotal = sections.length > 0 ? sections.length : stageEl.children.length > 0 ? 1 : 0;
    const totalChanged = nextTotal !== totalPages;
    totalPages = nextTotal;
    if (totalPages > 0) {
      currentPage = clamp(currentPage, 1, totalPages);
    }

    // Read untransformed layout dimensions from scaleLayerEl (not affected by CSS transform)
    const maxSectionWidth = sections.reduce((max, sec) => Math.max(max, sec.offsetWidth), 0);
    const rawWidth = Math.max(scaleLayerEl.offsetWidth, maxSectionWidth > 0 ? maxSectionWidth + 24 : 0);
    const rawHeight = scaleLayerEl.offsetHeight;

    if (rawWidth <= 0 || rawHeight <= 0) {
      if (totalChanged) emitState();
      return;
    }

    const viewportWidth = viewportEl.clientWidth;
    let nextScale = manualZoom;
    if (fitMode) {
      if (viewportWidth <= 0) return;
      const availableWidth = Math.max(120, viewportWidth - 8);
      nextScale = Math.min(1.0, Math.max(0.2, availableWidth / rawWidth));
    } else {
      nextScale = clamp(manualZoom, MIN_ZOOM, MAX_ZOOM);
    }

    const scaledW = Math.ceil(rawWidth * nextScale);
    const scaledH = Math.ceil(rawHeight * nextScale);

    sizerEl.style.width = `${scaledW}px`;
    sizerEl.style.height = `${scaledH}px`;
    scaleLayerEl.style.transform = `scale(${nextScale})`;

    const zoomChanged = Math.abs(nextScale - effectiveZoom) > 0.002;
    effectiveZoom = nextScale;

    if (zoomChanged || totalChanged) {
      emitState();
    }
  }

  function scheduleLayout(): void {
    if (destroyed || rafLayoutId !== null) return;
    rafLayoutId = requestAnimationFrame(() => {
      rafLayoutId = null;
      updateLayout();
    });
  }

  function updateCurrentPageFromScroll(): void {
    if (destroyed || !viewportEl) return;
    const sections = getPageSections();
    if (sections.length === 0) return;

    const viewportRect = viewportEl.getBoundingClientRect();
    const probeY = viewportRect.top + Math.min(120, viewportRect.height * 0.35);

    let bestIdx = 0;
    let bestDist = Infinity;
    for (let i = 0; i < sections.length; i++) {
      const rect = sections[i].getBoundingClientRect();
      if (rect.top <= probeY && rect.bottom >= probeY) {
        bestIdx = i;
        break;
      }
      const dist = Math.min(Math.abs(rect.top - probeY), Math.abs(rect.bottom - probeY));
      if (dist < bestDist) {
        bestDist = dist;
        bestIdx = i;
      }
    }

    const nextPage = bestIdx + 1;
    if (nextPage !== currentPage) {
      currentPage = nextPage;
      emitState();
    }
  }

  function handleViewportScroll(): void {
    if (destroyed || rafScrollId !== null) return;
    rafScrollId = requestAnimationFrame(() => {
      rafScrollId = null;
      updateCurrentPageFromScroll();
    });
  }

  async function renderCurrentBuffer(): Promise<void> {
    if (destroyed || !stageEl || !lastBuffer) return;
    if (lastBuffer.byteLength === 0) {
      throw new Error('Word 文档内容为空，无法渲染');
    }

    const gen = ++loadGeneration;
    stageEl.innerHTML = '';
    totalPages = 0;
    currentPage = 1;
    if (sizerEl) {
      sizerEl.style.width = '';
      sizerEl.style.height = '';
    }
    if (scaleLayerEl) {
      scaleLayerEl.style.transform = 'scale(1)';
    }

    await renderAsync(lastBuffer, stageEl, undefined, {
      inWrapper: true,
      breakPages: true,
      ignoreLastRenderedPageBreak: false,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: true,
      useBase64URL: false,
    });

    if (destroyed || gen !== loadGeneration) {
      return;
    }

    if (viewportEl) {
      viewportEl.scrollTop = 0;
      viewportEl.scrollLeft = 0;
    }

    updateLayout();
    // Re-check once images/fonts settle on next frame
    scheduleLayout();
    emitState();
  }

  return {
    mount(container: HTMLElement): void {
      if (destroyed) {
        destroyed = false;
      }
      container.innerHTML = '';

      rootEl = document.createElement('div');
      rootEl.className = 'fp-word-root';

      viewportEl = document.createElement('div');
      viewportEl.className = 'fp-word-viewport';

      sizerEl = document.createElement('div');
      sizerEl.className = 'fp-word-sizer';

      scaleLayerEl = document.createElement('div');
      scaleLayerEl.className = 'fp-word-scale-layer';

      stageEl = document.createElement('div');
      stageEl.className = 'fp-word-stage';

      scaleLayerEl.appendChild(stageEl);
      sizerEl.appendChild(scaleLayerEl);
      viewportEl.appendChild(sizerEl);
      rootEl.appendChild(viewportEl);
      container.appendChild(rootEl);

      viewportEl.addEventListener('scroll', handleViewportScroll, { passive: true });

      if (typeof ResizeObserver !== 'undefined') {
        resizeObserver = new ResizeObserver(() => {
          scheduleLayout();
        });
        resizeObserver.observe(viewportEl);
        resizeObserver.observe(scaleLayerEl);
      }
    },

    async loadBuffer(buffer: ArrayBuffer, meta: { name: string; path: string; ext: string }): Promise<void> {
      lastBuffer = buffer;
      lastMeta = meta;
      await renderCurrentBuffer();
    },

    async refresh(): Promise<void> {
      if (!lastBuffer || !lastMeta) return;
      await renderCurrentBuffer();
    },

    setZoom(zoom: number | 'fit'): void {
      if (zoom === 'fit') {
        fitMode = true;
        updateLayout();
        emitState();
        return;
      }
      if (!Number.isFinite(zoom)) return;
      fitMode = false;
      manualZoom = clamp(zoom, MIN_ZOOM, MAX_ZOOM);
      updateLayout();
      emitState();
    },

    nextPage(): void {
      if (totalPages <= 0) return;
      this.goToPage?.(currentPage + 1);
    },

    prevPage(): void {
      if (totalPages <= 0) return;
      this.goToPage?.(currentPage - 1);
    },

    goToPage(page: number): void {
      if (!viewportEl || totalPages <= 0 || !Number.isFinite(page)) return;
      const target = clamp(Math.round(page), 1, totalPages);
      const sections = getPageSections();
      const targetSection = sections[target - 1];
      if (targetSection) {
        const vpRect = viewportEl.getBoundingClientRect();
        const secRect = targetSection.getBoundingClientRect();
        const targetTop = viewportEl.scrollTop + (secRect.top - vpRect.top) - 12;
        viewportEl.scrollTo({ top: Math.max(0, targetTop), behavior: 'smooth' });
      }
      currentPage = target;
      emitState();
    },

    onStateChange(cb: (info: PreviewStateInfo) => void): void {
      stateListeners.add(cb);
    },

    destroy(): void {
      destroyed = true;
      loadGeneration++;

      if (rafLayoutId !== null) {
        cancelAnimationFrame(rafLayoutId);
        rafLayoutId = null;
      }
      if (rafScrollId !== null) {
        cancelAnimationFrame(rafScrollId);
        rafScrollId = null;
      }
      if (resizeObserver) {
        resizeObserver.disconnect();
        resizeObserver = null;
      }
      if (viewportEl) {
        viewportEl.removeEventListener('scroll', handleViewportScroll);
      }
      if (stageEl) {
        stageEl.innerHTML = '';
      }
      if (rootEl) {
        rootEl.remove();
      }

      rootEl = null;
      viewportEl = null;
      sizerEl = null;
      scaleLayerEl = null;
      stageEl = null;
      lastBuffer = null;
      lastMeta = null;
      totalPages = 0;
      currentPage = 1;
      stateListeners.clear();
    },
  };
}
