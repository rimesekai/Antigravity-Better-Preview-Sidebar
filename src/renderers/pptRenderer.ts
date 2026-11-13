import {
  PptxViewer,
  RECOMMENDED_ZIP_LIMITS,
  type SlideHandle,
  type TextSearchResult,
} from '@aiden0z/pptx-renderer/browser';
import type { PptLayoutMode, PreviewRenderer, PreviewStateInfo } from '../types/preview';

const MIN_ZOOM = 0.5;
const MAX_ZOOM = 2.0;

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function roundZoom(value: number): number {
  return Math.round(value * 100) / 100;
}

function isEditableTarget(target: EventTarget | null): boolean {
  if (!target || !(target instanceof HTMLElement)) {
    return false;
  }
  const tag = target.tagName.toUpperCase();
  if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') {
    return true;
  }
  if (target.isContentEditable) {
    return true;
  }
  return Boolean(target.closest('[contenteditable="true"], [role="textbox"]'));
}

export function createPptRenderer(): PreviewRenderer {
  let rootEl: HTMLElement | null = null;
  let thumbsPanelEl: HTMLElement | null = null;
  let thumbsListEl: HTMLElement | null = null;
  let scrollEl: HTMLElement | null = null;
  let stageEl: HTMLElement | null = null;

  let viewer: PptxViewer | null = null;
  let loadAbortController: AbortController | null = null;
  let thumbObserver: IntersectionObserver | null = null;
  const thumbHandles = new Map<number, SlideHandle>();
  const thumbItems = new Map<number, HTMLElement>();

  let destroyed = false;
  let loadGeneration = 0;

  let currentPage = 1;
  let totalPages = 0;
  let currentZoom = 1.0;
  let fitMode = true;
  let layoutMode: PptLayoutMode = 'slide';
  let thumbnailsVisible = false;
  let slideErrorCount = 0;

  let searchResults: TextSearchResult[] = [];
  let activeMatchIndex = -1;

  let lastBuffer: ArrayBuffer | null = null;
  let lastMeta: { name: string; path: string; ext: string } | null = null;

  const stateListeners = new Set<(info: PreviewStateInfo) => void>();

  function emitState(): void {
    const warning =
      slideErrorCount > 0
        ? `部分幻灯片（${slideErrorCount} 页）渲染异常，已跳过异常元素`
        : undefined;
    const info: PreviewStateInfo = {
      kind: 'ppt',
      page: totalPages > 0 ? currentPage : undefined,
      totalPages: totalPages > 0 ? totalPages : undefined,
      zoom: roundZoom(currentZoom),
      fitMode,
      layoutMode,
      thumbnailsVisible,
      searchMatches: searchResults.length,
      activeMatch: activeMatchIndex >= 0 ? activeMatchIndex + 1 : 0,
      warning,
    };
    for (const cb of stateListeners) {
      try {
        cb(info);
      } catch (err) {
        console.error('[pptRenderer] state listener error:', err);
      }
    }
  }

  function disposeThumbnails(): void {
    if (thumbObserver) {
      thumbObserver.disconnect();
      thumbObserver = null;
    }
    for (const handle of thumbHandles.values()) {
      try {
        handle.dispose();
      } catch {
        // ignore disposal errors
      }
    }
    thumbHandles.clear();
    thumbItems.clear();
    if (thumbsListEl) {
      thumbsListEl.innerHTML = '';
    }
  }

  function updateActiveThumbnail(): void {
    if (!thumbnailsVisible) return;
    const activeIdx = currentPage - 1;
    for (const [idx, item] of thumbItems.entries()) {
      const isActive = idx === activeIdx;
      item.classList.toggle('is-active', isActive);
      item.setAttribute('aria-current', isActive ? 'true' : 'false');
      if (isActive) {
        item.scrollIntoView({ block: 'nearest', inline: 'nearest' });
      }
    }
  }

  function buildThumbnails(): void {
    disposeThumbnails();
    if (!thumbnailsVisible || !thumbsListEl || !viewer || !viewer.presentationData) {
      return;
    }

    const count = viewer.slideCount;
    const slots: Array<{ index: number; item: HTMLElement; previewBox: HTMLElement }> = [];

    for (let i = 0; i < count; i++) {
      const item = document.createElement('button');
      item.type = 'button';
      item.className = 'fp-ppt-thumb-item';
      item.dataset.slideIndex = String(i);
      if (i === currentPage - 1) {
        item.classList.add('is-active');
        item.setAttribute('aria-current', 'true');
      }

      const previewBox = document.createElement('div');
      previewBox.className = 'fp-ppt-thumb-preview';

      const label = document.createElement('span');
      label.className = 'fp-ppt-thumb-label';
      label.textContent = String(i + 1);

      item.appendChild(previewBox);
      item.appendChild(label);
      item.addEventListener('click', () => {
        goToPage(i + 1);
      });

      thumbsListEl.appendChild(item);
      thumbItems.set(i, item);
      slots.push({ index: i, item, previewBox });
    }

    const mountThumb = (index: number, previewBox: HTMLElement): void => {
      if (destroyed || !viewer || thumbHandles.has(index)) return;
      const handle = viewer.renderThumbnailToContainer(index, previewBox, { width: 112 });
      if (handle) {
        thumbHandles.set(index, handle);
      }
    };

    const unmountThumb = (index: number, previewBox: HTMLElement): void => {
      const handle = thumbHandles.get(index);
      if (handle) {
        try {
          handle.dispose();
        } catch {
          // ignore
        }
        thumbHandles.delete(index);
        previewBox.innerHTML = '';
      }
    };

    if (typeof IntersectionObserver !== 'undefined') {
      thumbObserver = new IntersectionObserver(
        (entries) => {
          for (const entry of entries) {
            const target = entry.target as HTMLElement;
            const idx = Number(target.dataset.slideIndex ?? '-1');
            if (idx < 0 || Number.isNaN(idx)) continue;
            const slot = slots[idx];
            if (!slot) continue;
            if (entry.isIntersecting) {
              mountThumb(idx, slot.previewBox);
            } else {
              unmountThumb(idx, slot.previewBox);
            }
          }
        },
        { root: thumbsListEl, rootMargin: '160px 160px', threshold: 0 }
      );
      for (const slot of slots) {
        thumbObserver.observe(slot.item);
      }
    } else {
      for (const slot of slots.slice(0, 12)) {
        mountThumb(slot.index, slot.previewBox);
      }
    }

    updateActiveThumbnail();
  }

  async function highlightActiveMatch(): Promise<void> {
    if (destroyed || !viewer || !viewer.presentationData) return;
    viewer.clearSearchHighlights();
    const match = activeMatchIndex >= 0 ? searchResults[activeMatchIndex] : undefined;
    if (!match) return;

    await viewer.highlightSearchResult(match, {
      className: 'fp-ppt-search-hit',
      borderColor: '#2563eb',
      backgroundColor: 'rgba(37, 99, 235, 0.18)',
      borderWidth: 2,
      borderRadius: 4,
      padding: 3,
      scrollIntoView: { behavior: 'smooth', block: 'center' },
    });

    if (!destroyed) {
      currentPage = match.slideIndex + 1;
      updateActiveThumbnail();
    }
  }

  function goToPage(page: number): void {
    if (destroyed || !viewer || totalPages <= 0 || !Number.isFinite(page)) return;
    const clamped = clamp(Math.round(page), 1, totalPages);
    currentPage = clamped;
    updateActiveThumbnail();
    emitState();
    void viewer.goToSlide(clamped - 1, { behavior: 'smooth', block: 'center' });
  }

  function nextPage(): void {
    if (totalPages <= 0 || currentPage >= totalPages) return;
    goToPage(currentPage + 1);
  }

  function prevPage(): void {
    if (totalPages <= 0 || currentPage <= 1) return;
    goToPage(currentPage - 1);
  }

  function handleKeyDown(e: KeyboardEvent): void {
    if (destroyed || !viewer || totalPages <= 0) return;
    if (e.defaultPrevented || e.isComposing) return;
    if (e.ctrlKey || e.metaKey || e.altKey || e.shiftKey) return;
    if (isEditableTarget(e.target)) return;

    if (e.key === 'ArrowLeft') {
      e.preventDefault();
      prevPage();
    } else if (e.key === 'ArrowRight') {
      e.preventDefault();
      nextPage();
    }
  }

  async function openCurrentBuffer(): Promise<void> {
    if (destroyed || !stageEl || !scrollEl || !lastBuffer) return;
    if (lastBuffer.byteLength === 0) {
      throw new Error('PPTX 文件内容为空，无法渲染');
    }

    const gen = ++loadGeneration;

    if (loadAbortController) {
      loadAbortController.abort();
      loadAbortController = null;
    }
    loadAbortController = new AbortController();
    const signal = loadAbortController.signal;

    disposeThumbnails();
    if (viewer) {
      viewer.destroy();
      viewer = null;
    }

    stageEl.innerHTML = '';
    searchResults = [];
    activeMatchIndex = -1;
    slideErrorCount = 0;
    currentPage = 1;
    totalPages = 0;

    const instance = new PptxViewer(stageEl, {
      fitMode: 'contain',
      zoomPercent: Math.round(currentZoom * 100),
      scrollContainer: scrollEl,
      zipLimits: RECOMMENDED_ZIP_LIMITS,
      externalMedia: 'block',
      lazySlides: true,
      lazyMedia: true,
      pdfjs: false,
      onSlideChange: (index) => {
        if (destroyed || gen !== loadGeneration) return;
        currentPage = index + 1;
        updateActiveThumbnail();
        emitState();
      },
      onSlideError: (index, err) => {
        if (destroyed || gen !== loadGeneration) return;
        console.warn(`[pptRenderer] slide ${index + 1} error:`, err);
        slideErrorCount++;
        emitState();
      },
    });
    viewer = instance;

    await instance.open(lastBuffer, {
      renderMode: layoutMode,
      listOptions: {
        windowed: true,
        initialSlides: 4,
        batchSize: 6,
        overscanViewport: 1.5,
        showSlideLabels: true,
      },
      signal,
      lazySlides: true,
      lazyMedia: true,
    });

    if (destroyed || gen !== loadGeneration) {
      instance.destroy();
      return;
    }

    totalPages = instance.slideCount;
    currentPage = totalPages > 0 ? instance.currentSlideIndex + 1 : 0;

    if (scrollEl) {
      scrollEl.scrollTop = 0;
      scrollEl.scrollLeft = 0;
    }

    if (thumbnailsVisible) {
      buildThumbnails();
    }

    emitState();
  }

  return {
    mount(container: HTMLElement): void {
      if (destroyed) {
        destroyed = false;
      }
      container.innerHTML = '';

      rootEl = document.createElement('div');
      rootEl.className = 'fp-ppt-root';

      const bodyEl = document.createElement('div');
      bodyEl.className = 'fp-ppt-body';

      thumbsPanelEl = document.createElement('aside');
      thumbsPanelEl.className = 'fp-ppt-thumbs';
      thumbsPanelEl.hidden = !thumbnailsVisible;

      const thumbsHeaderEl = document.createElement('div');
      thumbsHeaderEl.className = 'fp-ppt-thumbs-header';

      const thumbsTitleEl = document.createElement('span');
      thumbsTitleEl.textContent = '幻灯片缩略图';

      const thumbsCloseBtn = document.createElement('button');
      thumbsCloseBtn.type = 'button';
      thumbsCloseBtn.className = 'fp-ppt-thumbs-close';
      thumbsCloseBtn.title = '收起缩略图';
      thumbsCloseBtn.setAttribute('aria-label', '收起缩略图');
      thumbsCloseBtn.textContent = '×';
      thumbsCloseBtn.addEventListener('click', () => {
        this.toggleThumbnails?.(false);
      });

      thumbsHeaderEl.appendChild(thumbsTitleEl);
      thumbsHeaderEl.appendChild(thumbsCloseBtn);

      thumbsListEl = document.createElement('div');
      thumbsListEl.className = 'fp-ppt-thumbs-list';

      thumbsPanelEl.appendChild(thumbsHeaderEl);
      thumbsPanelEl.appendChild(thumbsListEl);

      scrollEl = document.createElement('div');
      scrollEl.className = 'fp-ppt-scroll';

      stageEl = document.createElement('div');
      stageEl.className = 'fp-ppt-stage';

      scrollEl.appendChild(stageEl);
      bodyEl.appendChild(thumbsPanelEl);
      bodyEl.appendChild(scrollEl);
      rootEl.appendChild(bodyEl);
      container.appendChild(rootEl);

      window.addEventListener('keydown', handleKeyDown);
    },

    async loadBuffer(buffer: ArrayBuffer, meta: { name: string; path: string; ext: string }): Promise<void> {
      lastBuffer = buffer;
      lastMeta = meta;
      await openCurrentBuffer();
    },

    async refresh(): Promise<void> {
      if (!lastBuffer || !lastMeta) return;
      await openCurrentBuffer();
    },

    setZoom(zoom: number | 'fit'): void {
      if (zoom === 'fit') {
        fitMode = true;
        currentZoom = 1.0;
      } else if (Number.isFinite(zoom)) {
        fitMode = false;
        currentZoom = clamp(zoom, MIN_ZOOM, MAX_ZOOM);
      } else {
        return;
      }

      if (viewer && viewer.presentationData) {
        const targetPercent = Math.round(currentZoom * 100);
        void viewer
          .setFitMode('contain')
          .then(() => viewer?.setZoom(targetPercent))
          .then(() => {
            if (!destroyed && activeMatchIndex >= 0) {
              return highlightActiveMatch();
            }
          });
      }
      emitState();
    },

    nextPage,
    prevPage,
    goToPage,

    setLayoutMode(mode: PptLayoutMode): void {
      if (mode !== 'slide' && mode !== 'list') return;
      if (layoutMode === mode) return;
      layoutMode = mode;

      if (viewer && viewer.presentationData) {
        const targetIndex = clamp(currentPage - 1, 0, Math.max(0, totalPages - 1));
        viewer.clearSearchHighlights();
        const switchTask =
          mode === 'slide'
            ? viewer.renderSlide(targetIndex)
            : viewer
                .renderList({
                  windowed: true,
                  initialSlides: 4,
                  batchSize: 6,
                  overscanViewport: 1.5,
                  showSlideLabels: true,
                })
                .then(() => viewer?.goToSlide(targetIndex, { behavior: 'instant', block: 'center' }));

        void switchTask.then(() => {
          if (!destroyed && activeMatchIndex >= 0) {
            return highlightActiveMatch();
          }
        });
      }
      emitState();
    },

    toggleThumbnails(visible?: boolean): void {
      const next = visible !== undefined ? visible : !thumbnailsVisible;
      if (next === thumbnailsVisible) return;
      thumbnailsVisible = next;
      if (thumbsPanelEl) {
        thumbsPanelEl.hidden = !thumbnailsVisible;
      }
      if (rootEl) {
        rootEl.classList.toggle('has-thumbs', thumbnailsVisible);
      }
      if (thumbnailsVisible) {
        buildThumbnails();
      } else {
        disposeThumbnails();
      }
      emitState();
    },

    async search(query: string): Promise<{ total: number; current: number }> {
      if (!viewer || !viewer.presentationData) {
        searchResults = [];
        activeMatchIndex = -1;
        emitState();
        return { total: 0, current: 0 };
      }

      viewer.clearSearchHighlights();
      const trimmed = query.trim();
      if (!trimmed) {
        searchResults = [];
        activeMatchIndex = -1;
        emitState();
        return { total: 0, current: 0 };
      }

      searchResults = viewer.searchText(trimmed, { matchCase: false });
      if (searchResults.length === 0) {
        activeMatchIndex = -1;
        emitState();
        return { total: 0, current: 0 };
      }

      activeMatchIndex = 0;
      await highlightActiveMatch();
      emitState();
      return { total: searchResults.length, current: 1 };
    },

    nextMatch(): void {
      if (searchResults.length === 0) return;
      activeMatchIndex = (activeMatchIndex + 1) % searchResults.length;
      void highlightActiveMatch().then(() => {
        if (!destroyed) emitState();
      });
    },

    prevMatch(): void {
      if (searchResults.length === 0) return;
      activeMatchIndex = (activeMatchIndex - 1 + searchResults.length) % searchResults.length;
      void highlightActiveMatch().then(() => {
        if (!destroyed) emitState();
      });
    },

    onStateChange(cb: (info: PreviewStateInfo) => void): void {
      stateListeners.add(cb);
    },

    destroy(): void {
      destroyed = true;
      loadGeneration++;

      window.removeEventListener('keydown', handleKeyDown);

      if (loadAbortController) {
        loadAbortController.abort();
        loadAbortController = null;
      }

      disposeThumbnails();

      if (viewer) {
        viewer.destroy();
        viewer = null;
      }

      if (rootEl) {
        rootEl.remove();
      }

      rootEl = null;
      thumbsPanelEl = null;
      thumbsListEl = null;
      scrollEl = null;
      stageEl = null;
      lastBuffer = null;
      lastMeta = null;
      searchResults = [];
      activeMatchIndex = -1;
      totalPages = 0;
      currentPage = 1;
      stateListeners.clear();
    },
  };
}
