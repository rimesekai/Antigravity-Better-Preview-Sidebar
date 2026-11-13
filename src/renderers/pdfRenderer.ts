import * as pdfjsLib from 'pdfjs-dist';
import type {
  PDFDocumentLoadingTask,
  PDFDocumentProxy,
  PDFPageProxy,
  RenderTask,
  TextLayerRenderTask,
} from 'pdfjs-dist';
import type {
  PdfMountInfo,
  PdfViewMode,
  PptLayoutMode,
  PreviewRenderer,
  PreviewStateInfo,
} from '../types/preview';

const MIN_ZOOM = 0.5;
const MAX_ZOOM = 2.0;
const RANGE_CHUNK_SIZE = 65536;
const PAGE_GAP_PX = 16;
const PAGE_HORIZONTAL_PADDING = 28;
const FAR_PAGE_DISTANCE = 4;

const { GlobalWorkerOptions, getDocument, renderTextLayer } = pdfjsLib;

type PdfInitParameters = Exclude<Parameters<typeof getDocument>[0], string | URL | ArrayBuffer | ArrayLike<number>>;

// 使用服务端精确路由托管的本地 Worker 脚本（无外网依赖）
GlobalWorkerOptions.workerSrc = '/pdf.worker.js';

interface PageSlot {
  pageNum: number;
  boxEl: HTMLElement;
  canvasWrapEl: HTMLElement;
  textLayerEl: HTMLElement;
  badgeEl: HTMLElement;
  renderedScale: number | null;
}

interface ThumbSlot {
  pageNum: number;
  itemEl: HTMLButtonElement;
  previewEl: HTMLElement;
  rendered: boolean;
}

interface PdfSearchMatch {
  page: number;
  matchIndexInPage: number;
}

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

function isNonWorkerFatalError(err: unknown): boolean {
  const name = err && typeof err === 'object' && 'name' in err ? String((err as { name: unknown }).name) : '';
  return (
    name === 'PasswordException' ||
    name === 'InvalidPDFException' ||
    name === 'MissingPDFException' ||
    name === 'UnexpectedResponseException'
  );
}

function formatPdfLoadError(err: unknown): Error {
  const name = err && typeof err === 'object' && 'name' in err ? String((err as { name: unknown }).name) : '';
  if (name === 'PasswordException') {
    return new Error('此 PDF 文档受密码保护，暂不支持密码解锁预览。');
  }
  if (name === 'InvalidPDFException') {
    return new Error('PDF 文件格式无效或已损坏，无法解析。');
  }
  if (name === 'MissingPDFException') {
    return new Error('未找到 PDF 文件内容，请确认文件未被移动或删除。');
  }
  if (err instanceof Error) {
    return err;
  }
  return new Error('PDF 加载或渲染失败。');
}

export function createPdfRenderer(): PreviewRenderer {
  let rootEl: HTMLElement | null = null;
  let bodyEl: HTMLElement | null = null;
  let thumbsPanelEl: HTMLElement | null = null;
  let thumbsListEl: HTMLElement | null = null;
  let scrollEl: HTMLElement | null = null;
  let stageEl: HTMLElement | null = null;
  let nativeFrameEl: HTMLIFrameElement | null = null;

  let loadingTask: PDFDocumentLoadingTask | null = null;
  let pdfDoc: PDFDocumentProxy | null = null;
  let pageObserver: IntersectionObserver | null = null;
  let thumbObserver: IntersectionObserver | null = null;
  let resizeObserver: ResizeObserver | null = null;
  let resizeTimer: number | null = null;
  let scrollSyncRaf: number | null = null;

  let currentMountInfo: PdfMountInfo | null = null;
  let currentBuffer: ArrayBuffer | null = null;
  let currentMeta: { name: string; path: string; ext: string; size?: number } | null = null;
  let nativeBlobUrl: string | null = null;

  let destroyed = false;
  let loadGeneration = 0;

  let pdfMode: PdfViewMode = 'canvas';
  let layoutMode: PptLayoutMode = 'list';
  let thumbnailsVisible = false;
  let fitMode = true;
  let currentZoom = 1.0;
  let currentPage = 1;
  let totalPages = 0;
  let activeWarning: string | undefined;

  let defaultPageSize = { width: 595, height: 842 };
  const pageSizeCache = new Map<number, { width: number; height: number }>();
  const pageSlots = new Map<number, PageSlot>();
  const activeRenderTasks = new Map<number, RenderTask>();
  const activeTextTasks = new Map<number, TextLayerRenderTask>();
  const visiblePageSet = new Set<number>();

  const thumbSlots = new Map<number, ThumbSlot>();
  const thumbRenderSet = new Set<number>();

  const pageTextCache = new Map<number, string>();
  let searchQuery = '';
  let searchMatches: PdfSearchMatch[] = [];
  let activeMatchIndex = -1;

  const stateListeners = new Set<(info: PreviewStateInfo) => void>();

  function emitState(): void {
    const fileSize = currentMeta?.size ?? currentMountInfo?.size ?? currentBuffer?.byteLength;
    const info: PreviewStateInfo = {
      kind: 'pdf',
      pdfMode,
      fileSize,
      page: totalPages > 0 ? currentPage : undefined,
      totalPages: totalPages > 0 ? totalPages : undefined,
      zoom: roundZoom(currentZoom),
      fitMode,
      layoutMode,
      thumbnailsVisible,
      searchMatches: searchMatches.length,
      activeMatch: activeMatchIndex >= 0 ? activeMatchIndex + 1 : 0,
      warning: activeWarning,
    };
    for (const cb of stateListeners) {
      try {
        cb(info);
      } catch (err) {
        console.error('[pdfRenderer] state listener error:', err);
      }
    }
  }

  function revokeNativeBlobUrl(): void {
    if (nativeBlobUrl) {
      URL.revokeObjectURL(nativeBlobUrl);
      nativeBlobUrl = null;
    }
  }

  function cancelPageTasks(pageNum: number): void {
    const renderTask = activeRenderTasks.get(pageNum);
    if (renderTask) {
      try {
        renderTask.cancel();
      } catch {
        // ignore cancel error
      }
      activeRenderTasks.delete(pageNum);
    }
    const textTask = activeTextTasks.get(pageNum);
    if (textTask) {
      try {
        textTask.cancel();
      } catch {
        // ignore cancel error
      }
      activeTextTasks.delete(pageNum);
    }
  }

  function cancelAllPageTasks(): void {
    for (const pageNum of Array.from(activeRenderTasks.keys())) {
      cancelPageTasks(pageNum);
    }
    for (const pageNum of Array.from(activeTextTasks.keys())) {
      cancelPageTasks(pageNum);
    }
  }

  function releasePageSlot(slot: PageSlot): void {
    cancelPageTasks(slot.pageNum);
    const canvases = slot.canvasWrapEl.querySelectorAll('canvas');
    canvases.forEach((canvas) => {
      canvas.width = 0;
      canvas.height = 0;
      canvas.remove();
    });
    slot.canvasWrapEl.innerHTML = '';
    slot.textLayerEl.innerHTML = '';
    slot.renderedScale = null;
  }

  function disposeStage(): void {
    if (pageObserver) {
      pageObserver.disconnect();
      pageObserver = null;
    }
    cancelAllPageTasks();
    for (const slot of pageSlots.values()) {
      releasePageSlot(slot);
    }
    pageSlots.clear();
    visiblePageSet.clear();
    if (stageEl) {
      stageEl.innerHTML = '';
    }
  }

  function disposeThumbnails(): void {
    if (thumbObserver) {
      thumbObserver.disconnect();
      thumbObserver = null;
    }
    thumbRenderSet.clear();
    for (const slot of thumbSlots.values()) {
      const canvases = slot.previewEl.querySelectorAll('canvas');
      canvases.forEach((canvas) => {
        canvas.width = 0;
        canvas.height = 0;
        canvas.remove();
      });
      slot.previewEl.innerHTML = '';
      slot.rendered = false;
    }
    thumbSlots.clear();
    if (thumbsListEl) {
      thumbsListEl.innerHTML = '';
    }
  }

  async function disposeDocument(): Promise<void> {
    disposeStage();
    disposeThumbnails();
    pageSizeCache.clear();
    pageTextCache.clear();

    if (loadingTask) {
      try {
        await loadingTask.destroy();
      } catch {
        // ignore
      }
      loadingTask = null;
    }
    if (pdfDoc) {
      try {
        await pdfDoc.destroy();
      } catch {
        // ignore
      }
      pdfDoc = null;
    }
  }

  function computeFitScale(pageWidthPts: number): number {
    if (!scrollEl || pageWidthPts <= 0) return 1.0;
    const availWidth = Math.max(200, scrollEl.clientWidth - PAGE_HORIZONTAL_PADDING);
    return clamp(roundZoom(availWidth / pageWidthPts), MIN_ZOOM, MAX_ZOOM);
  }

  function getPageBoxDimensions(pageNum: number): { width: number; height: number; scale: number } {
    const base = pageSizeCache.get(pageNum) ?? defaultPageSize;
    const scale = fitMode ? computeFitScale(base.width) : currentZoom;
    return {
      width: Math.max(1, Math.round(base.width * scale)),
      height: Math.max(1, Math.round(base.height * scale)),
      scale,
    };
  }

  function updateAllSlotSizes(): void {
    for (const slot of pageSlots.values()) {
      const { width, height } = getPageBoxDimensions(slot.pageNum);
      slot.boxEl.style.width = `${width}px`;
      slot.boxEl.style.height = `${height}px`;
    }
    const base1 = pageSizeCache.get(1) ?? defaultPageSize;
    if (fitMode) {
      currentZoom = computeFitScale(base1.width);
    }
  }

  function applyHighlightsToPage(pageNum: number): void {
    const slot = pageSlots.get(pageNum);
    if (!slot) return;

    // 清除旧高亮并还原纯文本节点
    const marks = slot.textLayerEl.querySelectorAll('mark.fp-pdf-search-hit');
    marks.forEach((mark) => {
      const parent = mark.parentNode;
      if (!parent) return;
      parent.replaceChild(document.createTextNode(mark.textContent ?? ''), mark);
      parent.normalize();
    });

    if (!searchQuery) return;
    const queryLower = searchQuery.toLowerCase();
    const activeMatch = activeMatchIndex >= 0 ? searchMatches[activeMatchIndex] : undefined;
    const targetMatchInPage =
      activeMatch && activeMatch.page === pageNum ? activeMatch.matchIndexInPage : -1;

    let pageMatchCursor = 0;
    const spans = slot.textLayerEl.querySelectorAll('span');

    spans.forEach((span) => {
      const text = span.textContent ?? '';
      if (!text) return;
      const lower = text.toLowerCase();
      if (!lower.includes(queryLower)) return;

      const frag = document.createDocumentFragment();
      let pos = 0;
      while (pos < text.length) {
        const idx = lower.indexOf(queryLower, pos);
        if (idx === -1) {
          frag.appendChild(document.createTextNode(text.slice(pos)));
          break;
        }
        if (idx > pos) {
          frag.appendChild(document.createTextNode(text.slice(pos, idx)));
        }
        const mark = document.createElement('mark');
        mark.className = 'fp-pdf-search-hit';
        if (pageMatchCursor === targetMatchInPage) {
          mark.classList.add('is-active');
        }
        mark.textContent = text.slice(idx, idx + searchQuery.length);
        frag.appendChild(mark);
        pageMatchCursor++;
        pos = idx + searchQuery.length;
      }

      span.textContent = '';
      span.appendChild(frag);
    });
  }

  async function renderPageSlot(slot: PageSlot, gen: number): Promise<void> {
    if (destroyed || gen !== loadGeneration || !pdfDoc || pdfMode !== 'canvas') {
      return;
    }

    const base = pageSizeCache.get(slot.pageNum) ?? defaultPageSize;
    const targetScale = fitMode ? computeFitScale(base.width) : currentZoom;

    if (
      slot.renderedScale !== null &&
      Math.abs(slot.renderedScale - targetScale) < 0.005 &&
      slot.canvasWrapEl.childElementCount > 0
    ) {
      applyHighlightsToPage(slot.pageNum);
      return;
    }

    cancelPageTasks(slot.pageNum);

    let page: PDFPageProxy;
    try {
      page = await pdfDoc.getPage(slot.pageNum);
    } catch {
      return;
    }
    if (destroyed || gen !== loadGeneration || !pdfDoc || pdfMode !== 'canvas') {
      return;
    }

    const baseViewport = page.getViewport({ scale: 1.0 });
    pageSizeCache.set(slot.pageNum, {
      width: baseViewport.width,
      height: baseViewport.height,
    });

    const effectiveScale = fitMode ? computeFitScale(baseViewport.width) : currentZoom;
    if (fitMode && slot.pageNum === currentPage) {
      currentZoom = effectiveScale;
    }

    const viewport = page.getViewport({ scale: effectiveScale });
    const cssWidth = Math.max(1, Math.round(viewport.width));
    const cssHeight = Math.max(1, Math.round(viewport.height));

    slot.boxEl.style.width = `${cssWidth}px`;
    slot.boxEl.style.height = `${cssHeight}px`;

    const dpr = Math.min(window.devicePixelRatio || 1, 2);
    const canvas = document.createElement('canvas');
    canvas.className = 'fp-pdf-page-canvas';
    canvas.width = Math.max(1, Math.floor(viewport.width * dpr));
    canvas.height = Math.max(1, Math.floor(viewport.height * dpr));
    canvas.style.width = `${cssWidth}px`;
    canvas.style.height = `${cssHeight}px`;

    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    slot.canvasWrapEl.innerHTML = '';
    slot.canvasWrapEl.appendChild(canvas);
    slot.textLayerEl.innerHTML = '';
    slot.textLayerEl.style.setProperty('--scale-factor', String(viewport.scale));

    const renderTask = page.render({
      canvasContext: ctx,
      viewport,
      transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : undefined,
      background: 'rgb(255,255,255)',
    });
    activeRenderTasks.set(slot.pageNum, renderTask);

    try {
      await renderTask.promise;
    } catch (err) {
      const name = err && typeof err === 'object' && 'name' in err ? String((err as { name: unknown }).name) : '';
      if (name === 'RenderingCancelledException') {
        return;
      }
      console.warn(`[pdfRenderer] 第 ${slot.pageNum} 页渲染失败:`, err);
      return;
    } finally {
      if (activeRenderTasks.get(slot.pageNum) === renderTask) {
        activeRenderTasks.delete(slot.pageNum);
      }
    }

    if (destroyed || gen !== loadGeneration || pdfMode !== 'canvas') {
      return;
    }

    slot.renderedScale = effectiveScale;

    // 渲染可选择/高亮文本层
    try {
      const textContent = await page.getTextContent();
      if (destroyed || gen !== loadGeneration || pdfMode !== 'canvas') return;
      const textTask = renderTextLayer({
        textContentSource: textContent,
        container: slot.textLayerEl,
        viewport,
        textDivs: [],
      });
      activeTextTasks.set(slot.pageNum, textTask);
      await textTask.promise;
      if (activeTextTasks.get(slot.pageNum) === textTask) {
        activeTextTasks.delete(slot.pageNum);
      }
      applyHighlightsToPage(slot.pageNum);
    } catch {
      // 取消或文本层异常不影响主画布预览
    }
  }

  function releaseFarPages(anchorPage: number): void {
    if (layoutMode !== 'list') return;
    for (const [pageNum, slot] of pageSlots.entries()) {
      if (visiblePageSet.has(pageNum)) continue;
      if (Math.abs(pageNum - anchorPage) > FAR_PAGE_DISTANCE && slot.renderedScale !== null) {
        releasePageSlot(slot);
      }
    }
  }

  function updateActiveThumbnail(): void {
    if (!thumbnailsVisible) return;
    for (const [pageNum, slot] of thumbSlots.entries()) {
      const isActive = pageNum === currentPage;
      slot.itemEl.classList.toggle('is-active', isActive);
      slot.itemEl.setAttribute('aria-current', isActive ? 'true' : 'false');
      if (isActive) {
        slot.itemEl.scrollIntoView({ block: 'nearest', inline: 'nearest' });
      }
    }
  }

  async function renderThumbnail(slot: ThumbSlot, gen: number): Promise<void> {
    if (destroyed || gen !== loadGeneration || !pdfDoc || slot.rendered || thumbRenderSet.has(slot.pageNum)) {
      return;
    }
    thumbRenderSet.add(slot.pageNum);

    try {
      const page = await pdfDoc.getPage(slot.pageNum);
      if (destroyed || gen !== loadGeneration || !thumbnailsVisible) return;

      const baseViewport = page.getViewport({ scale: 1.0 });
      const targetWidth = 112;
      const scale = targetWidth / Math.max(1, baseViewport.width);
      const viewport = page.getViewport({ scale });
      const dpr = Math.min(window.devicePixelRatio || 1, 2);

      const canvas = document.createElement('canvas');
      canvas.className = 'fp-pdf-thumb-canvas';
      canvas.width = Math.max(1, Math.floor(viewport.width * dpr));
      canvas.height = Math.max(1, Math.floor(viewport.height * dpr));
      canvas.style.width = `${Math.round(viewport.width)}px`;
      canvas.style.height = `${Math.round(viewport.height)}px`;

      const ctx = canvas.getContext('2d');
      if (!ctx) return;

      const task = page.render({
        canvasContext: ctx,
        viewport,
        transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : undefined,
        background: 'rgb(255,255,255)',
      });
      await task.promise;

      if (destroyed || gen !== loadGeneration || !thumbnailsVisible) {
        canvas.width = 0;
        canvas.height = 0;
        return;
      }

      slot.previewEl.innerHTML = '';
      slot.previewEl.appendChild(canvas);
      slot.rendered = true;
    } catch {
      // ignore thumbnail cancellation/errors
    } finally {
      thumbRenderSet.delete(slot.pageNum);
    }
  }

  function releaseThumbnail(slot: ThumbSlot): void {
    if (!slot.rendered) return;
    const canvases = slot.previewEl.querySelectorAll('canvas');
    canvases.forEach((canvas) => {
      canvas.width = 0;
      canvas.height = 0;
      canvas.remove();
    });
    slot.previewEl.innerHTML = '';
    slot.rendered = false;
  }

  function buildThumbnails(): void {
    disposeThumbnails();
    if (!thumbnailsVisible || !thumbsListEl || !pdfDoc || pdfMode !== 'canvas') {
      return;
    }

    const gen = loadGeneration;
    const slots: ThumbSlot[] = [];

    for (let p = 1; p <= totalPages; p++) {
      const itemEl = document.createElement('button');
      itemEl.type = 'button';
      itemEl.className = 'fp-pdf-thumb-item';
      itemEl.dataset.page = String(p);
      if (p === currentPage) {
        itemEl.classList.add('is-active');
        itemEl.setAttribute('aria-current', 'true');
      }

      const previewEl = document.createElement('div');
      previewEl.className = 'fp-pdf-thumb-preview';

      const labelEl = document.createElement('span');
      labelEl.className = 'fp-pdf-thumb-label';
      labelEl.textContent = String(p);

      itemEl.append(previewEl, labelEl);
      itemEl.addEventListener('click', () => {
        goToPage(p);
      });

      thumbsListEl.appendChild(itemEl);
      const slot: ThumbSlot = { pageNum: p, itemEl, previewEl, rendered: false };
      thumbSlots.set(p, slot);
      slots.push(slot);
    }

    if (typeof IntersectionObserver !== 'undefined') {
      thumbObserver = new IntersectionObserver(
        (entries) => {
          for (const entry of entries) {
            const target = entry.target as HTMLElement;
            const pageNum = Number(target.dataset.page ?? '0');
            const slot = thumbSlots.get(pageNum);
            if (!slot) continue;
            if (entry.isIntersecting) {
              void renderThumbnail(slot, gen);
            } else {
              releaseThumbnail(slot);
            }
          }
        },
        { root: thumbsListEl, rootMargin: '160px 160px', threshold: 0 },
      );
      for (const slot of slots) {
        thumbObserver.observe(slot.itemEl);
      }
    } else {
      for (const slot of slots.slice(0, 12)) {
        void renderThumbnail(slot, gen);
      }
    }

    updateActiveThumbnail();
  }

  function createPageSlotElement(pageNum: number): PageSlot {
    const boxEl = document.createElement('div');
    boxEl.className = 'fp-pdf-page-box';
    boxEl.dataset.page = String(pageNum);

    const { width, height } = getPageBoxDimensions(pageNum);
    boxEl.style.width = `${width}px`;
    boxEl.style.height = `${height}px`;

    const canvasWrapEl = document.createElement('div');
    canvasWrapEl.className = 'fp-pdf-canvas-wrap';

    const textLayerEl = document.createElement('div');
    textLayerEl.className = 'fp-pdf-text-layer textLayer';

    const badgeEl = document.createElement('div');
    badgeEl.className = 'fp-pdf-page-badge';
    badgeEl.textContent = `${pageNum} / ${totalPages}`;

    boxEl.append(canvasWrapEl, textLayerEl, badgeEl);
    return {
      pageNum,
      boxEl,
      canvasWrapEl,
      textLayerEl,
      badgeEl,
      renderedScale: null,
    };
  }

  function syncVisiblePageFromScroll(): void {
    if (!scrollEl || layoutMode !== 'list' || pageSlots.size === 0) return;
    const viewportRect = scrollEl.getBoundingClientRect();
    const centerY = viewportRect.top + viewportRect.height * 0.35;

    let bestPage = currentPage;
    let bestDistance = Number.POSITIVE_INFINITY;

    const candidates = visiblePageSet.size > 0 ? Array.from(visiblePageSet) : [currentPage];
    for (const p of candidates) {
      const slot = pageSlots.get(p);
      if (!slot) continue;
      const rect = slot.boxEl.getBoundingClientRect();
      const pageCenter = rect.top + rect.height / 2;
      const dist = Math.abs(pageCenter - centerY);
      if (dist < bestDistance) {
        bestDistance = dist;
        bestPage = p;
      }
    }

    if (bestPage !== currentPage) {
      currentPage = bestPage;
      updateActiveThumbnail();
      releaseFarPages(currentPage);
      emitState();
    }
  }

  function renderCanvasStage(): void {
    if (destroyed || !stageEl || !scrollEl || !pdfDoc || pdfMode !== 'canvas') return;

    disposeStage();
    const gen = loadGeneration;

    if (nativeFrameEl) {
      nativeFrameEl.src = 'about:blank';
      nativeFrameEl.remove();
      nativeFrameEl = null;
    }
    revokeNativeBlobUrl();

    if (thumbsPanelEl) {
      thumbsPanelEl.hidden = !thumbnailsVisible;
    }
    scrollEl.hidden = false;

    const base1 = pageSizeCache.get(1) ?? defaultPageSize;
    if (fitMode) {
      currentZoom = computeFitScale(base1.width);
    }

    if (layoutMode === 'slide') {
      const pageToShow = clamp(currentPage, 1, Math.max(1, totalPages));
      currentPage = pageToShow;
      const slot = createPageSlotElement(pageToShow);
      pageSlots.set(pageToShow, slot);
      stageEl.appendChild(slot.boxEl);
      visiblePageSet.add(pageToShow);
      void renderPageSlot(slot, gen);
    } else {
      const frag = document.createDocumentFragment();
      for (let p = 1; p <= totalPages; p++) {
        const slot = createPageSlotElement(p);
        pageSlots.set(p, slot);
        frag.appendChild(slot.boxEl);
      }
      stageEl.appendChild(frag);

      if (typeof IntersectionObserver !== 'undefined') {
        pageObserver = new IntersectionObserver(
          (entries) => {
            if (destroyed || gen !== loadGeneration) return;
            for (const entry of entries) {
              const target = entry.target as HTMLElement;
              const pageNum = Number(target.dataset.page ?? '0');
              const slot = pageSlots.get(pageNum);
              if (!slot) continue;
              if (entry.isIntersecting) {
                visiblePageSet.add(pageNum);
                void renderPageSlot(slot, gen);
              } else {
                visiblePageSet.delete(pageNum);
                if (Math.abs(pageNum - currentPage) > FAR_PAGE_DISTANCE && slot.renderedScale !== null) {
                  releasePageSlot(slot);
                }
              }
            }
            syncVisiblePageFromScroll();
          },
          { root: scrollEl, rootMargin: '360px 0px', threshold: 0 },
        );
        for (const slot of pageSlots.values()) {
          pageObserver.observe(slot.boxEl);
        }
      } else {
        const start = Math.max(1, currentPage - 1);
        const end = Math.min(totalPages, currentPage + 2);
        for (let p = start; p <= end; p++) {
          const slot = pageSlots.get(p);
          if (slot) {
            visiblePageSet.add(p);
            void renderPageSlot(slot, gen);
          }
        }
      }
    }

    if (thumbnailsVisible) {
      buildThumbnails();
    }
  }

  function renderNativeStage(): void {
    if (destroyed || !bodyEl || !scrollEl) return;

    disposeStage();
    disposeThumbnails();
    if (thumbsPanelEl) {
      thumbsPanelEl.hidden = true;
    }
    scrollEl.hidden = true;

    if (nativeFrameEl) {
      nativeFrameEl.src = 'about:blank';
      nativeFrameEl.remove();
      nativeFrameEl = null;
    }
    revokeNativeBlobUrl();

    const iframe = document.createElement('iframe');
    iframe.className = 'fp-pdf-native-frame';
    iframe.title = currentMeta ? `PDF 原生预览: ${currentMeta.name}` : 'PDF 原生预览';

    if (currentMountInfo?.url) {
      iframe.src = currentMountInfo.url;
    } else if (currentBuffer) {
      nativeBlobUrl = URL.createObjectURL(new Blob([currentBuffer], { type: 'application/pdf' }));
      iframe.src = nativeBlobUrl;
    }

    nativeFrameEl = iframe;
    bodyEl.appendChild(iframe);
  }

  function renderActiveMode(): void {
    if (pdfMode === 'native') {
      renderNativeStage();
    } else {
      renderCanvasStage();
    }
  }

  async function openPdfWithFallback(
    mountInfo: PdfMountInfo | null,
    buffer: ArrayBuffer | null,
  ): Promise<PDFDocumentProxy> {
    let workerDisabledFlag = false;

    const buildParams = (disableWorker: boolean): PdfInitParameters => {
      const base: PdfInitParameters & { disableWorker?: boolean } = mountInfo?.url
        ? {
            url: mountInfo.url,
            rangeChunkSize: RANGE_CHUNK_SIZE,
            disableAutoFetch: true,
            disableStream: false,
          }
        : buffer
          ? {
              data: new Uint8Array(buffer.slice(0)),
            }
          : (() => {
              throw new Error('缺少 PDF 数据源或挂载 URL');
            })();

      if (disableWorker) {
        base.disableWorker = true;
      }
      return base;
    };

    const runAttempt = async (disableWorker: boolean): Promise<PDFDocumentProxy> => {
      const task = getDocument(buildParams(disableWorker));
      loadingTask = task;
      const doc = await task.promise;
      if (loadingTask === task) {
        loadingTask = null;
      }
      return doc;
    };

    try {
      return await runAttempt(false);
    } catch (err) {
      const workerUtil = (
        pdfjsLib as unknown as { PDFWorkerUtil?: { isWorkerDisabled?: boolean } }
      ).PDFWorkerUtil;
      workerDisabledFlag = Boolean(workerUtil?.isWorkerDisabled);

      if (!workerDisabledFlag && !isNonWorkerFatalError(err)) {
        console.warn('[pdfRenderer] PDF Worker 加载失败，回退主线程模式重试:', err);
        if (workerUtil) {
          workerUtil.isWorkerDisabled = true;
        }
        try {
          return await runAttempt(true);
        } catch (retryErr) {
          throw formatPdfLoadError(retryErr);
        }
      }
      throw formatPdfLoadError(err);
    }
  }

  async function openCurrentSource(): Promise<void> {
    if (destroyed || !stageEl || !scrollEl) return;
    const gen = ++loadGeneration;

    await disposeDocument();
    if (destroyed || gen !== loadGeneration) return;

    searchQuery = '';
    searchMatches = [];
    activeMatchIndex = -1;
    activeWarning = undefined;
    currentPage = 1;
    totalPages = 0;

    if (currentBuffer && currentBuffer.byteLength === 0) {
      throw new Error('PDF 文件为空（0 字节），无法渲染');
    }

    const doc = await openPdfWithFallback(currentMountInfo, currentBuffer);
    if (destroyed || gen !== loadGeneration) {
      await doc.destroy();
      return;
    }

    pdfDoc = doc;
    totalPages = doc.numPages;
    currentPage = totalPages > 0 ? 1 : 0;

    if (totalPages > 0) {
      try {
        const firstPage = await doc.getPage(1);
        if (destroyed || gen !== loadGeneration) return;
        const vp = firstPage.getViewport({ scale: 1.0 });
        defaultPageSize = { width: vp.width, height: vp.height };
        pageSizeCache.set(1, { ...defaultPageSize });
      } catch {
        // 保留默认 A4 比例占位尺寸
      }
    }

    renderActiveMode();
    emitState();
  }

  function goToPage(page: number): void {
    if (destroyed || totalPages <= 0 || !Number.isFinite(page)) return;
    const clamped = clamp(Math.round(page), 1, totalPages);
    currentPage = clamped;

    if (pdfMode === 'canvas') {
      if (layoutMode === 'slide') {
        renderCanvasStage();
      } else {
        const slot = pageSlots.get(clamped);
        if (slot) {
          slot.boxEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
          void renderPageSlot(slot, loadGeneration);
        }
        releaseFarPages(clamped);
      }
      updateActiveThumbnail();
    }

    emitState();
  }

  function nextPage(): void {
    if (totalPages <= 0 || currentPage >= totalPages) return;
    goToPage(currentPage + 1);
  }

  function prevPage(): void {
    if (totalPages <= 0 || currentPage <= 1) return;
    goToPage(currentPage - 1);
  }

  async function extractPageSearchText(pageNum: number): Promise<string> {
    const cached = pageTextCache.get(pageNum);
    if (cached !== undefined) return cached;
    if (!pdfDoc) return '';

    const page = await pdfDoc.getPage(pageNum);
    const textContent = await page.getTextContent();
    let text = '';
    for (const item of textContent.items) {
      if ('str' in item) {
        text += item.str;
        if (item.hasEOL && !text.endsWith(' ') && !text.endsWith('\n')) {
          text += ' ';
        }
      }
    }
    pageTextCache.set(pageNum, text);
    return text;
  }

  async function focusMatch(index: number): Promise<void> {
    if (destroyed || index < 0 || index >= searchMatches.length) return;
    const match = searchMatches[index];
    currentPage = match.page;

    if (pdfMode === 'canvas') {
      if (layoutMode === 'slide') {
        renderCanvasStage();
      } else {
        const slot = pageSlots.get(match.page);
        if (slot) {
          slot.boxEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
      }
      const targetSlot = pageSlots.get(match.page);
      if (targetSlot) {
        await renderPageSlot(targetSlot, loadGeneration);
        applyHighlightsToPage(match.page);
        const activeMark = targetSlot.textLayerEl.querySelector('mark.fp-pdf-search-hit.is-active');
        if (activeMark) {
          activeMark.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
      }
      for (const p of visiblePageSet) {
        if (p !== match.page) {
          applyHighlightsToPage(p);
        }
      }
      updateActiveThumbnail();
    }
    emitState();
  }

  function handleKeyDown(e: KeyboardEvent): void {
    if (destroyed || pdfMode !== 'canvas' || totalPages <= 0) return;
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

  return {
    mount(container: HTMLElement): void {
      if (destroyed) {
        destroyed = false;
      }
      container.innerHTML = '';

      rootEl = document.createElement('div');
      rootEl.className = 'fp-pdf-root';

      bodyEl = document.createElement('div');
      bodyEl.className = 'fp-pdf-body';

      thumbsPanelEl = document.createElement('aside');
      thumbsPanelEl.className = 'fp-pdf-thumbs';
      thumbsPanelEl.hidden = !thumbnailsVisible;

      const thumbsHeaderEl = document.createElement('div');
      thumbsHeaderEl.className = 'fp-pdf-thumbs-header';

      const thumbsTitleEl = document.createElement('span');
      thumbsTitleEl.textContent = '页面缩略图';

      const thumbsCloseBtn = document.createElement('button');
      thumbsCloseBtn.type = 'button';
      thumbsCloseBtn.className = 'fp-pdf-thumbs-close';
      thumbsCloseBtn.title = '收起缩略图';
      thumbsCloseBtn.setAttribute('aria-label', '收起缩略图');
      thumbsCloseBtn.textContent = '×';
      thumbsCloseBtn.addEventListener('click', () => {
        this.toggleThumbnails?.(false);
      });

      thumbsHeaderEl.append(thumbsTitleEl, thumbsCloseBtn);

      thumbsListEl = document.createElement('div');
      thumbsListEl.className = 'fp-pdf-thumbs-list';

      thumbsPanelEl.append(thumbsHeaderEl, thumbsListEl);

      scrollEl = document.createElement('div');
      scrollEl.className = 'fp-pdf-scroll';
      scrollEl.addEventListener('scroll', () => {
        if (scrollSyncRaf !== null) return;
        scrollSyncRaf = window.requestAnimationFrame(() => {
          scrollSyncRaf = null;
          if (!destroyed && pdfMode === 'canvas') {
            syncVisiblePageFromScroll();
          }
        });
      });

      stageEl = document.createElement('div');
      stageEl.className = 'fp-pdf-stage';
      stageEl.style.gap = `${PAGE_GAP_PX}px`;

      scrollEl.appendChild(stageEl);
      bodyEl.append(thumbsPanelEl, scrollEl);
      rootEl.appendChild(bodyEl);
      container.appendChild(rootEl);

      if (typeof ResizeObserver !== 'undefined') {
        resizeObserver = new ResizeObserver(() => {
          if (destroyed || !fitMode || pdfMode !== 'canvas' || !pdfDoc) return;
          if (resizeTimer !== null) {
            window.clearTimeout(resizeTimer);
          }
          resizeTimer = window.setTimeout(() => {
            resizeTimer = null;
            if (destroyed || !fitMode || pdfMode !== 'canvas' || !pdfDoc) return;
            updateAllSlotSizes();
            for (const p of visiblePageSet) {
              const slot = pageSlots.get(p);
              if (slot) {
                void renderPageSlot(slot, loadGeneration);
              }
            }
            emitState();
          }, 120);
        });
        resizeObserver.observe(scrollEl);
      }

      window.addEventListener('keydown', handleKeyDown);
    },

    async loadPdf(
      mountInfo: PdfMountInfo | null,
      buffer: ArrayBuffer | null,
      meta: { name: string; path: string; ext: string; size?: number },
    ): Promise<void> {
      currentMountInfo = mountInfo;
      currentBuffer = buffer;
      currentMeta = meta;
      await openCurrentSource();
    },

    async refresh(): Promise<void> {
      if (!currentMeta) return;
      await openCurrentSource();
    },

    setZoom(zoom: number | 'fit'): void {
      if (zoom === 'fit') {
        fitMode = true;
        const base = pageSizeCache.get(currentPage) ?? defaultPageSize;
        currentZoom = computeFitScale(base.width);
      } else if (Number.isFinite(zoom)) {
        fitMode = false;
        currentZoom = clamp(roundZoom(zoom), MIN_ZOOM, MAX_ZOOM);
      } else {
        return;
      }

      if (pdfMode === 'canvas' && pdfDoc) {
        updateAllSlotSizes();
        for (const p of visiblePageSet) {
          const slot = pageSlots.get(p);
          if (slot) {
            void renderPageSlot(slot, loadGeneration);
          }
        }
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

      if (pdfMode === 'canvas' && pdfDoc) {
        const target = currentPage;
        renderCanvasStage();
        if (mode === 'list' && target > 1) {
          const slot = pageSlots.get(target);
          slot?.boxEl.scrollIntoView({ behavior: 'instant', block: 'start' });
        }
      }
      emitState();
    },

    toggleThumbnails(visible?: boolean): void {
      const next = visible !== undefined ? visible : !thumbnailsVisible;
      if (next === thumbnailsVisible) return;
      thumbnailsVisible = next;
      if (thumbsPanelEl) {
        thumbsPanelEl.hidden = !thumbnailsVisible || pdfMode !== 'canvas';
      }
      if (pdfMode === 'canvas' && pdfDoc) {
        if (thumbnailsVisible) {
          buildThumbnails();
        } else {
          disposeThumbnails();
        }
      }
      emitState();
    },

    setPdfMode(mode: PdfViewMode): void {
      if (mode !== 'canvas' && mode !== 'native') return;
      if (pdfMode === mode) return;
      pdfMode = mode;

      if (pdfDoc || currentMountInfo || currentBuffer) {
        renderActiveMode();
        if (mode === 'canvas' && layoutMode === 'list' && currentPage > 1) {
          const slot = pageSlots.get(currentPage);
          slot?.boxEl.scrollIntoView({ behavior: 'instant', block: 'start' });
        }
      }
      emitState();
    },

    async search(query: string): Promise<{ total: number; current: number }> {
      const trimmed = query.trim();
      searchQuery = trimmed;
      searchMatches = [];
      activeMatchIndex = -1;

      if (!trimmed || !pdfDoc || pdfMode !== 'canvas') {
        for (const p of visiblePageSet) {
          applyHighlightsToPage(p);
        }
        emitState();
        return { total: 0, current: 0 };
      }

      const gen = loadGeneration;
      const queryLower = trimmed.toLowerCase();
      const found: PdfSearchMatch[] = [];

      for (let p = 1; p <= totalPages; p++) {
        if (destroyed || gen !== loadGeneration) {
          return { total: 0, current: 0 };
        }
        const text = await extractPageSearchText(p);
        const lower = text.toLowerCase();
        let pos = 0;
        let inPageIdx = 0;
        while (pos < lower.length) {
          const idx = lower.indexOf(queryLower, pos);
          if (idx === -1) break;
          found.push({ page: p, matchIndexInPage: inPageIdx++ });
          pos = idx + queryLower.length;
        }
      }

      if (destroyed || gen !== loadGeneration) {
        return { total: 0, current: 0 };
      }

      searchMatches = found;
      if (found.length === 0) {
        for (const p of visiblePageSet) {
          applyHighlightsToPage(p);
        }
        emitState();
        return { total: 0, current: 0 };
      }

      activeMatchIndex = 0;
      await focusMatch(0);
      return { total: searchMatches.length, current: 1 };
    },

    nextMatch(): void {
      if (searchMatches.length === 0) return;
      activeMatchIndex = (activeMatchIndex + 1) % searchMatches.length;
      void focusMatch(activeMatchIndex);
    },

    prevMatch(): void {
      if (searchMatches.length === 0) return;
      activeMatchIndex = (activeMatchIndex - 1 + searchMatches.length) % searchMatches.length;
      void focusMatch(activeMatchIndex);
    },

    onStateChange(cb: (info: PreviewStateInfo) => void): void {
      stateListeners.add(cb);
    },

    destroy(): void {
      destroyed = true;
      loadGeneration++;

      window.removeEventListener('keydown', handleKeyDown);

      if (resizeTimer !== null) {
        window.clearTimeout(resizeTimer);
        resizeTimer = null;
      }
      if (scrollSyncRaf !== null) {
        window.cancelAnimationFrame(scrollSyncRaf);
        scrollSyncRaf = null;
      }
      if (resizeObserver) {
        resizeObserver.disconnect();
        resizeObserver = null;
      }

      if (nativeFrameEl) {
        nativeFrameEl.src = 'about:blank';
        nativeFrameEl.remove();
        nativeFrameEl = null;
      }
      revokeNativeBlobUrl();

      void disposeDocument();

      if (rootEl) {
        rootEl.remove();
      }

      rootEl = null;
      bodyEl = null;
      thumbsPanelEl = null;
      thumbsListEl = null;
      scrollEl = null;
      stageEl = null;
      currentMountInfo = null;
      currentBuffer = null;
      currentMeta = null;
      searchQuery = '';
      searchMatches = [];
      activeMatchIndex = -1;
      totalPages = 0;
      currentPage = 1;
      activeWarning = undefined;
      stateListeners.clear();
    },
  };
}
