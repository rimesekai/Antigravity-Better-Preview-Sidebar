import type {
  HtmlPreviewMode,
  PptLayoutMode,
  PreviewKind,
  PreviewStateInfo,
} from '../types/preview';
import type { ThemePreference } from '../services/themeSync';
import { formatKindBadge } from '../utils/fileMeta';
import {
  type LocalePreference,
  getCurrentLocale,
  getLocalePreference,
  getLocaleToggleDisplay,
  onLocaleChange,
  tr,
} from '../utils/i18n';

const LARGE_PDF_THRESHOLD_BYTES = 20 * 1024 * 1024;

export interface ToolbarFileMeta {
  name: string;
  path: string;
  ext: string;
  kind: PreviewKind;
  size?: number;
  isLocalDrop: boolean;
}

export interface ToolbarRenderState {
  visible: boolean;
  file: ToolbarFileMeta | null;
  preview: PreviewStateInfo | null;
  loading: boolean;
  hasError: boolean;
  autoRefresh: boolean;
  drawerOpen: boolean;
  themePreference?: ThemePreference;
  localePreference?: LocalePreference;
}

export interface ToolbarCallbacks {
  onToggleDrawer: () => void;
  onRefresh: () => void;
  onRevealInExplorer?: () => void;
  onToggleAutoRefresh: () => void;
  onCycleTheme?: () => void;
  onCycleLocale?: () => void;
  onZoomChange: (zoom: number | 'fit') => void;
  onPrevPage: () => void;
  onNextPage: () => void;
  onGoToPage: (page: number) => void;
  onLayoutModeChange: (mode: PptLayoutMode) => void;
  onToggleThumbnails: () => void;
  onSearch: (query: string) => Promise<void>;
  onNextMatch: () => void;
  onPrevMatch: () => void;
  onHtmlModeChange: (mode: HtmlPreviewMode) => void;
}

export interface ToolbarController {
  update: (state: ToolbarRenderState) => void;
  resetSearch: () => void;
  destroy: () => void;
}

const ZOOM_STEPS = [0.5, 0.67, 0.75, 0.9, 1.0, 1.25, 1.5, 1.75, 2.0];

function clampZoom(value: number): number {
  return Math.min(2.0, Math.max(0.5, Math.round(value * 100) / 100));
}

function stepZoom(current: number, direction: -1 | 1): number {
  const normalized = clampZoom(current || 1.0);
  if (direction > 0) {
    const next = ZOOM_STEPS.find((step) => step > normalized + 0.02);
    return next ?? clampZoom(normalized + 0.15);
  }
  const reversed = [...ZOOM_STEPS].reverse();
  const prev = reversed.find((step) => step < normalized - 0.02);
  return prev ?? clampZoom(normalized - 0.15);
}

function createSvgIcon(pathD: string, viewBox = '0 0 16 16'): SVGSVGElement {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('viewBox', viewBox);
  svg.setAttribute('width', '14');
  svg.setAttribute('height', '14');
  svg.setAttribute('fill', 'none');
  svg.setAttribute('stroke', 'currentColor');
  svg.setAttribute('stroke-width', '1.6');
  svg.setAttribute('stroke-linecap', 'round');
  svg.setAttribute('stroke-linejoin', 'round');
  svg.setAttribute('aria-hidden', 'true');
  const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  path.setAttribute('d', pathD);
  svg.appendChild(path);
  return svg;
}

/**
 * 创建顶部自适应工具栏控制器（持久 DOM 节点，避免输入时失焦）。
 */
export function createToolbar(
  container: HTMLElement,
  callbacks: ToolbarCallbacks,
): ToolbarController {
  let latestState: ToolbarRenderState = {
    visible: false,
    file: null,
    preview: null,
    loading: false,
    hasError: false,
    autoRefresh: true,
    drawerOpen: false,
  };

  interface SearchClusterController {
    clusterEl: HTMLElement;
    inputEl: HTMLInputElement;
    setPlaceholder: (placeholder: string, ariaLabel: string) => void;
    refreshLocale: () => void;
    updateMatches: (totalMatches: number, activeMatch: number) => void;
    reset: () => void;
  }

  /** 创建紧凑搜索控件组（PPT / PDF / Excel / Word 复用）。 */
  function createSearchCluster(placeholder: string, ariaLabel: string): SearchClusterController {
    let lastQuery = '';

    const clusterEl = document.createElement('div');
    clusterEl.className = 'fp-cluster fp-ppt-search';

    const inputEl = document.createElement('input');
    inputEl.type = 'search';
    inputEl.className = 'fp-search-input';
    inputEl.placeholder = placeholder;
    inputEl.setAttribute('aria-label', ariaLabel);

    const triggerSearch = () => {
      const query = inputEl.value.trim();
      if (query === lastQuery && query) {
        callbacks.onNextMatch();
        return;
      }
      lastQuery = query;
      void callbacks.onSearch(query);
    };

    inputEl.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        if (e.shiftKey && inputEl.value.trim() === lastQuery && lastQuery) {
          callbacks.onPrevMatch();
          return;
        }
        triggerSearch();
      } else if (e.key === 'Escape') {
        inputEl.value = '';
        lastQuery = '';
        void callbacks.onSearch('');
      }
    });

    inputEl.addEventListener('input', () => {
      if (!inputEl.value.trim() && lastQuery) {
        lastQuery = '';
        void callbacks.onSearch('');
      }
    });

    const countEl = document.createElement('span');
    countEl.className = 'fp-search-count';
    countEl.hidden = true;

    const prevBtn = document.createElement('button');
    prevBtn.type = 'button';
    prevBtn.className = 'fp-btn fp-btn-icon';
    prevBtn.textContent = '↑';
    prevBtn.title = tr('上一个匹配项', 'Previous match');
    prevBtn.hidden = true;
    prevBtn.addEventListener('click', () => callbacks.onPrevMatch());

    const nextBtn = document.createElement('button');
    nextBtn.type = 'button';
    nextBtn.className = 'fp-btn fp-btn-icon';
    nextBtn.textContent = '↓';
    nextBtn.title = tr('下一个匹配项', 'Next match');
    nextBtn.hidden = true;
    nextBtn.addEventListener('click', () => callbacks.onNextMatch());

    clusterEl.append(inputEl, countEl, prevBtn, nextBtn);

    return {
      clusterEl,
      inputEl,
      setPlaceholder(nextPlaceholder: string, nextAriaLabel: string): void {
        inputEl.placeholder = nextPlaceholder;
        inputEl.setAttribute('aria-label', nextAriaLabel);
      },
      refreshLocale(): void {
        prevBtn.title = tr('上一个匹配项', 'Previous match');
        nextBtn.title = tr('下一个匹配项', 'Next match');
      },
      updateMatches(totalMatches: number, activeMatch: number): void {
        const hasQuery = Boolean(lastQuery);
        countEl.hidden = !hasQuery && totalMatches === 0;
        if (hasQuery || totalMatches > 0) {
          countEl.textContent = totalMatches > 0 ? `${activeMatch}/${totalMatches}` : '0/0';
        }
        const navVisible = totalMatches > 0;
        prevBtn.hidden = !navVisible;
        nextBtn.hidden = !navVisible;
      },
      reset(): void {
        lastQuery = '';
        inputEl.value = '';
        countEl.hidden = true;
        prevBtn.hidden = true;
        nextBtn.hidden = true;
      },
    };
  }

  const root = document.createElement('div');
  root.className = 'fp-toolbar';
  root.hidden = true;

  // 第一行：文件导航、标题与刷新控制
  const primaryRow = document.createElement('div');
  primaryRow.className = 'fp-toolbar-row fp-toolbar-primary';

  const leftGroup = document.createElement('div');
  leftGroup.className = 'fp-toolbar-group fp-toolbar-left';

  const drawerBtn = document.createElement('button');
  drawerBtn.type = 'button';
  drawerBtn.className = 'fp-btn fp-btn-ghost fp-btn-files';
  drawerBtn.title = '打开或切换工作区文件列表';
  drawerBtn.appendChild(createSvgIcon('M2.5 4.5h4l1.2 1.5h5.8v6.5h-11z'));
  const drawerBtnText = document.createElement('span');
  drawerBtnText.textContent = '文件';
  drawerBtn.appendChild(drawerBtnText);
  drawerBtn.addEventListener('click', () => callbacks.onToggleDrawer());

  const fileBadge = document.createElement('span');
  fileBadge.className = 'fp-kind-badge';

  const fileNameEl = document.createElement('span');
  fileNameEl.className = 'fp-toolbar-filename';

  const largePdfBadge = document.createElement('span');
  largePdfBadge.className = 'fp-large-pdf-badge';
  largePdfBadge.textContent = '>20MB 流式预览';
  largePdfBadge.hidden = true;

  const dropTag = document.createElement('span');
  dropTag.className = 'fp-drop-tag';
  dropTag.textContent = '本地拖入';
  dropTag.hidden = true;

  leftGroup.append(drawerBtn, fileBadge, fileNameEl, largePdfBadge, dropTag);

  const rightGroup = document.createElement('div');
  rightGroup.className = 'fp-toolbar-group fp-toolbar-right';

  const revealBtn = document.createElement('button');
  revealBtn.type = 'button';
  revealBtn.className = 'fp-btn fp-btn-ghost fp-btn-icon-label';
  revealBtn.title = '在文件资源管理器中查看';
  revealBtn.appendChild(
    createSvgIcon('M2.2 4.5a1.3 1.3 0 0 1 1.3-1.3h3l1.4 1.6h4.6a1.3 1.3 0 0 1 1.3 1.3v5.4a1.3 1.3 0 0 1-1.3 1.3h-9a1.3 1.3 0 0 1-1.3-1.3z'),
  );
  const revealLabel = document.createElement('span');
  revealLabel.className = 'fp-hide-xs';
  revealLabel.textContent = '打开位置';
  revealBtn.appendChild(revealLabel);
  revealBtn.addEventListener('click', () => callbacks.onRevealInExplorer?.());

  const refreshBtn = document.createElement('button');
  refreshBtn.type = 'button';
  refreshBtn.className = 'fp-btn fp-btn-ghost fp-btn-icon-label';
  refreshBtn.title = '立即重新加载当前文件';
  refreshBtn.appendChild(
    createSvgIcon('M13.2 6.5A5.5 5.5 0 1 0 13.7 10M13.5 3.2v3.3h-3.3'),
  );
  const refreshLabel = document.createElement('span');
  refreshLabel.className = 'fp-hide-xs';
  refreshLabel.textContent = '刷新';
  refreshBtn.appendChild(refreshLabel);
  refreshBtn.addEventListener('click', () => callbacks.onRefresh());

  const autoRefreshBtn = document.createElement('button');
  autoRefreshBtn.type = 'button';
  autoRefreshBtn.className = 'fp-btn fp-btn-toggle';
  autoRefreshBtn.title = '每 2.5 秒检测文件修改时间并自动刷新';
  const autoDot = document.createElement('span');
  autoDot.className = 'fp-auto-dot';
  const autoText = document.createElement('span');
  autoText.textContent = '自动刷新';
  autoRefreshBtn.append(autoDot, autoText);
  autoRefreshBtn.addEventListener('click', () => callbacks.onToggleAutoRefresh());

  const themeBtn = document.createElement('button');
  themeBtn.type = 'button';
  themeBtn.className = 'fp-btn fp-btn-ghost';
  themeBtn.addEventListener('click', () => callbacks.onCycleTheme?.());

  const langBtn = document.createElement('button');
  langBtn.type = 'button';
  langBtn.className = 'fp-btn fp-btn-ghost';
  langBtn.addEventListener('click', () => callbacks.onCycleLocale?.());

  rightGroup.append(langBtn, themeBtn, revealBtn, refreshBtn, autoRefreshBtn);
  primaryRow.append(leftGroup, rightGroup);

  // 第二行：格式专属工具条
  const secondaryRow = document.createElement('div');
  secondaryRow.className = 'fp-toolbar-row fp-toolbar-secondary';

  // --- PPT / PDF 分页工具条 ---
  const pptGroup = document.createElement('div');
  pptGroup.className = 'fp-format-controls fp-format-ppt';
  pptGroup.hidden = true;

  const pptLayoutSeg = document.createElement('div');
  pptLayoutSeg.className = 'fp-segmented';
  const pptSlideBtn = document.createElement('button');
  pptSlideBtn.type = 'button';
  pptSlideBtn.className = 'fp-seg-btn';
  pptSlideBtn.textContent = '单页';
  pptSlideBtn.addEventListener('click', () => callbacks.onLayoutModeChange('slide'));

  const pptListBtn = document.createElement('button');
  pptListBtn.type = 'button';
  pptListBtn.className = 'fp-seg-btn';
  pptListBtn.textContent = '滚动';
  pptListBtn.addEventListener('click', () => callbacks.onLayoutModeChange('list'));
  pptLayoutSeg.append(pptSlideBtn, pptListBtn);

  const pptThumbBtn = document.createElement('button');
  pptThumbBtn.type = 'button';
  pptThumbBtn.className = 'fp-btn fp-btn-ghost';
  pptThumbBtn.title = '切换幻灯片缩略图栏';
  pptThumbBtn.textContent = '缩略图';
  pptThumbBtn.addEventListener('click', () => callbacks.onToggleThumbnails());

  const pptPageCluster = document.createElement('div');
  pptPageCluster.className = 'fp-cluster';

  const prevPageBtn = document.createElement('button');
  prevPageBtn.type = 'button';
  prevPageBtn.className = 'fp-btn fp-btn-icon';
  prevPageBtn.title = '上一页';
  prevPageBtn.textContent = '‹';
  prevPageBtn.addEventListener('click', () => callbacks.onPrevPage());

  const pageInput = document.createElement('input');
  pageInput.type = 'number';
  pageInput.min = '1';
  pageInput.className = 'fp-page-input';
  pageInput.setAttribute('aria-label', '当前页码');

  const commitPageInput = () => {
    const total = latestState.preview?.totalPages ?? 1;
    const raw = Number.parseInt(pageInput.value, 10);
    if (!Number.isFinite(raw)) {
      pageInput.value = String(latestState.preview?.page ?? 1);
      return;
    }
    const clamped = Math.min(Math.max(1, raw), Math.max(1, total));
    pageInput.value = String(clamped);
    callbacks.onGoToPage(clamped);
  };

  pageInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      commitPageInput();
      pageInput.blur();
    }
  });
  pageInput.addEventListener('change', commitPageInput);

  const pageTotalEl = document.createElement('span');
  pageTotalEl.className = 'fp-page-total';
  pageTotalEl.textContent = '/ 1';

  const nextPageBtn = document.createElement('button');
  nextPageBtn.type = 'button';
  nextPageBtn.className = 'fp-btn fp-btn-icon';
  nextPageBtn.title = '下一页';
  nextPageBtn.textContent = '›';
  nextPageBtn.addEventListener('click', () => callbacks.onNextPage());

  pptPageCluster.append(prevPageBtn, pageInput, pageTotalEl, nextPageBtn);

  const sheetNamePill = document.createElement('span');
  sheetNamePill.className = 'fp-sheet-name-pill';
  sheetNamePill.hidden = true;

  // PPT 缩放控件
  const pptZoomCluster = document.createElement('div');
  pptZoomCluster.className = 'fp-cluster';

  const pptFitBtn = document.createElement('button');
  pptFitBtn.type = 'button';
  pptFitBtn.className = 'fp-btn fp-btn-ghost';
  pptFitBtn.textContent = '适应';
  pptFitBtn.title = '自适应侧边栏宽度';
  pptFitBtn.addEventListener('click', () => callbacks.onZoomChange('fit'));

  const pptZoomOutBtn = document.createElement('button');
  pptZoomOutBtn.type = 'button';
  pptZoomOutBtn.className = 'fp-btn fp-btn-icon';
  pptZoomOutBtn.textContent = '−';
  pptZoomOutBtn.title = '缩小';
  pptZoomOutBtn.addEventListener('click', () => {
    const current = latestState.preview?.zoom ?? 1.0;
    callbacks.onZoomChange(stepZoom(current, -1));
  });

  const pptZoomText = document.createElement('span');
  pptZoomText.className = 'fp-zoom-value';
  pptZoomText.textContent = '100%';

  const pptZoomInBtn = document.createElement('button');
  pptZoomInBtn.type = 'button';
  pptZoomInBtn.className = 'fp-btn fp-btn-icon';
  pptZoomInBtn.textContent = '+';
  pptZoomInBtn.title = '放大';
  pptZoomInBtn.addEventListener('click', () => {
    const current = latestState.preview?.zoom ?? 1.0;
    callbacks.onZoomChange(stepZoom(current, 1));
  });

  pptZoomCluster.append(pptFitBtn, pptZoomOutBtn, pptZoomText, pptZoomInBtn);

  // PPT / PDF / Excel 紧凑搜索框
  const pptSearch = createSearchCluster('搜索幻灯片 (回车)', '搜索幻灯片文本');

  pptGroup.append(
    pptLayoutSeg,
    pptThumbBtn,
    pptPageCluster,
    sheetNamePill,
    pptZoomCluster,
    pptSearch.clusterEl,
  );

  // --- Word 工具条 ---
  const wordGroup = document.createElement('div');
  wordGroup.className = 'fp-format-controls fp-format-word';
  wordGroup.hidden = true;

  const wordZoomCluster = document.createElement('div');
  wordZoomCluster.className = 'fp-cluster';

  const wordFitBtn = document.createElement('button');
  wordFitBtn.type = 'button';
  wordFitBtn.className = 'fp-btn fp-btn-ghost';
  wordFitBtn.textContent = '自适应宽度';
  wordFitBtn.addEventListener('click', () => callbacks.onZoomChange('fit'));

  const wordZoomOutBtn = document.createElement('button');
  wordZoomOutBtn.type = 'button';
  wordZoomOutBtn.className = 'fp-btn fp-btn-icon';
  wordZoomOutBtn.textContent = '−';
  wordZoomOutBtn.title = '缩小';
  wordZoomOutBtn.addEventListener('click', () => {
    const current = latestState.preview?.zoom ?? 1.0;
    callbacks.onZoomChange(stepZoom(current, -1));
  });

  const wordZoomText = document.createElement('span');
  wordZoomText.className = 'fp-zoom-value';
  wordZoomText.textContent = '100%';

  const wordZoomInBtn = document.createElement('button');
  wordZoomInBtn.type = 'button';
  wordZoomInBtn.className = 'fp-btn fp-btn-icon';
  wordZoomInBtn.textContent = '+';
  wordZoomInBtn.title = '放大';
  wordZoomInBtn.addEventListener('click', () => {
    const current = latestState.preview?.zoom ?? 1.0;
    callbacks.onZoomChange(stepZoom(current, 1));
  });

  wordZoomCluster.append(wordFitBtn, wordZoomOutBtn, wordZoomText, wordZoomInBtn);

  const wordSearch = createSearchCluster('搜索文档内容 (回车)', '搜索 Word 文档内容');

  wordGroup.append(wordZoomCluster, wordSearch.clusterEl);

  // --- HTML 工具条 ---
  const htmlGroup = document.createElement('div');
  htmlGroup.className = 'fp-format-controls fp-format-html';
  htmlGroup.hidden = true;

  const htmlModeSeg = document.createElement('div');
  htmlModeSeg.className = 'fp-segmented';

  const htmlSandboxBtn = document.createElement('button');
  htmlSandboxBtn.type = 'button';
  htmlSandboxBtn.className = 'fp-seg-btn';
  htmlSandboxBtn.textContent = '沙箱预览';
  htmlSandboxBtn.addEventListener('click', () => callbacks.onHtmlModeChange('sandbox'));

  const htmlInteractiveBtn = document.createElement('button');
  htmlInteractiveBtn.type = 'button';
  htmlInteractiveBtn.className = 'fp-seg-btn';
  htmlInteractiveBtn.textContent = '交互模式';
  htmlInteractiveBtn.addEventListener('click', () => callbacks.onHtmlModeChange('interactive'));

  const htmlSourceBtn = document.createElement('button');
  htmlSourceBtn.type = 'button';
  htmlSourceBtn.className = 'fp-seg-btn';
  htmlSourceBtn.textContent = '源码';
  htmlSourceBtn.addEventListener('click', () => callbacks.onHtmlModeChange('source'));

  htmlModeSeg.append(htmlSandboxBtn, htmlInteractiveBtn, htmlSourceBtn);

  const htmlStatusBanner = document.createElement('div');
  htmlStatusBanner.className = 'fp-html-status-pill';

  htmlGroup.append(htmlModeSeg, htmlStatusBanner);

  secondaryRow.append(pptGroup, wordGroup, htmlGroup);
  root.append(primaryRow, secondaryRow);
  container.appendChild(root);

  function updatePagedControls(preview: PreviewStateInfo | null, kind: PreviewKind): void {
    const isPdf = kind === 'pdf';
    const isExcel = kind === 'excel';

    pptLayoutSeg.hidden = isExcel;
    pptThumbBtn.hidden = isExcel;
    pptPageCluster.hidden = false;
    pptZoomCluster.hidden = false;
    pptSearch.clusterEl.hidden = false;

    if (isExcel) {
      prevPageBtn.title = tr('上一工作表', 'Previous sheet');
      nextPageBtn.title = tr('下一工作表', 'Next sheet');
      pageInput.setAttribute('aria-label', tr('当前工作表序号', 'Current sheet index'));
      pptFitBtn.title = tr('重置为标准视图', 'Reset to 100% view');
      pptSearch.setPlaceholder(tr('搜索表格内容 (回车)', 'Search sheet (Enter)'), tr('搜索表格内容', 'Search sheet'));
    } else if (isPdf) {
      prevPageBtn.title = tr('上一页', 'Previous page');
      nextPageBtn.title = tr('下一页', 'Next page');
      pageInput.setAttribute('aria-label', tr('当前页码', 'Current page'));
      pptFitBtn.title = tr('自适应侧边栏宽度', 'Fit sidebar width');
      pptThumbBtn.title = tr('切换页面缩略图栏', 'Toggle page thumbnails');
      pptSearch.setPlaceholder(tr('搜索 PDF 文本 (回车)', 'Search PDF (Enter)'), tr('搜索 PDF 文本', 'Search PDF text'));
    } else {
      prevPageBtn.title = tr('上一页', 'Previous page');
      nextPageBtn.title = tr('下一页', 'Next page');
      pageInput.setAttribute('aria-label', tr('当前页码', 'Current page'));
      pptFitBtn.title = tr('自适应侧边栏宽度', 'Fit sidebar width');
      pptThumbBtn.title = tr('切换幻灯片缩略图栏', 'Toggle slide thumbnails');
      pptSearch.setPlaceholder(tr('搜索幻灯片 (回车)', 'Search slides (Enter)'), tr('搜索幻灯片文本', 'Search slide text'));
    }

    const activeSheet = preview?.activeSheet?.trim();
    if (isExcel && activeSheet) {
      sheetNamePill.hidden = false;
      sheetNamePill.textContent = activeSheet;
      sheetNamePill.title = tr(`当前工作表：${activeSheet}`, `Active sheet: ${activeSheet}`);
    } else {
      sheetNamePill.hidden = true;
      sheetNamePill.textContent = '';
      sheetNamePill.title = '';
    }

    const page = preview?.page ?? 1;
    const totalPages = Math.max(1, preview?.totalPages ?? 1);
    const layoutMode = preview?.layoutMode ?? 'slide';
    const thumbs = Boolean(preview?.thumbnailsVisible);
    const fitMode = preview?.fitMode ?? true;
    const zoom = preview?.zoom ?? 1.0;

    pptSlideBtn.classList.toggle('is-active', layoutMode === 'slide');
    pptListBtn.classList.toggle('is-active', layoutMode === 'list');
    pptThumbBtn.classList.toggle('is-active', thumbs);

    if (document.activeElement !== pageInput) {
      pageInput.value = String(page);
    }
    pageInput.max = String(totalPages);
    pageTotalEl.textContent = `/ ${totalPages}`;
    prevPageBtn.disabled = page <= 1;
    nextPageBtn.disabled = page >= totalPages;

    pptFitBtn.classList.toggle('is-active', fitMode);
    pptZoomText.textContent = `${Math.round(zoom * 100)}%`;
    pptZoomOutBtn.disabled = zoom <= 0.5;
    pptZoomInBtn.disabled = zoom >= 2.0;

    const totalMatches = preview?.searchMatches ?? 0;
    const activeMatch = preview?.activeMatch ?? 0;
    pptSearch.updateMatches(totalMatches, activeMatch);
  }

  function updateWordControls(preview: PreviewStateInfo | null): void {
    const fitMode = preview?.fitMode ?? true;
    const zoom = preview?.zoom ?? 1.0;
    wordFitBtn.classList.toggle('is-active', fitMode);
    wordZoomText.textContent = `${Math.round(zoom * 100)}%`;
    wordZoomOutBtn.disabled = zoom <= 0.5;
    wordZoomInBtn.disabled = zoom >= 2.0;

    const totalMatches = preview?.searchMatches ?? 0;
    const activeMatch = preview?.activeMatch ?? 0;
    wordSearch.updateMatches(totalMatches, activeMatch);
  }

  function updateHtmlControls(preview: PreviewStateInfo | null, file: ToolbarFileMeta): void {
    const mode: HtmlPreviewMode = preview?.htmlMode ?? 'sandbox';
    htmlSandboxBtn.classList.toggle('is-active', mode === 'sandbox');
    htmlInteractiveBtn.classList.toggle('is-active', mode === 'interactive');
    htmlSourceBtn.classList.toggle('is-active', mode === 'source');

    htmlStatusBanner.classList.remove('is-warn', 'is-safe', 'is-interactive');
    if (preview?.warning) {
      htmlStatusBanner.classList.add('is-warn');
      htmlStatusBanner.textContent = preview.warning;
    } else if (mode === 'sandbox') {
      htmlStatusBanner.classList.add('is-safe');
      htmlStatusBanner.textContent = tr('安全沙箱：已禁止脚本执行', 'Sandbox: scripts disabled');
    } else if (mode === 'interactive') {
      htmlStatusBanner.classList.add('is-interactive');
      htmlStatusBanner.textContent = file.isLocalDrop
        ? tr('仅工作区路径文件支持完整交互脚本', 'Full scripts require a workspace file path')
        : tr('交互模式：跨源隔离执行页面脚本', 'Interactive: scripts run in isolated origin');
    } else {
      htmlStatusBanner.textContent = tr('源码视图：只读文本展示', 'Source View: read-only HTML');
    }
  }

  function applyStaticLocaleText(): void {
    drawerBtn.title = tr('打开或切换工作区文件列表', 'Toggle workspace file list');
    drawerBtnText.textContent = tr('文件', 'Files');
    largePdfBadge.textContent = tr('>20MB 流式预览', '>20MB Streaming');
    dropTag.textContent = tr('本地拖入', 'Local Drop');
    revealBtn.title = tr('在文件资源管理器中查看当前文件', 'Reveal current file in Explorer');
    revealLabel.textContent = tr('打开位置', 'Reveal');
    refreshBtn.title = tr('立即重新加载当前文件', 'Reload current file');
    refreshLabel.textContent = tr('刷新', 'Refresh');
    autoText.textContent = tr('自动刷新', 'Auto Refresh');

    pptSlideBtn.textContent = tr('单页', 'Slide');
    pptListBtn.textContent = tr('滚动', 'Scroll');
    pptThumbBtn.textContent = tr('缩略图', 'Thumbs');
    pptFitBtn.textContent = tr('适应', 'Fit');
    pptZoomOutBtn.title = tr('缩小', 'Zoom out');
    pptZoomInBtn.title = tr('放大', 'Zoom in');
    pptSearch.refreshLocale();

    wordFitBtn.textContent = tr('自适应宽度', 'Fit Width');
    wordZoomOutBtn.title = tr('缩小', 'Zoom out');
    wordZoomInBtn.title = tr('放大', 'Zoom in');
    wordSearch.setPlaceholder(
      tr('搜索文档内容 (回车)', 'Search document (Enter)'),
      tr('搜索 Word 文档内容', 'Search Word document'),
    );
    wordSearch.refreshLocale();

    htmlSandboxBtn.textContent = tr('沙箱预览', 'Sandbox');
    htmlInteractiveBtn.textContent = tr('交互模式', 'Interactive');
    htmlSourceBtn.textContent = tr('源码', 'Source');
  }

  applyStaticLocaleText();
  const stopLocaleSync = onLocaleChange(() => {
    applyStaticLocaleText();
    if (latestState.visible) {
      update(latestState);
    }
  });

  function update(state: ToolbarRenderState): void {
    latestState = state;
    root.hidden = !state.visible;
    if (!state.visible) {
      return;
    }

    applyStaticLocaleText();
    drawerBtn.classList.toggle('is-active', state.drawerOpen);

    const file = state.file;
    if (file) {
      fileBadge.hidden = false;
      fileBadge.dataset.kind = file.kind;
      fileBadge.textContent = formatKindBadge(file.kind, file.ext);
      fileNameEl.textContent = file.name;
      fileNameEl.title = file.isLocalDrop
        ? tr(`${file.name}（本地拖拽文件）`, `${file.name} (Local dropped file)`)
        : file.path;
      largePdfBadge.hidden = !(
        (file.kind === 'pdf' || file.kind === 'mp4') &&
        (file.size ?? 0) > LARGE_PDF_THRESHOLD_BYTES
      );
      dropTag.hidden = !file.isLocalDrop;
    } else {
      fileBadge.hidden = true;
      delete fileBadge.dataset.kind;
      fileBadge.textContent = '';
      fileNameEl.textContent = tr('未选择文件', 'No file selected');
      fileNameEl.title = '';
      largePdfBadge.hidden = true;
      dropTag.hidden = true;
    }

    const isDiskFile = Boolean(file && !file.isLocalDrop);
    revealBtn.disabled = !isDiskFile;
    refreshBtn.disabled = !isDiskFile || state.loading;
    autoRefreshBtn.disabled = !isDiskFile;
    autoRefreshBtn.classList.toggle('is-active', isDiskFile && state.autoRefresh);
    autoRefreshBtn.title = isDiskFile
      ? tr('每 2.5 秒检查一次磁盘修改时间并自动刷新', 'Check file changes every 2.5s and auto-reload')
      : tr('拖拽的本地内存文件无磁盘路径，暂不支持自动监听刷新', 'Dropped local files have no disk path for auto-refresh');

    const localePref = state.localePreference ?? getLocalePreference();
    const langDisplay = getLocaleToggleDisplay(localePref, getCurrentLocale());
    langBtn.textContent = langDisplay.label;
    langBtn.title = langDisplay.title;

    const themePref = state.themePreference ?? 'auto';
    themeBtn.textContent =
      themePref === 'light'
        ? tr('浅色', 'Light')
        : themePref === 'dark'
          ? tr('深色', 'Dark')
          : tr('自动', 'Auto');
    themeBtn.title =
      themePref === 'light'
        ? tr('当前：浅色模式（点击切换为深色）', 'Theme: Light (click for Dark)')
        : themePref === 'dark'
          ? tr('当前：深色模式（点击切换为跟随宿主）', 'Theme: Dark (click for Auto)')
          : tr('当前：跟随宿主主题（点击切换为浅色）', 'Theme: Auto (click for Light)');

    const showFormatRow = Boolean(file && !state.hasError && file.kind !== 'mp4');
    secondaryRow.hidden = !showFormatRow;
    if (!showFormatRow || !file) {
      pptGroup.hidden = true;
      wordGroup.hidden = true;
      htmlGroup.hidden = true;
      return;
    }

    const isPaged = file.kind === 'ppt' || file.kind === 'pdf' || file.kind === 'excel';
    pptGroup.hidden = !isPaged;
    wordGroup.hidden = file.kind !== 'word';
    htmlGroup.hidden = file.kind !== 'html';

    if (isPaged) {
      updatePagedControls(state.preview, file.kind);
    } else if (file.kind === 'word') {
      updateWordControls(state.preview);
    } else {
      updateHtmlControls(state.preview, file);
    }
  }

  function handleGlobalFindShortcut(e: KeyboardEvent): void {
    if (!latestState.visible || latestState.hasError || !latestState.file) return;
    if (e.defaultPrevented || e.isComposing) return;
    if (!(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey) return;
    if (e.key.toLowerCase() !== 'f') return;

    const kind = latestState.file.kind;
    let targetInput: HTMLInputElement | null = null;
    if (kind === 'ppt' || kind === 'pdf' || kind === 'excel') {
      targetInput = pptSearch.inputEl;
    } else if (kind === 'word') {
      targetInput = wordSearch.inputEl;
    }
    if (!targetInput) return;

    e.preventDefault();
    targetInput.focus();
    targetInput.select();
  }

  window.addEventListener('keydown', handleGlobalFindShortcut);

  function resetSearch(): void {
    pptSearch.reset();
    wordSearch.reset();
  }

  function destroy(): void {
    stopLocaleSync();
    window.removeEventListener('keydown', handleGlobalFindShortcut);
    root.remove();
  }

  return {
    update,
    resetSearch,
    destroy,
  };
}
