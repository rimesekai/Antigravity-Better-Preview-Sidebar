import type { PreviewKind, WorkspaceFileItem } from '../types/preview';
import {
  cycleLocalePreference,
  getCurrentLocale,
  getLocalePreference,
  getLocaleToggleDisplay,
  onLocaleChange,
  tr,
} from '../utils/i18n';

const LARGE_PDF_THRESHOLD_BYTES = 20 * 1024 * 1024;

export type BrowserKindFilter = 'all' | PreviewKind;

export interface WorkspaceBrowserState {
  mode: 'welcome' | 'drawer';
  files: WorkspaceFileItem[];
  loading: boolean;
  error: string | null;
  activePath: string | null;
}

export interface WorkspaceBrowserCallbacks {
  onOpenDiskFile: (path: string) => void;
  onOpenLocalFile: (file: File) => void;
  onRefreshList: () => void;
  onCloseDrawer?: () => void;
}

export interface WorkspaceBrowserController {
  update: (state: WorkspaceBrowserState) => void;
  focusPathInput: () => void;
  destroy: () => void;
}

function formatBytes(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes < 0) return '0 B';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function formatModifiedTime(mtimeMs: number): string {
  if (!Number.isFinite(mtimeMs) || mtimeMs <= 0) return '';
  const diff = Date.now() - mtimeMs;
  if (diff >= 0 && diff < 60_000) return tr('刚刚', 'Just now');
  if (diff >= 0 && diff < 3600_000) {
    const mins = Math.floor(diff / 60_000);
    return tr(`${mins} 分钟前`, `${mins}m ago`);
  }
  if (diff >= 0 && diff < 86400_000) {
    const hours = Math.floor(diff / 3600_000);
    return tr(`${hours} 小时前`, `${hours}h ago`);
  }

  const date = new Date(mtimeMs);
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hour = String(date.getHours()).padStart(2, '0');
  const minute = String(date.getMinutes()).padStart(2, '0');
  return `${month}-${day} ${hour}:${minute}`;
}

function getKindShortLabel(kind: PreviewKind, ext: string): string {
  if (kind === 'ppt') return 'PPTX';
  if (kind === 'word') return 'DOCX';
  if (kind === 'excel') {
    const normalized = ext.toLowerCase();
    if (normalized === '.csv') return 'CSV';
    if (normalized === '.tsv') return 'TSV';
    if (normalized === '.xls') return 'XLS';
    return 'XLSX';
  }
  if (kind === 'pdf') return 'PDF';
  if (kind === 'mp4') return 'MP4';
  return ext.toLowerCase() === '.htm' ? 'HTM' : 'HTML';
}

/**
 * 创建工作区文档浏览器与欢迎页组件（支持全页模式与抽屉模式）。
 */
export function createWorkspaceBrowser(
  container: HTMLElement,
  callbacks: WorkspaceBrowserCallbacks,
): WorkspaceBrowserController {
  let currentFilter: BrowserKindFilter = 'all';
  let currentKeyword = '';
  let latestState: WorkspaceBrowserState = {
    mode: 'welcome',
    files: [],
    loading: false,
    error: null,
    activePath: null,
  };

  const root = document.createElement('div');
  root.className = 'fp-browser';

  // 欢迎区（仅在 welcome 模式显示）
  const hero = document.createElement('section');
  hero.className = 'fp-hero';

  const heroBrand = document.createElement('div');
  heroBrand.className = 'fp-hero-brand';

  const logoImg = document.createElement('img');
  logoImg.src = '/assets/logo.svg';
  logoImg.alt = 'File Preview';
  logoImg.className = 'fp-hero-logo';
  logoImg.addEventListener('error', () => {
    logoImg.hidden = true;
  });

  const heroTitleBox = document.createElement('div');
  const heroTitle = document.createElement('h1');
  heroTitle.className = 'fp-hero-title';
  heroTitle.textContent = 'File Preview';
  const heroSub = document.createElement('p');
  heroSub.className = 'fp-hero-subtitle';
  heroTitleBox.append(heroTitle, heroSub);

  heroBrand.append(logoImg, heroTitleBox);

  const heroLangBtn = document.createElement('button');
  heroLangBtn.type = 'button';
  heroLangBtn.className = 'fp-btn fp-btn-ghost fp-hero-lang-btn';
  heroLangBtn.addEventListener('click', () => {
    cycleLocalePreference();
  });

  hero.append(heroBrand, heroLangBtn);

  // 抽屉模式标题栏（仅在 drawer 模式显示）
  const drawerHeader = document.createElement('div');
  drawerHeader.className = 'fp-drawer-header';
  const drawerTitle = document.createElement('h2');
  drawerTitle.className = 'fp-drawer-title';
  const drawerCloseBtn = document.createElement('button');
  drawerCloseBtn.type = 'button';
  drawerCloseBtn.className = 'fp-btn fp-btn-ghost fp-btn-icon';
  drawerCloseBtn.textContent = '✕';
  drawerCloseBtn.addEventListener('click', () => callbacks.onCloseDrawer?.());
  drawerHeader.append(drawerTitle, drawerCloseBtn);

  // 快捷打开区：路径输入 + 本地文件选择 + 拖拽提示
  const quickOpenCard = document.createElement('section');
  quickOpenCard.className = 'fp-quick-open';

  const pathForm = document.createElement('form');
  pathForm.className = 'fp-path-form';

  const pathInput = document.createElement('input');
  pathInput.type = 'text';
  pathInput.className = 'fp-input fp-path-input';

  const pathSubmitBtn = document.createElement('button');
  pathSubmitBtn.type = 'submit';
  pathSubmitBtn.className = 'fp-btn fp-btn-primary';

  pathForm.append(pathInput, pathSubmitBtn);
  pathForm.addEventListener('submit', (e) => {
    e.preventDefault();
    const target = pathInput.value.trim();
    if (!target) {
      pathInput.focus();
      return;
    }
    callbacks.onOpenDiskFile(target);
  });

  const localActionsRow = document.createElement('div');
  localActionsRow.className = 'fp-local-actions';

  const hiddenFileInput = document.createElement('input');
  hiddenFileInput.type = 'file';
  hiddenFileInput.accept = '.pptx,.docx,.xlsx,.xls,.csv,.tsv,.pdf,.mp4,.html,.htm';
  hiddenFileInput.hidden = true;
  hiddenFileInput.addEventListener('change', () => {
    const picked = hiddenFileInput.files?.[0];
    hiddenFileInput.value = '';
    if (picked) {
      callbacks.onOpenLocalFile(picked);
    }
  });

  const pickFileBtn = document.createElement('button');
  pickFileBtn.type = 'button';
  pickFileBtn.className = 'fp-btn fp-btn-secondary';
  pickFileBtn.addEventListener('click', () => hiddenFileInput.click());

  const dragTip = document.createElement('span');
  dragTip.className = 'fp-drag-tip';

  localActionsRow.append(hiddenFileInput, pickFileBtn, dragTip);
  quickOpenCard.append(pathForm, localActionsRow);

  // 工作区列表区：过滤标签 + 搜索 + 刷新
  const listSection = document.createElement('section');
  listSection.className = 'fp-list-section';

  const listHeader = document.createElement('div');
  listHeader.className = 'fp-list-controls';

  const filterTabs = document.createElement('div');
  filterTabs.className = 'fp-filter-tabs';

  const getFilterLabel = (key: BrowserKindFilter): string => {
    switch (key) {
      case 'all':
        return tr('全部', 'All');
      case 'ppt':
        return 'PPT';
      case 'word':
        return 'Word';
      case 'excel':
        return 'Excel';
      case 'pdf':
        return 'PDF';
      case 'mp4':
        return 'MP4';
      case 'html':
        return 'HTML';
    }
  };

  const filterKeys: BrowserKindFilter[] = ['all', 'ppt', 'word', 'excel', 'pdf', 'mp4', 'html'];

  const filterButtons = new Map<BrowserKindFilter, HTMLButtonElement>();
  for (const key of filterKeys) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'fp-filter-tab';
    btn.textContent = getFilterLabel(key);
    btn.addEventListener('click', () => {
      currentFilter = key;
      renderFileList();
    });
    filterButtons.set(key, btn);
    filterTabs.appendChild(btn);
  }

  const searchAndRefresh = document.createElement('div');
  searchAndRefresh.className = 'fp-list-search-row';

  const filterSearchInput = document.createElement('input');
  filterSearchInput.type = 'search';
  filterSearchInput.className = 'fp-input fp-list-search';
  filterSearchInput.addEventListener('input', () => {
    currentKeyword = filterSearchInput.value.trim().toLowerCase();
    renderFileList();
  });

  const refreshListBtn = document.createElement('button');
  refreshListBtn.type = 'button';
  refreshListBtn.className = 'fp-btn fp-btn-ghost';
  refreshListBtn.addEventListener('click', () => callbacks.onRefreshList());

  searchAndRefresh.append(filterSearchInput, refreshListBtn);
  listHeader.append(filterTabs, searchAndRefresh);

  const listBody = document.createElement('div');
  listBody.className = 'fp-file-list';

  listSection.append(listHeader, listBody);
  root.append(hero, drawerHeader, quickOpenCard, listSection);
  container.appendChild(root);

  function renderFilterCounts(): void {
    const counts: Record<BrowserKindFilter, number> = {
      all: latestState.files.length,
      ppt: 0,
      word: 0,
      excel: 0,
      pdf: 0,
      mp4: 0,
      html: 0,
    };
    for (const file of latestState.files) {
      counts[file.kind] += 1;
    }
    for (const key of filterKeys) {
      const btn = filterButtons.get(key);
      if (!btn) continue;
      btn.textContent = `${getFilterLabel(key)} (${counts[key]})`;
      btn.classList.toggle('is-active', currentFilter === key);
    }
  }

  function renderFileList(): void {
    renderFilterCounts();
    listBody.innerHTML = '';

    if (latestState.loading && latestState.files.length === 0) {
      const loadingEl = document.createElement('div');
      loadingEl.className = 'fp-list-empty';
      loadingEl.textContent = tr('正在扫描工作区可预览文档...', 'Scanning workspace for previewable files...');
      listBody.appendChild(loadingEl);
      return;
    }

    if (latestState.error) {
      const errEl = document.createElement('div');
      errEl.className = 'fp-list-error';
      errEl.textContent = latestState.error;
      listBody.appendChild(errEl);
      return;
    }

    const filtered = latestState.files.filter((file) => {
      if (currentFilter !== 'all' && file.kind !== currentFilter) {
        return false;
      }
      if (currentKeyword) {
        const hay = `${file.name} ${file.relPath}`.toLowerCase();
        if (!hay.includes(currentKeyword)) {
          return false;
        }
      }
      return true;
    });

    if (filtered.length === 0) {
      const emptyEl = document.createElement('div');
      emptyEl.className = 'fp-list-empty';
      if (latestState.files.length === 0) {
        emptyEl.textContent = tr(
          '当前工作区暂未扫描到 .pptx / .docx / .xlsx / .xls / .csv / .tsv / .pdf / .mp4 / .html 文件，可使用上方路径输入或直接拖拽文件打开。',
          'No .pptx / .docx / .xlsx / .xls / .csv / .tsv / .pdf / .mp4 / .html files found in the workspace. Paste a path above or drag & drop a file.',
        );
      } else {
        emptyEl.textContent = tr('没有符合当前筛选条件的文件。', 'No files match the current filter.');
      }
      listBody.appendChild(emptyEl);
      return;
    }

    const fragment = document.createDocumentFragment();
    for (const file of filtered) {
      const itemBtn = document.createElement('button');
      itemBtn.type = 'button';
      itemBtn.className = 'fp-file-item';
      if (latestState.activePath && latestState.activePath === file.path) {
        itemBtn.classList.add('is-current');
      }
      itemBtn.title = file.path;
      itemBtn.addEventListener('click', () => callbacks.onOpenDiskFile(file.path));

      const badge = document.createElement('span');
      badge.className = 'fp-kind-badge';
      badge.dataset.kind = file.kind;
      badge.textContent = getKindShortLabel(file.kind, file.ext);

      const info = document.createElement('div');
      info.className = 'fp-file-info';

      const nameRow = document.createElement('div');
      nameRow.className = 'fp-file-name-row';

      const nameEl = document.createElement('div');
      nameEl.className = 'fp-file-name';
      nameEl.textContent = file.name;
      nameRow.appendChild(nameEl);

      if ((file.kind === 'pdf' || file.kind === 'mp4') && file.size > LARGE_PDF_THRESHOLD_BYTES) {
        const largeBadge = document.createElement('span');
        largeBadge.className = 'fp-large-pdf-badge';
        largeBadge.textContent = tr('>20MB 流式可预览', '>20MB Streaming');
        nameRow.appendChild(largeBadge);
      }

      const metaEl = document.createElement('div');
      metaEl.className = 'fp-file-meta';
      const timeText = formatModifiedTime(file.mtimeMs);
      metaEl.textContent = timeText
        ? `${file.relPath} · ${formatBytes(file.size)} · ${timeText}`
        : `${file.relPath} · ${formatBytes(file.size)}`;

      info.append(nameRow, metaEl);
      itemBtn.append(badge, info);
      fragment.appendChild(itemBtn);
    }
    listBody.appendChild(fragment);
  }

  function applyLocaleText(): void {
    const langDisplay = getLocaleToggleDisplay(getLocalePreference(), getCurrentLocale());
    heroLangBtn.textContent = langDisplay.label;
    heroLangBtn.title = langDisplay.title;

    heroSub.textContent = tr(
      '侧边栏本地文档高保真预览 · 支持 .pptx / .docx / .xlsx / .xls / .csv / .tsv / .pdf / .mp4 / .html',
      'High-fidelity sidebar preview · .pptx / .docx / .xlsx / .xls / .csv / .tsv / .pdf / .mp4 / .html',
    );
    drawerTitle.textContent = tr('切换工作区文档', 'Workspace Files');
    drawerCloseBtn.title = tr('关闭文件列表', 'Close file list');
    pathInput.placeholder = tr(
      '输入本地文件路径或 file:// URI，回车打开...',
      'Enter local file path or file:// URI and press Enter...',
    );
    pathInput.setAttribute('aria-label', tr('输入本地文件路径', 'Enter local file path'));
    pathSubmitBtn.textContent = tr('打开', 'Open');
    pickFileBtn.textContent = tr('选择电脑本地文件', 'Choose Local File');
    dragTip.textContent = tr(
      '或直接拖拽 .pptx / .docx / .xlsx / .xls / .csv / .tsv / .pdf / .mp4 / .html 到此窗口',
      'or drag & drop .pptx / .docx / .xlsx / .xls / .csv / .tsv / .pdf / .mp4 / .html files here',
    );
    filterSearchInput.placeholder = tr('按文件名或路径筛选...', 'Filter by file name or path...');
    filterSearchInput.setAttribute('aria-label', tr('筛选工作区文件', 'Filter workspace files'));
    refreshListBtn.textContent = latestState.loading
      ? tr('扫描中...', 'Scanning...')
      : tr('刷新列表', 'Refresh');
  }

  applyLocaleText();
  const stopLocaleSync = onLocaleChange(() => {
    applyLocaleText();
    renderFileList();
  });

  function update(state: WorkspaceBrowserState): void {
    latestState = state;
    root.dataset.mode = state.mode;
    hero.hidden = state.mode !== 'welcome';
    drawerHeader.hidden = state.mode !== 'drawer';
    refreshListBtn.disabled = state.loading;
    applyLocaleText();
    renderFileList();
  }

  function focusPathInput(): void {
    pathInput.focus();
  }

  function destroy(): void {
    stopLocaleSync();
    root.remove();
  }

  return {
    update,
    focusPathInput,
    destroy,
  };
}
