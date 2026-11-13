import type { PreviewKind, WorkspaceFileItem } from '../types/preview';

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
  if (diff >= 0 && diff < 60_000) return '刚刚';
  if (diff >= 0 && diff < 3600_000) return `${Math.floor(diff / 60_000)} 分钟前`;
  if (diff >= 0 && diff < 86400_000) return `${Math.floor(diff / 3600_000)} 小时前`;

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
  heroSub.textContent = '侧边栏本地文档高保真预览 · 支持 .pptx / .docx / .html';
  heroTitleBox.append(heroTitle, heroSub);

  heroBrand.append(logoImg, heroTitleBox);
  hero.appendChild(heroBrand);

  // 抽屉模式标题栏（仅在 drawer 模式显示）
  const drawerHeader = document.createElement('div');
  drawerHeader.className = 'fp-drawer-header';
  const drawerTitle = document.createElement('h2');
  drawerTitle.className = 'fp-drawer-title';
  drawerTitle.textContent = '切换工作区文档';
  const drawerCloseBtn = document.createElement('button');
  drawerCloseBtn.type = 'button';
  drawerCloseBtn.className = 'fp-btn fp-btn-ghost fp-btn-icon';
  drawerCloseBtn.title = '关闭文件列表';
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
  pathInput.placeholder = '输入本地文件路径或 file:// URI，回车打开...';
  pathInput.setAttribute('aria-label', '输入本地文件路径');

  const pathSubmitBtn = document.createElement('button');
  pathSubmitBtn.type = 'submit';
  pathSubmitBtn.className = 'fp-btn fp-btn-primary';
  pathSubmitBtn.textContent = '打开';

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
  hiddenFileInput.accept = '.pptx,.docx,.html,.htm';
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
  pickFileBtn.textContent = '选择电脑本地文件';
  pickFileBtn.addEventListener('click', () => hiddenFileInput.click());

  const dragTip = document.createElement('span');
  dragTip.className = 'fp-drag-tip';
  dragTip.textContent = '或直接拖拽 .pptx / .docx / .html 到此窗口';

  localActionsRow.append(hiddenFileInput, pickFileBtn, dragTip);
  quickOpenCard.append(pathForm, localActionsRow);

  // 工作区列表区：过滤标签 + 搜索 + 刷新
  const listSection = document.createElement('section');
  listSection.className = 'fp-list-section';

  const listHeader = document.createElement('div');
  listHeader.className = 'fp-list-controls';

  const filterTabs = document.createElement('div');
  filterTabs.className = 'fp-filter-tabs';

  const filterMeta: Array<{ key: BrowserKindFilter; label: string }> = [
    { key: 'all', label: '全部' },
    { key: 'ppt', label: 'PPT' },
    { key: 'word', label: 'Word' },
    { key: 'html', label: 'HTML' },
  ];

  const filterButtons = new Map<BrowserKindFilter, HTMLButtonElement>();
  for (const item of filterMeta) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'fp-filter-tab';
    btn.textContent = item.label;
    btn.addEventListener('click', () => {
      currentFilter = item.key;
      renderFileList();
    });
    filterButtons.set(item.key, btn);
    filterTabs.appendChild(btn);
  }

  const searchAndRefresh = document.createElement('div');
  searchAndRefresh.className = 'fp-list-search-row';

  const filterSearchInput = document.createElement('input');
  filterSearchInput.type = 'search';
  filterSearchInput.className = 'fp-input fp-list-search';
  filterSearchInput.placeholder = '按文件名或路径筛选...';
  filterSearchInput.setAttribute('aria-label', '筛选工作区文件');
  filterSearchInput.addEventListener('input', () => {
    currentKeyword = filterSearchInput.value.trim().toLowerCase();
    renderFileList();
  });

  const refreshListBtn = document.createElement('button');
  refreshListBtn.type = 'button';
  refreshListBtn.className = 'fp-btn fp-btn-ghost';
  refreshListBtn.textContent = '刷新列表';
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
      html: 0,
    };
    for (const file of latestState.files) {
      counts[file.kind] += 1;
    }
    for (const item of filterMeta) {
      const btn = filterButtons.get(item.key);
      if (!btn) continue;
      btn.textContent = `${item.label} (${counts[item.key]})`;
      btn.classList.toggle('is-active', currentFilter === item.key);
    }
  }

  function renderFileList(): void {
    renderFilterCounts();
    listBody.innerHTML = '';

    if (latestState.loading && latestState.files.length === 0) {
      const loadingEl = document.createElement('div');
      loadingEl.className = 'fp-list-empty';
      loadingEl.textContent = '正在扫描工作区可预览文档...';
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
        emptyEl.textContent = '当前工作区暂未扫描到 .pptx / .docx / .html 文档，可使用上方路径输入或直接拖拽文件打开。';
      } else {
        emptyEl.textContent = '没有符合当前筛选条件的文件。';
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

      const nameEl = document.createElement('div');
      nameEl.className = 'fp-file-name';
      nameEl.textContent = file.name;

      const metaEl = document.createElement('div');
      metaEl.className = 'fp-file-meta';
      const timeText = formatModifiedTime(file.mtimeMs);
      metaEl.textContent = timeText
        ? `${file.relPath} · ${formatBytes(file.size)} · ${timeText}`
        : `${file.relPath} · ${formatBytes(file.size)}`;

      info.append(nameEl, metaEl);
      itemBtn.append(badge, info);
      fragment.appendChild(itemBtn);
    }
    listBody.appendChild(fragment);
  }

  function update(state: WorkspaceBrowserState): void {
    latestState = state;
    root.dataset.mode = state.mode;
    hero.hidden = state.mode !== 'welcome';
    drawerHeader.hidden = state.mode !== 'drawer';
    refreshListBtn.disabled = state.loading;
    refreshListBtn.textContent = state.loading ? '扫描中...' : '刷新列表';
    renderFileList();
  }

  function focusPathInput(): void {
    pathInput.focus();
  }

  function destroy(): void {
    root.remove();
  }

  return {
    update,
    focusPathInput,
    destroy,
  };
}
