import './styles.css';
import './renderers/renderers.css';
import type {
  FileStatInfo,
  HtmlPreviewMode,
  PptLayoutMode,
  PreviewKind,
  PreviewRenderer,
  PreviewStateInfo,
  WorkspaceFileItem,
} from './types/preview';
import { createWordRenderer } from './renderers/wordRenderer';
import { createPptRenderer } from './renderers/pptRenderer';
import { createHtmlRenderer } from './renderers/htmlRenderer';
import { createPdfRenderer } from './renderers/pdfRenderer';
import { createVideoRenderer } from './renderers/videoRenderer';
import {
  SidecarApiError,
  listWorkspaceFiles,
  mountHtmlFile,
  mountPdfFile,
  mountVideoFile,
  readFileBuffer,
  readFileText,
  statFile,
} from './services/sidecarClient';
import {
  cycleThemePreference,
  getThemePreference,
  startThemeSync,
  subscribeWorkspaceChanges,
} from './services/themeSync';
import { createToolbar, type ToolbarFileMeta } from './components/Toolbar';
import { createWorkspaceBrowser } from './components/WorkspaceBrowser';

type ActiveSource =
  | { type: 'disk'; path: string; stat: FileStatInfo }
  | { type: 'local'; name: string; ext: string; kind: PreviewKind; size: number };

const AUTO_REFRESH_INTERVAL_MS = 2500;

function classifyDroppedFile(fileName: string): {
  kind: PreviewKind | null;
  ext: string;
  errorTitle?: string;
  errorMessage?: string;
} {
  const dotIdx = fileName.lastIndexOf('.');
  const ext = dotIdx >= 0 ? fileName.slice(dotIdx).toLowerCase() : '';
  if (ext === '.pptx') return { kind: 'ppt', ext };
  if (ext === '.docx') return { kind: 'word', ext };
  if (ext === '.pdf') return { kind: 'pdf', ext };
  if (ext === '.mp4') return { kind: 'mp4', ext };
  if (ext === '.html' || ext === '.htm') return { kind: 'html', ext };

  if (ext === '.ppt') {
    return {
      kind: null,
      ext,
      errorTitle: '暂不支持旧版 .ppt 格式',
      errorMessage: '请先在 PowerPoint 或 WPS 中将文件另存为 .pptx 格式后再拖入预览。',
    };
  }
  if (ext === '.doc') {
    return {
      kind: null,
      ext,
      errorTitle: '暂不支持旧版 .doc 格式',
      errorMessage: '请先在 Word 或 WPS 中将文件另存为 .docx 格式后再拖入预览。',
    };
  }
  return {
    kind: null,
    ext,
    errorTitle: '不支持的文件格式',
    errorMessage: `当前仅支持 .pptx、.docx、.pdf、.mp4、.html 和 .htm 文件（收到：${ext || '无扩展名'}）。`,
  };
}

function createRendererForKind(kind: PreviewKind): PreviewRenderer {
  switch (kind) {
    case 'word':
      return createWordRenderer();
    case 'ppt':
      return createPptRenderer();
    case 'pdf':
      return createPdfRenderer();
    case 'mp4':
      return createVideoRenderer();
    case 'html':
      return createHtmlRenderer();
  }
}

function bootstrap(): void {
  const appRoot = document.getElementById('app');
  if (!appRoot) {
    return;
  }

  appRoot.innerHTML = '';

  // 主壳结构
  const shell = document.createElement('div');
  shell.className = 'fp-shell';

  const toolbarHost = document.createElement('header');
  toolbarHost.className = 'fp-toolbar-host';

  const mainStage = document.createElement('main');
  mainStage.className = 'fp-main';

  const welcomeHost = document.createElement('div');
  welcomeHost.className = 'fp-welcome-host';

  const previewHost = document.createElement('div');
  previewHost.className = 'fp-preview-host';
  previewHost.hidden = true;

  const rendererContainer = document.createElement('div');
  rendererContainer.className = 'fp-renderer-container';

  const loadingOverlay = document.createElement('div');
  loadingOverlay.className = 'fp-loading-overlay';
  loadingOverlay.hidden = true;
  loadingOverlay.innerHTML =
    '<div class="fp-spinner" aria-hidden="true"></div><div class="fp-loading-text">正在加载并渲染文档...</div>';
  const loadingTextEl = loadingOverlay.querySelector('.fp-loading-text') as HTMLElement;

  const errorCardHost = document.createElement('div');
  errorCardHost.className = 'fp-error-host';
  errorCardHost.hidden = true;

  previewHost.append(rendererContainer, loadingOverlay, errorCardHost);

  // 文件列表抽屉
  const drawerBackdrop = document.createElement('div');
  drawerBackdrop.className = 'fp-drawer-backdrop';
  drawerBackdrop.hidden = true;

  const drawerPanel = document.createElement('aside');
  drawerPanel.className = 'fp-drawer-panel';
  drawerPanel.setAttribute('aria-label', '工作区文档列表');
  drawerPanel.hidden = true;

  // 全屏拖拽遮罩
  const dropOverlay = document.createElement('div');
  dropOverlay.className = 'fp-drop-overlay';
  dropOverlay.hidden = true;
  dropOverlay.innerHTML =
    '<div class="fp-drop-box"><div class="fp-drop-title">松开鼠标即可预览</div><div class="fp-drop-sub">支持 .pptx / .docx / .pdf / .html / .htm 本地文件</div></div>';

  // 错误页内使用的隐藏文件选择器
  const fallbackFileInput = document.createElement('input');
  fallbackFileInput.type = 'file';
  fallbackFileInput.accept = '.pptx,.docx,.pdf,.html,.htm';
  fallbackFileInput.hidden = true;

  mainStage.append(welcomeHost, previewHost, drawerBackdrop, drawerPanel);
  shell.append(toolbarHost, mainStage, dropOverlay, fallbackFileInput);
  appRoot.appendChild(shell);

  // 应用状态
  let viewMode: 'welcome' | 'preview' = 'welcome';
  let drawerOpen = false;
  let workspaceFiles: WorkspaceFileItem[] = [];
  let workspaceLoading = false;
  let workspaceError: string | null = null;

  let currentSource: ActiveSource | null = null;
  let activeRenderer: PreviewRenderer | null = null;
  let previewState: PreviewStateInfo | null = null;
  let isLoading = false;
  let activeError: { title: string; message: string; code?: string; retryPath?: string } | null = null;
  let autoRefreshEnabled = true;
  let isPollingStat = false;
  let loadRequestSeq = 0;

  // 工具栏实例
  const toolbar = createToolbar(toolbarHost, {
    onToggleDrawer: () => {
      drawerOpen = !drawerOpen;
      if (drawerOpen && workspaceFiles.length === 0 && !workspaceLoading) {
        void refreshWorkspaceList();
      }
      syncUi();
    },
    onRefresh: () => {
      if (currentSource?.type === 'disk') {
        void openDiskFile(currentSource.path, { isRefresh: true });
      }
    },
    onToggleAutoRefresh: () => {
      if (currentSource?.type !== 'disk') {
        return;
      }
      autoRefreshEnabled = !autoRefreshEnabled;
      syncUi();
    },
    onCycleTheme: () => {
      cycleThemePreference();
      syncUi();
    },
    onZoomChange: (zoom) => {
      activeRenderer?.setZoom?.(zoom);
    },
    onPrevPage: () => {
      activeRenderer?.prevPage?.();
    },
    onNextPage: () => {
      activeRenderer?.nextPage?.();
    },
    onGoToPage: (page) => {
      activeRenderer?.goToPage?.(page);
    },
    onLayoutModeChange: (mode: PptLayoutMode) => {
      activeRenderer?.setLayoutMode?.(mode);
    },
    onToggleThumbnails: () => {
      activeRenderer?.toggleThumbnails?.();
    },
    onSearch: async (query) => {
      if (!activeRenderer?.search) return;
      const result = await activeRenderer.search(query);
      previewState = {
        ...(previewState ?? { kind: currentSource?.type === 'disk' ? currentSource.stat.kind : 'ppt' }),
        searchMatches: result.total,
        activeMatch: result.current,
      };
      syncUi();
    },
    onNextMatch: () => {
      activeRenderer?.nextMatch?.();
    },
    onPrevMatch: () => {
      activeRenderer?.prevMatch?.();
    },
    onHtmlModeChange: (mode: HtmlPreviewMode) => {
      activeRenderer?.setHtmlMode?.(mode);
    },
  });

  const browserCallbacks = {
    onOpenDiskFile: (filePath: string) => {
      drawerOpen = false;
      void openDiskFile(filePath);
    },
    onOpenLocalFile: (file: File) => {
      drawerOpen = false;
      void openLocalDroppedFile(file);
    },
    onRefreshList: () => {
      void refreshWorkspaceList();
    },
    onCloseDrawer: () => {
      drawerOpen = false;
      syncUi();
    },
  };

  const welcomeBrowser = createWorkspaceBrowser(welcomeHost, browserCallbacks);
  const drawerBrowser = createWorkspaceBrowser(drawerPanel, browserCallbacks);

  drawerBackdrop.addEventListener('click', () => {
    drawerOpen = false;
    syncUi();
  });

  window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && drawerOpen) {
      drawerOpen = false;
      syncUi();
    }
  });

  fallbackFileInput.addEventListener('change', () => {
    const picked = fallbackFileInput.files?.[0];
    fallbackFileInput.value = '';
    if (picked) {
      void openLocalDroppedFile(picked);
    }
  });

  function renderErrorCard(): void {
    errorCardHost.hidden = !activeError;
    if (!activeError) {
      errorCardHost.innerHTML = '';
      return;
    }

    errorCardHost.innerHTML = '';
    const card = document.createElement('div');
    card.className = 'fp-error-card';

    const titleEl = document.createElement('h2');
    titleEl.className = 'fp-error-title';
    titleEl.textContent = activeError.title;

    const msgEl = document.createElement('p');
    msgEl.className = 'fp-error-message';
    msgEl.textContent = activeError.message;

    card.append(titleEl, msgEl);

    if (activeError.code) {
      const codeEl = document.createElement('div');
      codeEl.className = 'fp-error-code';
      codeEl.textContent = `错误代码：${activeError.code}`;
      card.appendChild(codeEl);
    }

    const actions = document.createElement('div');
    actions.className = 'fp-error-actions';

    const backBtn = document.createElement('button');
    backBtn.type = 'button';
    backBtn.className = 'fp-btn fp-btn-primary';
    backBtn.textContent = '返回文件列表';
    backBtn.addEventListener('click', () => {
      destroyActiveRenderer();
      currentSource = null;
      activeError = null;
      viewMode = 'welcome';
      drawerOpen = false;
      syncUi();
      void refreshWorkspaceList();
    });
    actions.appendChild(backBtn);

    if (activeError.retryPath) {
      const retryPath = activeError.retryPath;
      const retryBtn = document.createElement('button');
      retryBtn.type = 'button';
      retryBtn.className = 'fp-btn fp-btn-secondary';
      retryBtn.textContent = '重试加载';
      retryBtn.addEventListener('click', () => {
        void openDiskFile(retryPath);
      });
      actions.appendChild(retryBtn);
    }

    const pickBtn = document.createElement('button');
    pickBtn.type = 'button';
    pickBtn.className = 'fp-btn fp-btn-ghost';
    pickBtn.textContent = '选择其他本地文件';
    pickBtn.addEventListener('click', () => fallbackFileInput.click());
    actions.appendChild(pickBtn);

    card.appendChild(actions);
    errorCardHost.appendChild(card);
  }

  function getToolbarFileMeta(): ToolbarFileMeta | null {
    if (!currentSource) return null;
    if (currentSource.type === 'disk') {
      return {
        name: currentSource.stat.name,
        path: currentSource.stat.path,
        ext: currentSource.stat.ext,
        kind: currentSource.stat.kind,
        size: currentSource.stat.size,
        isLocalDrop: false,
      };
    }
    return {
      name: currentSource.name,
      path: currentSource.name,
      ext: currentSource.ext,
      kind: currentSource.kind,
      size: currentSource.size,
      isLocalDrop: true,
    };
  }

  function syncUi(): void {
    welcomeHost.hidden = viewMode !== 'welcome';
    previewHost.hidden = viewMode !== 'preview';
    loadingOverlay.hidden = !isLoading;

    const showDrawer = viewMode === 'preview' && drawerOpen;
    drawerBackdrop.hidden = !showDrawer;
    drawerPanel.hidden = !showDrawer;

    renderErrorCard();

    const activePath = currentSource?.type === 'disk' ? currentSource.path : null;
    if (viewMode === 'welcome') {
      welcomeBrowser.update({
        mode: 'welcome',
        files: workspaceFiles,
        loading: workspaceLoading,
        error: workspaceError,
        activePath,
      });
    } else if (showDrawer) {
      drawerBrowser.update({
        mode: 'drawer',
        files: workspaceFiles,
        loading: workspaceLoading,
        error: workspaceError,
        activePath,
      });
    }

    toolbar.update({
      visible: viewMode === 'preview',
      file: getToolbarFileMeta(),
      preview: previewState,
      loading: isLoading,
      hasError: Boolean(activeError),
      autoRefresh: autoRefreshEnabled,
      drawerOpen: showDrawer,
      themePreference: getThemePreference(),
    });
  }

  function destroyActiveRenderer(): void {
    if (activeRenderer) {
      try {
        activeRenderer.destroy();
      } catch (err) {
        console.error('[file-preview] 销毁渲染器失败:', err);
      }
      activeRenderer = null;
    }
    rendererContainer.innerHTML = '';
    previewState = null;
  }

  function mountRendererWithState(
    kind: PreviewKind,
    preservedState?: PreviewStateInfo | null,
  ): PreviewRenderer {
    destroyActiveRenderer();
    const renderer = createRendererForKind(kind);
    activeRenderer = renderer;
    previewState = preservedState
      ? { ...preservedState, kind }
      : {
          kind,
          ...(kind === 'html' ? { htmlMode: 'sandbox' } : {}),
        };

    renderer.onStateChange?.((nextState) => {
      if (activeRenderer !== renderer) return;
      previewState = nextState;
      syncUi();
    });

    renderer.mount(rendererContainer);
    return renderer;
  }

  function applyPreservedViewOptions(
    renderer: PreviewRenderer,
    kind: PreviewKind,
    preserved: PreviewStateInfo | null,
  ): void {
    if (!preserved || preserved.kind !== kind) return;
    if (kind === 'html' && preserved.htmlMode && preserved.htmlMode !== 'sandbox') {
      renderer.setHtmlMode?.(preserved.htmlMode);
    } else if (kind === 'ppt') {
      if (preserved.layoutMode) {
        renderer.setLayoutMode?.(preserved.layoutMode);
      }
      if (preserved.thumbnailsVisible !== undefined) {
        renderer.toggleThumbnails?.(preserved.thumbnailsVisible);
      }
      if (preserved.fitMode) {
        renderer.setZoom?.('fit');
      } else if (typeof preserved.zoom === 'number') {
        renderer.setZoom?.(preserved.zoom);
      }
      if (typeof preserved.page === 'number' && preserved.page > 1) {
        renderer.goToPage?.(preserved.page);
      }
    } else if (kind === 'pdf') {
      if (preserved.layoutMode) {
        renderer.setLayoutMode?.(preserved.layoutMode);
      }
      if (preserved.thumbnailsVisible !== undefined) {
        renderer.toggleThumbnails?.(preserved.thumbnailsVisible);
      }
      if (preserved.fitMode) {
        renderer.setZoom?.('fit');
      } else if (typeof preserved.zoom === 'number') {
        renderer.setZoom?.(preserved.zoom);
      }
      if (typeof preserved.page === 'number' && preserved.page > 1) {
        renderer.goToPage?.(preserved.page);
      }
    } else if (kind === 'word') {
      if (preserved.fitMode === false && typeof preserved.zoom === 'number') {
        renderer.setZoom?.(preserved.zoom);
      }
    }
  }

  function setErrorState(
    title: string,
    message: string,
    code?: string,
    retryPath?: string,
  ): void {
    destroyActiveRenderer();
    isLoading = false;
    activeError = { title, message, code, retryPath };
    viewMode = 'preview';
    syncUi();
  }

  async function refreshWorkspaceList(): Promise<void> {
    workspaceLoading = true;
    workspaceError = null;
    syncUi();

    try {
      workspaceFiles = await listWorkspaceFiles();
    } catch (err) {
      workspaceError =
        err instanceof SidecarApiError
          ? err.userMessage
          : '扫描工作区文件失败，请稍后重试。';
    } finally {
      workspaceLoading = false;
      syncUi();
    }
  }

  async function openDiskFile(
    rawPath: string,
    options: { isRefresh?: boolean; isAutoRefresh?: boolean } = {},
  ): Promise<void> {
    const reqId = ++loadRequestSeq;
    const isSameDiskFile =
      currentSource?.type === 'disk' &&
      (currentSource.path === rawPath || options.isRefresh || options.isAutoRefresh);
    const preserved = isSameDiskFile && previewState ? { ...previewState } : null;

    if (!options.isAutoRefresh) {
      viewMode = 'preview';
      activeError = null;
      isLoading = true;
      loadingTextEl.textContent = options.isRefresh ? '正在重新加载文档...' : '正在读取并渲染文档...';
      if (!isSameDiskFile) {
        toolbar.resetSearch();
      }
      syncUi();
    }

    try {
      const stat = await statFile(rawPath);
      if (reqId !== loadRequestSeq) return;

      if (stat.kind === 'html') {
        const [mountInfo, sourceText] = await Promise.all([
          mountHtmlFile(stat.path),
          readFileText(stat.path),
        ]);
        if (reqId !== loadRequestSeq) return;

        currentSource = { type: 'disk', path: stat.path, stat };
        const renderer = mountRendererWithState('html', preserved);
        await renderer.loadHtml?.(mountInfo, sourceText, {
          name: stat.name,
          path: stat.path,
          ext: stat.ext,
        });
        if (reqId !== loadRequestSeq) return;
        applyPreservedViewOptions(renderer, 'html', preserved);
      } else if (stat.kind === 'pdf') {
        const mountInfo = await mountPdfFile(stat.path);
        if (reqId !== loadRequestSeq) return;

        currentSource = { type: 'disk', path: stat.path, stat };
        const renderer = mountRendererWithState('pdf', preserved);
        await renderer.loadPdf?.(mountInfo, null, {
          name: stat.name,
          path: stat.path,
          ext: stat.ext,
          size: stat.size,
        });
        if (reqId !== loadRequestSeq) return;
        applyPreservedViewOptions(renderer, 'pdf', preserved);
      } else if (stat.kind === 'mp4') {
        const mountInfo = await mountVideoFile(stat.path);
        if (reqId !== loadRequestSeq) return;

        currentSource = { type: 'disk', path: stat.path, stat };
        const renderer = mountRendererWithState('mp4', preserved);
        await renderer.loadVideo?.(mountInfo, null, {
          name: stat.name,
          path: stat.path,
          ext: stat.ext,
          size: stat.size,
        });
        if (reqId !== loadRequestSeq) return;
      } else {
        const buffer = await readFileBuffer(stat.path);
        if (reqId !== loadRequestSeq) return;

        currentSource = { type: 'disk', path: stat.path, stat };
        const renderer = mountRendererWithState(stat.kind, preserved);
        await renderer.loadBuffer?.(buffer, {
          name: stat.name,
          path: stat.path,
          ext: stat.ext,
        });
        if (reqId !== loadRequestSeq) return;
        applyPreservedViewOptions(renderer, stat.kind, preserved);
      }

      activeError = null;
      isLoading = false;
      syncUi();
    } catch (err) {
      if (reqId !== loadRequestSeq) return;
      if (options.isAutoRefresh) {
        // 自动刷新遇到原子写入短暂锁定或缺失时静默跳过本轮
        console.debug('[file-preview] 自动刷新本轮跳过:', err);
        return;
      }

      const apiErr = err instanceof SidecarApiError ? err : null;
      const title =
        apiErr?.status === 415
          ? '暂不支持的文件格式'
          : apiErr?.status === 404
            ? '未找到目标文件'
            : apiErr?.status === 403
              ? '无法访问该路径'
              : '文档加载失败';
      const message =
        apiErr?.userMessage ||
        (err instanceof Error ? err.message : '加载或渲染文件时出现未知错误。');

      currentSource = {
        type: 'disk',
        path: rawPath,
        stat: {
          path: rawPath,
          name: rawPath.split(/[\\/]/).pop() || rawPath,
          ext: '',
          kind: 'word',
          size: 0,
          mtimeMs: 0,
        },
      };
      setErrorState(title, message, apiErr?.code, rawPath);
    }
  }

  async function openLocalDroppedFile(file: File): Promise<void> {
    const reqId = ++loadRequestSeq;
    const classified = classifyDroppedFile(file.name);

    viewMode = 'preview';
    toolbar.resetSearch();

    if (!classified.kind) {
      currentSource = null;
      setErrorState(
        classified.errorTitle || '不支持的文件格式',
        classified.errorMessage || '仅支持 .pptx、.docx、.pdf、.html 和 .htm 文件。',
      );
      return;
    }

    currentSource = {
      type: 'local',
      name: file.name,
      ext: classified.ext,
      kind: classified.kind,
      size: file.size,
    };
    activeError = null;
    isLoading = true;
    loadingTextEl.textContent = `正在解析本地文件 ${file.name}...`;
    syncUi();

    try {
      if (classified.kind === 'html') {
        const sourceText = await file.text();
        if (reqId !== loadRequestSeq) return;

        const renderer = mountRendererWithState('html');
        await renderer.loadHtml?.(null, sourceText, {
          name: file.name,
          path: file.name,
          ext: classified.ext,
        });
      } else if (classified.kind === 'pdf') {
        const buffer = await file.arrayBuffer();
        if (reqId !== loadRequestSeq) return;

        const renderer = mountRendererWithState('pdf');
        await renderer.loadPdf?.(null, buffer, {
          name: file.name,
          path: file.name,
          ext: classified.ext,
          size: file.size,
        });
      } else if (classified.kind === 'mp4') {
        const objectUrl = URL.createObjectURL(file);
        const renderer = mountRendererWithState('mp4');
        await renderer.loadVideo?.(null, objectUrl, {
          name: file.name,
          path: file.name,
          ext: classified.ext,
          size: file.size,
        });
      } else {
        const buffer = await file.arrayBuffer();
        if (reqId !== loadRequestSeq) return;

        const renderer = mountRendererWithState(classified.kind);
        await renderer.loadBuffer?.(buffer, {
          name: file.name,
          path: file.name,
          ext: classified.ext,
        });
      }

      if (reqId !== loadRequestSeq) return;
      isLoading = false;
      syncUi();
    } catch (err) {
      if (reqId !== loadRequestSeq) return;
      const message =
        err instanceof Error ? err.message : '本地文件解析失败，请确认文件未损坏。';
      setErrorState('本地文件渲染失败', message);
    }
  }

  // 全屏文件拖拽支持
  let dragDepth = 0;
  const hasDraggedFiles = (event: DragEvent): boolean =>
    Boolean(event.dataTransfer && Array.from(event.dataTransfer.types).includes('Files'));

  window.addEventListener('dragenter', (e) => {
    if (!hasDraggedFiles(e)) return;
    e.preventDefault();
    dragDepth += 1;
    dropOverlay.hidden = false;
  });

  window.addEventListener('dragover', (e) => {
    if (!hasDraggedFiles(e)) return;
    e.preventDefault();
    if (e.dataTransfer) {
      e.dataTransfer.dropEffect = 'copy';
    }
  });

  window.addEventListener('dragleave', (e) => {
    if (!hasDraggedFiles(e)) return;
    e.preventDefault();
    dragDepth = Math.max(0, dragDepth - 1);
    if (dragDepth === 0) {
      dropOverlay.hidden = true;
    }
  });

  window.addEventListener('drop', (e) => {
    if (!hasDraggedFiles(e)) return;
    e.preventDefault();
    dragDepth = 0;
    dropOverlay.hidden = true;
    const droppedFile = e.dataTransfer?.files?.[0];
    if (droppedFile) {
      void openLocalDroppedFile(droppedFile);
    }
  });

  // 自动刷新监听（每 2.5s 检查一次磁盘文件 mtimeMs）
  const pollTimer = window.setInterval(async () => {
    if (!autoRefreshEnabled || isLoading || isPollingStat) {
      return;
    }
    if (currentSource?.type !== 'disk' || currentSource.stat.mtimeMs <= 0) {
      return;
    }

    const targetPath = currentSource.path;
    const prevMtime = currentSource.stat.mtimeMs;
    const prevSize = currentSource.stat.size;

    isPollingStat = true;
    try {
      const nextStat = await statFile(targetPath);
      if (
        currentSource?.type === 'disk' &&
        currentSource.path === targetPath &&
        (nextStat.mtimeMs !== prevMtime || nextStat.size !== prevSize)
      ) {
        await openDiskFile(targetPath, { isAutoRefresh: true });
      }
    } catch {
      // 轮询期间文件可能正处于写入替换中，忽略单次检测错误
    } finally {
      isPollingStat = false;
    }
  }, AUTO_REFRESH_INTERVAL_MS);

  // 启动主题与工作区事件同步
  const stopThemeSync = startThemeSync();
  const stopWorkspaceSync = subscribeWorkspaceChanges(() => {
    void refreshWorkspaceList();
  });

  window.addEventListener('beforeunload', () => {
    window.clearInterval(pollTimer);
    stopThemeSync();
    stopWorkspaceSync();
    destroyActiveRenderer();
    toolbar.destroy();
    welcomeBrowser.destroy();
    drawerBrowser.destroy();
  });

  // 启动路由：检查 URL ?file= 参数
  const initialFile = new URLSearchParams(window.location.search).get('file')?.trim();
  if (initialFile) {
    void openDiskFile(initialFile);
    void refreshWorkspaceList();
  } else {
    syncUi();
    void refreshWorkspaceList();
  }
}

bootstrap();
