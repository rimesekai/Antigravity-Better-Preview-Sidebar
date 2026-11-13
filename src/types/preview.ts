export type PreviewKind = 'ppt' | 'word' | 'html';

export interface FileStatInfo {
  path: string;
  name: string;
  ext: string;
  kind: PreviewKind;
  size: number;
  mtimeMs: number;
}

export interface WorkspaceFileItem extends FileStatInfo {
  relPath: string;
  workspaceRoot: string;
}

export interface HtmlMountInfo {
  url: string;
  entryRelPath: string;
  mountId: string;
  mountRoot: string;
}

export type HtmlPreviewMode = 'sandbox' | 'interactive' | 'source';
export type PptLayoutMode = 'slide' | 'list';

export interface PreviewStateInfo {
  kind: PreviewKind;
  page?: number;
  totalPages?: number;
  zoom?: number;
  fitMode?: boolean;
  layoutMode?: PptLayoutMode;
  thumbnailsVisible?: boolean;
  searchMatches?: number;
  activeMatch?: number;
  htmlMode?: HtmlPreviewMode;
  warning?: string;
}

export interface PreviewRenderer {
  mount(container: HTMLElement): void;
  loadBuffer?(buffer: ArrayBuffer, meta: { name: string; path: string; ext: string }): Promise<void>;
  loadHtml?(
    mountInfo: HtmlMountInfo | null,
    sourceText: string,
    meta: { name: string; path: string; ext: string }
  ): Promise<void>;
  refresh?(): Promise<void>;
  setZoom?(zoom: number | 'fit'): void;
  nextPage?(): void;
  prevPage?(): void;
  goToPage?(page: number): void;
  setLayoutMode?(mode: PptLayoutMode): void;
  toggleThumbnails?(visible?: boolean): void;
  search?(query: string): Promise<{ total: number; current: number }>;
  nextMatch?(): void;
  prevMatch?(): void;
  setHtmlMode?(mode: HtmlPreviewMode): void;
  onStateChange?(cb: (info: PreviewStateInfo) => void): void;
  destroy(): void;
}
