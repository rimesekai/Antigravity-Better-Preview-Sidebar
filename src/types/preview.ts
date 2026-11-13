export type PreviewKind = 'ppt' | 'word' | 'html' | 'pdf' | 'mp4';

export const LARGE_PDF_THRESHOLD_BYTES = 20 * 1024 * 1024;

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

export interface PdfMountInfo {
  url: string;
  mountId: string;
  fileName: string;
  size: number;
}

export interface VideoMountInfo {
  url: string;
  mountId: string;
  fileName: string;
  size: number;
}

export type HtmlPreviewMode = 'sandbox' | 'interactive' | 'source';
export type PptLayoutMode = 'slide' | 'list';
export type PdfViewMode = 'canvas' | 'native';

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
  pdfMode?: PdfViewMode;
  fileSize?: number;
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
  loadPdf?(
    mountInfo: PdfMountInfo | null,
    buffer: ArrayBuffer | null,
    meta: { name: string; path: string; ext: string; size?: number }
  ): Promise<void>;
  loadVideo?(
    mountInfo: VideoMountInfo | null,
    objectUrl: string | null,
    meta: { name: string; path: string; ext: string; size?: number }
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
  setPdfMode?(mode: PdfViewMode): void;
  onStateChange?(cb: (info: PreviewStateInfo) => void): void;
  destroy(): void;
}
