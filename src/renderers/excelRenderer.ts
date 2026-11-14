import * as XLSX from 'xlsx';
import type { PreviewRenderer, PreviewStateInfo } from '../types/preview';
import { copyTextToClipboard } from '../utils/clipboard';
import { onLocaleChange, tr } from '../utils/i18n';

const MIN_ZOOM = 0.5;
const MAX_ZOOM = 2.0;
const MAX_INITIAL_ROWS = 2000;
const ROW_BATCH_SIZE = 2000;
const MAX_RENDER_ROWS_HARD = 20000;
const MAX_INITIAL_COLS = 100;
const COL_BATCH_SIZE = 100;
const MAX_RENDER_COLS_HARD = 500;
const DEFAULT_COL_WIDTH_PX = 110;
const MIN_COL_WIDTH_PX = 64;
const MAX_COL_WIDTH_PX = 360;
const ROW_HEADER_WIDTH_PX = 52;

interface MergeMasterInfo {
  rowSpan: number;
  colSpan: number;
  rangeRef: string;
}

interface MergeSpanInfo {
  endR: number;
  endC: number;
  rangeRef: string;
}

interface SearchMatchItem {
  sheetIndex: number;
  r: number;
  c: number;
  addr: string;
}

interface SelectedCellInfo {
  sheetIndex: number;
  r: number;
  c: number;
  endR: number;
  endC: number;
  addr: string;
  displayRef: string;
}

function isEditableTarget(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) return false;
  if (target.isContentEditable) return true;
  const tag = target.tagName;
  return tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT';
}

function hasCellContent(cell: XLSX.CellObject | undefined): boolean {
  if (!cell) return false;
  if (cell.f !== undefined && cell.f !== '') return true;
  if (cell.v !== undefined && cell.v !== null && cell.v !== '') return true;
  if (cell.w !== undefined && cell.w !== null && cell.w !== '') return true;
  return false;
}

/**
 * 裁剪工作表稀疏 `!ref` 声明尾部（如仅含空样式的 A1:XFD1048576），返回真实非空边界。
 */
function getEffectiveSheetBounds(sheet: XLSX.WorkSheet | undefined): {
  totalRows: number;
  totalCols: number;
} {
  if (!sheet || !sheet['!ref']) {
    return { totalRows: 0, totalCols: 0 };
  }

  let declared: XLSX.Range;
  try {
    declared = XLSX.utils.decode_range(sheet['!ref']);
  } catch {
    return { totalRows: 0, totalCols: 0 };
  }

  let maxRow = -1;
  let maxCol = -1;
  for (const key of Object.keys(sheet)) {
    if (key.startsWith('!')) continue;
    const cell = sheet[key] as XLSX.CellObject | undefined;
    if (!hasCellContent(cell)) continue;
    try {
      const pos = XLSX.utils.decode_cell(key);
      if (pos.r > maxRow) maxRow = pos.r;
      if (pos.c > maxCol) maxCol = pos.c;
    } catch {
      // 忽略非法单元格键名
    }
  }

  // 合并单元格仅在接近真实内容边界且跨度合理时计入，防止病态全列/全行合并撑爆渲染边界
  const merges = sheet['!merges'] ?? [];
  for (const m of merges) {
    if (m.e.r < 0 || m.e.c < 0) continue;
    const rowSpan = m.e.r - m.s.r + 1;
    const colSpan = m.e.c - m.s.c + 1;
    if (rowSpan <= 1000 && (maxRow < 0 || m.s.r <= maxRow + 50)) {
      if (m.e.r > maxRow) maxRow = m.e.r;
    }
    if (colSpan <= 200 && (maxCol < 0 || m.s.c <= maxCol + 20)) {
      if (m.e.c > maxCol) maxCol = m.e.c;
    }
  }

  if (maxRow < 0 || maxCol < 0) {
    return { totalRows: 0, totalCols: 0 };
  }

  return {
    totalRows: Math.max(0, Math.min(declared.e.r + 1, maxRow + 1)),
    totalCols: Math.max(0, Math.min(declared.e.c + 1, maxCol + 1)),
  };
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function roundZoom(value: number): number {
  return Math.round(value * 100) / 100;
}

function cellKey(r: number, c: number): string {
  return `${r},${c}`;
}

function decodeDelimitedText(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let text: string;
  try {
    text = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    text = new TextDecoder('gb18030').decode(bytes);
  }
  if (text.charCodeAt(0) === 0xfeff) {
    text = text.slice(1);
  }
  return text;
}

function getColumnWidthPx(sheet: XLSX.WorkSheet, colIndex: number): number {
  const colInfo = sheet['!cols']?.[colIndex];
  let px = DEFAULT_COL_WIDTH_PX;
  if (colInfo) {
    if (typeof colInfo.wpx === 'number' && Number.isFinite(colInfo.wpx) && colInfo.wpx > 0) {
      px = colInfo.wpx;
    } else if (typeof colInfo.wch === 'number' && Number.isFinite(colInfo.wch) && colInfo.wch > 0) {
      px = colInfo.wch * 8 + 16;
    }
  }
  return Math.round(clamp(px, MIN_COL_WIDTH_PX, MAX_COL_WIDTH_PX));
}

function getCellDisplayText(cell: XLSX.CellObject | undefined): string {
  if (!cell) return '';
  if (cell.w !== undefined && cell.w !== null && cell.w !== '') {
    return String(cell.w);
  }
  try {
    const formatted = XLSX.utils.format_cell(cell);
    if (formatted !== undefined && formatted !== null && formatted !== '') {
      return String(formatted);
    }
  } catch {
    // 忽略格式化异常并回退到原始字段
  }
  if (cell.v !== undefined && cell.v !== null) {
    return String(cell.v);
  }
  return '';
}

function getNormalizedExtension(meta: { name: string; path: string; ext: string }): string {
  const raw = (meta.ext || '').trim().toLowerCase();
  if (raw) {
    return raw.startsWith('.') ? raw : `.${raw}`;
  }
  const idx = meta.name.lastIndexOf('.');
  return idx >= 0 ? meta.name.slice(idx).toLowerCase() : '';
}

export function createExcelRenderer(): PreviewRenderer {
  let rootEl: HTMLElement | null = null;
  let formulaBarEl: HTMLElement | null = null;
  let cellRefEl: HTMLElement | null = null;
  let sheetStatsEl: HTMLElement | null = null;
  let formulaValueEl: HTMLElement | null = null;
  let copyBtnEl: HTMLButtonElement | null = null;
  let bannerEl: HTMLElement | null = null;
  let viewportEl: HTMLElement | null = null;
  let sheetBarEl: HTMLElement | null = null;

  let tableEl: HTMLTableElement | null = null;
  let rowHeaderColEl: HTMLTableColElement | null = null;
  let dataColEls: HTMLTableColElement[] = [];
  let baseColWidths: number[] = [];
  let cellElementMap = new Map<string, HTMLTableCellElement>();
  let colHeaderEls: HTMLTableCellElement[] = [];
  let rowHeaderEls: HTMLTableCellElement[] = [];
  let mergeSpanByKey = new Map<string, MergeSpanInfo>();

  let lastBuffer: ArrayBuffer | null = null;
  let lastMeta: { name: string; path: string; ext: string } | null = null;
  let workbook: XLSX.WorkBook | null = null;

  let currentSheetIndex = 0;
  let stopLocaleSync: (() => void) | null = null;
  let currentRowCount = 0;
  let currentColCount = 0;
  let fitMode = true;
  let currentZoom = 1.0;

  let selectedCell: SelectedCellInfo | null = null;
  let copyFeedbackTimer: number | null = null;
  let copyFeedbackText = '';

  let searchQuery = '';
  let searchMatches: SearchMatchItem[] = [];
  let activeMatchIndex = -1;

  const expandedRowsLimitBySheet = new Map<string, number>();
  const expandedColsLimitBySheet = new Map<string, number>();
  const stateListeners = new Set<(info: PreviewStateInfo) => void>();
  let destroyed = false;

  function emitState(warning?: string): void {
    const sheetNames = workbook?.SheetNames ?? [];
    const activeSheet = sheetNames[currentSheetIndex];
    const info: PreviewStateInfo = {
      kind: 'excel',
      fileSize: lastBuffer?.byteLength,
      page: sheetNames.length > 0 ? currentSheetIndex + 1 : undefined,
      totalPages: sheetNames.length > 0 ? sheetNames.length : undefined,
      zoom: roundZoom(currentZoom),
      fitMode,
      searchMatches: searchMatches.length,
      activeMatch: activeMatchIndex >= 0 ? activeMatchIndex + 1 : 0,
      sheetNames,
      activeSheet,
      rowCount: currentRowCount,
      colCount: currentColCount,
      warning,
    };
    for (const cb of stateListeners) {
      try {
        cb(info);
      } catch (err) {
        console.error('[excelRenderer] state listener error:', err);
      }
    }
  }

  function clearCopyFeedback(): void {
    if (copyFeedbackTimer !== null) {
      window.clearTimeout(copyFeedbackTimer);
      copyFeedbackTimer = null;
    }
    copyFeedbackText = '';
    if (copyBtnEl) {
      copyBtnEl.classList.remove('is-copied', 'is-copy-failed');
      copyBtnEl.textContent = tr('复制', 'Copy');
    }
  }

  function getActiveSheet(): XLSX.WorkSheet | undefined {
    if (!workbook || workbook.SheetNames.length === 0) return undefined;
    const name = workbook.SheetNames[currentSheetIndex];
    return name ? workbook.Sheets[name] : undefined;
  }

  function updateFormulaBar(): void {
    if (!cellRefEl || !sheetStatsEl || !formulaValueEl || !copyBtnEl) return;

    const sheetNames = workbook?.SheetNames ?? [];
    if (sheetNames.length === 0) {
      cellRefEl.textContent = '—';
      sheetStatsEl.textContent = tr('0 行 × 0 列', '0 rows × 0 cols');
      formulaValueEl.textContent = tr('当前工作簿无可用工作表', 'No sheets available in workbook');
      formulaValueEl.classList.add('is-placeholder');
      copyBtnEl.disabled = true;
      copyFeedbackText = '';
      return;
    }

    sheetStatsEl.textContent = tr(
      `${currentRowCount.toLocaleString()} 行 × ${currentColCount} 列`,
      `${currentRowCount.toLocaleString()} rows × ${currentColCount} cols`,
    );

    if (!selectedCell || selectedCell.sheetIndex !== currentSheetIndex) {
      cellRefEl.textContent = '—';
      formulaValueEl.textContent =
        currentRowCount === 0 && currentColCount === 0
          ? tr('当前工作表暂无单元格数据', 'This sheet has no cell data')
          : tr('点击单元格查看完整内容或公式', 'Click a cell to inspect its value or formula');
      formulaValueEl.classList.add('is-placeholder');
      copyBtnEl.disabled = true;
      copyFeedbackText = '';
      return;
    }

    const sheet = getActiveSheet();
    const cell = sheet?.[selectedCell.addr] as XLSX.CellObject | undefined;
    const displayValue = getCellDisplayText(cell);

    cellRefEl.textContent = selectedCell.displayRef;

    if (cell?.f) {
      const fullText = tr(
        `=${cell.f}  (值: ${displayValue})`,
        `=${cell.f}  (Value: ${displayValue})`,
      );
      formulaValueEl.textContent = fullText;
      formulaValueEl.classList.remove('is-placeholder');
      copyFeedbackText = displayValue !== '' ? displayValue : `=${cell.f}`;
      copyBtnEl.disabled = false;
    } else if (displayValue !== '') {
      formulaValueEl.textContent = displayValue;
      formulaValueEl.classList.remove('is-placeholder');
      copyFeedbackText = displayValue;
      copyBtnEl.disabled = false;
    } else {
      formulaValueEl.textContent = tr('（空单元格）', '(Empty cell)');
      formulaValueEl.classList.add('is-placeholder');
      copyFeedbackText = '';
      copyBtnEl.disabled = true;
    }
  }

  function updateHeaderHighlight(): void {
    colHeaderEls.forEach((th, idx) => {
      const active =
        selectedCell !== null &&
        selectedCell.sheetIndex === currentSheetIndex &&
        idx >= selectedCell.c &&
        idx <= selectedCell.endC;
      th.classList.toggle('is-active-header', active);
    });
    rowHeaderEls.forEach((th, idx) => {
      const active =
        selectedCell !== null &&
        selectedCell.sheetIndex === currentSheetIndex &&
        idx >= selectedCell.r &&
        idx <= selectedCell.endR;
      th.classList.toggle('is-active-header', active);
    });
  }

  function selectCellAt(r: number, c: number, syncScroll = false): void {
    const key = cellKey(r, c);
    const prevSelected = cellElementMap.get(
      selectedCell && selectedCell.sheetIndex === currentSheetIndex
        ? cellKey(selectedCell.r, selectedCell.c)
        : '',
    );
    if (prevSelected) {
      prevSelected.classList.remove('is-selected');
    }

    const addr = XLSX.utils.encode_cell({ r, c });
    const mergeSpan = mergeSpanByKey.get(key);
    selectedCell = {
      sheetIndex: currentSheetIndex,
      r,
      c,
      endR: mergeSpan?.endR ?? r,
      endC: mergeSpan?.endC ?? c,
      addr,
      displayRef: mergeSpan?.rangeRef ?? addr,
    };

    const td = cellElementMap.get(key);
    if (td) {
      td.classList.add('is-selected');
      if (syncScroll) {
        td.scrollIntoView({ block: 'center', inline: 'center' });
      }
    }

    clearCopyFeedback();
    updateFormulaBar();
    updateHeaderHighlight();
  }

  function applyZoomToDOM(): void {
    if (!viewportEl) return;
    viewportEl.style.setProperty('--fp-excel-zoom', String(currentZoom));

    if (rowHeaderColEl) {
      rowHeaderColEl.style.width = `${Math.round(ROW_HEADER_WIDTH_PX * currentZoom)}px`;
    }
    let totalTableWidth = Math.round(ROW_HEADER_WIDTH_PX * currentZoom);
    for (let c = 0; c < dataColEls.length; c++) {
      const scaledW = Math.round((baseColWidths[c] ?? DEFAULT_COL_WIDTH_PX) * currentZoom);
      dataColEls[c].style.width = `${scaledW}px`;
      totalTableWidth += scaledW;
    }
    if (tableEl) {
      tableEl.style.width = `${totalTableWidth}px`;
    }
  }

  function applySearchHighlightsToCurrentSheet(): void {
    for (const td of cellElementMap.values()) {
      td.classList.remove('fp-excel-search-hit', 'is-active');
    }

    if (!searchQuery || searchMatches.length === 0) return;

    const activeMatch = activeMatchIndex >= 0 ? searchMatches[activeMatchIndex] : undefined;
    for (const match of searchMatches) {
      if (match.sheetIndex !== currentSheetIndex) continue;
      const td = cellElementMap.get(cellKey(match.r, match.c));
      if (!td) continue;
      td.classList.add('fp-excel-search-hit');
      if (
        activeMatch &&
        activeMatch.sheetIndex === match.sheetIndex &&
        activeMatch.r === match.r &&
        activeMatch.c === match.c
      ) {
        td.classList.add('is-active');
      }
    }
  }

  function renderEmptyState(title: string, desc: string): void {
    if (!viewportEl) return;
    viewportEl.innerHTML = '';
    const box = document.createElement('div');
    box.className = 'fp-excel-empty';

    const titleEl = document.createElement('div');
    titleEl.className = 'fp-excel-empty-title';
    titleEl.textContent = title;

    const descEl = document.createElement('div');
    descEl.className = 'fp-excel-empty-desc';
    descEl.textContent = desc;

    box.append(titleEl, descEl);
    viewportEl.appendChild(box);
  }

  function renderSheetBar(): void {
    if (!sheetBarEl) return;
    sheetBarEl.innerHTML = '';

    const sheetNames = workbook?.SheetNames ?? [];
    if (sheetNames.length === 0) {
      sheetBarEl.hidden = true;
      return;
    }

    sheetBarEl.hidden = false;
    sheetNames.forEach((name, idx) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'fp-excel-sheet-tab';
      btn.dataset.sheetIndex = String(idx);
      const isActive = idx === currentSheetIndex;
      btn.classList.toggle('is-active', isActive);
      btn.setAttribute('aria-selected', isActive ? 'true' : 'false');
      btn.textContent = name;
      btn.title = name;
      sheetBarEl?.appendChild(btn);
    });

    const activeTab = sheetBarEl.querySelector<HTMLElement>('.fp-excel-sheet-tab.is-active');
    activeTab?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  function renderBanner(
    sheetName: string,
    totalRows: number,
    totalCols: number,
    renderRows: number,
    renderCols: number,
  ): void {
    if (!bannerEl) return;
    const rowsTruncated = renderRows < totalRows;
    const colsTruncated = renderCols < totalCols;
    if (!rowsTruncated && !colsTruncated) {
      bannerEl.hidden = true;
      bannerEl.innerHTML = '';
      return;
    }

    const rowCap = Math.min(totalRows, MAX_RENDER_ROWS_HARD);
    const colCap = Math.min(totalCols, MAX_RENDER_COLS_HARD);
    const canExpandRows = renderRows < rowCap;
    const canExpandCols = renderCols < colCap;
    const hitHardCap = renderRows >= MAX_RENDER_ROWS_HARD || renderCols >= MAX_RENDER_COLS_HARD;

    bannerEl.hidden = false;
    bannerEl.innerHTML = '';

    const msgEl = document.createElement('span');
    msgEl.className = 'fp-excel-truncation-text';
    const renderedParts: string[] = [];
    if (rowsTruncated) {
      renderedParts.push(tr(`前 ${renderRows.toLocaleString()} 行`, `first ${renderRows.toLocaleString()} rows`));
    }
    if (colsTruncated) {
      renderedParts.push(tr(`前 ${renderCols} 列`, `first ${renderCols} cols`));
    }
    let msg = tr(
      `当前工作表共 ${totalRows.toLocaleString()} 行 × ${totalCols} 列，已渲染${renderedParts.join('、')}`,
      `Sheet has ${totalRows.toLocaleString()} rows × ${totalCols} cols; showing ${renderedParts.join(' and ')}`,
    );
    if (hitHardCap) {
      msg += tr(
        '（已达安全渲染上限，可使用搜索精准定位）',
        ' (safe render cap reached; use search to jump directly)',
      );
    }
    msgEl.textContent = msg;
    bannerEl.appendChild(msgEl);

    if (canExpandRows || canExpandCols) {
      const actionsEl = document.createElement('div');
      actionsEl.className = 'fp-excel-banner-actions';

      if (canExpandRows) {
        const nextRowLimit = Math.min(rowCap, renderRows + ROW_BATCH_SIZE);
        const deltaRows = nextRowLimit - renderRows;
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'fp-btn fp-btn-secondary fp-excel-load-all-btn';
        btn.textContent =
          nextRowLimit >= totalRows
            ? tr(`加载全部行 (+${deltaRows.toLocaleString()})`, `Load all rows (+${deltaRows.toLocaleString()})`)
            : tr(`再加载 ${deltaRows.toLocaleString()} 行`, `Load ${deltaRows.toLocaleString()} more rows`);
        btn.addEventListener('click', () => {
          if (destroyed) return;
          expandedRowsLimitBySheet.set(sheetName, nextRowLimit);
          renderCurrentSheet(true);
          emitState();
        });
        actionsEl.appendChild(btn);
      }

      if (canExpandCols) {
        const nextColLimit = Math.min(colCap, renderCols + COL_BATCH_SIZE);
        const deltaCols = nextColLimit - renderCols;
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'fp-btn fp-btn-secondary fp-excel-load-all-btn';
        btn.textContent =
          nextColLimit >= totalCols
            ? tr(`展开全部列 (+${deltaCols})`, `Expand all cols (+${deltaCols})`)
            : tr(`再加载 ${deltaCols} 列`, `Load ${deltaCols} more cols`);
        btn.addEventListener('click', () => {
          if (destroyed) return;
          expandedColsLimitBySheet.set(sheetName, nextColLimit);
          renderCurrentSheet(true);
          emitState();
        });
        actionsEl.appendChild(btn);
      }

      bannerEl.appendChild(actionsEl);
    }
  }

  function resetGridRefs(): void {
    tableEl = null;
    rowHeaderColEl = null;
    dataColEls = [];
    baseColWidths = [];
    cellElementMap.clear();
    colHeaderEls = [];
    rowHeaderEls = [];
    mergeSpanByKey.clear();
  }

  function renderCurrentSheet(preserveScroll = false): void {
    if (destroyed || !viewportEl) return;
    const prevScrollTop = preserveScroll ? viewportEl.scrollTop : 0;
    const prevScrollLeft = preserveScroll ? viewportEl.scrollLeft : 0;
    resetGridRefs();

    const sheetNames = workbook?.SheetNames ?? [];
    if (sheetNames.length === 0) {
      currentRowCount = 0;
      currentColCount = 0;
      if (bannerEl) {
        bannerEl.hidden = true;
        bannerEl.innerHTML = '';
      }
      renderEmptyState(tr('工作簿为空', 'Empty Workbook'), tr('当前文件不包含任何工作表', 'This file contains no sheets.'));
      renderSheetBar();
      updateFormulaBar();
      return;
    }

    currentSheetIndex = clamp(currentSheetIndex, 0, sheetNames.length - 1);
    const sheetName = sheetNames[currentSheetIndex];
    const sheet = workbook?.Sheets[sheetName];

    renderSheetBar();

    const { totalRows, totalCols } = getEffectiveSheetBounds(sheet);
    currentRowCount = totalRows;
    currentColCount = totalCols;

    if (!sheet || totalRows === 0 || totalCols === 0) {
      if (bannerEl) {
        bannerEl.hidden = true;
        bannerEl.innerHTML = '';
      }
      renderEmptyState(
        tr('空工作表', 'Empty Sheet'),
        tr(`工作表「${sheetName}」暂无单元格数据`, `Sheet "${sheetName}" has no cell data.`),
      );
      updateFormulaBar();
      return;
    }

    const rowLimit = Math.min(
      MAX_RENDER_ROWS_HARD,
      expandedRowsLimitBySheet.get(sheetName) ?? MAX_INITIAL_ROWS,
    );
    const colLimit = Math.min(
      MAX_RENDER_COLS_HARD,
      expandedColsLimitBySheet.get(sheetName) ?? MAX_INITIAL_COLS,
    );
    const renderRows = Math.min(totalRows, rowLimit);
    const renderCols = Math.min(totalCols, colLimit);

    renderBanner(sheetName, totalRows, totalCols, renderRows, renderCols);

    // 解析合并单元格
    const mergeMasterMap = new Map<string, MergeMasterInfo>();
    const coveredCellSet = new Set<string>();
    const merges = sheet['!merges'] ?? [];
    for (const m of merges) {
      if (m.e.r < 0 || m.e.c < 0 || m.s.r >= renderRows || m.s.c >= renderCols) {
        continue;
      }
      const startR = Math.max(0, m.s.r);
      const startC = Math.max(0, m.s.c);
      const endR = Math.min(renderRows - 1, m.e.r);
      const endC = Math.min(renderCols - 1, m.e.c);
      const rowSpan = endR - startR + 1;
      const colSpan = endC - startC + 1;
      if (rowSpan < 1 || colSpan < 1) continue;

      const masterKey = cellKey(startR, startC);
      if (coveredCellSet.has(masterKey)) continue;

      const rangeRef = XLSX.utils.encode_range(m);
      if (rowSpan > 1 || colSpan > 1) {
        mergeMasterMap.set(masterKey, { rowSpan, colSpan, rangeRef });
      }
      mergeSpanByKey.set(masterKey, { endR, endC, rangeRef });

      for (let r = startR; r <= endR; r++) {
        for (let c = startC; c <= endC; c++) {
          if (r === startR && c === startC) continue;
          coveredCellSet.add(cellKey(r, c));
        }
      }
    }

    const table = document.createElement('table');
    table.className = 'fp-excel-table';
    tableEl = table;

    const colgroup = document.createElement('colgroup');
    const rowCol = document.createElement('col');
    rowCol.className = 'fp-excel-col-rowhdr';
    rowHeaderColEl = rowCol;
    colgroup.appendChild(rowCol);

    for (let c = 0; c < renderCols; c++) {
      const w = getColumnWidthPx(sheet, c);
      baseColWidths.push(w);
      const colEl = document.createElement('col');
      dataColEls.push(colEl);
      colgroup.appendChild(colEl);
    }
    table.appendChild(colgroup);

    // 表头列标 (A, B, C...)
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    const cornerTh = document.createElement('th');
    cornerTh.className = 'fp-excel-corner';
    cornerTh.setAttribute('scope', 'col');
    headerRow.appendChild(cornerTh);

    for (let c = 0; c < renderCols; c++) {
      const th = document.createElement('th');
      th.className = 'fp-excel-col-header';
      th.setAttribute('scope', 'col');
      th.textContent = XLSX.utils.encode_col(c);
      colHeaderEls.push(th);
      headerRow.appendChild(th);
    }
    thead.appendChild(headerRow);
    table.appendChild(thead);

    // 表体数据行
    const tbody = document.createElement('tbody');
    const rowsFrag = document.createDocumentFragment();

    for (let r = 0; r < renderRows; r++) {
      const tr = document.createElement('tr');

      const rowTh = document.createElement('th');
      rowTh.className = 'fp-excel-row-header';
      rowTh.setAttribute('scope', 'row');
      rowTh.textContent = String(r + 1);
      rowHeaderEls.push(rowTh);
      tr.appendChild(rowTh);

      for (let c = 0; c < renderCols; c++) {
        const key = cellKey(r, c);
        if (coveredCellSet.has(key)) {
          continue;
        }

        const td = document.createElement('td');
        td.className = 'fp-excel-cell';
        td.dataset.r = String(r);
        td.dataset.c = String(c);

        const mergeInfo = mergeMasterMap.get(key);
        if (mergeInfo) {
          if (mergeInfo.rowSpan > 1) td.rowSpan = mergeInfo.rowSpan;
          if (mergeInfo.colSpan > 1) td.colSpan = mergeInfo.colSpan;
          td.classList.add('fp-excel-cell-merged');
        }

        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = sheet[addr] as XLSX.CellObject | undefined;
        const displayText = getCellDisplayText(cell);
        td.textContent = displayText;

        if (cell) {
          if (cell.t === 'n') {
            td.classList.add('fp-excel-cell-num');
          } else if (cell.t === 'b') {
            td.classList.add('fp-excel-cell-bool');
          } else if (cell.t === 'e') {
            td.classList.add('fp-excel-cell-err');
          }
          if (cell.f) {
            td.classList.add('fp-excel-cell-formula');
          }
          if (displayText) {
            td.title = cell.f ? `=${cell.f} (${displayText})` : displayText;
          } else if (cell.f) {
            td.title = `=${cell.f}`;
          }
        }

        if (
          selectedCell &&
          selectedCell.sheetIndex === currentSheetIndex &&
          selectedCell.r === r &&
          selectedCell.c === c
        ) {
          td.classList.add('is-selected');
        }

        cellElementMap.set(key, td);
        tr.appendChild(td);
      }

      rowsFrag.appendChild(tr);
    }

    tbody.appendChild(rowsFrag);
    table.appendChild(tbody);

    viewportEl.innerHTML = '';
    viewportEl.appendChild(table);
    viewportEl.scrollTop = prevScrollTop;
    viewportEl.scrollLeft = prevScrollLeft;

    applyZoomToDOM();
    applySearchHighlightsToCurrentSheet();
    updateFormulaBar();
    updateHeaderHighlight();
  }

  function focusActiveMatch(): void {
    if (activeMatchIndex < 0 || activeMatchIndex >= searchMatches.length || !workbook) {
      applySearchHighlightsToCurrentSheet();
      return;
    }

    const match = searchMatches[activeMatchIndex];
    const targetSheetName = workbook.SheetNames[match.sheetIndex];
    let needRerender = false;

    if (targetSheetName) {
      const currentRowLimit = Math.min(
        MAX_RENDER_ROWS_HARD,
        expandedRowsLimitBySheet.get(targetSheetName) ?? MAX_INITIAL_ROWS,
      );
      if (match.r >= currentRowLimit && match.r < MAX_RENDER_ROWS_HARD) {
        const neededRows = Math.min(
          MAX_RENDER_ROWS_HARD,
          Math.ceil((match.r + 1) / ROW_BATCH_SIZE) * ROW_BATCH_SIZE,
        );
        expandedRowsLimitBySheet.set(targetSheetName, neededRows);
        needRerender = true;
      }

      const currentColLimit = Math.min(
        MAX_RENDER_COLS_HARD,
        expandedColsLimitBySheet.get(targetSheetName) ?? MAX_INITIAL_COLS,
      );
      if (match.c >= currentColLimit && match.c < MAX_RENDER_COLS_HARD) {
        const neededCols = Math.min(
          MAX_RENDER_COLS_HARD,
          Math.ceil((match.c + 1) / COL_BATCH_SIZE) * COL_BATCH_SIZE,
        );
        expandedColsLimitBySheet.set(targetSheetName, neededCols);
        needRerender = true;
      }
    }

    if (match.sheetIndex !== currentSheetIndex) {
      currentSheetIndex = match.sheetIndex;
      renderCurrentSheet();
    } else if (needRerender) {
      renderCurrentSheet(true);
    } else {
      applySearchHighlightsToCurrentSheet();
    }

    selectCellAt(match.r, match.c, false);

    const td = cellElementMap.get(cellKey(match.r, match.c));
    if (td) {
      td.scrollIntoView({ block: 'center', inline: 'center' });
    }
  }

  function collectSheetMatches(
    sheet: XLSX.WorkSheet,
    sheetIndex: number,
    queryLower: string,
    out: SearchMatchItem[],
  ): void {
    const { totalRows, totalCols } = getEffectiveSheetBounds(sheet);
    if (totalRows === 0 || totalCols === 0) return;

    // 搜索边界与最大可渲染行列上限对齐，确保每条命中均能展开网格并高亮聚焦
    const endR = Math.min(totalRows - 1, MAX_RENDER_ROWS_HARD - 1);
    const endC = Math.min(totalCols - 1, MAX_RENDER_COLS_HARD - 1);
    const totalCells = (endR + 1) * (endC + 1);

    const checkCell = (r: number, c: number, addr: string, cell: XLSX.CellObject | undefined) => {
      if (!cell) return;
      const displayText = getCellDisplayText(cell);
      const formulaText = cell.f ? `=${cell.f}` : '';
      if (
        (displayText && displayText.toLowerCase().includes(queryLower)) ||
        (formulaText && formulaText.toLowerCase().includes(queryLower))
      ) {
        out.push({ sheetIndex, r, c, addr });
      }
    };

    if (totalCells > 250_000) {
      const coords: Array<{ r: number; c: number; addr: string }> = [];
      for (const key of Object.keys(sheet)) {
        if (key.startsWith('!') || !/^[A-Z]+\d+$/i.test(key)) continue;
        const pos = XLSX.utils.decode_cell(key);
        if (pos.r < 0 || pos.r > endR || pos.c < 0 || pos.c > endC) continue;
        coords.push({ r: pos.r, c: pos.c, addr: key });
      }
      coords.sort((a, b) => (a.r !== b.r ? a.r - b.r : a.c - b.c));
      for (const item of coords) {
        checkCell(item.r, item.c, item.addr, sheet[item.addr] as XLSX.CellObject | undefined);
      }
      return;
    }

    for (let r = 0; r <= endR; r++) {
      for (let c = 0; c <= endC; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        checkCell(r, c, addr, sheet[addr] as XLSX.CellObject | undefined);
      }
    }
  }

  function parseWorkbook(buffer: ArrayBuffer, meta: { name: string; path: string; ext: string }): XLSX.WorkBook {
    const ext = getNormalizedExtension(meta);
    if (ext === '.csv' || ext === '.tsv') {
      const text = decodeDelimitedText(buffer);
      if (!text.trim()) {
        return { SheetNames: [], Sheets: {} };
      }
      return XLSX.read(text, {
        type: 'string',
        cellFormula: true,
        cellDates: false,
        cellNF: true,
        cellStyles: true,
        ...(ext === '.tsv' ? { FS: '\t' } : {}),
      });
    }

    if (buffer.byteLength === 0) {
      return { SheetNames: [], Sheets: {} };
    }

    return XLSX.read(new Uint8Array(buffer), {
      type: 'array',
      cellFormula: true,
      cellDates: false,
      cellNF: true,
      cellStyles: true,
    });
  }

  function handleViewportClick(e: MouseEvent): void {
    const target = e.target as HTMLElement | null;
    if (!target) return;
    const td = target.closest<HTMLTableCellElement>('td.fp-excel-cell[data-r][data-c]');
    if (!td) return;
    const r = Number.parseInt(td.dataset.r ?? '', 10);
    const c = Number.parseInt(td.dataset.c ?? '', 10);
    if (!Number.isFinite(r) || !Number.isFinite(c)) return;
    selectCellAt(r, c, false);
  }

  function handleSheetBarClick(e: MouseEvent): void {
    const target = e.target as HTMLElement | null;
    if (!target) return;
    const tab = target.closest<HTMLButtonElement>('button.fp-excel-sheet-tab[data-sheet-index]');
    if (!tab) return;
    const idx = Number.parseInt(tab.dataset.sheetIndex ?? '', 10);
    if (!Number.isFinite(idx)) return;
    goToPage(idx + 1);
  }

  function showCopyStatus(ok: boolean): void {
    if (destroyed || !copyBtnEl) return;
    copyBtnEl.classList.toggle('is-copied', ok);
    copyBtnEl.classList.toggle('is-copy-failed', !ok);
    copyBtnEl.textContent = ok ? tr('已复制', 'Copied') : tr('复制失败', 'Failed');
    if (copyFeedbackTimer !== null) {
      window.clearTimeout(copyFeedbackTimer);
    }
    copyFeedbackTimer = window.setTimeout(() => {
      copyFeedbackTimer = null;
      if (!copyBtnEl || destroyed) return;
      copyBtnEl.classList.remove('is-copied', 'is-copy-failed');
      copyBtnEl.textContent = tr('复制', 'Copy');
    }, 1500);
  }

  async function handleCopyClick(): Promise<void> {
    if (!copyFeedbackText || !copyBtnEl) return;
    const ok = await copyTextToClipboard(copyFeedbackText);
    if (!ok) {
      console.warn('[excelRenderer] 复制到剪贴板失败');
    }
    showCopyStatus(ok);
  }

  function handleKeyDown(e: KeyboardEvent): void {
    if (destroyed || !rootEl) return;
    if (e.defaultPrevented || e.isComposing) return;
    if (!(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey) return;
    if (e.key.toLowerCase() !== 'c') return;
    if (isEditableTarget(e.target)) return;

    // 若用户手动拖选了局部文本，保留浏览器原生选区复制行为
    const selectionText = window.getSelection()?.toString() ?? '';
    if (selectionText.length > 0) return;

    if (!copyFeedbackText) return;
    e.preventDefault();
    void handleCopyClick();
  }

  function goToPage(page: number): void {
    if (!workbook || workbook.SheetNames.length === 0 || !Number.isFinite(page)) return;
    const total = workbook.SheetNames.length;
    const targetIdx = clamp(Math.round(page), 1, total) - 1;
    if (targetIdx === currentSheetIndex) return;
    currentSheetIndex = targetIdx;
    clearCopyFeedback();
    renderCurrentSheet();
    emitState();
  }

  function nextPage(): void {
    if (!workbook || workbook.SheetNames.length === 0) return;
    goToPage(currentSheetIndex + 2);
  }

  function prevPage(): void {
    if (!workbook || workbook.SheetNames.length === 0) return;
    goToPage(currentSheetIndex);
  }

  return {
    mount(container: HTMLElement): void {
      if (destroyed) {
        destroyed = false;
      }
      container.innerHTML = '';

      rootEl = document.createElement('div');
      rootEl.className = 'fp-excel-root';

      formulaBarEl = document.createElement('div');
      formulaBarEl.className = 'fp-excel-formulabar';

      const formulaLeft = document.createElement('div');
      formulaLeft.className = 'fp-excel-formulabar-left';

      cellRefEl = document.createElement('span');
      cellRefEl.className = 'fp-excel-cell-ref';
      cellRefEl.textContent = '—';

      sheetStatsEl = document.createElement('span');
      sheetStatsEl.className = 'fp-excel-sheet-stats';
      sheetStatsEl.textContent = tr('0 行 × 0 列', '0 rows × 0 cols');

      formulaLeft.append(cellRefEl, sheetStatsEl);

      formulaValueEl = document.createElement('div');
      formulaValueEl.className = 'fp-excel-formula-value is-placeholder';
      formulaValueEl.textContent = tr(
        '点击单元格查看完整内容或公式',
        'Click a cell to inspect its value or formula',
      );

      copyBtnEl = document.createElement('button');
      copyBtnEl.type = 'button';
      copyBtnEl.className = 'fp-btn fp-btn-ghost fp-excel-copy-btn';
      copyBtnEl.textContent = tr('复制', 'Copy');
      copyBtnEl.disabled = true;
      copyBtnEl.addEventListener('click', () => {
        void handleCopyClick();
      });

      stopLocaleSync?.();
      stopLocaleSync = onLocaleChange(() => {
        if (destroyed) return;
        if (copyFeedbackTimer === null && copyBtnEl) {
          copyBtnEl.textContent = tr('复制', 'Copy');
        }
        if (workbook) {
          renderCurrentSheet(true);
        } else {
          updateFormulaBar();
        }
      });

      formulaBarEl.append(formulaLeft, formulaValueEl, copyBtnEl);

      bannerEl = document.createElement('div');
      bannerEl.className = 'fp-excel-truncation-banner';
      bannerEl.hidden = true;

      viewportEl = document.createElement('div');
      viewportEl.className = 'fp-excel-viewport';
      viewportEl.addEventListener('click', handleViewportClick);

      sheetBarEl = document.createElement('div');
      sheetBarEl.className = 'fp-excel-sheetbar';
      sheetBarEl.setAttribute('role', 'tablist');
      sheetBarEl.hidden = true;
      sheetBarEl.addEventListener('click', handleSheetBarClick);

      rootEl.append(formulaBarEl, bannerEl, viewportEl, sheetBarEl);
      container.appendChild(rootEl);
      window.addEventListener('keydown', handleKeyDown);
    },

    async loadBuffer(buffer: ArrayBuffer, meta: { name: string; path: string; ext: string }): Promise<void> {
      lastBuffer = buffer;
      lastMeta = meta;
      if (destroyed) return;

      workbook = parseWorkbook(buffer, meta);
      currentSheetIndex = 0;
      selectedCell = null;
      expandedRowsLimitBySheet.clear();
      expandedColsLimitBySheet.clear();
      searchQuery = '';
      searchMatches = [];
      activeMatchIndex = -1;
      clearCopyFeedback();

      renderCurrentSheet();
      emitState();
    },

    async reloadBuffer(buffer: ArrayBuffer, meta: { name: string; path: string; ext: string }): Promise<void> {
      lastBuffer = buffer;
      lastMeta = meta;
      await this.refresh?.();
    },

    async refresh(): Promise<void> {
      if (!lastBuffer || !lastMeta || destroyed) return;

      const prevSheetCount = workbook?.SheetNames.length ?? 0;
      const prevSheetName = workbook?.SheetNames[currentSheetIndex];
      const prevSelected = selectedCell ? { ...selectedCell } : null;

      workbook = parseWorkbook(lastBuffer, lastMeta);
      const nextSheetNames = workbook.SheetNames;

      if (nextSheetNames.length === 0) {
        currentSheetIndex = 0;
        selectedCell = null;
      } else if (prevSheetName && nextSheetNames.includes(prevSheetName)) {
        currentSheetIndex = nextSheetNames.indexOf(prevSheetName);
      } else if (prevSheetCount > 0) {
        currentSheetIndex = clamp(currentSheetIndex, 0, nextSheetNames.length - 1);
      } else {
        currentSheetIndex = 0;
      }

      if (nextSheetNames.length > 0 && prevSelected && prevSelected.sheetIndex === currentSheetIndex) {
        const activeSheet = workbook.Sheets[nextSheetNames[currentSheetIndex]];
        const bounds = getEffectiveSheetBounds(activeSheet);
        if (prevSelected.r < bounds.totalRows && prevSelected.c < bounds.totalCols) {
          selectedCell = {
            ...prevSelected,
            sheetIndex: currentSheetIndex,
          };
        } else {
          selectedCell = null;
        }
      } else {
        selectedCell = null;
      }

      if (searchQuery) {
        const queryLower = searchQuery.toLowerCase();
        const found: SearchMatchItem[] = [];
        for (let i = 0; i < nextSheetNames.length; i++) {
          const sheet = workbook.Sheets[nextSheetNames[i]];
          if (sheet) {
            collectSheetMatches(sheet, i, queryLower, found);
          }
        }
        searchMatches = found;
        activeMatchIndex = found.length > 0 ? clamp(activeMatchIndex, 0, found.length - 1) : -1;
      }

      clearCopyFeedback();
      renderCurrentSheet(true);
      if (selectedCell && selectedCell.sheetIndex === currentSheetIndex) {
        // 同步更新刷新后的单元格选中区间与公式栏内容
        selectCellAt(selectedCell.r, selectedCell.c, false);
      }
      emitState();
    },

    setZoom(zoom: number | 'fit'): void {
      if (zoom === 'fit') {
        fitMode = true;
        currentZoom = 1.0;
      } else if (Number.isFinite(zoom)) {
        fitMode = false;
        currentZoom = clamp(roundZoom(zoom), MIN_ZOOM, MAX_ZOOM);
      } else {
        return;
      }
      applyZoomToDOM();
      emitState();
    },

    nextPage,
    prevPage,
    goToPage,

    async search(query: string): Promise<{ total: number; current: number }> {
      const trimmed = query.trim();
      searchQuery = trimmed;
      searchMatches = [];
      activeMatchIndex = -1;

      if (!trimmed || !workbook || workbook.SheetNames.length === 0) {
        applySearchHighlightsToCurrentSheet();
        emitState();
        return { total: 0, current: 0 };
      }

      const queryLower = trimmed.toLowerCase();
      const found: SearchMatchItem[] = [];
      for (let i = 0; i < workbook.SheetNames.length; i++) {
        const sheetName = workbook.SheetNames[i];
        const sheet = workbook.Sheets[sheetName];
        if (sheet) {
          collectSheetMatches(sheet, i, queryLower, found);
        }
      }

      searchMatches = found;
      if (found.length === 0) {
        applySearchHighlightsToCurrentSheet();
        emitState();
        return { total: 0, current: 0 };
      }

      activeMatchIndex = 0;
      focusActiveMatch();
      emitState();
      return { total: searchMatches.length, current: 1 };
    },

    nextMatch(): void {
      if (searchMatches.length === 0) return;
      activeMatchIndex = (activeMatchIndex + 1) % searchMatches.length;
      focusActiveMatch();
      emitState();
    },

    prevMatch(): void {
      if (searchMatches.length === 0) return;
      activeMatchIndex = (activeMatchIndex - 1 + searchMatches.length) % searchMatches.length;
      focusActiveMatch();
      emitState();
    },

    onStateChange(cb: (info: PreviewStateInfo) => void): void {
      stateListeners.add(cb);
    },

    destroy(): void {
      destroyed = true;
      stopLocaleSync?.();
      stopLocaleSync = null;
      clearCopyFeedback();
      window.removeEventListener('keydown', handleKeyDown);

      if (viewportEl) {
        viewportEl.removeEventListener('click', handleViewportClick);
      }
      if (sheetBarEl) {
        sheetBarEl.removeEventListener('click', handleSheetBarClick);
      }
      if (rootEl) {
        rootEl.remove();
      }

      resetGridRefs();
      rootEl = null;
      formulaBarEl = null;
      cellRefEl = null;
      sheetStatsEl = null;
      formulaValueEl = null;
      copyBtnEl = null;
      bannerEl = null;
      viewportEl = null;
      sheetBarEl = null;

      lastBuffer = null;
      lastMeta = null;
      workbook = null;
      currentSheetIndex = 0;
      currentRowCount = 0;
      currentColCount = 0;
      selectedCell = null;
      searchQuery = '';
      searchMatches = [];
      activeMatchIndex = -1;
      expandedRowsLimitBySheet.clear();
      expandedColsLimitBySheet.clear();
      stateListeners.clear();
    },
  };
}
