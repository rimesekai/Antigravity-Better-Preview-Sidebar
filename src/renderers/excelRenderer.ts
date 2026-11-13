import * as XLSX from 'xlsx';
import type { PreviewRenderer, PreviewStateInfo } from '../types/preview';

const MIN_ZOOM = 0.5;
const MAX_ZOOM = 2.0;
const MAX_INITIAL_ROWS = 2000;
const MAX_COLS = 100;
const DEFAULT_COL_WIDTH_PX = 110;
const MIN_COL_WIDTH_PX = 64;
const MAX_COL_WIDTH_PX = 360;
const ROW_HEADER_WIDTH_PX = 52;

interface MergeMasterInfo {
  rowSpan: number;
  colSpan: number;
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
  addr: string;
  displayRef: string;
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
  try {
    const formatted = XLSX.utils.format_cell(cell);
    if (formatted !== undefined && formatted !== null && formatted !== '') {
      return String(formatted);
    }
  } catch {
    // 忽略格式化异常并回退到原始字段
  }
  if (cell.w !== undefined && cell.w !== null) {
    return String(cell.w);
  }
  if (cell.v !== undefined && cell.v !== null) {
    if (cell.v instanceof Date) {
      return Number.isNaN(cell.v.getTime()) ? '' : cell.v.toLocaleString();
    }
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
  let mergeRangeByKey = new Map<string, string>();

  let lastBuffer: ArrayBuffer | null = null;
  let lastMeta: { name: string; path: string; ext: string } | null = null;
  let workbook: XLSX.WorkBook | null = null;

  let currentSheetIndex = 0;
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

  const expandedRowsSheets = new Set<string>();
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
      copyBtnEl.classList.remove('is-copied');
      copyBtnEl.textContent = '复制';
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
      sheetStatsEl.textContent = '0 行 × 0 列';
      formulaValueEl.textContent = '当前工作簿无可用工作表';
      formulaValueEl.classList.add('is-placeholder');
      copyBtnEl.disabled = true;
      copyFeedbackText = '';
      return;
    }

    sheetStatsEl.textContent = `${currentRowCount.toLocaleString()} 行 × ${currentColCount} 列`;

    if (!selectedCell || selectedCell.sheetIndex !== currentSheetIndex) {
      cellRefEl.textContent = '—';
      formulaValueEl.textContent =
        currentRowCount === 0 && currentColCount === 0
          ? '当前工作表暂无单元格数据'
          : '点击单元格查看完整内容或公式';
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
      const fullText = `=${cell.f}  (值: ${displayValue})`;
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
      formulaValueEl.textContent = '（空单元格）';
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
        selectedCell.c === idx;
      th.classList.toggle('is-active-header', active);
    });
    rowHeaderEls.forEach((th, idx) => {
      const active =
        selectedCell !== null &&
        selectedCell.sheetIndex === currentSheetIndex &&
        selectedCell.r === idx;
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
    const mergeRef = mergeRangeByKey.get(key);
    selectedCell = {
      sheetIndex: currentSheetIndex,
      r,
      c,
      addr,
      displayRef: mergeRef ?? addr,
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
    rowsTruncated: boolean,
    colsTruncated: boolean,
  ): void {
    if (!bannerEl) return;
    if (!rowsTruncated && !colsTruncated) {
      bannerEl.hidden = true;
      bannerEl.innerHTML = '';
      return;
    }

    bannerEl.hidden = false;
    bannerEl.innerHTML = '';

    const msgEl = document.createElement('span');
    msgEl.className = 'fp-excel-truncation-text';
    if (rowsTruncated && colsTruncated) {
      msgEl.textContent = `当前工作表共 ${totalRows.toLocaleString()} 行 × ${totalCols} 列，已渲染前 ${renderRows.toLocaleString()} 行（最多 ${MAX_COLS} 列）`;
    } else if (rowsTruncated) {
      msgEl.textContent = `当前工作表共 ${totalRows.toLocaleString()} 行 × ${totalCols} 列，已渲染前 ${renderRows.toLocaleString()} 行`;
    } else {
      msgEl.textContent = `当前工作表共 ${totalRows.toLocaleString()} 行 × ${totalCols} 列，已渲染前 ${renderCols} 列（最多 ${MAX_COLS} 列）`;
    }
    bannerEl.appendChild(msgEl);

    if (rowsTruncated) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'fp-btn fp-btn-secondary fp-excel-load-all-btn';
      btn.textContent = '加载全部行';
      btn.addEventListener('click', () => {
        if (destroyed) return;
        expandedRowsSheets.add(sheetName);
        renderCurrentSheet();
        emitState();
      });
      bannerEl.appendChild(btn);
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
    mergeRangeByKey.clear();
  }

  function renderCurrentSheet(): void {
    if (destroyed || !viewportEl) return;
    resetGridRefs();

    const sheetNames = workbook?.SheetNames ?? [];
    if (sheetNames.length === 0) {
      currentRowCount = 0;
      currentColCount = 0;
      if (bannerEl) {
        bannerEl.hidden = true;
        bannerEl.innerHTML = '';
      }
      renderEmptyState('工作簿为空', '当前文件不包含任何工作表');
      renderSheetBar();
      updateFormulaBar();
      return;
    }

    currentSheetIndex = clamp(currentSheetIndex, 0, sheetNames.length - 1);
    const sheetName = sheetNames[currentSheetIndex];
    const sheet = workbook?.Sheets[sheetName];

    renderSheetBar();

    if (!sheet || !sheet['!ref']) {
      currentRowCount = 0;
      currentColCount = 0;
      if (bannerEl) {
        bannerEl.hidden = true;
        bannerEl.innerHTML = '';
      }
      renderEmptyState('空工作表', `工作表「${sheetName}」暂无单元格数据`);
      updateFormulaBar();
      return;
    }

    const range = XLSX.utils.decode_range(sheet['!ref']);
    const totalRows = Math.max(0, range.e.r + 1);
    const totalCols = Math.max(0, range.e.c + 1);
    currentRowCount = totalRows;
    currentColCount = totalCols;

    if (totalRows === 0 || totalCols === 0) {
      if (bannerEl) {
        bannerEl.hidden = true;
        bannerEl.innerHTML = '';
      }
      renderEmptyState('空工作表', `工作表「${sheetName}」暂无单元格数据`);
      updateFormulaBar();
      return;
    }

    const isExpanded = expandedRowsSheets.has(sheetName);
    const renderRows = isExpanded ? totalRows : Math.min(totalRows, MAX_INITIAL_ROWS);
    const renderCols = Math.min(totalCols, MAX_COLS);
    const rowsTruncated = !isExpanded && totalRows > MAX_INITIAL_ROWS;
    const colsTruncated = totalCols > MAX_COLS;

    renderBanner(sheetName, totalRows, totalCols, renderRows, renderCols, rowsTruncated, colsTruncated);

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
      mergeRangeByKey.set(masterKey, rangeRef);

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
    viewportEl.scrollTop = 0;
    viewportEl.scrollLeft = 0;

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

    if (match.r >= MAX_INITIAL_ROWS && targetSheetName && !expandedRowsSheets.has(targetSheetName)) {
      expandedRowsSheets.add(targetSheetName);
      if (match.sheetIndex === currentSheetIndex) {
        needRerender = true;
      }
    }

    if (match.sheetIndex !== currentSheetIndex) {
      currentSheetIndex = match.sheetIndex;
      renderCurrentSheet();
    } else if (needRerender) {
      renderCurrentSheet();
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
    if (!sheet['!ref']) return;
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const startR = Math.max(0, range.s.r);
    const startC = Math.max(0, range.s.c);
    const endR = Math.max(startR, range.e.r);
    const endC = Math.max(startC, range.e.c);
    const totalCells = (endR - startR + 1) * (endC - startC + 1);

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
        coords.push({ r: pos.r, c: pos.c, addr: key });
      }
      coords.sort((a, b) => (a.r !== b.r ? a.r - b.r : a.c - b.c));
      for (const item of coords) {
        checkCell(item.r, item.c, item.addr, sheet[item.addr] as XLSX.CellObject | undefined);
      }
      return;
    }

    for (let r = startR; r <= endR; r++) {
      for (let c = startC; c <= endC; c++) {
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
        cellDates: true,
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
      cellDates: true,
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

  async function handleCopyClick(): Promise<void> {
    if (!copyFeedbackText || !copyBtnEl) return;
    try {
      if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(copyFeedbackText);
        if (destroyed || !copyBtnEl) return;
        copyBtnEl.classList.add('is-copied');
        copyBtnEl.textContent = '已复制';
        if (copyFeedbackTimer !== null) {
          window.clearTimeout(copyFeedbackTimer);
        }
        copyFeedbackTimer = window.setTimeout(() => {
          copyFeedbackTimer = null;
          if (!copyBtnEl || destroyed) return;
          copyBtnEl.classList.remove('is-copied');
          copyBtnEl.textContent = '复制';
        }, 1500);
      }
    } catch (err) {
      console.warn('[excelRenderer] 复制到剪贴板失败:', err);
    }
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
      sheetStatsEl.textContent = '0 行 × 0 列';

      formulaLeft.append(cellRefEl, sheetStatsEl);

      formulaValueEl = document.createElement('div');
      formulaValueEl.className = 'fp-excel-formula-value is-placeholder';
      formulaValueEl.textContent = '点击单元格查看完整内容或公式';

      copyBtnEl = document.createElement('button');
      copyBtnEl.type = 'button';
      copyBtnEl.className = 'fp-btn fp-btn-ghost fp-excel-copy-btn';
      copyBtnEl.textContent = '复制';
      copyBtnEl.disabled = true;
      copyBtnEl.addEventListener('click', () => {
        void handleCopyClick();
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
    },

    async loadBuffer(buffer: ArrayBuffer, meta: { name: string; path: string; ext: string }): Promise<void> {
      lastBuffer = buffer;
      lastMeta = meta;
      if (destroyed) return;

      workbook = parseWorkbook(buffer, meta);
      currentSheetIndex = 0;
      selectedCell = null;
      expandedRowsSheets.clear();
      searchQuery = '';
      searchMatches = [];
      activeMatchIndex = -1;
      clearCopyFeedback();

      renderCurrentSheet();
      emitState();
    },

    async refresh(): Promise<void> {
      if (!lastBuffer || !lastMeta || destroyed) return;

      const prevSheetCount = workbook?.SheetNames.length ?? 0;
      const prevSheetName = workbook?.SheetNames[currentSheetIndex];

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
      renderCurrentSheet();
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
      clearCopyFeedback();

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
      expandedRowsSheets.clear();
      stateListeners.clear();
    },
  };
}
