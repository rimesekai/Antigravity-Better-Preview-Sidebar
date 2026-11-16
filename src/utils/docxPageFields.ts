export const DOCX_PAGE_FIELD_TOKEN = '\uE001FP_FIELD:PAGE\uE001';
export const DOCX_NUMPAGES_FIELD_TOKEN = '\uE001FP_FIELD:NUMPAGES\uE001';
export const DOCX_SECTIONPAGES_FIELD_TOKEN = '\uE001FP_FIELD:SECTIONPAGES\uE001';
export const DOCX_SECTION_END_CLASS_PREFIX = 'fp-sect-end-';
export const DOCX_TAB_PARA_CLASS_PREFIX = 'fp-tabs-';
const EMPTY_FIRST_HEADER_FOOTER_REL_ID = '__fp_empty_first__';

export type PageFieldKind = 'PAGE' | 'NUMPAGES' | 'SECTIONPAGES';

export interface DocxSectionPageConfig {
  format?: string;
  start?: number;
}

export interface DocxTabStop {
  position: string;
  leader?: string;
  style?: string;
}

export interface DocxTabStopContext {
  tabParagraphs: Map<string, DocxTabStop[]>;
  defaultTabStop?: string;
}

interface ComplexFieldFrame {
  instr: string;
  hasSeparate: boolean;
  resultRunIndices: number[];
}

const PAGE_FIELD_REGEX = /^\s*(PAGE|NUMPAGES|SECTIONPAGES)(?:\s|\\|$)/i;

const FIELD_TOKEN_MAP: Record<PageFieldKind, string> = {
  PAGE: DOCX_PAGE_FIELD_TOKEN,
  NUMPAGES: DOCX_NUMPAGES_FIELD_TOKEN,
  SECTIONPAGES: DOCX_SECTIONPAGES_FIELD_TOKEN,
};

/**
 * 解析 Word 域指令中的页码域类型，非页码域返回 null。
 */
export function parsePageFieldKind(instruction?: string | null): PageFieldKind | null {
  if (!instruction) return null;
  const match = PAGE_FIELD_REGEX.exec(instruction);
  if (!match) return null;
  const upper = match[1].toUpperCase();
  if (upper === 'PAGE' || upper === 'NUMPAGES' || upper === 'SECTIONPAGES') {
    return upper;
  }
  return null;
}

function toRoman(num: number, upper: boolean): string {
  const table: Array<[number, string]> = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ];
  let remaining = Math.floor(num);
  let res = '';
  for (const [val, sym] of table) {
    while (remaining >= val) {
      res += sym;
      remaining -= val;
    }
  }
  return upper ? res : res.toLowerCase();
}

function toLetter(num: number, upper: boolean): string {
  let n = Math.floor(num);
  let res = '';
  while (n > 0) {
    n -= 1;
    res = String.fromCharCode((upper ? 65 : 97) + (n % 26)) + res;
    n = Math.floor(n / 26);
  }
  return res || String(num);
}

/**
 * 按 Word w:pgNumType/@w:fmt 格式规范格式化页码数字。
 */
export function formatPageNumber(num: number, format?: string): string {
  if (!Number.isFinite(num) || num <= 0) {
    return String(num);
  }
  switch (format) {
    case 'upperRoman':
      return toRoman(num, true);
    case 'lowerRoman':
      return toRoman(num, false);
    case 'upperLetter':
      return toLetter(num, true);
    case 'lowerLetter':
      return toLetter(num, false);
    case 'decimal':
    default:
      return String(Math.floor(num));
  }
}

/**
 * 原地重写单个段落 children 中的页码域（SimpleField 与 ComplexField），
 * 用带哨兵的文本占位符替换静态缓存数字（避免 docx-preview 每页重复渲染同一缓存值）。
 */
export function rewriteParagraphPageFields(children: unknown[]): boolean {
  if (!Array.isArray(children) || children.length === 0) return false;

  let changed = false;
  const fieldStack: ComplexFieldFrame[] = [];
  const nextChildren: Array<unknown | null> = [...children];

  for (let i = 0; i < nextChildren.length; i += 1) {
    const child = nextChildren[i] as Record<string, any> | null;
    if (!child || typeof child !== 'object') continue;

    if (child.type === 'simpleField') {
      const kind = parsePageFieldKind(child.instruction);
      if (kind) {
        nextChildren[i] = {
          type: 'run',
          children: [{ type: 'text', text: FIELD_TOKEN_MAP[kind] }],
        };
        changed = true;
      }
      continue;
    }

    if (child.type !== 'run' || !Array.isArray(child.children)) {
      continue;
    }

    let hasFieldMarker = Boolean(child.fieldRun);
    for (const sub of child.children as Array<Record<string, any>>) {
      if (!sub || typeof sub !== 'object') continue;
      if (sub.type === 'complexField') {
        hasFieldMarker = true;
        if (sub.charType === 'begin') {
          fieldStack.push({ instr: '', hasSeparate: false, resultRunIndices: [] });
        } else if (sub.charType === 'separate') {
          if (fieldStack.length > 0) {
            fieldStack[fieldStack.length - 1].hasSeparate = true;
          }
        } else if (sub.charType === 'end') {
          const frame = fieldStack.pop();
          if (frame) {
            const kind = parsePageFieldKind(frame.instr);
            if (kind) {
              const token = FIELD_TOKEN_MAP[kind];
              if (frame.resultRunIndices.length > 0) {
                const firstIdx = frame.resultRunIndices[0];
                const firstRun = nextChildren[firstIdx] as Record<string, any> | null;
                if (firstRun && typeof firstRun === 'object') {
                  firstRun.children = [{ type: 'text', text: token }];
                }
                for (let k = 1; k < frame.resultRunIndices.length; k += 1) {
                  nextChildren[frame.resultRunIndices[k]] = null;
                }
              } else {
                // 无缓存结果 run 时，直接将 end 占位替换为可渲染的普通 run
                nextChildren[i] = {
                  type: 'run',
                  children: [{ type: 'text', text: token }],
                };
              }
              changed = true;
            }
          }
        }
      } else if (sub.type === 'instruction') {
        hasFieldMarker = true;
        if (fieldStack.length > 0) {
          const top = fieldStack[fieldStack.length - 1];
          if (!top.hasSeparate && typeof sub.text === 'string') {
            top.instr += sub.text;
          }
        }
      }
    }

    if (!hasFieldMarker && fieldStack.length > 0) {
      const top = fieldStack[fieldStack.length - 1];
      if (top.hasSeparate) {
        top.resultRunIndices.push(i);
      }
    }
  }

  if (changed) {
    children.length = 0;
    for (const item of nextChildren) {
      if (item !== null) {
        children.push(item);
      }
    }
  }
  return changed;
}

function walkAstForParagraphs(root: unknown, visited: WeakSet<object>): void {
  if (!root || typeof root !== 'object') return;
  if (visited.has(root)) return;
  visited.add(root);

  if (Array.isArray(root)) {
    for (const item of root) {
      walkAstForParagraphs(item, visited);
    }
    return;
  }

  const node = root as Record<string, any>;
  if (node.type === 'paragraph' && Array.isArray(node.children)) {
    rewriteParagraphPageFields(node.children);
  }
  if (Array.isArray(node.children)) {
    for (const child of node.children) {
      walkAstForParagraphs(child, visited);
    }
  }
}

/**
 * 修复 docx-preview 的 titlePage（首页不同）回退缺陷：
 * 当某节开启 titlePage 但未提供 type === 'first' 的页眉/页脚引用时，
 * docx-preview 会错误地回退到 default 页眉/页脚，导致封面首页多出页码。
 * 注入一个不存在的 first relId 可使 findPartByRelId 返回 undefined，从而跳过首页渲染。
 */
export function patchTitlePageMissingFirstRefs(sectProps: any): void {
  if (!sectProps || typeof sectProps !== 'object' || !sectProps.titlePage) return;

  for (const key of ['footerRefs', 'headerRefs'] as const) {
    const refs = sectProps[key];
    if (Array.isArray(refs) && refs.length > 0 && !refs.some((r: any) => r?.type === 'first')) {
      refs.push({ id: EMPTY_FIRST_HEADER_FOOTER_REL_ID, type: 'first' });
    }
  }
}

function paragraphHasPageBreak(para: Record<string, any>): boolean {
  if (!Array.isArray(para.children)) return false;
  for (const r of para.children) {
    if (!r || typeof r !== 'object' || !Array.isArray(r.children)) continue;
    for (const sub of r.children) {
      if (
        sub &&
        typeof sub === 'object' &&
        sub.type === 'break' &&
        (sub.break === 'page' || sub.break === 'lastRenderedPageBreak')
      ) {
        return true;
      }
    }
  }
  return false;
}

/**
 * 确保非连续分节符（nextPage/evenPage/oddPage/默认）在 docx-preview 中触发分页：
 * docx-preview 的 splitBySection 仅在检测到 break 时才会设置 current.pageBreak = true，
 * 否则相邻同尺寸分节会被合并到同一页，导致首页错误继承下一节的页脚。
 */
export function ensureSectionBreakPageBreak(para: Record<string, any>): void {
  if (!para || para.type !== 'paragraph' || !para.sectionProps) return;
  if (para.sectionProps.type === 'continuous') return;
  if (paragraphHasPageBreak(para)) return;

  if (!Array.isArray(para.children)) {
    para.children = [];
  }
  para.children.push({
    type: 'run',
    children: [{ type: 'break', break: 'page' }],
  });
}

/**
 * 若紧跟在分节符之后的段落开头存在前导换页符（常见于 Word 在新节首段写入的 <w:lastRenderedPageBreak/>），
 * 且该换页符之前没有任何可见内容，则移除该前导换页符，避免在分节分页之后再多出一页空白页。
 */
export function stripLeadingPageBreakIfEmptyBefore(para: Record<string, any>): void {
  if (!para || para.type !== 'paragraph' || !Array.isArray(para.children)) return;

  for (let rIdx = 0; rIdx < para.children.length; rIdx += 1) {
    const child = para.children[rIdx];
    if (!child || typeof child !== 'object') continue;
    if (child.type !== 'run') {
      // 遇到非 run 节点（如超链接、简单域等），停止检查
      return;
    }
    if (!Array.isArray(child.children)) continue;

    for (let sIdx = 0; sIdx < child.children.length; sIdx += 1) {
      const sub = child.children[sIdx];
      if (!sub || typeof sub !== 'object') continue;
      if (
        sub.type === 'break' &&
        (sub.break === 'lastRenderedPageBreak' || sub.break === 'page')
      ) {
        child.children.splice(sIdx, 1);
        if (child.children.length === 0) {
          para.children.splice(rIdx, 1);
        }
        return;
      }
      if (sub.type === 'text' && typeof sub.text === 'string' && sub.text.length === 0) {
        continue;
      }
      if (sub.type === 'complexField' || sub.type === 'instruction') {
        continue;
      }
      // 遇到任何实际可见内容，不能移除后续换页符
      return;
    }
  }
}

/**
 * 在 docx-preview parseAsync 之后、renderDocument 之前预处理页码域并采集分节页码配置。
 */
export function preprocessDocxPageFields(doc: any): DocxSectionPageConfig[] {
  if (!doc || typeof doc !== 'object') return [];

  const visited = new WeakSet<object>();
  if (Array.isArray(doc.parts)) {
    for (const part of doc.parts) {
      if (part?.rootElement) {
        walkAstForParagraphs(part.rootElement, visited);
      }
    }
  }
  if (doc.documentPart?.body) {
    walkAstForParagraphs(doc.documentPart.body, visited);
  }

  const sectionConfigs: DocxSectionPageConfig[] = [];
  const body = doc.documentPart?.body;
  if (body && Array.isArray(body.children)) {
    for (let i = 0; i < body.children.length; i += 1) {
      const child = body.children[i];
      if (child?.type === 'paragraph' && child.sectionProps) {
        patchTitlePageMissingFirstRefs(child.sectionProps);
        ensureSectionBreakPageBreak(child);
        if (child.sectionProps.type !== 'continuous' && i + 1 < body.children.length) {
          stripLeadingPageBreakIfEmptyBefore(body.children[i + 1]);
        }
        const sectIdx = sectionConfigs.length;
        const markerClass = `${DOCX_SECTION_END_CLASS_PREFIX}${sectIdx}`;
        child.className = child.className ? `${child.className} ${markerClass}` : markerClass;
        sectionConfigs.push({
          format: child.sectionProps.pageNumber?.format,
          start: child.sectionProps.pageNumber?.start,
        });
      }
    }
  }

  if (body?.props) {
    patchTitlePageMissingFirstRefs(body.props);
    sectionConfigs.push({
      format: body.props.pageNumber?.format,
      start: body.props.pageNumber?.start,
    });
  }

  return sectionConfigs;
}

function findMaxEndedSectionIndex(sectionEl: HTMLElement): number | null {
  const markers = sectionEl.querySelectorAll(`[class*="${DOCX_SECTION_END_CLASS_PREFIX}"]`);
  let maxIdx: number | null = null;
  for (const el of markers) {
    for (const cls of el.classList) {
      if (!cls.startsWith(DOCX_SECTION_END_CLASS_PREFIX)) continue;
      const parsed = Number.parseInt(cls.slice(DOCX_SECTION_END_CLASS_PREFIX.length), 10);
      if (Number.isFinite(parsed) && (maxIdx === null || parsed > maxIdx)) {
        maxIdx = parsed;
      }
    }
  }
  return maxIdx;
}

/**
 * 在 DOM 挂载到 stageEl 之后，将各页中的页码域占位符替换为该页实际计算出的页码/总页数。
 */
export function resolveDocxPageFields(
  stageEl: HTMLElement,
  sectionConfigs: DocxSectionPageConfig[],
): void {
  const sections = Array.from(stageEl.querySelectorAll<HTMLElement>('.docx-wrapper > section.docx'));
  const targetSections =
    sections.length > 0 ? sections : Array.from(stageEl.querySelectorAll<HTMLElement>('section.docx'));
  if (targetSections.length === 0) return;

  const totalPages = targetSections.length;
  const maxSectIdx = Math.max(0, sectionConfigs.length - 1);

  const pageSectionIndices: number[] = [];
  let currentSectIdx = 0;
  for (const secEl of targetSections) {
    const endedIdx = findMaxEndedSectionIndex(secEl);
    if (endedIdx !== null) {
      const clampedEnded = Math.min(endedIdx, maxSectIdx);
      pageSectionIndices.push(clampedEnded);
      currentSectIdx = Math.min(clampedEnded + 1, maxSectIdx);
    } else {
      pageSectionIndices.push(currentSectIdx);
    }
  }

  const sectionPageCounts = new Map<number, number>();
  for (const sectIdx of pageSectionIndices) {
    sectionPageCounts.set(sectIdx, (sectionPageCounts.get(sectIdx) ?? 0) + 1);
  }

  const pageNumbers: number[] = [];
  let runningPageNum = 1;
  let prevSectIdx = -1;
  for (let i = 0; i < totalPages; i += 1) {
    const sectIdx = pageSectionIndices[i];
    const cfg = sectionConfigs[sectIdx];
    if (sectIdx !== prevSectIdx && typeof cfg?.start === 'number' && Number.isFinite(cfg.start)) {
      runningPageNum = cfg.start;
    }
    pageNumbers.push(runningPageNum);
    runningPageNum += 1;
    prevSectIdx = sectIdx;
  }

  for (let i = 0; i < totalPages; i += 1) {
    const secEl = targetSections[i];
    const sectIdx = pageSectionIndices[i];
    const format = sectionConfigs[sectIdx]?.format;
    const pageText = formatPageNumber(pageNumbers[i], format);
    const numPagesText = formatPageNumber(totalPages, format);
    const sectionPagesText = formatPageNumber(sectionPageCounts.get(sectIdx) ?? totalPages, format);

    const walker = document.createTreeWalker(secEl, NodeFilter.SHOW_TEXT);
    let node = walker.nextNode();
    while (node) {
      const val = node.nodeValue;
      if (val && val.includes('\uE001FP_FIELD:')) {
        node.nodeValue = val
          .replaceAll(DOCX_PAGE_FIELD_TOKEN, pageText)
          .replaceAll(DOCX_NUMPAGES_FIELD_TOKEN, numPagesText)
          .replaceAll(DOCX_SECTIONPAGES_FIELD_TOKEN, sectionPagesText);
      }
      node = walker.nextNode();
    }
  }
}

function paragraphContainsTab(para: Record<string, any>): boolean {
  if (!Array.isArray(para.children)) return false;
  for (const child of para.children) {
    if (!child || typeof child !== 'object') continue;
    if (child.type === 'tab') return true;
    if (child.type === 'run' && Array.isArray(child.children)) {
      for (const sub of child.children) {
        if (sub && typeof sub === 'object' && sub.type === 'tab') {
          return true;
        }
      }
    }
  }
  return false;
}

/**
 * 收集文档及页眉/页脚中包含制表符（<w:tab/>）的段落及其生效的 tabs 定义，
 * 用于在 DOM 挂载到 stageEl 后重新计算制表位宽度（修复 docx-preview 挂载前计算失效导致的左偏问题）。
 */
export function collectDocxTabStops(doc: any): DocxTabStopContext {
  const tabParagraphs = new Map<string, DocxTabStop[]>();
  const defaultTabStop: string | undefined = doc?.settingsPart?.settings?.defaultTabStop;

  if (!doc || typeof doc !== 'object') {
    return { tabParagraphs, defaultTabStop };
  }

  const styleMap = new Map<string, any>();
  const styles = doc.stylesPart?.styles;
  if (Array.isArray(styles)) {
    for (const st of styles) {
      if (st?.id) {
        styleMap.set(st.id, st);
      }
    }
  }

  const resolveStyleTabs = (styleName?: string): DocxTabStop[] | undefined => {
    let currId = styleName;
    const seen = new Set<string>();
    while (currId && !seen.has(currId)) {
      seen.add(currId);
      const st = styleMap.get(currId);
      if (!st) break;
      if (Array.isArray(st.paragraphProps?.tabs) && st.paragraphProps.tabs.length > 0) {
        return st.paragraphProps.tabs;
      }
      currId = st.basedOn;
    }
    return undefined;
  };

  let nextId = 0;
  const visited = new WeakSet<object>();

  const walk = (root: unknown): void => {
    if (!root || typeof root !== 'object') return;
    if (visited.has(root)) return;
    visited.add(root);

    if (Array.isArray(root)) {
      for (const item of root) walk(item);
      return;
    }

    const node = root as Record<string, any>;
    if (node.type === 'paragraph' && paragraphContainsTab(node)) {
      const effectiveTabs =
        (Array.isArray(node.tabs) && node.tabs.length > 0 ? node.tabs : undefined) ??
        resolveStyleTabs(node.styleName) ??
        [];
      const markerClass = `${DOCX_TAB_PARA_CLASS_PREFIX}${nextId++}`;
      node.className = node.className ? `${node.className} ${markerClass}` : markerClass;
      tabParagraphs.set(markerClass, effectiveTabs);
    }

    if (Array.isArray(node.children)) {
      for (const child of node.children) walk(child);
    }
  };

  if (Array.isArray(doc.parts)) {
    for (const part of doc.parts) {
      if (part?.rootElement) walk(part.rootElement);
    }
  }
  if (doc.documentPart?.body) {
    walk(doc.documentPart.body);
  }

  return { tabParagraphs, defaultTabStop };
}

function measurePixelToPoint(container: HTMLElement): number {
  const probe = document.createElement('div');
  probe.style.position = 'absolute';
  probe.style.visibility = 'hidden';
  probe.style.width = '100pt';
  container.appendChild(probe);
  const px = probe.offsetWidth;
  probe.remove();
  return px > 0 ? 100 / px : 72 / 96;
}

/**
 * 在 DOM 挂载到 stageEl 且页码占位符已替换后，重新计算制表位宽度。
 */
export function recalculateDocxTabStops(
  stageEl: HTMLElement,
  tabCtx: DocxTabStopContext,
): void {
  if (tabCtx.tabParagraphs.size === 0) return;

  const pixelToPoint = measurePixelToPoint(stageEl);
  const defaultSizePt = Number.parseFloat(tabCtx.defaultTabStop ?? '36pt') || 36;
  const maxTabs = 50;

  for (const [markerClass, rawTabs] of tabCtx.tabParagraphs) {
    const paragraphs = stageEl.querySelectorAll<HTMLElement>(`p.${markerClass}`);
    for (const p of paragraphs) {
      const stopEls = Array.from(p.querySelectorAll<HTMLElement>('.docx-tab-stop'));
      if (stopEls.length === 0) continue;

      for (const el of stopEls) {
        el.innerHTML = '\u00a0';
        el.style.wordSpacing = '';
      }

      const pbb = p.getBoundingClientRect();
      if (pbb.width <= 0) continue;
      const pcs = getComputedStyle(p);
      const marginLeft = Number.parseFloat(pcs.marginLeft) || 0;
      const pOffset = pbb.left + marginLeft;
      const pWidthPt = pbb.width * pixelToPoint;

      const tabStops =
        rawTabs.length > 0
          ? rawTabs
              .map((t) => ({
                pos: Number.parseFloat(t.position) || 0,
                leader: t.leader ?? 'none',
                style: t.style ?? 'left',
              }))
              .sort((a, b) => a.pos - b.pos)
          : [{ pos: 0, leader: 'none', style: 'left' }];

      const lastTab = tabStops[tabStops.length - 1];
      for (let pos = lastTab.pos + defaultSizePt; pos < pWidthPt && tabStops.length < maxTabs; pos += defaultSizePt) {
        tabStops.push({ pos, leader: 'none', style: 'left' });
      }

      for (let idx = 0; idx < stopEls.length; idx += 1) {
        const elem = stopEls[idx];
        const ebb = elem.getBoundingClientRect();
        const left = (ebb.left - pOffset) * pixelToPoint;
        const tab = tabStops.find((t) => t.style !== 'clear' && t.pos > left);
        if (!tab) continue;

        let width = 1;
        if (tab.style === 'right' || tab.style === 'center') {
          const range = document.createRange();
          range.setStartAfter(elem);
          if (idx + 1 < stopEls.length) {
            range.setEndBefore(stopEls[idx + 1]);
          } else {
            range.setEndAfter(p);
          }
          const mul = tab.style === 'center' ? 0.5 : 1;
          const nextBB = range.getBoundingClientRect();
          const offset = nextBB.left + mul * nextBB.width - pOffset;
          width = tab.pos - offset * pixelToPoint;
        } else {
          width = tab.pos - left;
        }

        if (Number.isFinite(width) && width > 0) {
          elem.style.wordSpacing = `${width.toFixed(1)}pt`;
        }
      }
    }
  }
}

type FrameXAlign = 'left' | 'center' | 'right';

function normalizeFrameXAlign(val?: string | null): FrameXAlign | null {
  if (!val) return null;
  const lower = val.toLowerCase();
  if (lower === 'center' || lower === 'right' || lower === 'left') {
    return lower;
  }
  return null;
}

/**
 * 从页眉/页脚 XML 字符串中按段落顺序提取 <w:framePr w:xAlign="..."/> 水平对齐方式。
 * docx-preview 未解析 <w:framePr>，导致 Word 中通过图文框居中的页码退化为左对齐。
 */
export function extractFrameAlignmentsFromXml(xmlText: string): Array<FrameXAlign | null> {
  if (!xmlText) return [];
  const matches = xmlText.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>|<w:p\s*\/>/g);
  if (!matches) return [];

  return matches.map((pXml) => {
    const pPrMatch = /<w:pPr(?:\s[^>]*)?>([\s\S]*?)<\/w:pPr>/.exec(pXml);
    if (!pPrMatch) return null;
    const frameMatch = /<w:framePr\b([^>]*)\/?>/.exec(pPrMatch[1]);
    if (!frameMatch) return null;
    const xAlignMatch = /\bw:xAlign="([^"]+)"/.exec(frameMatch[1]);
    return normalizeFrameXAlign(xAlignMatch?.[1]);
  });
}

function collectAstParagraphs(root: unknown, out: Array<Record<string, any>>, visited: WeakSet<object>): void {
  if (!root || typeof root !== 'object' || visited.has(root)) return;
  visited.add(root);

  if (Array.isArray(root)) {
    for (const item of root) collectAstParagraphs(item, out, visited);
    return;
  }

  const node = root as Record<string, any>;
  if (node.type === 'paragraph') {
    out.push(node);
  }
  if (Array.isArray(node.children)) {
    for (const child of node.children) collectAstParagraphs(child, out, visited);
  }
}

export function applyFrameAlignmentsToParagraphs(
  paragraphs: Array<Record<string, any>>,
  alignments: Array<FrameXAlign | null>,
): void {
  const count = Math.min(paragraphs.length, alignments.length);
  for (let i = 0; i < count; i += 1) {
    const align = alignments[i];
    if (!align) continue;
    const para = paragraphs[i];
    if (!para.cssStyle || typeof para.cssStyle !== 'object') {
      para.cssStyle = {};
    }
    if (!para.cssStyle['text-align']) {
      para.cssStyle['text-align'] = align;
    }
  }
}

/**
 * 读取页眉/页脚原始 XML 中的 <w:framePr w:xAlign="..."/> 并应用到对应 AST 段落的 text-align。
 */
export async function applyHeaderFooterFrameAlignment(doc: any): Promise<void> {
  if (!doc || !Array.isArray(doc.parts)) return;

  for (const part of doc.parts) {
    const rootType = part?.rootElement?.type;
    if (rootType !== 'footer' && rootType !== 'header') continue;
    if (!part.path || typeof part._package?.load !== 'function') continue;

    let xmlText: string | null = null;
    try {
      xmlText = await part._package.load(part.path, 'string');
    } catch {
      continue;
    }
    if (!xmlText) continue;

    const alignments = extractFrameAlignmentsFromXml(xmlText);
    if (!alignments.some(Boolean)) continue;

    const paragraphs: Array<Record<string, any>> = [];
    collectAstParagraphs(part.rootElement, paragraphs, new WeakSet());
    applyFrameAlignmentsToParagraphs(paragraphs, alignments);
  }
}
