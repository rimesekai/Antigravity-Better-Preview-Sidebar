export const DOCX_CJK_PARA_CLASS_PREFIX = 'fp-cjk-p-';

const DEFAULT_FONT_SIZE_PT = 10.5;
const FALLBACK_LINE_HEIGHT_RATIO = 1.3;

export type IndentChannelValue =
  | { kind: 'dxa' }
  | { kind: 'chars'; value: number }
  | { kind: 'zero' };

export type SpacingChannelValue =
  | { kind: 'dxa' }
  | { kind: 'lines'; value: number }
  | { kind: 'zero' };

export interface ParagraphLayoutLayer {
  textIndent?: IndentChannelValue;
  marginStart?: IndentChannelValue;
  marginEnd?: IndentChannelValue;
  marginTop?: SpacingChannelValue;
  marginBottom?: SpacingChannelValue;
  fontSizePt?: number;
}

export interface DocxCjkParagraphPatch {
  textIndent?: Exclude<IndentChannelValue, { kind: 'dxa' }>;
  marginStart?: Exclude<IndentChannelValue, { kind: 'dxa' }>;
  marginEnd?: Exclude<IndentChannelValue, { kind: 'dxa' }>;
  marginTop?: Exclude<SpacingChannelValue, { kind: 'dxa' }>;
  marginBottom?: Exclude<SpacingChannelValue, { kind: 'dxa' }>;
  /** 段落自身 <w:pPr><w:rPr><w:sz> 声明的字号（pt） */
  directFontSizePt?: number;
  /** 段落样式链 / docDefaults 回退字号（pt） */
  fallbackFontSizePt?: number;
  /** 所在分节 <w:docGrid w:linePitch> 行高（pt） */
  linePitchPt?: number;
}

export interface DocxCjkLayoutContext {
  patches: Map<string, DocxCjkParagraphPatch>;
}

export interface XmlAttrLike {
  localName: string;
  value: string;
}

export interface XmlElementLike {
  localName: string;
  nodeType?: number;
  childNodes: ArrayLike<any>;
  attributes: ArrayLike<XmlAttrLike>;
}

interface ParsedStyleLayout {
  id: string;
  isDefault: boolean;
  basedOn?: string;
  pLayer: ParagraphLayoutLayer;
  runFontSizePt?: number;
  numbering?: { id: string; level: number };
}

interface ParsedStylesRegistry {
  docDefaultP: ParagraphLayoutLayer;
  docDefaultFontSizePt?: number;
  defaultParagraphStyleId?: string;
  stylesById: Map<string, ParsedStyleLayout>;
}

function getXmlAttr(elem: XmlElementLike, localName: string): string | null {
  const attrs = elem.attributes;
  if (!attrs) return null;
  for (let i = 0; i < attrs.length; i += 1) {
    const a = attrs[i];
    if (a && a.localName === localName) {
      return a.value;
    }
  }
  return null;
}

function getXmlChildElements(elem: XmlElementLike, localName?: string): XmlElementLike[] {
  const result: XmlElementLike[] = [];
  const children = elem.childNodes;
  if (!children) return result;
  for (let i = 0; i < children.length; i += 1) {
    const c = children[i];
    if (!c || typeof c !== 'object') continue;
    // Node.ELEMENT_NODE === 1；测试 mock 允许省略 nodeType
    if (c.nodeType !== undefined && c.nodeType !== 1) continue;
    if (!c.localName) continue;
    if (localName === undefined || c.localName === localName) {
      result.push(c as XmlElementLike);
    }
  }
  return result;
}

function getXmlFirstChild(elem: XmlElementLike, localName: string): XmlElementLike | null {
  const children = elem.childNodes;
  if (!children) return null;
  for (let i = 0; i < children.length; i += 1) {
    const c = children[i];
    if (!c || typeof c !== 'object') continue;
    if (c.nodeType !== undefined && c.nodeType !== 1) continue;
    if (c.localName === localName) {
      return c as XmlElementLike;
    }
  }
  return null;
}

function parseNumberValue(raw: string | null | undefined): number | null {
  if (raw == null) return null;
  const trimmed = raw.trim();
  if (!trimmed) return null;
  const parsed = Number.parseFloat(trimmed);
  return Number.isFinite(parsed) ? parsed : null;
}

function parseBoolAttr(raw: string | null | undefined): boolean {
  if (!raw) return false;
  const v = raw.trim().toLowerCase();
  return v === '1' || v === 'true' || v === 'on';
}

function resolveDxaAndChars(
  dxaVal: number | null,
  charsVal: number | null,
  sign: 1 | -1 = 1,
): IndentChannelValue | undefined {
  if (charsVal !== null) {
    if (dxaVal !== null && dxaVal !== 0) {
      return { kind: 'dxa' };
    }
    if (charsVal !== 0) {
      return { kind: 'chars', value: (charsVal / 100) * sign };
    }
    return { kind: 'zero' };
  }
  if (dxaVal !== null) {
    return { kind: 'dxa' };
  }
  return undefined;
}

function resolveDxaAndLines(
  dxaVal: number | null,
  linesVal: number | null,
): SpacingChannelValue | undefined {
  if (linesVal !== null) {
    if (dxaVal !== null && dxaVal !== 0) {
      return { kind: 'dxa' };
    }
    if (linesVal !== 0) {
      return { kind: 'lines', value: linesVal / 100 };
    }
    return { kind: 'zero' };
  }
  if (dxaVal !== null) {
    return { kind: 'dxa' };
  }
  return undefined;
}

/**
 * 解析 <w:ind> 元素中的字符单位与绝对单位缩进属性。
 * 规则：
 * - hanging / hangingChars 优先于 firstLine / firstLineChars（与 OOXML §17.3.1.12 及 docx-preview 一致）；
 * - start / startChars 优先于 left / leftChars；end / endChars 优先于 right / rightChars；
 * - 若绝对值（dxa）非零，则由 docx-preview 处理并标记 kind: 'dxa' 阻断继承；
 * - 若绝对值缺失或为 0 且存在 *Chars，则以 *Chars 为准（非零转 chars，0 转 zero）。
 */
export function parseIndAttributes(
  getAttr: (localName: string) => string | null,
): Pick<ParagraphLayoutLayer, 'textIndent' | 'marginStart' | 'marginEnd'> {
  const result: Pick<ParagraphLayoutLayer, 'textIndent' | 'marginStart' | 'marginEnd'> = {};

  const hanging = parseNumberValue(getAttr('hanging'));
  const hangingChars = parseNumberValue(getAttr('hangingChars'));
  const firstLine = parseNumberValue(getAttr('firstLine'));
  const firstLineChars = parseNumberValue(getAttr('firstLineChars'));

  const hangingResolved = resolveDxaAndChars(hanging, hangingChars, -1);
  const firstLineResolved = resolveDxaAndChars(firstLine, firstLineChars, 1);

  const hasNonZeroHanging =
    (hanging !== null && hanging !== 0) || (hangingChars !== null && hangingChars !== 0);
  const hasNonZeroFirstLine =
    (firstLine !== null && firstLine !== 0) || (firstLineChars !== null && firstLineChars !== 0);

  if (hasNonZeroHanging) {
    result.textIndent = hangingResolved;
  } else if (hasNonZeroFirstLine) {
    result.textIndent = firstLineResolved;
  } else if (hangingResolved !== undefined) {
    result.textIndent = hangingResolved;
  } else if (firstLineResolved !== undefined) {
    result.textIndent = firstLineResolved;
  }

  const start = parseNumberValue(getAttr('start'));
  const left = parseNumberValue(getAttr('left'));
  const startChars = parseNumberValue(getAttr('startChars'));
  const leftChars = parseNumberValue(getAttr('leftChars'));
  const marginStart = resolveDxaAndChars(start ?? left, startChars ?? leftChars, 1);
  if (marginStart !== undefined) {
    result.marginStart = marginStart;
  }

  const end = parseNumberValue(getAttr('end'));
  const right = parseNumberValue(getAttr('right'));
  const endChars = parseNumberValue(getAttr('endChars'));
  const rightChars = parseNumberValue(getAttr('rightChars'));
  const marginEnd = resolveDxaAndChars(end ?? right, endChars ?? rightChars, 1);
  if (marginEnd !== undefined) {
    result.marginEnd = marginEnd;
  }

  return result;
}

/**
 * 解析 <w:spacing> 元素中的行单位（beforeLines / afterLines）与绝对单位段距属性。
 */
export function parseSpacingAttributes(
  getAttr: (localName: string) => string | null,
): Pick<ParagraphLayoutLayer, 'marginTop' | 'marginBottom'> {
  const result: Pick<ParagraphLayoutLayer, 'marginTop' | 'marginBottom'> = {};

  const before = parseNumberValue(getAttr('before'));
  const beforeLines = parseNumberValue(getAttr('beforeLines'));
  const marginTop = resolveDxaAndLines(before, beforeLines);
  if (marginTop !== undefined) {
    result.marginTop = marginTop;
  }

  const after = parseNumberValue(getAttr('after'));
  const afterLines = parseNumberValue(getAttr('afterLines'));
  const marginBottom = resolveDxaAndLines(after, afterLines);
  if (marginBottom !== undefined) {
    result.marginBottom = marginBottom;
  }

  return result;
}

function parseRunFontSizePt(rPrElem: XmlElementLike | null): number | undefined {
  if (!rPrElem) return undefined;
  const szElem = getXmlFirstChild(rPrElem, 'sz');
  if (!szElem) return undefined;
  const val = parseNumberValue(getXmlAttr(szElem, 'val'));
  if (val === null || val <= 0) return undefined;
  return val / 2;
}

function parseNumberingRef(pPrElem: XmlElementLike): { id: string; level: number } | undefined {
  const numPr = getXmlFirstChild(pPrElem, 'numPr');
  if (!numPr) return undefined;
  const numIdEl = getXmlFirstChild(numPr, 'numId');
  const id = numIdEl ? getXmlAttr(numIdEl, 'val') : null;
  if (!id || id === '0') return undefined;
  const ilvlEl = getXmlFirstChild(numPr, 'ilvl');
  const parsedLevel = ilvlEl ? parseNumberValue(getXmlAttr(ilvlEl, 'val')) : 0;
  const level = parsedLevel !== null && parsedLevel >= 0 ? Math.floor(parsedLevel) : 0;
  return { id, level };
}

function parseSectionLinePitchPt(sectPrElem: XmlElementLike | null): number | undefined {
  if (!sectPrElem) return undefined;
  const docGrid = getXmlFirstChild(sectPrElem, 'docGrid');
  if (!docGrid) return undefined;
  const linePitch = parseNumberValue(getXmlAttr(docGrid, 'linePitch'));
  if (linePitch === null || linePitch <= 0) return undefined;
  return linePitch / 20;
}

function parsePPrLayout(pPrElem: XmlElementLike | null): {
  layer: ParagraphLayoutLayer;
  numbering?: { id: string; level: number };
  sectionLinePitchPt?: number;
  hasSectionBreak: boolean;
} {
  const layer: ParagraphLayoutLayer = {};
  if (!pPrElem) {
    return { layer, hasSectionBreak: false };
  }

  const indElem = getXmlFirstChild(pPrElem, 'ind');
  if (indElem) {
    Object.assign(layer, parseIndAttributes((name) => getXmlAttr(indElem, name)));
  }

  const spacingElem = getXmlFirstChild(pPrElem, 'spacing');
  if (spacingElem) {
    Object.assign(layer, parseSpacingAttributes((name) => getXmlAttr(spacingElem, name)));
  }

  const rPrElem = getXmlFirstChild(pPrElem, 'rPr');
  const fontSizePt = parseRunFontSizePt(rPrElem);
  if (fontSizePt !== undefined) {
    layer.fontSizePt = fontSizePt;
  }

  const numbering = parseNumberingRef(pPrElem);
  const sectPrElem = getXmlFirstChild(pPrElem, 'sectPr');
  const sectionLinePitchPt = parseSectionLinePitchPt(sectPrElem);

  return {
    layer,
    numbering,
    sectionLinePitchPt,
    hasSectionBreak: sectPrElem !== null,
  };
}

function hasAnyLayoutLayerData(layer: ParagraphLayoutLayer): boolean {
  return (
    layer.textIndent !== undefined ||
    layer.marginStart !== undefined ||
    layer.marginEnd !== undefined ||
    layer.marginTop !== undefined ||
    layer.marginBottom !== undefined
  );
}

/**
 * 按优先级从高到低合并多个段落样式层，并产出需要修补的 DocxCjkParagraphPatch。
 * 若 5 个通道均不需要字符/行单位修补或显式归零，则返回 null。
 */
export function resolveCjkParagraphPatch(
  layersHighToLow: ParagraphLayoutLayer[],
  options: {
    directFontSizePt?: number;
    fallbackFontSizePt?: number;
    linePitchPt?: number;
  } = {},
): DocxCjkParagraphPatch | null {
  let textIndent: IndentChannelValue | undefined;
  let marginStart: IndentChannelValue | undefined;
  let marginEnd: IndentChannelValue | undefined;
  let marginTop: SpacingChannelValue | undefined;
  let marginBottom: SpacingChannelValue | undefined;

  for (const layer of layersHighToLow) {
    if (textIndent === undefined && layer.textIndent !== undefined) {
      textIndent = layer.textIndent;
    }
    if (marginStart === undefined && layer.marginStart !== undefined) {
      marginStart = layer.marginStart;
    }
    if (marginEnd === undefined && layer.marginEnd !== undefined) {
      marginEnd = layer.marginEnd;
    }
    if (marginTop === undefined && layer.marginTop !== undefined) {
      marginTop = layer.marginTop;
    }
    if (marginBottom === undefined && layer.marginBottom !== undefined) {
      marginBottom = layer.marginBottom;
    }
  }

  const patch: DocxCjkParagraphPatch = {};
  let hasPatch = false;

  if (textIndent && textIndent.kind !== 'dxa') {
    patch.textIndent = textIndent;
    hasPatch = true;
  }
  if (marginStart && marginStart.kind !== 'dxa') {
    patch.marginStart = marginStart;
    hasPatch = true;
  }
  if (marginEnd && marginEnd.kind !== 'dxa') {
    patch.marginEnd = marginEnd;
    hasPatch = true;
  }
  if (marginTop && marginTop.kind !== 'dxa') {
    patch.marginTop = marginTop;
    hasPatch = true;
  }
  if (marginBottom && marginBottom.kind !== 'dxa') {
    patch.marginBottom = marginBottom;
    hasPatch = true;
  }

  if (!hasPatch) return null;

  if (options.directFontSizePt !== undefined) {
    patch.directFontSizePt = options.directFontSizePt;
  }
  if (options.fallbackFontSizePt !== undefined) {
    patch.fallbackFontSizePt = options.fallbackFontSizePt;
  }
  if (options.linePitchPt !== undefined) {
    patch.linePitchPt = options.linePitchPt;
  }

  return patch;
}

export function parseStylesXmlRoot(stylesRoot: XmlElementLike | null): ParsedStylesRegistry {
  const registry: ParsedStylesRegistry = {
    docDefaultP: {},
    stylesById: new Map(),
  };
  if (!stylesRoot) return registry;

  for (const child of getXmlChildElements(stylesRoot)) {
    if (child.localName === 'docDefaults') {
      const pPrDefault = getXmlFirstChild(child, 'pPrDefault');
      const pPr = pPrDefault ? getXmlFirstChild(pPrDefault, 'pPr') : null;
      if (pPr) {
        registry.docDefaultP = parsePPrLayout(pPr).layer;
      }
      const rPrDefault = getXmlFirstChild(child, 'rPrDefault');
      const rPr = rPrDefault ? getXmlFirstChild(rPrDefault, 'rPr') : null;
      const defSz = parseRunFontSizePt(rPr);
      if (defSz !== undefined) {
        registry.docDefaultFontSizePt = defSz;
      }
      continue;
    }

    if (child.localName !== 'style') continue;
    if (getXmlAttr(child, 'type') !== 'paragraph') continue;

    const id = getXmlAttr(child, 'styleId');
    if (!id) continue;

    const isDefault = parseBoolAttr(getXmlAttr(child, 'default'));
    const basedOnEl = getXmlFirstChild(child, 'basedOn');
    const basedOn = basedOnEl ? getXmlAttr(basedOnEl, 'val') ?? undefined : undefined;
    const pPr = getXmlFirstChild(child, 'pPr');
    const parsedPPr = parsePPrLayout(pPr);
    const rPr = getXmlFirstChild(child, 'rPr');
    const runFontSizePt = parseRunFontSizePt(rPr) ?? parsedPPr.layer.fontSizePt;

    registry.stylesById.set(id, {
      id,
      isDefault,
      basedOn,
      pLayer: parsedPPr.layer,
      runFontSizePt,
      numbering: parsedPPr.numbering,
    });

    if (isDefault && !registry.defaultParagraphStyleId) {
      registry.defaultParagraphStyleId = id;
    }
  }

  return registry;
}

export function parseNumberingXmlRoot(
  numberingRoot: XmlElementLike | null,
): Map<string, ParagraphLayoutLayer> {
  const result = new Map<string, ParagraphLayoutLayer>();
  if (!numberingRoot) return result;

  const abstractLevels = new Map<string, Map<number, ParagraphLayoutLayer>>();

  for (const child of getXmlChildElements(numberingRoot, 'abstractNum')) {
    const absId = getXmlAttr(child, 'abstractNumId');
    if (!absId) continue;
    const lvlMap = new Map<number, ParagraphLayoutLayer>();
    for (const lvl of getXmlChildElements(child, 'lvl')) {
      const ilvl = parseNumberValue(getXmlAttr(lvl, 'ilvl'));
      if (ilvl === null) continue;
      const pPr = getXmlFirstChild(lvl, 'pPr');
      const { layer } = parsePPrLayout(pPr);
      if (hasAnyLayoutLayerData(layer)) {
        lvlMap.set(Math.floor(ilvl), layer);
      }
    }
    abstractLevels.set(absId, lvlMap);
  }

  for (const child of getXmlChildElements(numberingRoot, 'num')) {
    const numId = getXmlAttr(child, 'numId');
    if (!numId) continue;
    const absRef = getXmlFirstChild(child, 'abstractNumId');
    const absId = absRef ? getXmlAttr(absRef, 'val') : null;
    const baseLevels = absId ? abstractLevels.get(absId) : undefined;
    if (baseLevels) {
      for (const [ilvl, layer] of baseLevels) {
        result.set(`${numId}:${ilvl}`, layer);
      }
    }

    for (const ov of getXmlChildElements(child, 'lvlOverride')) {
      const ilvl = parseNumberValue(getXmlAttr(ov, 'ilvl'));
      if (ilvl === null) continue;
      const lvl = getXmlFirstChild(ov, 'lvl');
      if (!lvl) continue;
      const pPr = getXmlFirstChild(lvl, 'pPr');
      const { layer } = parsePPrLayout(pPr);
      if (hasAnyLayoutLayerData(layer)) {
        result.set(`${numId}:${Math.floor(ilvl)}`, layer);
      }
    }
  }

  return result;
}

function collectStyleChain(
  styleId: string | undefined,
  registry: ParsedStylesRegistry,
): ParsedStyleLayout[] {
  const chain: ParsedStyleLayout[] = [];
  const seen = new Set<string>();

  const walkStyle = (startId?: string) => {
    let curr = startId;
    while (curr && !seen.has(curr)) {
      seen.add(curr);
      const st = registry.stylesById.get(curr);
      if (!st) break;
      chain.push(st);
      curr = st.basedOn;
    }
  };

  if (styleId) {
    walkStyle(styleId);
  }
  if (registry.defaultParagraphStyleId && !seen.has(registry.defaultParagraphStyleId)) {
    walkStyle(registry.defaultParagraphStyleId);
  }

  return chain;
}

function flattenBodyXmlBlocks(xmlContainer: XmlElementLike, allowSdt: boolean): XmlElementLike[] {
  const out: XmlElementLike[] = [];
  for (const child of getXmlChildElements(xmlContainer)) {
    switch (child.localName) {
      case 'p':
      case 'tbl':
        out.push(child);
        break;
      case 'altChunk':
        if (allowSdt) {
          out.push(child);
        }
        break;
      case 'sdt':
        if (allowSdt) {
          const sdtContent = getXmlFirstChild(child, 'sdtContent');
          if (sdtContent) {
            out.push(...flattenBodyXmlBlocks(sdtContent, true));
          }
        }
        break;
      default:
        break;
    }
  }
  return out;
}

function collectXmlTxbxContents(node: XmlElementLike, out: XmlElementLike[]): void {
  for (const child of getXmlChildElements(node)) {
    if (child.localName === 'txbxContent') {
      out.push(child);
      continue;
    }
    // 跳过 AlternateContent 中的 Fallback，避免与 docx-preview 只选 Choice 冲突
    if (child.localName === 'AlternateContent') {
      const choice = getXmlFirstChild(child, 'Choice');
      const fallback = getXmlFirstChild(child, 'Fallback');
      const branch = choice ?? fallback;
      if (branch) {
        collectXmlTxbxContents(branch, out);
      }
      continue;
    }
    collectXmlTxbxContents(child, out);
  }
}

function collectAstVmlTextBoxes(node: unknown, out: Array<Record<string, any>>): void {
  if (!node || typeof node !== 'object') return;
  if (Array.isArray(node)) {
    for (const item of node) collectAstVmlTextBoxes(item, out);
    return;
  }
  const rec = node as Record<string, any>;
  if (
    rec.type === 'vmlElement' &&
    Array.isArray(rec.children) &&
    rec.children.some((c: any) => c?.type === 'paragraph' || c?.type === 'table')
  ) {
    out.push(rec);
    return;
  }
  if (Array.isArray(rec.children)) {
    for (const child of rec.children) {
      collectAstVmlTextBoxes(child, out);
    }
  }
}

function computeBodySectionLinePitches(
  bodyBlocks: XmlElementLike[],
  defaultLinePitchPt?: number,
): Array<number | undefined> {
  // 每个段落所属分节的 sectPr 出现在该节末尾段落（或 body 末尾）
  const sectionEndIndices: Array<{ blockIdx: number; linePitchPt?: number }> = [];
  for (let i = 0; i < bodyBlocks.length; i += 1) {
    const blk = bodyBlocks[i];
    if (blk.localName !== 'p') continue;
    const pPr = getXmlFirstChild(blk, 'pPr');
    const sectPr = pPr ? getXmlFirstChild(pPr, 'sectPr') : null;
    if (sectPr) {
      sectionEndIndices.push({
        blockIdx: i,
        linePitchPt: parseSectionLinePitchPt(sectPr),
      });
    }
  }

  const pitches: Array<number | undefined> = new Array(bodyBlocks.length);
  let cursor = 0;
  for (const sec of sectionEndIndices) {
    const pitch = sec.linePitchPt ?? defaultLinePitchPt;
    while (cursor <= sec.blockIdx) {
      pitches[cursor] = pitch;
      cursor += 1;
    }
  }
  while (cursor < bodyBlocks.length) {
    pitches[cursor] = defaultLinePitchPt;
    cursor += 1;
  }
  return pitches;
}

async function loadPartXmlRoot(part: any): Promise<XmlElementLike | null> {
  if (!part) return null;
  if (part._xmlDocument?.firstElementChild) {
    return part._xmlDocument.firstElementChild as XmlElementLike;
  }
  if (!part.path || typeof part._package?.load !== 'function') return null;

  try {
    const xmlText = await part._package.load(part.path, 'string');
    if (!xmlText) return null;
    if (typeof part._package.parseXmlDocument === 'function') {
      const xmlDoc = part._package.parseXmlDocument(xmlText);
      return (xmlDoc?.firstElementChild as XmlElementLike) ?? null;
    }
    if (typeof DOMParser !== 'undefined') {
      const xmlDoc = new DOMParser().parseFromString(xmlText, 'application/xml');
      return (xmlDoc?.firstElementChild as XmlElementLike) ?? null;
    }
  } catch {
    return null;
  }
  return null;
}

/**
 * 同步遍历 XML 容器节点与 docx-preview AST 子节点，为需要字符缩进或行距补丁的段落打标记并采集补丁。
 */
export function collectPatchesFromXmlContainer(
  xmlContainer: XmlElementLike,
  astChildren: Array<Record<string, any>>,
  registry: ParsedStylesRegistry,
  numberingLayers: Map<string, ParagraphLayoutLayer>,
  patches: Map<string, DocxCjkParagraphPatch>,
  counter: { nextId: number },
  options: {
    allowSdt: boolean;
    defaultLinePitchPt?: number;
    sectionLinePitches?: Array<number | undefined>;
  },
): void {
  if (!Array.isArray(astChildren) || astChildren.length === 0) return;

  const xmlBlocks = flattenBodyXmlBlocks(xmlContainer, options.allowSdt);
  if (xmlBlocks.length !== astChildren.length) return;

  for (let i = 0; i < xmlBlocks.length; i += 1) {
    const xmlNode = xmlBlocks[i];
    const astNode = astChildren[i];
    if (!astNode || typeof astNode !== 'object') continue;

    const linePitchPt = options.sectionLinePitches
      ? options.sectionLinePitches[i]
      : options.defaultLinePitchPt;

    if (xmlNode.localName === 'p' && astNode.type === 'paragraph') {
      const pPr = getXmlFirstChild(xmlNode, 'pPr');
      const parsedPPr = parsePPrLayout(pPr);
      const styleChain = collectStyleChain(astNode.styleName, registry);

      let effectiveNum = parsedPPr.numbering;
      if (!effectiveNum) {
        for (const st of styleChain) {
          if (st.numbering) {
            effectiveNum = st.numbering;
            break;
          }
        }
      }
      const numLayer = effectiveNum
        ? numberingLayers.get(`${effectiveNum.id}:${effectiveNum.level}`)
        : undefined;

      const layersHighToLow: ParagraphLayoutLayer[] = [];
      if (hasAnyLayoutLayerData(parsedPPr.layer)) {
        layersHighToLow.push(parsedPPr.layer);
      }
      if (numLayer) {
        layersHighToLow.push(numLayer);
      }
      for (const st of styleChain) {
        if (hasAnyLayoutLayerData(st.pLayer)) {
          layersHighToLow.push(st.pLayer);
        }
      }
      if (hasAnyLayoutLayerData(registry.docDefaultP)) {
        layersHighToLow.push(registry.docDefaultP);
      }

      let fallbackFontSizePt: number | undefined;
      for (const st of styleChain) {
        if (st.runFontSizePt !== undefined) {
          fallbackFontSizePt = st.runFontSizePt;
          break;
        }
      }
      if (fallbackFontSizePt === undefined) {
        fallbackFontSizePt = registry.docDefaultFontSizePt;
      }

      const patch = resolveCjkParagraphPatch(layersHighToLow, {
        directFontSizePt: parsedPPr.layer.fontSizePt,
        fallbackFontSizePt,
        linePitchPt,
      });

      if (patch) {
        const markerClass = `${DOCX_CJK_PARA_CLASS_PREFIX}${counter.nextId++}`;
        astNode.className = astNode.className ? `${astNode.className} ${markerClass}` : markerClass;
        patches.set(markerClass, patch);
      }

      // 同步处理段落内嵌套的 VML 文本框段落（仅在数量一致时处理，避免错位）
      const xmlTxbxs: XmlElementLike[] = [];
      collectXmlTxbxContents(xmlNode, xmlTxbxs);
      if (xmlTxbxs.length > 0 && Array.isArray(astNode.children)) {
        const astTxbxs: Array<Record<string, any>> = [];
        collectAstVmlTextBoxes(astNode.children, astTxbxs);
        if (xmlTxbxs.length === astTxbxs.length) {
          for (let t = 0; t < xmlTxbxs.length; t += 1) {
            collectPatchesFromXmlContainer(
              xmlTxbxs[t],
              astTxbxs[t].children,
              registry,
              numberingLayers,
              patches,
              counter,
              { allowSdt: true, defaultLinePitchPt: linePitchPt },
            );
          }
        }
      }
      continue;
    }

    if (xmlNode.localName === 'tbl' && astNode.type === 'table' && Array.isArray(astNode.children)) {
      const xmlRows = getXmlChildElements(xmlNode, 'tr');
      if (xmlRows.length !== astNode.children.length) continue;

      for (let r = 0; r < xmlRows.length; r += 1) {
        const astRow = astNode.children[r];
        if (!astRow || astRow.type !== 'row' || !Array.isArray(astRow.children)) continue;

        const xmlCells = getXmlChildElements(xmlRows[r], 'tc');
        if (xmlCells.length !== astRow.children.length) continue;

        for (let c = 0; c < xmlCells.length; c += 1) {
          const astCell = astRow.children[c];
          if (!astCell || astCell.type !== 'cell' || !Array.isArray(astCell.children)) continue;

          // 与 docx-preview parseTableCell 一致：单元格内直接遍历 p 与 tbl
          collectPatchesFromXmlContainer(
            xmlCells[c],
            astCell.children,
            registry,
            numberingLayers,
            patches,
            counter,
            { allowSdt: false, defaultLinePitchPt: linePitchPt },
          );
        }
      }
    }
  }
}

/**
 * 在 docx-preview parseAsync 之后、renderDocument 之前采集文档中所有字符单位缩进
 *（firstLineChars / hangingChars / leftChars / startChars / rightChars / endChars）
 * 与行单位段距（beforeLines / afterLines），并清理临时 _xmlDocument 引用。
 */
export async function collectDocxCjkLayout(doc: any): Promise<DocxCjkLayoutContext> {
  const patches = new Map<string, DocxCjkParagraphPatch>();
  if (!doc || typeof doc !== 'object') {
    return { patches };
  }

  try {
    const stylesRoot = await loadPartXmlRoot(doc.stylesPart);
    const registry = parseStylesXmlRoot(stylesRoot);

    const numberingRoot = await loadPartXmlRoot(doc.numberingPart);
    const numberingLayers = parseNumberingXmlRoot(numberingRoot);

    const counter = { nextId: 0 };
    let defaultBodyLinePitchPt: number | undefined;

    // 1. 正文 document.xml
    const docRoot = await loadPartXmlRoot(doc.documentPart);
    const xmlBody = docRoot ? getXmlFirstChild(docRoot, 'body') : null;
    if (xmlBody && Array.isArray(doc.documentPart?.body?.children)) {
      const bodySectPr = getXmlFirstChild(xmlBody, 'sectPr');
      defaultBodyLinePitchPt = parseSectionLinePitchPt(bodySectPr);
      const bodyBlocks = flattenBodyXmlBlocks(xmlBody, true);
      const sectionLinePitches = computeBodySectionLinePitches(bodyBlocks, defaultBodyLinePitchPt);

      collectPatchesFromXmlContainer(
        xmlBody,
        doc.documentPart.body.children,
        registry,
        numberingLayers,
        patches,
        counter,
        {
          allowSdt: true,
          defaultLinePitchPt: defaultBodyLinePitchPt,
          sectionLinePitches,
        },
      );
    }

    // 2. 页眉 / 页脚
    if (Array.isArray(doc.parts)) {
      for (const part of doc.parts) {
        const rootType = part?.rootElement?.type;
        if (rootType !== 'header' && rootType !== 'footer') continue;
        if (!Array.isArray(part.rootElement.children)) continue;
        const partRoot = await loadPartXmlRoot(part);
        if (!partRoot) continue;
        collectPatchesFromXmlContainer(
          partRoot,
          part.rootElement.children,
          registry,
          numberingLayers,
          patches,
          counter,
          { allowSdt: true, defaultLinePitchPt: defaultBodyLinePitchPt },
        );
      }
    }

    // 3. 脚注 / 尾注
    for (const notesPart of [doc.footnotesPart, doc.endnotesPart]) {
      if (!notesPart || !Array.isArray(notesPart.notes) || notesPart.notes.length === 0) continue;
      const notesRoot = await loadPartXmlRoot(notesPart);
      if (!notesRoot) continue;
      const noteMap = new Map<string, any>();
      for (const n of notesPart.notes) {
        if (n?.id != null) {
          noteMap.set(String(n.id), n);
        }
      }
      for (const noteEl of getXmlChildElements(notesRoot)) {
        if (noteEl.localName !== 'footnote' && noteEl.localName !== 'endnote') continue;
        const id = getXmlAttr(noteEl, 'id');
        if (!id) continue;
        const astNote = noteMap.get(id);
        if (!astNote || !Array.isArray(astNote.children)) continue;
        collectPatchesFromXmlContainer(
          noteEl,
          astNote.children,
          registry,
          numberingLayers,
          patches,
          counter,
          { allowSdt: true, defaultLinePitchPt: defaultBodyLinePitchPt },
        );
      }
    }
  } finally {
    // 释放 keepOrigin 缓存的原始 XML Document，避免常驻内存
    if (Array.isArray(doc.parts)) {
      for (const part of doc.parts) {
        if (part && '_xmlDocument' in part) {
          part._xmlDocument = undefined;
        }
      }
    }
  }

  return { patches };
}

function measurePointToPixel(container: HTMLElement): number {
  if (typeof document === 'undefined') return 96 / 72;
  const probe = document.createElement('div');
  probe.style.position = 'absolute';
  probe.style.visibility = 'hidden';
  probe.style.width = '100pt';
  container.appendChild(probe);
  const px = probe.offsetWidth;
  probe.remove();
  return px > 0 ? px / 100 : 96 / 72;
}

function hasVisibleParagraphContent(p: HTMLElement): boolean {
  const text = p.textContent;
  if (text && text.replace(/[\u200B\uFEFF]/g, '').trim().length > 0) {
    return true;
  }
  return Boolean(p.querySelector('img, svg, math, canvas, table'));
}

function measureParagraphCharSizePx(
  p: HTMLElement,
  patch: DocxCjkParagraphPatch,
  ptToPx: number,
): number {
  if (typeof document !== 'undefined' && typeof getComputedStyle === 'function') {
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    let firstAnyFontSizePx: number | null = null;
    let node = walker.nextNode();

    while (node) {
      const text = node.nodeValue;
      const parent = node.parentElement;
      if (
        parent &&
        text &&
        text.replace(/[\s\u00a0\u200B\uFEFF]/g, '').length > 0 &&
        parent.closest('p') === p &&
        !parent.closest('.docx-tab-stop, math, svg')
      ) {
        const cs = getComputedStyle(parent);
        const fs = Number.parseFloat(cs.fontSize);
        if (Number.isFinite(fs) && fs > 0) {
          if (firstAnyFontSizePx === null) {
            firstAnyFontSizePx = fs;
          }
          const vAlign = cs.verticalAlign;
          const isScript =
            vAlign === 'super' ||
            vAlign === 'sub' ||
            Boolean(parent.closest('sup, sub'));
          if (!isScript) {
            return fs;
          }
        }
      }
      node = walker.nextNode();
    }

    if (firstAnyFontSizePx !== null) {
      return firstAnyFontSizePx;
    }
  }

  const fallbackPt = patch.directFontSizePt ?? patch.fallbackFontSizePt ?? DEFAULT_FONT_SIZE_PT;
  return fallbackPt * ptToPx;
}

function applyPatchToParagraphElement(
  p: HTMLElement,
  patch: DocxCjkParagraphPatch,
  ptToPx: number,
  flags: { applyFirstLine: boolean; applyTop: boolean; applyBottom: boolean },
  measureCharPx?: (p: HTMLElement, patch: DocxCjkParagraphPatch, ptToPx: number) => number,
): void {
  const needsCharSize =
    (flags.applyFirstLine && patch.textIndent?.kind === 'chars') ||
    patch.marginStart?.kind === 'chars' ||
    patch.marginEnd?.kind === 'chars' ||
    (flags.applyTop && patch.marginTop?.kind === 'lines' && patch.linePitchPt === undefined) ||
    (flags.applyBottom && patch.marginBottom?.kind === 'lines' && patch.linePitchPt === undefined);

  const charSizePx = needsCharSize
    ? (measureCharPx ?? measureParagraphCharSizePx)(p, patch, ptToPx)
    : DEFAULT_FONT_SIZE_PT * ptToPx;

  const lineHeightPx =
    patch.linePitchPt !== undefined ? patch.linePitchPt * ptToPx : charSizePx * FALLBACK_LINE_HEIGHT_RATIO;

  if (flags.applyFirstLine && patch.textIndent) {
    p.style.textIndent =
      patch.textIndent.kind === 'zero'
        ? '0'
        : `${(patch.textIndent.value * charSizePx).toFixed(2)}px`;
  }

  if (patch.marginStart) {
    p.style.marginInlineStart =
      patch.marginStart.kind === 'zero'
        ? '0'
        : `${(patch.marginStart.value * charSizePx).toFixed(2)}px`;
  }

  if (patch.marginEnd) {
    p.style.marginInlineEnd =
      patch.marginEnd.kind === 'zero'
        ? '0'
        : `${(patch.marginEnd.value * charSizePx).toFixed(2)}px`;
  }

  if (flags.applyTop && patch.marginTop) {
    p.style.marginTop =
      patch.marginTop.kind === 'zero'
        ? '0'
        : `${(patch.marginTop.value * lineHeightPx).toFixed(2)}px`;
  }

  if (flags.applyBottom && patch.marginBottom) {
    p.style.marginBottom =
      patch.marginBottom.kind === 'zero'
        ? '0'
        : `${(patch.marginBottom.value * lineHeightPx).toFixed(2)}px`;
  }
}

/**
 * 在 DOM 挂载到 stageEl 之后，根据各段落实际渲染字号与分节行距应用字符单位缩进与行单位段距。
 */
export function applyDocxCjkLayout(
  stageEl: HTMLElement,
  ctx: DocxCjkLayoutContext,
  options?: {
    ptToPx?: number;
    measureCharPx?: (p: HTMLElement, patch: DocxCjkParagraphPatch, ptToPx: number) => number;
  },
): void {
  if (!ctx || ctx.patches.size === 0) return;

  const ptToPx = options?.ptToPx ?? measurePointToPixel(stageEl);

  for (const [markerClass, patch] of ctx.patches) {
    const matched = Array.from(stageEl.querySelectorAll<HTMLElement>(`p.${markerClass}`));
    if (matched.length === 0) continue;

    const headerFooterParas: HTMLElement[] = [];
    const bodyParas: HTMLElement[] = [];

    for (const p of matched) {
      if (typeof p.closest === 'function' && p.closest('header, footer')) {
        headerFooterParas.push(p);
      } else {
        bodyParas.push(p);
      }
    }

    // 页眉/页脚在每一页都会独立克隆，每个实例均完整应用
    for (const p of headerFooterParas) {
      applyPatchToParagraphElement(
        p,
        patch,
        ptToPx,
        { applyFirstLine: true, applyTop: true, applyBottom: true },
        options?.measureCharPx,
      );
    }

    // 正文段落在跨页拆分时可能对应多个 <p>，首行缩进与段前距仅用于首个可见片段，段后距仅用于末片段
    if (bodyParas.length > 0) {
      let firstContentIdx = bodyParas.findIndex((p) => hasVisibleParagraphContent(p));
      if (firstContentIdx < 0) firstContentIdx = 0;
      const lastIdx = bodyParas.length - 1;

      for (let i = 0; i < bodyParas.length; i += 1) {
        applyPatchToParagraphElement(
          bodyParas[i],
          patch,
          ptToPx,
          {
            applyFirstLine: i === firstContentIdx,
            applyTop: i === firstContentIdx,
            applyBottom: i === lastIdx,
          },
          options?.measureCharPx,
        );
      }
    }
  }
}
