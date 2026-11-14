/**
 * 放大后鼠标拖拽平移控制器：
 * - 允许在可滚动容器空白处按住鼠标左键或任意区域按住中键拖拽平移视图；
 * - 鼠标停留在文字上时自动放行左键原生划选，并提供悬停光标状态类名（is-over-text）；
 * - 自动忽略交互式控件与原生滚动条轨道上的按下事件；
 * - 拖拽完成后抑制紧随其后的一次 click 事件，防止误触链接或幻灯片元素。
 */

export interface DragPannerOptions {
  /** 获取当前滚动容器 */
  getScrollContainer: () => HTMLElement | null;
  /** 判断当前是否允许拖拽平移 */
  isEnabled: () => boolean;
  /** 可选判断目标元素是否允许触发拖拽平移（返回 false 时放行原生交互） */
  shouldPanTarget?: (target: EventTarget | null, e: MouseEvent) => boolean;
  /** 触发拖拽平移的最小移动距离（像素），默认 3 */
  dragThresholdPx?: number;
  /** 全局事件监听目标，默认在浏览器环境取 window */
  eventTarget?: Pick<Window, 'addEventListener' | 'removeEventListener'>;
}

export interface DragPannerController {
  handleMouseDown: (e: MouseEvent) => void;
  handleMouseMove: (e: MouseEvent) => void;
  handleMouseUp: (e?: MouseEvent) => void;
  handleHoverMove: (e: MouseEvent) => void;
  handleMouseLeave: () => void;
  handleClickCapture: (e: MouseEvent) => void;
  syncState: () => void;
  reset: () => void;
  destroy: () => void;
  isPanning: () => boolean;
}

interface TargetElementLike {
  nodeType?: number;
  tagName?: string;
  className?: unknown;
  classList?: { contains?: (token: string) => boolean };
  isContentEditable?: boolean;
  closest?: (selector: string) => unknown;
  getAttribute?: (name: string) => string | null;
  parentElement?: TargetElementLike | null;
  parentNode?: TargetElementLike | null;
  childNodes?: Iterable<unknown> | ArrayLike<unknown>;
}

interface RectLike {
  left: number;
  right: number;
  top: number;
  bottom: number;
  width?: number;
  height?: number;
}

interface CaretPositionLike {
  offsetNode: Node;
  offset: number;
  getClientRect?: () => RectLike | null;
}

interface DocumentWithCaretApis {
  caretRangeFromPoint?: (x: number, y: number) => Range | null;
  caretPositionFromPoint?: (x: number, y: number) => CaretPositionLike | null;
  createRange?: () => Range;
}

const DEFAULT_DRAG_THRESHOLD_PX = 3;
const CLICK_SUPPRESS_TIMEOUT_MS = 400;
const TEXT_HIT_TOLERANCE_PX = 3;
const CLASS_PANNABLE = 'is-pannable';
const CLASS_PANNING = 'is-panning';
const CLASS_OVER_TEXT = 'is-over-text';

const INTERACTIVE_TAGS = new Set([
  'BUTTON',
  'INPUT',
  'TEXTAREA',
  'SELECT',
  'A',
  'LABEL',
  'SUMMARY',
  'OPTION',
]);

const NON_TEXT_TAGS = new Set([
  'IMG',
  'CANVAS',
  'VIDEO',
  'AUDIO',
  'IFRAME',
  'OBJECT',
  'EMBED',
  'HR',
  'BR',
  'SOURCE',
  'TRACK',
]);

const INTERACTIVE_SELECTOR =
  'button, input, textarea, select, a, label, summary, option, [role="button"], [role="link"], [contenteditable="true"], [contenteditable=""]';

const EXCLUDED_UI_CLASSES = ['fp-selection-bar', 'fp-ppt-thumbs', 'fp-pdf-thumbs'] as const;
const EXCLUDED_UI_SELECTOR = '.fp-selection-bar, .fp-ppt-thumbs, .fp-pdf-thumbs';

function addClass(el: HTMLElement | null | undefined, className: string): void {
  el?.classList?.add(className);
}

function removeClass(el: HTMLElement | null | undefined, className: string): void {
  el?.classList?.remove(className);
}

function resolveTargetElement(target: EventTarget | null): TargetElementLike | null {
  if (!target || typeof target !== 'object') {
    return null;
  }
  const candidate = target as TargetElementLike;
  if (candidate.nodeType === 3) {
    return candidate.parentElement && typeof candidate.parentElement === 'object'
      ? candidate.parentElement
      : null;
  }
  if (
    candidate.nodeType === 1 ||
    typeof candidate.tagName === 'string' ||
    typeof candidate.closest === 'function' ||
    typeof candidate.isContentEditable === 'boolean' ||
    typeof candidate.className === 'string' ||
    Boolean(candidate.classList) ||
    Boolean(candidate.childNodes)
  ) {
    return candidate;
  }
  if (candidate.parentElement && typeof candidate.parentElement === 'object') {
    return candidate.parentElement;
  }
  return null;
}

function isInteractiveElementNode(el: TargetElementLike): boolean {
  if (el.isContentEditable === true) {
    return true;
  }
  if (typeof el.tagName === 'string' && INTERACTIVE_TAGS.has(el.tagName.toUpperCase())) {
    return true;
  }
  if (typeof el.getAttribute === 'function') {
    const role = el.getAttribute('role');
    if (role === 'button' || role === 'link') {
      return true;
    }
    const contentEditable = el.getAttribute('contenteditable');
    if (contentEditable === '' || contentEditable === 'true') {
      return true;
    }
  }
  return false;
}

function isInteractiveTarget(
  target: EventTarget | null,
  container?: HTMLElement | null,
): boolean {
  const el = resolveTargetElement(target);
  if (!el) {
    return false;
  }
  if (isInteractiveElementNode(el)) {
    return true;
  }
  if (typeof el.closest === 'function') {
    return Boolean(el.closest(INTERACTIVE_SELECTOR));
  }

  // 在无 closest 实现的轻量测试替身中沿 parentElement 向上检查
  let cursor: TargetElementLike | null | undefined = el.parentElement;
  let depth = 0;
  while (cursor && cursor !== (container as unknown as TargetElementLike) && depth < 32) {
    if (isInteractiveElementNode(cursor)) {
      return true;
    }
    cursor = cursor.parentElement;
    depth += 1;
  }
  return false;
}

function hasExcludedUiClass(el: TargetElementLike): boolean {
  if (el.classList && typeof el.classList.contains === 'function') {
    for (const cls of EXCLUDED_UI_CLASSES) {
      if (el.classList.contains(cls)) {
        return true;
      }
    }
  }
  if (typeof el.className === 'string' && el.className.length > 0) {
    const tokens = el.className.split(/\s+/);
    for (const cls of EXCLUDED_UI_CLASSES) {
      if (tokens.includes(cls)) {
        return true;
      }
    }
  }
  return false;
}

function isExcludedUiTarget(
  el: TargetElementLike | null,
  container?: HTMLElement | null,
): boolean {
  if (!el) {
    return false;
  }
  if (hasExcludedUiClass(el)) {
    return true;
  }
  if (typeof el.closest === 'function') {
    return Boolean(el.closest(EXCLUDED_UI_SELECTOR));
  }

  let cursor: TargetElementLike | null | undefined = el.parentElement;
  let depth = 0;
  while (cursor && cursor !== (container as unknown as TargetElementLike) && depth < 32) {
    if (hasExcludedUiClass(cursor)) {
      return true;
    }
    cursor = cursor.parentElement;
    depth += 1;
  }
  return false;
}

function isNonTextElement(el: TargetElementLike | null): boolean {
  if (!el || typeof el.tagName !== 'string') {
    return false;
  }
  return NON_TEXT_TAGS.has(el.tagName.toUpperCase());
}

function isNonEmptyTextNode(node: unknown): node is Text {
  if (!node || typeof node !== 'object') {
    return false;
  }
  const candidate = node as { nodeType?: unknown; textContent?: unknown };
  return (
    candidate.nodeType === 3 &&
    typeof candidate.textContent === 'string' &&
    candidate.textContent.trim().length > 0
  );
}

function isNodeInsideContainer(
  node: unknown,
  container: HTMLElement | null | undefined,
): boolean {
  if (!container || !node || typeof node !== 'object') {
    return true;
  }
  if (typeof container.contains === 'function') {
    return container.contains(node as Node);
  }
  let cursor: { parentElement?: unknown; parentNode?: unknown } | null = node as {
    parentElement?: unknown;
    parentNode?: unknown;
  };
  let depth = 0;
  while (cursor && depth < 64) {
    if (cursor === (container as unknown)) {
      return true;
    }
    cursor = (cursor.parentElement ?? cursor.parentNode ?? null) as {
      parentElement?: unknown;
      parentNode?: unknown;
    } | null;
    depth += 1;
  }
  const raw = node as { parentElement?: unknown; parentNode?: unknown };
  if (raw.parentElement === undefined && raw.parentNode === undefined) {
    return true;
  }
  return false;
}

function isValidRect(rect: RectLike | null | undefined): rect is RectLike {
  if (!rect) {
    return false;
  }
  const width = typeof rect.width === 'number' ? rect.width : rect.right - rect.left;
  const height = typeof rect.height === 'number' ? rect.height : rect.bottom - rect.top;
  return (
    Number.isFinite(rect.left) &&
    Number.isFinite(rect.right) &&
    Number.isFinite(rect.top) &&
    Number.isFinite(rect.bottom) &&
    (width > 0 || height > 0)
  );
}

function isPointInRectList(
  clientX: number,
  clientY: number,
  rects: Iterable<RectLike> | ArrayLike<RectLike> | null | undefined,
  tolerance = TEXT_HIT_TOLERANCE_PX,
): boolean {
  if (!rects || !Number.isFinite(clientX) || !Number.isFinite(clientY)) {
    return false;
  }
  const list = Array.from(rects);
  for (const rect of list) {
    if (!isValidRect(rect)) {
      continue;
    }
    if (
      clientX >= rect.left - tolerance &&
      clientX <= rect.right + tolerance &&
      clientY >= rect.top - tolerance &&
      clientY <= rect.bottom + tolerance
    ) {
      return true;
    }
  }
  return false;
}

function hasAnyValidRect(
  rects: Iterable<RectLike> | ArrayLike<RectLike> | null | undefined,
): boolean {
  if (!rects) {
    return false;
  }
  for (const rect of Array.from(rects)) {
    if (isValidRect(rect)) {
      return true;
    }
  }
  return false;
}

function hitTestCaretNode(
  doc: DocumentWithCaretApis,
  node: Node | null | undefined,
  offset: number,
  fallbackRects: (() => Iterable<RectLike> | ArrayLike<RectLike> | null | undefined) | null,
  clientX: number,
  clientY: number,
  container?: HTMLElement | null,
): boolean {
  if (!isNonEmptyTextNode(node)) {
    return false;
  }
  if (!isNodeInsideContainer(node, container)) {
    return false;
  }
  const parentEl = resolveTargetElement(node as unknown as EventTarget);
  if (
    parentEl &&
    (isInteractiveTarget(parentEl as unknown as EventTarget, container) ||
      isExcludedUiTarget(parentEl, container) ||
      isNonTextElement(parentEl))
  ) {
    return false;
  }

  const len = node.textContent?.length ?? 0;
  if (len <= 0) {
    return false;
  }

  if (typeof doc.createRange === 'function') {
    try {
      const charRange = doc.createRange();
      const clampedOffset = Math.max(0, Math.min(offset, len));
      const start = Math.max(0, Math.min(clampedOffset - 1, len - 1));
      const end = Math.min(len, Math.max(start + 1, clampedOffset + 1));
      charRange.setStart(node, start);
      charRange.setEnd(node, end);
      const charRects = charRange.getClientRects?.();
      if (hasAnyValidRect(charRects)) {
        return isPointInRectList(clientX, clientY, charRects);
      }
    } catch {
      // 忽略测试替身或非标准节点导致的 Range 异常，继续尝试 fallbackRects
    }
  }

  if (fallbackRects) {
    const rects = fallbackRects();
    if (hasAnyValidRect(rects)) {
      return isPointInRectList(clientX, clientY, rects);
    }
  }

  return false;
}

function collectDirectChildTextNodes(el: TargetElementLike): Text[] {
  if (!el.childNodes) {
    return [];
  }
  const nodes: Text[] = [];
  for (const child of Array.from(el.childNodes)) {
    if (isNonEmptyTextNode(child)) {
      nodes.push(child);
    }
  }
  return nodes;
}

/**
 * 判断指针位置 `(clientX, clientY)` 是否命中可划选文本：
 * - 排除交互控件、选区工具条/缩略图侧栏及纯非文本容器；
 * - 优先使用 `caretRangeFromPoint` / `caretPositionFromPoint` 并校验字符矩形（含 3px 容差），避免吸附到远处空白行；
 * - 降级检查目标元素直接子 `Text` 节点的 `Range.getClientRects()`（无 `createRange` 的测试替身环境直接检查 `nodeType === 3`）。
 */
export function isPointOverText(
  target: EventTarget | null,
  clientX: number,
  clientY: number,
  container?: HTMLElement | null,
): boolean {
  if (!target || typeof target !== 'object') {
    return false;
  }
  if (container && target === (container as unknown as EventTarget)) {
    return false;
  }

  const isDirectTextTarget = (target as { nodeType?: unknown }).nodeType === 3;
  const el = resolveTargetElement(target);
  if (!el && !isDirectTextTarget) {
    return false;
  }
  if (el && container && el === (container as unknown as TargetElementLike)) {
    return false;
  }

  if (el) {
    if (
      isInteractiveTarget(el as unknown as EventTarget, container) ||
      isExcludedUiTarget(el, container) ||
      isNonTextElement(el)
    ) {
      return false;
    }
  }

  const doc = (typeof document !== 'undefined' ? document : undefined) as
    | DocumentWithCaretApis
    | undefined;

  if (doc && Number.isFinite(clientX) && Number.isFinite(clientY)) {
    if (typeof doc.caretRangeFromPoint === 'function') {
      try {
        const range = doc.caretRangeFromPoint(clientX, clientY);
        if (
          range &&
          hitTestCaretNode(
            doc,
            range.startContainer,
            range.startOffset,
            () => range.getClientRects?.(),
            clientX,
            clientY,
            container,
          )
        ) {
          return true;
        }
      } catch {
        // 忽略底层 DOM API 异常并进入降级分支
      }
    } else if (typeof doc.caretPositionFromPoint === 'function') {
      try {
        const pos = doc.caretPositionFromPoint(clientX, clientY);
        if (
          pos &&
          hitTestCaretNode(
            doc,
            pos.offsetNode,
            pos.offset,
            pos.getClientRect ? () => {
              const rect = pos.getClientRect?.();
              return rect ? [rect] : [];
            } : null,
            clientX,
            clientY,
            container,
          )
        ) {
          return true;
        }
      } catch {
        // 忽略底层 DOM API 异常并进入降级分支
      }
    }
  }

  // 降级检查目标自身的直接子 Text 节点（或目标本身即为 Text 节点）
  const textNodes: Text[] = isDirectTextTarget
    ? isNonEmptyTextNode(target)
      ? [target]
      : []
    : el
      ? collectDirectChildTextNodes(el)
      : [];

  if (textNodes.length === 0) {
    return false;
  }

  if (!doc || typeof doc.createRange !== 'function') {
    return true;
  }

  let hasMeasurableRect = false;
  for (const textNode of textNodes) {
    try {
      const range = doc.createRange();
      range.selectNodeContents(textNode);
      const rects = range.getClientRects?.();
      if (hasAnyValidRect(rects)) {
        hasMeasurableRect = true;
        if (isPointInRectList(clientX, clientY, rects)) {
          return true;
        }
      }
    } catch {
      // 若在部分测试替身中无法对 mock Text 调用 selectNodeContents，则回退为命中
      if (!hasMeasurableRect) {
        return true;
      }
    }
  }

  return false;
}

function isOnScrollbarTrack(container: HTMLElement, e: MouseEvent): boolean {
  if (container.clientWidth <= 0 || typeof container.getBoundingClientRect !== 'function') {
    return false;
  }
  const rect = container.getBoundingClientRect();
  if (e.clientX >= rect.left + container.clientWidth) {
    return true;
  }
  if (container.clientHeight > 0 && e.clientY >= rect.top + container.clientHeight) {
    return true;
  }
  return false;
}

/**
 * 创建鼠标拖拽平移控制器。
 */
export function createDragPanner(options: DragPannerOptions): DragPannerController {
  const dragThresholdPx =
    typeof options.dragThresholdPx === 'number' &&
    Number.isFinite(options.dragThresholdPx) &&
    options.dragThresholdPx >= 0
      ? options.dragThresholdPx
      : DEFAULT_DRAG_THRESHOLD_PX;

  const eventTarget =
    options.eventTarget ?? (typeof window !== 'undefined' ? window : undefined);

  let destroyed = false;
  let isPressed = false;
  let hasDragged = false;
  let suppressNextClick = false;
  let listenersBound = false;

  let startX = 0;
  let startY = 0;
  let startScrollLeft = 0;
  let startScrollTop = 0;

  let activeContainer: HTMLElement | null = null;
  let lastContainer: HTMLElement | null = null;
  let clickSuppressTimer: ReturnType<typeof setTimeout> | null = null;

  function clearClickSuppressTimer(): void {
    if (clickSuppressTimer !== null) {
      clearTimeout(clickSuppressTimer);
      clickSuppressTimer = null;
    }
  }

  function scheduleClickSuppressReset(): void {
    clearClickSuppressTimer();
    clickSuppressTimer = setTimeout(() => {
      clickSuppressTimer = null;
      suppressNextClick = false;
    }, CLICK_SUPPRESS_TIMEOUT_MS);
    if (
      typeof clickSuppressTimer === 'object' &&
      clickSuppressTimer !== null &&
      'unref' in clickSuppressTimer &&
      typeof (clickSuppressTimer as { unref?: unknown }).unref === 'function'
    ) {
      (clickSuppressTimer as { unref: () => void }).unref();
    }
  }

  const onTargetMouseMove = (event: Event): void => {
    handleMouseMove(event as MouseEvent);
  };

  const onTargetMouseUp = (event: Event): void => {
    handleMouseUp(event as MouseEvent);
  };

  const onTargetBlur = (): void => {
    clearClickSuppressTimer();
    suppressNextClick = false;
    stopDrag(false);
  };

  function bindGlobalListeners(): void {
    if (!eventTarget || listenersBound) {
      return;
    }
    listenersBound = true;
    eventTarget.addEventListener('mousemove', onTargetMouseMove);
    eventTarget.addEventListener('mouseup', onTargetMouseUp);
    eventTarget.addEventListener('blur', onTargetBlur);
  }

  function unbindGlobalListeners(): void {
    if (!eventTarget || !listenersBound) {
      return;
    }
    listenersBound = false;
    eventTarget.removeEventListener('mousemove', onTargetMouseMove);
    eventTarget.removeEventListener('mouseup', onTargetMouseUp);
    eventTarget.removeEventListener('blur', onTargetBlur);
  }

  function clearContainerClasses(el: HTMLElement | null | undefined): void {
    removeClass(el, CLASS_PANNABLE);
    removeClass(el, CLASS_PANNING);
    removeClass(el, CLASS_OVER_TEXT);
  }

  function stopDrag(shouldSuppressClick: boolean): void {
    const didDrag = hasDragged;
    isPressed = false;
    hasDragged = false;
    unbindGlobalListeners();

    const currentContainer = options.getScrollContainer();
    removeClass(activeContainer, CLASS_PANNING);
    if (currentContainer && currentContainer !== activeContainer) {
      removeClass(currentContainer, CLASS_PANNING);
    }
    activeContainer = null;

    if (shouldSuppressClick && didDrag) {
      suppressNextClick = true;
      scheduleClickSuppressReset();
    }
  }

  function handleMouseDown(e: MouseEvent): void {
    if (destroyed) {
      return;
    }
    if (e.defaultPrevented) {
      return;
    }
    if (e.button !== 0 && e.button !== 1) {
      return;
    }
    if (e.ctrlKey || e.metaKey || e.altKey) {
      return;
    }

    const container = options.getScrollContainer();
    if (!options.isEnabled() || !container) {
      clearContainerClasses(container);
      if (lastContainer && lastContainer !== container) {
        clearContainerClasses(lastContainer);
      }
      stopDrag(false);
      return;
    }

    if (isInteractiveTarget(e.target, container)) {
      return;
    }
    if (isOnScrollbarTrack(container, e)) {
      return;
    }

    // 左键按在文字上或外部判定不应平移时，放行浏览器原生划选文字；中键保持任意区域可平移
    if (e.button === 0) {
      if (
        options.shouldPanTarget?.(e.target, e) === false ||
        isPointOverText(e.target, e.clientX, e.clientY, container)
      ) {
        return;
      }
    }

    // 阻止鼠标中键按下时浏览器默认的自动滚屏罗盘图标
    if (e.button === 1) {
      e.preventDefault?.();
    }

    clearClickSuppressTimer();
    suppressNextClick = false;

    if (activeContainer && activeContainer !== container) {
      removeClass(activeContainer, CLASS_PANNING);
    }

    lastContainer = container;
    activeContainer = container;
    isPressed = true;
    hasDragged = false;
    startX = e.clientX;
    startY = e.clientY;
    startScrollLeft = container.scrollLeft;
    startScrollTop = container.scrollTop;

    bindGlobalListeners();
  }

  function handleMouseMove(e: MouseEvent): void {
    if (!isPressed || destroyed) {
      return;
    }

    const container = options.getScrollContainer();
    const enabled = options.isEnabled();
    if (!container || !enabled || (activeContainer && container !== activeContainer)) {
      if (!enabled) {
        removeClass(container, CLASS_PANNABLE);
        removeClass(container, CLASS_OVER_TEXT);
        removeClass(activeContainer, CLASS_PANNABLE);
        removeClass(activeContainer, CLASS_OVER_TEXT);
      }
      stopDrag(false);
      return;
    }

    const dx = e.clientX - startX;
    const dy = e.clientY - startY;

    if (!hasDragged) {
      if (Math.hypot(dx, dy) < dragThresholdPx) {
        return;
      }
      hasDragged = true;
      removeClass(container, CLASS_OVER_TEXT);
      addClass(container, CLASS_PANNING);
    }

    e.preventDefault?.();
    container.scrollLeft = startScrollLeft - dx;
    container.scrollTop = startScrollTop - dy;
  }

  function handleMouseUp(_e?: MouseEvent): void {
    if (!isPressed) {
      unbindGlobalListeners();
      removeClass(activeContainer, CLASS_PANNING);
      removeClass(options.getScrollContainer(), CLASS_PANNING);
      activeContainer = null;
      return;
    }
    stopDrag(true);
  }

  function handleHoverMove(e: MouseEvent): void {
    if (destroyed || isPressed || (e.buttons & 1) !== 0) {
      return;
    }

    const container = options.getScrollContainer();
    if (lastContainer && lastContainer !== container) {
      clearContainerClasses(lastContainer);
    }
    lastContainer = container;

    if (!container || !options.isEnabled()) {
      removeClass(container, CLASS_OVER_TEXT);
      removeClass(container, CLASS_PANNABLE);
      return;
    }

    if (isPointOverText(e.target, e.clientX, e.clientY, container)) {
      addClass(container, CLASS_OVER_TEXT);
    } else {
      removeClass(container, CLASS_OVER_TEXT);
    }
  }

  function handleMouseLeave(): void {
    const container = options.getScrollContainer();
    removeClass(container, CLASS_OVER_TEXT);
    if (lastContainer && lastContainer !== container) {
      removeClass(lastContainer, CLASS_OVER_TEXT);
    }
  }

  function handleClickCapture(e: MouseEvent): void {
    if (!suppressNextClick) {
      return;
    }
    suppressNextClick = false;
    clearClickSuppressTimer();
    e.preventDefault?.();
    e.stopPropagation?.();
  }

  function syncState(): void {
    const container = options.getScrollContainer();
    if (lastContainer && lastContainer !== container) {
      clearContainerClasses(lastContainer);
    }
    lastContainer = container;

    if (destroyed || !container || !options.isEnabled()) {
      clearContainerClasses(container);
      if (isPressed || hasDragged || activeContainer) {
        stopDrag(false);
      }
      return;
    }

    addClass(container, CLASS_PANNABLE);
  }

  function reset(): void {
    clearClickSuppressTimer();
    suppressNextClick = false;
    const prevLast = lastContainer;
    stopDrag(false);
    removeClass(prevLast, CLASS_PANNING);
    removeClass(prevLast, CLASS_OVER_TEXT);
    removeClass(options.getScrollContainer(), CLASS_OVER_TEXT);
  }

  function destroy(): void {
    destroyed = true;
    clearClickSuppressTimer();
    suppressNextClick = false;
    const prevLast = lastContainer;
    const prevActive = activeContainer;
    stopDrag(false);

    const currentContainer = options.getScrollContainer();
    clearContainerClasses(currentContainer);
    clearContainerClasses(prevActive);
    clearContainerClasses(prevLast);
    lastContainer = null;
  }

  return {
    handleMouseDown,
    handleMouseMove,
    handleMouseUp,
    handleHoverMove,
    handleMouseLeave,
    handleClickCapture,
    syncState,
    reset,
    destroy,
    isPanning: () => isPressed && hasDragged,
  };
}
