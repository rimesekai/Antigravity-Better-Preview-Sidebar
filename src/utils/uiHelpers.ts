/**
 * UI 交互与可访问性通用辅助函数：
 * - 文件名主干与扩展名安全拆分（配合 single-line-overflow 保证扩展名永远可见）
 * - IME 组合输入态安全判断（遵循 ime-safe-enter-submit 指南）
 * - Toolbar / Tablist 键盘方向键漫游焦点回退（当原生 focusgroup 不可用时启用）
 */

export function splitFileNameParts(fileName: string): { stem: string; ext: string } {
  const dotIdx = fileName.lastIndexOf('.');
  if (dotIdx <= 0 || dotIdx >= fileName.length - 1) {
    return { stem: fileName, ext: '' };
  }
  return {
    stem: fileName.slice(0, dotIdx),
    ext: fileName.slice(dotIdx),
  };
}

/**
 * 将完整文件名渲染到单行容器内，保留扩展名在窄容器中不被省略号吞掉。
 */
export function renderTruncatableFileName(targetEl: HTMLElement, fileName: string): void {
  const { stem, ext } = splitFileNameParts(fileName);
  targetEl.textContent = '';
  const stemSpan = document.createElement('span');
  stemSpan.className = 'fp-filename-stem';
  stemSpan.textContent = stem;
  targetEl.appendChild(stemSpan);
  if (ext) {
    const extSpan = document.createElement('span');
    extSpan.className = 'fp-filename-ext';
    extSpan.textContent = ext;
    targetEl.appendChild(extSpan);
  }
}

/**
 * 判断键盘事件是否处于中文/日文/韩文等输入法选词组合态（含 Safari keyCode=229 兼容）。
 */
export function isImeComposing(event: KeyboardEvent): boolean {
  return event.isComposing || event.keyCode === 229;
}

/**
 * 为工具栏启用键盘漫游焦点（仅当浏览器不支持原生 focusgroup 时激活，且不拦截输入框内的方向键）。
 */
export function attachToolbarRovingFocus(toolbarEl: HTMLElement): () => void {
  toolbarEl.setAttribute('focusgroup', 'toolbar wrap');
  if (typeof HTMLElement !== 'undefined' && 'focusGroup' in HTMLElement.prototype) {
    return () => {};
  }

  const getItems = (): HTMLElement[] =>
    Array.from(
      toolbarEl.querySelectorAll<HTMLElement>('button:not([hidden]), input:not([hidden])'),
    ).filter((el) => el.offsetParent !== null && !el.closest('[hidden]'));

  const syncTabStops = (activeEl?: HTMLElement) => {
    const items = getItems();
    if (items.length === 0) return;
    const chosen = activeEl && items.includes(activeEl) ? activeEl : items.find((el) => el.tabIndex === 0) ?? items[0];
    for (const item of items) {
      item.tabIndex = item === chosen ? 0 : -1;
    }
  };

  syncTabStops();

  const onFocusIn = (e: FocusEvent) => {
    const target = e.target as HTMLElement | null;
    if (!target) return;
    const items = getItems();
    if (items.includes(target)) {
      syncTabStops(target);
    }
  };

  const onKeyDown = (e: KeyboardEvent) => {
    const active = document.activeElement as HTMLElement | null;
    if (!active || !toolbarEl.contains(active)) return;
    // 输入框内部保留左右箭头与 Home/End 编辑光标能力
    if (active.tagName === 'INPUT') {
      return;
    }

    const rtl = getComputedStyle(toolbarEl).direction === 'rtl';
    let step = 0;
    if (e.key === 'ArrowRight') step = rtl ? -1 : 1;
    else if (e.key === 'ArrowLeft') step = rtl ? 1 : -1;
    else if (e.key !== 'Home' && e.key !== 'End') return;

    const items = getItems();
    if (items.length === 0) return;

    const currentIdx = items.indexOf(active);
    let nextIdx = 0;
    if (e.key === 'Home') {
      nextIdx = 0;
    } else if (e.key === 'End') {
      nextIdx = items.length - 1;
    } else if (currentIdx >= 0) {
      nextIdx = (currentIdx + step + items.length) % items.length;
    }

    e.preventDefault();
    const nextItem = items[nextIdx];
    syncTabStops(nextItem);
    nextItem.focus();
  };

  toolbarEl.addEventListener('focusin', onFocusIn);
  toolbarEl.addEventListener('keydown', onKeyDown);

  return () => {
    toolbarEl.removeEventListener('focusin', onFocusIn);
    toolbarEl.removeEventListener('keydown', onKeyDown);
  };
}

/**
 * 为横排 Tablist 启用方向键切换（ArrowLeft / ArrowRight / Home / End）。
 */
export function attachTablistKeyboard(tablistEl: HTMLElement): () => void {
  tablistEl.setAttribute('focusgroup', 'tablist nomemory');

  const getTabs = (): HTMLButtonElement[] =>
    Array.from(tablistEl.querySelectorAll<HTMLButtonElement>('[role="tab"]:not([hidden])'));

  const syncTabs = () => {
    const tabs = getTabs();
    const selected = tabs.find((t) => t.getAttribute('aria-selected') === 'true') ?? tabs[0];
    for (const tab of tabs) {
      const isSelected = tab === selected;
      tab.toggleAttribute('focusgroupstart', isSelected);
      if (!('focusGroup' in HTMLElement.prototype)) {
        tab.tabIndex = isSelected ? 0 : -1;
      }
    }
  };

  syncTabs();

  const onKeyDown = (e: KeyboardEvent) => {
    if ('focusGroup' in HTMLElement.prototype) return;
    const active = document.activeElement as HTMLElement | null;
    if (!active || !tablistEl.contains(active)) return;

    const rtl = getComputedStyle(tablistEl).direction === 'rtl';
    let step = 0;
    if (e.key === 'ArrowRight') step = rtl ? -1 : 1;
    else if (e.key === 'ArrowLeft') step = rtl ? 1 : -1;
    else if (e.key !== 'Home' && e.key !== 'End') return;

    const tabs = getTabs();
    if (tabs.length === 0) return;

    const currentIdx = tabs.indexOf(active as HTMLButtonElement);
    let nextIdx = 0;
    if (e.key === 'Home') nextIdx = 0;
    else if (e.key === 'End') nextIdx = tabs.length - 1;
    else if (currentIdx >= 0) nextIdx = (currentIdx + step + tabs.length) % tabs.length;

    e.preventDefault();
    const targetTab = tabs[nextIdx];
    targetTab.focus();
    targetTab.click();
    syncTabs();
  };

  tablistEl.addEventListener('keydown', onKeyDown);
  return () => {
    tablistEl.removeEventListener('keydown', onKeyDown);
  };
}
