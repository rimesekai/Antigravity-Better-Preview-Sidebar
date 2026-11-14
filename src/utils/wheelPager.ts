export interface WheelPagerOptions {
  /** 获取主内容滚动容器，用于判断放大后是否到达顶部/底部边界。 */
  getScrollContainer: () => HTMLElement | null;
  /** 判断当前是否允许滚轮翻页（通常仅在单页模式启用）。 */
  isEnabled: () => boolean;
  /** 获取当前页码与总页数（1-based）。 */
  getPageInfo: () => { currentPage: number; totalPages: number };
  /** 切换到上一页。 */
  onPrevPage: () => void;
  /** 切换到下一页。 */
  onNextPage: () => void;
  /** 触发翻页的累计垂直滚动像素阈值，默认 45。 */
  deltaThreshold?: number;
  /** 两次翻页之间的最小冷却时间（毫秒），默认 120。 */
  cooldownMs?: number;
  /** 停止滚动后重置累计值与惯性状态的超时时间（毫秒），默认 180。 */
  resetTimeoutMs?: number;
  /** 可选时间戳提供器，便于单元测试注入虚拟时钟。 */
  now?: () => number;
}

export interface WheelPagerController {
  handleWheel: (event: WheelEvent) => void;
  reset: () => void;
  destroy: () => void;
}

const DEFAULT_DELTA_THRESHOLD = 45;
const DEFAULT_COOLDOWN_MS = 120;
const DEFAULT_RESET_TIMEOUT_MS = 180;
const EDGE_TOLERANCE_PX = 2;
const INERTIA_GAP_RESET_MS = 85;
const DISCRETE_STEP_MIN_DELTA = 80;
const DISCRETE_STEP_MIN_GAP_MS = 40;

function normalizeWheelDeltaY(event: WheelEvent): number {
  if (event.deltaMode === 1) {
    return event.deltaY * 32;
  }
  if (event.deltaMode === 2) {
    return event.deltaY * 400;
  }
  return event.deltaY;
}

function canScrollInDirection(container: HTMLElement, direction: -1 | 1): boolean {
  if (container.scrollHeight <= container.clientHeight + EDGE_TOLERANCE_PX) {
    return false;
  }
  if (direction > 0) {
    return container.scrollTop + container.clientHeight < container.scrollHeight - EDGE_TOLERANCE_PX;
  }
  return container.scrollTop > EDGE_TOLERANCE_PX;
}

/**
 * 创建单页模式鼠标滚轮翻页控制器：
 * - 放大出现垂直滚动条时先滚至顶/底边缘再触发翻页；
 * - 抑制触控板惯性余波导致的连续跳页，同时保留离散滚轮的跟手体验。
 */
export function createWheelPager(options: WheelPagerOptions): WheelPagerController {
  const deltaThreshold = options.deltaThreshold ?? DEFAULT_DELTA_THRESHOLD;
  const cooldownMs = options.cooldownMs ?? DEFAULT_COOLDOWN_MS;
  const resetTimeoutMs = options.resetTimeoutMs ?? DEFAULT_RESET_TIMEOUT_MS;
  const nowFn = options.now ?? (() => Date.now());

  let accum = 0;
  let lastDirection: -1 | 0 | 1 = 0;
  let lastTriggerTime = Number.NEGATIVE_INFINITY;
  let lastEventTime = Number.NEGATIVE_INFINITY;
  let inInertia = false;
  let resetTimer: ReturnType<typeof setTimeout> | null = null;

  function clearResetTimer(): void {
    if (resetTimer !== null) {
      clearTimeout(resetTimer);
      resetTimer = null;
    }
  }

  function reset(): void {
    clearResetTimer();
    accum = 0;
    lastDirection = 0;
    inInertia = false;
  }

  function scheduleReset(): void {
    clearResetTimer();
    resetTimer = setTimeout(() => {
      resetTimer = null;
      accum = 0;
      lastDirection = 0;
      inInertia = false;
    }, resetTimeoutMs);
  }

  function handleWheel(event: WheelEvent): void {
    if (!options.isEnabled()) return;
    if (event.defaultPrevented || event.ctrlKey || event.metaKey) return;

    const { currentPage, totalPages } = options.getPageInfo();
    if (totalPages <= 1) return;

    // 放大后左右滑动查看内容时不触发垂直翻页
    if (Math.abs(event.deltaX) > Math.abs(event.deltaY)) return;

    const dy = normalizeWheelDeltaY(event);
    if (!Number.isFinite(dy) || dy === 0) return;

    const dir: -1 | 1 = dy > 0 ? 1 : -1;
    const container = options.getScrollContainer();
    if (container && canScrollInDirection(container, dir)) {
      reset();
      return;
    }

    // 首末页边界不再继续拦截滚轮
    if ((dir < 0 && currentPage <= 1) || (dir > 0 && currentPage >= totalPages)) {
      reset();
      return;
    }

    const now = nowFn();
    const gap = now - lastEventTime;
    lastEventTime = now;

    if (dir !== lastDirection) {
      accum = 0;
      inInertia = false;
      lastDirection = dir;
    }

    if (inInertia) {
      const isNewGesture =
        gap > INERTIA_GAP_RESET_MS ||
        (Math.abs(dy) >= DISCRETE_STEP_MIN_DELTA && gap > DISCRETE_STEP_MIN_GAP_MS);
      if (!isNewGesture) {
        event.preventDefault();
        scheduleReset();
        return;
      }
      inInertia = false;
    }

    if (now - lastTriggerTime < cooldownMs) {
      event.preventDefault();
      scheduleReset();
      return;
    }

    accum += dy;
    scheduleReset();

    if (Math.abs(accum) < deltaThreshold) {
      return;
    }

    event.preventDefault();
    accum = 0;
    lastTriggerTime = now;
    inInertia = true;

    if (dir > 0) {
      options.onNextPage();
    } else {
      options.onPrevPage();
    }
  }

  return {
    handleWheel,
    reset,
    destroy: reset,
  };
}
