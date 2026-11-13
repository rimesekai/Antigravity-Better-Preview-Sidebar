import { updateCachedWorkspaceUris } from './sidecarClient';

export type ThemeMode = 'dark' | 'light';
export type ThemePreference = 'auto' | 'light' | 'dark';

const THEME_PREF_STORAGE_KEY = 'file-preview:theme-preference';

interface SidecarWorkspaceWindowLike {
  sidecar?: {
    onWorkspaceChange?: (listener: (uris: string[]) => void) => () => void;
  };
}

interface RgbColor {
  r: number;
  g: number;
  b: number;
}

/**
 * 解析 CSS 颜色值（支持 #rgb、#rrggbb、#rgba、#rrggbbaa、rgb(...)、rgba(...)、color(srgb ...)）。
 */
export function parseCssColorToRgb(raw: string): RgbColor | null {
  const value = raw.trim().toLowerCase();
  if (!value || value === 'transparent') {
    return null;
  }

  if (value.startsWith('#')) {
    const hex = value.slice(1);
    if (hex.length === 3 || hex.length === 4) {
      if (hex.length === 4) {
        const a = Number.parseInt(hex[3] + hex[3], 16) / 255;
        if (Number.isNaN(a) || a <= 0.05) return null;
      }
      const r = Number.parseInt(hex[0] + hex[0], 16);
      const g = Number.parseInt(hex[1] + hex[1], 16);
      const b = Number.parseInt(hex[2] + hex[2], 16);
      if ([r, g, b].some(Number.isNaN)) return null;
      return { r, g, b };
    }
    if (hex.length === 6 || hex.length === 8) {
      if (hex.length === 8) {
        const a = Number.parseInt(hex.slice(6, 8), 16) / 255;
        if (Number.isNaN(a) || a <= 0.05) return null;
      }
      const r = Number.parseInt(hex.slice(0, 2), 16);
      const g = Number.parseInt(hex.slice(2, 4), 16);
      const b = Number.parseInt(hex.slice(4, 6), 16);
      if ([r, g, b].some(Number.isNaN)) return null;
      return { r, g, b };
    }
    return null;
  }

  const rgbMatch = /^rgba?\(([^)]+)\)$/.exec(value);
  if (rgbMatch) {
    const [colorPart, slashAlpha] = rgbMatch[1].split('/');
    const inner = colorPart.trim();
    const parts = inner.includes(',')
      ? inner.split(',').map((item) => item.trim())
      : inner.split(/\s+/).map((item) => item.trim());

    if (parts.length < 3) {
      return null;
    }

    const alphaRaw = slashAlpha?.trim() ?? (parts.length >= 4 ? parts[3] : undefined);
    if (alphaRaw !== undefined) {
      const alpha = alphaRaw.endsWith('%')
        ? Number.parseFloat(alphaRaw.slice(0, -1)) / 100
        : Number.parseFloat(alphaRaw);
      if (!Number.isFinite(alpha) || alpha <= 0.05) {
        return null;
      }
    }

    const channels = parts.slice(0, 3).map((part) => {
      if (part.endsWith('%')) {
        const pct = Number.parseFloat(part.slice(0, -1));
        return Number.isFinite(pct) ? Math.round((pct / 100) * 255) : Number.NaN;
      }
      return Number.parseFloat(part);
    });

    if (channels.some((c) => !Number.isFinite(c))) {
      return null;
    }

    return {
      r: Math.min(255, Math.max(0, channels[0])),
      g: Math.min(255, Math.max(0, channels[1])),
      b: Math.min(255, Math.max(0, channels[2])),
    };
  }

  const srgbMatch = /^color\(srgb\s+([^)]+)\)$/.exec(value);
  if (srgbMatch) {
    const [colorPart, slashAlpha] = srgbMatch[1].split('/');
    if (slashAlpha !== undefined) {
      const alphaRaw = slashAlpha.trim();
      const alpha = alphaRaw.endsWith('%')
        ? Number.parseFloat(alphaRaw.slice(0, -1)) / 100
        : Number.parseFloat(alphaRaw);
      if (!Number.isFinite(alpha) || alpha <= 0.05) {
        return null;
      }
    }

    const inner = colorPart.trim();
    const parts = inner.split(/\s+/).map((item) => item.trim()).filter(Boolean);
    if (parts.length < 3) {
      return null;
    }

    const channels = parts.slice(0, 3).map((part) => {
      if (part.endsWith('%')) {
        const pct = Number.parseFloat(part.slice(0, -1));
        return Number.isFinite(pct) ? (pct / 100) * 255 : Number.NaN;
      }
      const unit = Number.parseFloat(part);
      return Number.isFinite(unit) ? unit * 255 : Number.NaN;
    });

    if (channels.some((c) => !Number.isFinite(c))) {
      return null;
    }

    return {
      r: Math.min(255, Math.max(0, channels[0])),
      g: Math.min(255, Math.max(0, channels[1])),
      b: Math.min(255, Math.max(0, channels[2])),
    };
  }

  return null;
}

function linearizeChannel(channel255: number): number {
  const srgb = channel255 / 255;
  return srgb <= 0.03928 ? srgb / 12.92 : Math.pow((srgb + 0.055) / 1.055, 2.4);
}

/**
 * 计算 sRGB 颜色的 WCAG 相对亮度（0 = 纯黑，1 = 纯白）。
 */
export function computeRelativeLuminance(rgb: RgbColor): number {
  const r = linearizeChannel(rgb.r);
  const g = linearizeChannel(rgb.g);
  const b = linearizeChannel(rgb.b);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

function resolveFallbackTheme(root: HTMLElement): ThemeMode {
  if (root.classList.contains('dark')) {
    return 'dark';
  }
  if (root.classList.contains('light')) {
    return 'light';
  }
  return 'light';
}

function probeCssVarColor(root: HTMLElement, cssVarName: string, rawValue: string): RgbColor | null {
  const probe = document.createElement('span');
  probe.style.cssText = `position:absolute;visibility:hidden;pointer-events:none;color:var(${cssVarName}, transparent)`;
  root.appendChild(probe);

  let computedColor = '';
  try {
    computedColor = getComputedStyle(probe).color;
  } finally {
    probe.remove();
  }

  return parseCssColorToRgb(computedColor) ?? parseCssColorToRgb(rawValue);
}

/**
 * 读取本地保存的主题偏好（auto 跟随宿主 / light 浅色 / dark 深色）。
 */
export function getThemePreference(): ThemePreference {
  try {
    const saved = window.localStorage.getItem(THEME_PREF_STORAGE_KEY);
    if (saved === 'light' || saved === 'dark' || saved === 'auto') {
      return saved;
    }
  } catch {
    // 忽略受限存储环境异常
  }
  return 'auto';
}

/**
 * 保存主题偏好并立即应用到根元素。
 */
export function setThemePreference(pref: ThemePreference): ThemeMode {
  try {
    window.localStorage.setItem(THEME_PREF_STORAGE_KEY, pref);
  } catch {
    // 忽略受限存储环境异常
  }
  return syncThemeAttribute();
}

/**
 * 在 auto -> light -> dark 之间循环切换主题偏好。
 */
export function cycleThemePreference(): { preference: ThemePreference; theme: ThemeMode } {
  const current = getThemePreference();
  const next: ThemePreference =
    current === 'auto' ? 'light' : current === 'light' ? 'dark' : 'auto';
  const theme = setThemePreference(next);
  return { preference: next, theme };
}

/**
 * 根据用户偏好或宿主 CSS 变量（--background / --card / --foreground）同步更新 data-theme。
 */
export function syncThemeAttribute(): ThemeMode {
  const root = document.documentElement;
  const pref = getThemePreference();

  if (pref === 'light' || pref === 'dark') {
    root.dataset.themeOverride = pref;
    if (root.dataset.theme !== pref) {
      root.dataset.theme = pref;
    }
    return pref;
  }

  delete root.dataset.themeOverride;

  const computed = getComputedStyle(root);
  const rawBackground = computed.getPropertyValue('--background').trim();
  const rawCard = computed.getPropertyValue('--card').trim();
  const rawForeground = computed.getPropertyValue('--foreground').trim();

  let nextTheme: ThemeMode | null = null;

  if (rawBackground) {
    const bgRgb = probeCssVarColor(root, '--background', rawBackground);
    if (bgRgb) {
      nextTheme = computeRelativeLuminance(bgRgb) < 0.5 ? 'dark' : 'light';
    }
  }

  if (!nextTheme && rawCard) {
    const cardRgb = probeCssVarColor(root, '--card', rawCard);
    if (cardRgb) {
      nextTheme = computeRelativeLuminance(cardRgb) < 0.5 ? 'dark' : 'light';
    }
  }

  if (!nextTheme && rawForeground) {
    const fgRgb = probeCssVarColor(root, '--foreground', rawForeground);
    if (fgRgb) {
      nextTheme = computeRelativeLuminance(fgRgb) >= 0.5 ? 'dark' : 'light';
    }
  }

  const resolvedTheme = nextTheme ?? resolveFallbackTheme(root);

  if (root.dataset.theme !== resolvedTheme) {
    root.dataset.theme = resolvedTheme;
  }

  return resolvedTheme;
}

/**
 * 启动宿主主题同步监听（监听 root style/class 属性变化与宿主 theme-change 消息）。
 */
export function startThemeSync(onThemeChange?: (theme: ThemeMode) => void): () => void {
  let lastTheme = syncThemeAttribute();
  onThemeChange?.(lastTheme);

  const handleUpdate = () => {
    const current = syncThemeAttribute();
    if (current !== lastTheme) {
      lastTheme = current;
      onThemeChange?.(current);
    }
  };

  const observer = new MutationObserver(() => {
    handleUpdate();
  });

  observer.observe(document.documentElement, {
    attributes: true,
    attributeFilter: ['style', 'class'],
  });

  const handleWindowMessage = (event: MessageEvent) => {
    const data = event.data as { type?: unknown } | null | undefined;
    if (data && typeof data === 'object' && data.type === 'theme-change') {
      handleUpdate();
    }
  };

  window.addEventListener('message', handleWindowMessage);

  return () => {
    observer.disconnect();
    window.removeEventListener('message', handleWindowMessage);
  };
}

/**
 * 订阅宿主工作区切换事件，更新本地缓存并触发列表刷新回调。
 */
export function subscribeWorkspaceChanges(
  onWorkspaceRefresh: (workspaceUris: string[]) => void,
): () => void {
  const sidecar = (window as unknown as SidecarWorkspaceWindowLike).sidecar;
  if (!sidecar || typeof sidecar.onWorkspaceChange !== 'function') {
    return () => {};
  }

  return sidecar.onWorkspaceChange((uris) => {
    const normalized = Array.isArray(uris) ? uris : [];
    updateCachedWorkspaceUris(normalized);
    onWorkspaceRefresh(normalized);
  });
}
