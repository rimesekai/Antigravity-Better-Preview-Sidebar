import { updateCachedWorkspaceUris } from './sidecarClient';

export type ThemeMode = 'dark' | 'light';
export type ThemePreference = 'auto' | 'light' | 'dark';
export type SeedId = 'auto-pair' | 'azure' | 'amber' | 'emerald' | 'violet' | 'rose' | 'cyan';
export type ColorStyle = 'vivid' | 'tinted' | 'soft' | 'monochrome';

export interface SeedPreset {
  id: Exclude<SeedId, 'auto-pair'>;
  nameZh: string;
  nameEn: string;
  hex: string;
  h: number;
  s: number;
  lLight: number;
  lDark: number;
}

export interface AppearancePrefs {
  seedId: SeedId;
  colorStyle: ColorStyle;
  pureBlack: boolean;
  liquidGlass: boolean;
}

export const SEED_PRESETS: readonly SeedPreset[] = [
  { id: 'azure', nameZh: '经典蓝', nameEn: 'Azure', hex: '#0A7AFF', h: 212, s: 100, lLight: 52, lDark: 63 },
  { id: 'amber', nameZh: '暖琥珀', nameEn: 'Amber', hex: '#F59E3D', h: 32, s: 90, lLight: 48, lDark: 62 },
  { id: 'emerald', nameZh: '翡翠绿', nameEn: 'Emerald', hex: '#10B981', h: 160, s: 84, lLight: 39, lDark: 55 },
  { id: 'violet', nameZh: '紫罗兰', nameEn: 'Violet', hex: '#7C3AED', h: 262, s: 83, lLight: 58, lDark: 69 },
  { id: 'rose', nameZh: '珊瑚红', nameEn: 'Rose', hex: '#F43F5E', h: 350, s: 89, lLight: 56, lDark: 65 },
  { id: 'cyan', nameZh: '北欧青', nameEn: 'Cyan', hex: '#06B6D4', h: 189, s: 94, lLight: 41, lDark: 56 },
];

const THEME_PREF_STORAGE_KEY = 'file-preview:theme-preference';
const APPEARANCE_PREFS_STORAGE_KEY = 'file-preview:appearance-prefs';

const DEFAULT_APPEARANCE_PREFS: AppearancePrefs = {
  seedId: 'auto-pair',
  colorStyle: 'vivid',
  pureBlack: false,
  liquidGlass: true,
};

type AppearanceListener = (prefs: AppearancePrefs, theme: ThemeMode, pref: ThemePreference) => void;
const appearanceListeners = new Set<AppearanceListener>();

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
  if (typeof window !== 'undefined' && typeof window.matchMedia === 'function') {
    try {
      if (window.matchMedia('(prefers-color-scheme: dark)').matches) {
        return 'dark';
      }
    } catch {
      // 忽略受限媒体查询异常
    }
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
 * 读取本地保存的外观定制偏好（种子色、色彩风格、纯黑 OLED、Liquid Glass）。
 */
export function getAppearancePrefs(): AppearancePrefs {
  try {
    const raw = window.localStorage.getItem(APPEARANCE_PREFS_STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw) as Partial<AppearancePrefs> | null;
      if (parsed && typeof parsed === 'object') {
        const validSeed =
          parsed.seedId === 'auto-pair' || SEED_PRESETS.some((p) => p.id === parsed.seedId);
        const validStyle =
          parsed.colorStyle === 'vivid' ||
          parsed.colorStyle === 'tinted' ||
          parsed.colorStyle === 'soft' ||
          parsed.colorStyle === 'monochrome';
        return {
          seedId: validSeed ? parsed.seedId! : DEFAULT_APPEARANCE_PREFS.seedId,
          colorStyle: validStyle ? parsed.colorStyle! : DEFAULT_APPEARANCE_PREFS.colorStyle,
          pureBlack:
            typeof parsed.pureBlack === 'boolean'
              ? parsed.pureBlack
              : DEFAULT_APPEARANCE_PREFS.pureBlack,
          liquidGlass:
            typeof parsed.liquidGlass === 'boolean'
              ? parsed.liquidGlass
              : DEFAULT_APPEARANCE_PREFS.liquidGlass,
        };
      }
    }
  } catch {
    // 忽略受限存储或 JSON 解析异常
  }
  return { ...DEFAULT_APPEARANCE_PREFS };
}

function notifyAppearanceListeners(resolvedTheme: ThemeMode): void {
  const prefs = getAppearancePrefs();
  const pref = getThemePreference();
  for (const listener of appearanceListeners) {
    try {
      listener(prefs, resolvedTheme, pref);
    } catch (err) {
      console.error('[themeSync] appearance listener error:', err);
    }
  }
}

/**
 * 将当前外观偏好（种子色动态变量、色彩衍生风格、纯黑 OLED 与 Liquid Glass）应用到根元素。
 */
export function applyAppearancePrefs(resolvedTheme: ThemeMode): AppearancePrefs {
  const root = document.documentElement;
  const prefs = getAppearancePrefs();

  root.dataset.seedId = prefs.seedId;
  root.dataset.colorStyle = prefs.colorStyle;
  root.dataset.pureBlack = String(prefs.pureBlack);
  root.dataset.liquidGlass = String(prefs.liquidGlass);

  const seed =
    prefs.seedId === 'auto-pair'
      ? resolvedTheme === 'dark'
        ? SEED_PRESETS[1]
        : SEED_PRESETS[0]
      : SEED_PRESETS.find((item) => item.id === prefs.seedId) ?? SEED_PRESETS[0];

  root.style.setProperty('--fc-seed-h', String(seed.h));
  root.style.setProperty('--fc-seed-s', `${seed.s}%`);
  root.style.setProperty('--fc-seed-l-light', `${seed.lLight}%`);
  root.style.setProperty('--fc-seed-l-dark', `${seed.lDark}%`);

  return prefs;
}

/**
 * 更新并持久化外观定制偏好，并立即同步到文档根节点。
 */
export function updateAppearancePrefs(partial: Partial<AppearancePrefs>): AppearancePrefs {
  const next: AppearancePrefs = {
    ...getAppearancePrefs(),
    ...partial,
  };
  try {
    window.localStorage.setItem(APPEARANCE_PREFS_STORAGE_KEY, JSON.stringify(next));
  } catch {
    // 忽略受限存储环境异常
  }
  const resolved = syncThemeAttribute();
  notifyAppearanceListeners(resolved);
  return next;
}

/**
 * 订阅外观与主题变更，返回取消订阅函数。
 */
export function subscribeAppearanceChanges(listener: AppearanceListener): () => void {
  appearanceListeners.add(listener);
  return () => {
    appearanceListeners.delete(listener);
  };
}

/**
 * 保存主题偏好并立即应用到根元素。
 */
export function setThemePreference(pref: ThemePreference): ThemeMode {
  try {
    if (pref === 'auto') {
      window.localStorage.removeItem(THEME_PREF_STORAGE_KEY);
    } else {
      window.localStorage.setItem(THEME_PREF_STORAGE_KEY, pref);
    }
  } catch {
    // 忽略受限存储环境异常
  }
  const resolved = syncThemeAttribute();
  notifyAppearanceListeners(resolved);
  return resolved;
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
 * 根据用户偏好或宿主 CSS 变量（--background / --card / --foreground）同步更新 data-theme、meta[color-scheme] 与设计令牌。
 */
export function syncThemeAttribute(): ThemeMode {
  const root = document.documentElement;
  const pref = getThemePreference();
  const meta = document.querySelector('meta[name="color-scheme"]');

  if (pref === 'light' || pref === 'dark') {
    root.dataset.themeOverride = pref;
    if (root.dataset.theme !== pref) {
      root.dataset.theme = pref;
    }
    if (root.dataset.resolvedTheme !== pref) {
      root.dataset.resolvedTheme = pref;
    }
    if (meta && meta.getAttribute('content') !== pref) {
      meta.setAttribute('content', pref);
    }
    applyAppearancePrefs(pref);
    return pref;
  }

  delete root.dataset.themeOverride;
  if (meta && meta.getAttribute('content') !== 'light dark') {
    meta.setAttribute('content', 'light dark');
  }

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
  if (root.dataset.resolvedTheme !== resolvedTheme) {
    root.dataset.resolvedTheme = resolvedTheme;
  }

  applyAppearancePrefs(resolvedTheme);
  return resolvedTheme;
}

/**
 * 启动宿主主题同步监听（监听 root style/class 属性变化、系统媒体偏好、跨标签存储与宿主 theme-change 消息）。
 */
export function startThemeSync(onThemeChange?: (theme: ThemeMode) => void): () => void {
  let lastTheme = syncThemeAttribute();
  onThemeChange?.(lastTheme);

  const handleUpdate = () => {
    const current = syncThemeAttribute();
    notifyAppearanceListeners(current);
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

  const handleStorage = (event: StorageEvent) => {
    if (
      event.key === THEME_PREF_STORAGE_KEY ||
      event.key === APPEARANCE_PREFS_STORAGE_KEY ||
      event.key === null
    ) {
      handleUpdate();
    }
  };

  let colorMedia: MediaQueryList | null = null;
  if (typeof window.matchMedia === 'function') {
    try {
      colorMedia = window.matchMedia('(prefers-color-scheme: dark)');
      colorMedia.addEventListener?.('change', handleUpdate);
    } catch {
      colorMedia = null;
    }
  }

  window.addEventListener('message', handleWindowMessage);
  window.addEventListener('storage', handleStorage);

  return () => {
    observer.disconnect();
    window.removeEventListener('message', handleWindowMessage);
    window.removeEventListener('storage', handleStorage);
    colorMedia?.removeEventListener?.('change', handleUpdate);
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
