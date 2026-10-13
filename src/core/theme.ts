// Theme colours. A style colour is either a hex colour ("1F3A5F") or a theme slot ("accent1"). The
// engine resolves slots once, before layout, against the style's `theme` (the slot colours of the
// deck it is drawn for), so contrast and tints are computed on real colours; the scene keeps the
// slot next to the hex as "1F3A5F@accent1". Renderers that can (the .pptx writer, the add-in) emit
// the slot, so a chart follows its presentation's theme; the SVG preview draws the hex.
import type { Style } from './types.ts';

/** Theme slots a colour may name: the six accents and the four base colours of a PowerPoint theme. */
export const THEME_SLOTS = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'dk1', 'lt1', 'dk2', 'lt2'] as const;
/**
 * Slide-mapped names (PowerPoint's own for text and background): `tx1`/`bg1`/`tx2`/`bg2` follow the
 * slide master's colour map, so text stays readable on a dark-background master. They resolve to
 * dk1/lt1/dk2/lt2 under the usual mapping.
 */
export const MAPPED_SLOTS = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' } as const;
export type ThemeSlot = typeof THEME_SLOTS[number];
export type ColorName = ThemeSlot | keyof typeof MAPPED_SLOTS;
export const COLOR_NAMES: readonly ColorName[] = [...THEME_SLOTS, ...(Object.keys(MAPPED_SLOTS) as (keyof typeof MAPPED_SLOTS)[])];

/** Hex colour of each theme slot. */
export type ThemeColors = Record<ThemeSlot, string>;

/** The open-charts theme: the colours slots resolve to in previews and in decks the CLI writes. */
export const DEFAULT_THEME: ThemeColors = {
  dk1: '222222', lt1: 'FFFFFF', dk2: '1F3A5F', lt2: 'EEEEEE',
  accent1: '1F3A5F', accent2: '3E6A9E', accent3: '7FA3CC', accent4: 'B9CDE5', accent5: 'A6A6A6', accent6: 'D9D9D9',
};

const HEX = /^#?[0-9A-Fa-f]{6}$/;
export const isHex = (v: unknown): v is string => typeof v === 'string' && HEX.test(v);
export const isColorName = (v: unknown): v is ColorName => typeof v === 'string' && (COLOR_NAMES as readonly string[]).includes(v);
/** A colour as the spec accepts it: hex, or a theme slot name. */
export const isColor = (v: unknown): v is string => isHex(v) || isColorName(v);

/** The theme slots a style resolves against: the open-charts theme, overridden by `style.theme`. */
export const themeOf = (style?: Partial<Pick<Style, 'theme'>>): ThemeColors => {
  const t = { ...DEFAULT_THEME };
  for (const k of THEME_SLOTS) {
    const v = style?.theme?.[k];
    if (isHex(v)) t[k] = v.replace('#', '').toUpperCase();
  }
  return t;
};

/**
 * Scene colour for a spec colour: "RRGGBB" for hex, "RRGGBB@slot" for a theme slot (resolved in
 * `theme`). Already resolved colours pass through, so resolving twice is harmless.
 */
export function resolveColor(c: string, theme: ThemeColors): string {
  if (isColorName(c)) {
    const slot = (MAPPED_SLOTS as Record<string, ThemeSlot>)[c] ?? (c as ThemeSlot);
    return `${theme[slot]}@${c}`;
  }
  return c.replace('#', '');
}

/** The hex part of a scene colour ("1F3A5F@accent1" → "1F3A5F"). */
export const colorHex = (c: string): string => c.replace('#', '').split('@')[0];
/** The theme slot of a scene colour, if it has one ("1F3A5F@accent1" → "accent1"). */
export const colorSlot = (c: string): ColorName | undefined => {
  const s = c.split('@')[1];
  return isColorName(s) ? s : undefined;
};

/** A style with every colour resolved for layout (see resolveColor). */
export function resolveStyle(style: Style): Style {
  const theme = themeOf(style);
  const r = (c: string) => resolveColor(c, theme);
  return {
    ...style,
    text: r(style.text), muted: r(style.muted), axis: r(style.axis), background: r(style.background),
    positive: r(style.positive), negative: r(style.negative), palette: style.palette.map(r),
  };
}

// ---- Element colours: theme colours with a tint / shade (per-element overrides, core/overrides.ts) ----

/** Hyperlink theme colours (not slots of `style.theme`): the defaults of a CLI deck's theme. */
export const LINK_SLOTS = { hlink: '0563C1', folHlink: '954F72' } as const;
export type LinkColors = Partial<Record<keyof typeof LINK_SLOTS, string>>;
/** Theme colour names an element colour may use, before an optional "-lighter<N>" / "-darker<N>" (N 0–100). */
export const ELEMENT_COLOR_NAMES: readonly string[] = [...COLOR_NAMES, ...Object.keys(LINK_SLOTS)];
const MODDED = /^([A-Za-z0-9]+)(?:-(lighter|darker)(\d{1,3}))?$/;

/** A theme colour with an optional tint / shade: "accent1", "accent1-lighter40", "tx1-darker25". */
export function parseThemeColor(v: string): { slot: string; mod?: 'lighter' | 'darker'; pct: number } | undefined {
  const m = MODDED.exec(v);
  if (!m || !ELEMENT_COLOR_NAMES.includes(m[1])) return undefined;
  const pct = m[3] === undefined ? 0 : +m[3];
  if (pct > 100) return undefined;
  return { slot: m[1], ...(m[2] ? { mod: m[2] as 'lighter' | 'darker' } : {}), pct };
}

/** An element colour: hex, a theme colour, or a theme colour with a tint / shade ("accent1-lighter40"). */
export const isElementColor = (v: unknown): v is string => isHex(v) || (typeof v === 'string' && !!parseThemeColor(v));

/**
 * Scene colour of an element colour: "RRGGBB@accent1-lighter40" — the hex the tint / shade gives on
 * `theme` (DrawingML lumMod / lumOff, applied to HSL lightness) with the theme name kept for renderers
 * that emit theme colours. Plain hex and plain slots resolve as `resolveColor` does.
 */
export function resolveElementColor(c: string, theme: ThemeColors & LinkColors): string {
  if (isHex(c)) return c.replace('#', '');
  const t = parseThemeColor(c);
  if (!t) return c;
  if (!t.mod && isColorName(c)) return resolveColor(c, theme);
  const link = t.slot as keyof typeof LINK_SLOTS;
  const base = t.slot in LINK_SLOTS ? theme[link] ?? LINK_SLOTS[link] : colorHex(resolveColor(t.slot, theme));
  return `${t.mod ? tintHex(base, t.mod === 'lighter' ? t.pct : -t.pct) : base}@${c}`;
}

/** Hex lightened (pct > 0: lumMod 1-p, lumOff p) or darkened (pct < 0: lumMod 1-p) in HSL lightness. */
export function tintHex(hex: string, pct: number): string {
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
  const max = Math.max(r, g, b), min = Math.min(r, g, b), l = (max + min) / 2, d = max - min;
  const s = d === 0 ? 0 : d / (1 - Math.abs(2 * l - 1));
  const h = d === 0 ? 0 : max === r ? ((g - b) / d + 6) % 6 : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  const p = Math.abs(pct) / 100;
  const l2 = pct >= 0 ? l * (1 - p) + p : l * (1 - p);
  const c = (1 - Math.abs(2 * l2 - 1)) * s, x = c * (1 - Math.abs((h % 2) - 1)), m = l2 - c / 2;
  const [r1, g1, b1] = h < 1 ? [c, x, 0] : h < 2 ? [x, c, 0] : h < 3 ? [0, c, x] : h < 4 ? [0, x, c] : h < 5 ? [x, 0, c] : [c, 0, x];
  return [r1, g1, b1].map((v) => Math.round(Math.min(1, Math.max(0, v + m)) * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/** The theme colour of a scene colour with its tint / shade ("…@accent1-lighter40"), if it has one. */
export const colorTheme = (c: string) => { const s = c.split('@')[1]; return s ? parseThemeColor(s) : undefined; };
