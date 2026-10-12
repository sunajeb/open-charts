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
