// Which font a chart's text is measured with: theme fonts resolved to the deck's families, a
// metric-compatible substitute when a font is missing, else a conservative wide estimate that is
// reported. Pure: the font files themselves are found by src/fonts.ts (Node) or the browser
// (src/officejs/measure.ts); both build on these rules, so a font measures the same everywhere.
import type { Measurer } from './text.ts';

/** The deck's theme fonts (DrawingML "+mj-lt" = major / headings, "+mn-lt" = minor / body). */
export interface ThemeFonts { major?: string; minor?: string }

export const isThemeFont = (name: string) => /^\+(mj|mn)-(lt|ea|cs)$/.test(name);

/**
 * The family a font name stands for: "+mj-lt" / "+mn-lt" become the deck's theme fonts (`themeFonts`,
 * else `font` — a CLI deck's theme uses the style font for both — else Arial).
 */
export function resolveFont(name: string | undefined, style: { font?: string; themeFonts?: ThemeFonts } = {}): string {
  const plain = style.font && !isThemeFont(style.font) ? style.font : undefined;
  const n = name ?? style.font ?? 'Arial';
  if (!isThemeFont(n)) return n;
  const t = style.themeFonts ?? {};
  return (n.startsWith('+mj') ? t.major ?? t.minor : t.minor ?? t.major) ?? plain ?? 'Arial';
}

/**
 * Fonts with the same advance widths (metric-compatible: drawn in one, a text measured in the other
 * fits exactly). Arial was made to match Helvetica; Liberation, Arimo/Tinos/Cousine, Carlito and
 * Caladea were made to match the Microsoft fonts.
 */
const COMPATIBLE: string[][] = [
  ['Arial', 'Liberation Sans', 'Arimo', 'Helvetica'],
  ['Times New Roman', 'Liberation Serif', 'Tinos'],
  ['Courier New', 'Liberation Mono', 'Cousine'],
  ['Calibri', 'Carlito'],
  ['Cambria', 'Caladea'],
];

/** Metric-compatible substitutes for `family`, best first (empty when none is known). */
export function substitutesFor(family: string): string[] {
  const key = family.trim().toLowerCase();
  const group = COMPATIBLE.find((g) => g.some((f) => f.toLowerCase() === key));
  return group ? group.filter((f) => f.toLowerCase() !== key) : [];
}

/**
 * Missing font without a substitute: measured as the base font (Arial or its equivalent) made 15 %
 * wider — wider than Verdana, the widest common UI font — with lines at least 1.3 em apart (Segoe
 * UI, among the tallest, is 1.33). Labels then get room to spare in any likely font.
 */
export const WIDE_FALLBACK = { width: 1.15, lineHeight: 1.3 } as const;

export function wideFallback(base: Measurer, requested: string): Measurer {
  const memo = new Map<string, number>();
  return {
    family: requested,
    width(text, size, bold) {
      const key = `${bold ? 'b' : 'r'}${size}|${text}`;
      let w = memo.get(key);
      if (w === undefined) {
        w = base.width(text, size, bold) * WIDE_FALLBACK.width;
        if (memo.size >= 20000) memo.delete(memo.keys().next().value!);
        memo.set(key, w);
      }
      return w;
    },
    lineHeight: (size) => Math.max(base.lineHeight(size), size * WIDE_FALLBACK.lineHeight),
    fallback: { requested, used: base.family ?? 'Arial', wide: true },
  };
}

/** The measurer for `family` (a resolved name: see resolveFont). */
export const measurerFor = (m: Measurer, family: string): Measurer => (m.forFont ? m.forFont(family) : m);

/** The scene warning for a font that is missing. */
export const fontWarning = (m: Measurer): string | undefined => m.fallback?.wide
  ? `font "${m.fallback.requested}" is not installed here: text was measured with a wider stand-in, so labels have extra room; install the font for an exact layout`
  : undefined;

/**
 * A measurer over several fonts: `find(family)` returns the measurer of an installed family or null.
 * Families resolve once each; a missing one goes to a metric-compatible substitute, else to the
 * wide fallback over `base`.
 */
export function fontSet(base: Measurer, find: (family: string) => Measurer | null): Measurer {
  const cache = new Map<string, Measurer>();
  const forFont = (family: string): Measurer => {
    const key = family.trim().toLowerCase();
    let m = cache.get(key);
    if (m) return m;
    if (base.family && base.family.toLowerCase() === key) m = base;
    else {
      m = find(family) ?? undefined;
      for (const sub of m ? [] : substitutesFor(family)) {
        const s = base.family?.toLowerCase() === sub.toLowerCase() ? base : find(sub);
        if (s) { m = { ...withoutForFont(s), family }; break; }
      }
      if (!m) return { ...wideFallback(base, family), forFont }; // not cached: the font may be installed later
    }
    const out = { ...withoutForFont(m), forFont };
    cache.set(key, out);
    return out;
  };
  return { ...withoutForFont(base), forFont };
}

function withoutForFont(m: Measurer): Measurer {
  const { forFont: _, ...rest } = m;
  // width / lineHeight may be methods on a prototype in other implementations: bind them.
  return { ...rest, width: m.width.bind(m), lineHeight: m.lineHeight.bind(m) };
}
