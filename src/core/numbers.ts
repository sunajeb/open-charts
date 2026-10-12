/**
 * "Nice" axis ticks: steps of 1, 2, 2.5 or 5 × 10^n. Picks the step that wastes the least
 * axis length while keeping close to the requested tick count (simplified Talbot et al. 2010).
 */
export function niceTicks(min: number, max: number, count = 5): number[] {
  if (min === max) max = min + 1;
  const span = max - min;
  let best: { lo: number; step: number; n: number; score: number } | null = null;
  const base = 10 ** Math.floor(Math.log10(span / count));
  for (const mult of [0.1, 1, 10]) {
    for (const q of [1, 2, 2.5, 5]) {
      const step = q * base * mult;
      const lo = Math.floor(min / step + 1e-9) * step;
      const hi = Math.ceil(max / step - 1e-9) * step;
      const n = Math.round((hi - lo) / step) + 1;
      if (n < 3 || n > 9) continue;
      const waste = (hi - lo - span) / (hi - lo);
      const score = waste * 2 + Math.abs(n - count) * 0.15 + (q === 2.5 ? 0.1 : 0);
      if (!best || score < best.score) best = { lo, step, n, score };
    }
  }
  const { lo, step, n } = best ?? { lo: min, step: span, n: 2 };
  return Array.from({ length: n }, (_, i) => round(lo + i * step, 10));
}

const round = (v: number, d: number) => Math.round(v * 10 ** d) / 10 ** d;

/**
 * Smallest number of decimals (0–2) that represents every value exactly; otherwise enough decimals
 * to show two significant digits of the smallest non-zero value (at most 3).
 */
export function inferDecimals(values: number[]): number {
  for (let d = 0; d <= 2; d++) {
    if (values.every((v) => Math.abs(round(v, d) - v) < 1e-9)) return d;
  }
  const small = Math.min(...values.filter((v) => v !== 0).map(Math.abs));
  if (!Number.isFinite(small) || small >= 10) return 0;
  return Math.min(3, Math.max(1, 1 - Math.floor(Math.log10(small))));
}

// ---- Locale ------------------------------------------------------------------------------------

export interface Locale {
  decimal: string;
  group: string;
  /** Group sizes from the right; the last one repeats (en-IN: 3 then 2s → 12,34,567). */
  grouping: number[];
  minus: string;
  /** Percent sign including any space the locale puts before it. */
  percent: string;
}

const NBSP = ' ';
const NNBSP = ' ';
const MINUS = '−';

export const LOCALES: Record<string, Locale> = {
  'en-US': { decimal: '.', group: ',', grouping: [3], minus: MINUS, percent: '%' },
  'en-GB': { decimal: '.', group: ',', grouping: [3], minus: MINUS, percent: '%' },
  'en-IN': { decimal: '.', group: ',', grouping: [3, 2], minus: MINUS, percent: '%' },
  'de-DE': { decimal: ',', group: '.', grouping: [3], minus: MINUS, percent: `${NBSP}%` },
  'fr-FR': { decimal: ',', group: NNBSP, grouping: [3], minus: MINUS, percent: `${NNBSP}%` },
};

/** Exact match, then language match ("de" → de-DE), then en-US. */
export function localeOf(id?: string): Locale {
  if (!id) return LOCALES['en-US'];
  if (LOCALES[id]) return LOCALES[id];
  const lang = id.split('-')[0].toLowerCase();
  const key = Object.keys(LOCALES).find((k) => k.toLowerCase().startsWith(`${lang}-`));
  return LOCALES[key ?? 'en-US'];
}

function groupDigits(int: string, loc: Locale): string {
  const parts: string[] = [];
  let rest = int;
  let k = 0;
  while (rest.length > 0) {
    const size = loc.grouping[Math.min(k, loc.grouping.length - 1)];
    parts.unshift(rest.slice(-size));
    rest = rest.slice(0, -size);
    k++;
  }
  return parts.join(loc.group);
}

/** Digits of |v| rounded to `decimals`, with locale grouping and decimal mark. */
function digits(abs: number, decimals: number, loc: Locale, group = true, minDecimals = decimals, minInt = 1): string {
  const [i, f = ''] = round(abs, decimals).toFixed(decimals).split('.');
  let frac = f;
  while (frac.length > minDecimals && frac.endsWith('0')) frac = frac.slice(0, -1);
  let int = i.replace(/^0+/, '').padStart(minInt, '0');
  if (group && int) int = groupDigits(int, loc);
  if (!int && !frac) int = '0';
  return frac ? `${int}${loc.decimal}${frac}` : int;
}

// ---- Units and magnitudes ----------------------------------------------------------------------

export interface Unit { text: string; position?: 'prefix' | 'suffix'; space?: boolean }
export type Scale = 'none' | 'auto' | 'k' | 'm' | 'bn';
export const MAGNITUDES: Record<'k' | 'm' | 'bn', number> = { k: 1e3, m: 1e6, bn: 1e9 };

/** One magnitude shared by the whole chart, so every label uses the same suffix. */
export function pickScale(values: number[], scale: Scale | undefined): { div: number; suffix: string } {
  if (!scale || scale === 'none') return { div: 1, suffix: '' };
  if (scale !== 'auto') return { div: MAGNITUDES[scale], suffix: scale };
  const max = Math.max(0, ...values.filter(Number.isFinite).map(Math.abs));
  for (const s of ['bn', 'm', 'k'] as const) if (max >= MAGNITUDES[s]) return { div: MAGNITUDES[s], suffix: s };
  return { div: 1, suffix: '' };
}

export const normalizeUnit = (u: string | Unit | undefined): Unit =>
  typeof u === 'string' || u == null
    ? { text: u ?? '', position: 'suffix', space: false }
    : { position: 'suffix', space: false, ...u };

// ---- Excel-style format strings ----------------------------------------------------------------

function splitSections(format: string): string[] {
  const out: string[] = [];
  let cur = '';
  let quoted = false;
  for (let i = 0; i < format.length; i++) {
    const ch = format[i];
    if (ch === '"') quoted = !quoted;
    if (ch === '\\' && !quoted) { cur += ch + (format[++i] ?? ''); continue; }
    if (ch === ';' && !quoted) { out.push(cur); cur = ''; continue; }
    cur += ch;
  }
  out.push(cur);
  return out;
}

/**
 * Format with an Excel number format string: up to three sections (positive;negative;zero),
 * `0 # ?` digit placeholders, `,` grouping, trailing `,` = ÷1000, `%` (×100), quoted text,
 * `\x` escapes, `_x` spaces; `[colour]` and `*` fill codes are ignored. The locale decides the
 * decimal mark, grouping character, minus sign and percent spacing.
 */
export function applyFormat(v: number, format: string, locale?: string | Locale): string {
  const loc = typeof locale === 'object' ? locale : localeOf(locale);
  const secs = splitSections(format);
  let sec = secs[0];
  let x = Math.abs(v);
  let autoMinus = false;
  if (v < 0) {
    if (secs[1] !== undefined && secs[1] !== '') sec = secs[1]; else autoMinus = true;
  } else if (v === 0 && secs[2] !== undefined) sec = secs[2];

  // Tokenise into literal text around the first numeric pattern.
  const pre: string[] = [], post: string[] = [];
  let pattern = '';
  let percent = 0;
  let i = 0;
  const lit = (s: string) => (pattern ? post : pre).push(s);
  while (i < sec.length) {
    const ch = sec[i];
    if (ch === '"') { const j = sec.indexOf('"', i + 1); lit(sec.slice(i + 1, j < 0 ? undefined : j)); i = j < 0 ? sec.length : j + 1; continue; }
    if (ch === '\\') { lit(sec[i + 1] ?? ''); i += 2; continue; }
    if (ch === '_') { lit(' '); i += 2; continue; }
    if (ch === '*') { i += 2; continue; }
    if (ch === '[') { const j = sec.indexOf(']', i); i = j < 0 ? sec.length : j + 1; continue; }
    if (ch === '%') { percent++; lit(loc.percent); i++; continue; }
    if (!pattern && (/[0#?]/.test(ch) || (ch === '.' && /[0#?]/.test(sec[i + 1] ?? '')))) {
      let j = i;
      while (j < sec.length && /[0#?,.]/.test(sec[j])) j++;
      pattern = sec.slice(i, j);
      i = j;
      continue;
    }
    lit(ch);
    i++;
  }
  if (!pattern) return pre.join('') + post.join('');

  const [intPat, fracPat = ''] = pattern.split('.');
  const tail = fracPat ? fracPat : intPat;
  const trailing = tail.match(/,+$/)?.[0].length ?? 0;
  const intDigits = fracPat ? intPat : intPat.replace(/,+$/, '');
  const fracDigits = fracPat.replace(/,/g, '');
  const decimals = fracDigits.length;
  const minDecimals = (fracDigits.match(/0/g) ?? []).length;
  const minInt = (intDigits.match(/0/g) ?? []).length;
  const group = /[0#?],[0#?]/.test(intDigits);
  const scaled = (x * 100 ** percent) / 1000 ** trailing;
  const out = pre.join('') + digits(scaled, decimals, loc, group, minDecimals, minInt) + post.join('');
  return autoMinus && round(scaled, decimals) !== 0 ? loc.minus + out : out;
}

// ---- The one entry point the engine uses -------------------------------------------------------

export interface NumberStyle {
  locale?: string;
  /** Fixed decimals (ignored when `format` is given). */
  decimals?: number;
  /** Excel-style format string, e.g. "#,##0.0", "0%", "#,##0;(#,##0)". */
  format?: string;
  unit?: string | Unit;
  /** "+" in front of positive values (growth, deltas). */
  signed?: boolean;
  /** Magnitude suffix ("k", "m", "bn") placed right after the number. */
  magnitude?: string;
}

/** Format one value: locale marks, sign, unit (prefix or suffix, optional space) and magnitude. */
export function formatValue(v: number, st: NumberStyle = {}): string {
  const loc = localeOf(st.locale);
  const unit = normalizeUnit(st.unit);
  const pct = unit.text === '%';
  const utext = pct ? '%' : unit.text;
  const gap = unit.space ? NBSP : pct ? loc.percent.slice(0, -1) : '';
  let body: string;
  let sign = '';
  if (st.format) {
    body = applyFormat(v, st.format, loc);
    if (st.signed && v > 0 && !body.startsWith('+')) sign = '+';
  } else {
    const d = st.decimals ?? 0;
    body = digits(Math.abs(v), d, loc);
    const zero = round(Math.abs(v), d) === 0;
    sign = v < 0 && !zero ? loc.minus : st.signed && v > 0 && !zero ? '+' : '';
  }
  const mag = st.magnitude ?? '';
  if (!utext) return `${sign}${body}${mag}`;
  return unit.position === 'prefix'
    ? `${sign}${utext}${gap}${body}${mag}`
    : `${sign}${body}${mag}${gap}${utext}`;
}

/** Shorthand kept for callers of the original API. */
export function formatNumber(v: number, decimals: number, unit: string | Unit = '', signed = false, locale?: string): string {
  return formatValue(v, { decimals, unit, signed, locale });
}

/**
 * Round values so the rounded figures add up exactly to the rounded total
 * (largest-remainder method), e.g. shares that always show 100%.
 */
export function roundToTotal(values: number[], decimals = 0): number[] {
  const f = 10 ** decimals;
  const scaled = values.map((v) => v * f);
  const floors = scaled.map(Math.floor);
  let short = Math.round(scaled.reduce((a, v) => a + v, 0)) - floors.reduce((a, v) => a + v, 0);
  const order = scaled.map((v, i) => i).sort((a, b) => (scaled[b] - floors[b]) - (scaled[a] - floors[a]));
  for (const i of order) {
    if (short <= 0) break;
    floors[i] += 1;
    short--;
  }
  return floors.map((v) => v / f);
}

/** Black or white text, whichever reads better on the given fill (WCAG relative luminance). */
export function textOn(fill: string, dark = '222222', light = 'FFFFFF'): string {
  const hex = fill.replace('#', '');
  const [r, g, b] = [0, 2, 4].map((i) => {
    const c = parseInt(hex.slice(i, i + 2), 16) / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  });
  const lum = 0.2126 * r + 0.7152 * g + 0.0722 * b;
  return lum > 0.4 ? dark : light;
}
