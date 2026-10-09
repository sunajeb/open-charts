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

/** Smallest number of decimals (0-2) that represents every value exactly enough to read. */
export function inferDecimals(values: number[]): number {
  for (let d = 0; d <= 2; d++) {
    if (values.every((v) => Math.abs(round(v, d) - v) < 1e-9)) return d;
  }
  return values.some((v) => v !== 0 && Math.abs(v) < 10) ? 1 : 0;
}

export function formatNumber(v: number, decimals: number, unit = '', signed = false): string {
  const s = Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
  const sign = v < 0 && Number(s.replace(/,/g, '')) !== 0 ? '−' : signed && v > 0 ? '+' : '';
  return `${sign}${s}${unit}`;
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
