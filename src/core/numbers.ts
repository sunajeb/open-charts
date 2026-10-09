/** "Nice" axis ticks (Heckbert, Graphics Gems 1990): steps of 1, 2 or 5 × 10^n. */
export function niceTicks(min: number, max: number, count = 5): number[] {
  if (min === max) max = min + 1;
  const range = niceNum(max - min, false);
  const step = niceNum(range / Math.max(1, count - 1), true);
  const lo = Math.floor(min / step) * step;
  const hi = Math.ceil(max / step) * step;
  const ticks: number[] = [];
  for (let v = lo; v <= hi + step / 2; v += step) ticks.push(round(v, 10));
  return ticks;
}

function niceNum(x: number, roundIt: boolean): number {
  const exp = Math.floor(Math.log10(x));
  const f = x / 10 ** exp;
  const nf = roundIt
    ? f < 1.5 ? 1 : f < 3 ? 2 : f < 7 ? 5 : 10
    : f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10;
  return nf * 10 ** exp;
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
