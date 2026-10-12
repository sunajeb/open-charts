// Dates read from category labels, for CAGR (B3). Each label maps to the first day of the period it
// names, so the distance between two labels of the same granularity is a whole number of periods.

export interface YMD { y: number; m: number; d: number }

const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

const year = (s: string) => {
  const n = Number(s);
  return s.length === 2 ? (n < 70 ? 2000 + n : 1900 + n) : n;
};

/**
 * Parse a category label as a date: years ("2024", "FY24", "CY2024", "2025E", "'24"), quarters
 * ("Q1 2023", "Q1 FY24", "2023 Q1", "2023-Q1", "1Q23"), halves ("H1 2023"), months ("Jan 2023",
 * "January 2023", "Mar-24", "2023-01", "01/2023") and ISO dates ("2023-01-15"). Returns null otherwise.
 */
export function parseCategoryDate(label: string): YMD | null {
  let s = label.trim().toLowerCase().replace(/\s+/g, ' ');
  s = s.replace(/^(.*\d)\s*(e|a|f|p|b|est\.?|act\.?|fcst\.?)$/, '$1'); // estimate / actual markers
  s = s.replace(/\b(fy|cy)\s*'?/g, '').replace(/'/g, '').trim();
  let r: RegExpMatchArray | null;
  if ((r = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/))) return ok(Number(r[1]), Number(r[2]), Number(r[3]));
  if ((r = s.match(/^(\d{4})[-/.](\d{1,2})$/))) return ok(Number(r[1]), Number(r[2]), 1);
  if ((r = s.match(/^(\d{1,2})[/.](\d{4})$/))) return ok(Number(r[2]), Number(r[1]), 1);
  if ((r = s.match(/^(\d{4}|\d{2})$/))) return ok(year(r[1]), 1, 1);
  const period = (kind: string, n: number, y: string) =>
    kind === 'q' ? (n >= 1 && n <= 4 ? ok(year(y), 3 * (n - 1) + 1, 1) : null) : n === 1 || n === 2 ? ok(year(y), n === 1 ? 1 : 7, 1) : null;
  if ((r = s.match(/^([qh])([1-4])[\s\-/]*(\d{4}|\d{2})$/))) return period(r[1], Number(r[2]), r[3]);
  if ((r = s.match(/^([1-4])([qh])[\s\-/]*(\d{4}|\d{2})$/))) return period(r[2], Number(r[1]), r[3]);
  if ((r = s.match(/^(\d{4}|\d{2})[\s\-/]*([qh])([1-4])$/))) return period(r[2], Number(r[3]), r[1]);
  if ((r = s.match(/^([a-z]{3,9})\.?[\s\-/]*(\d{4}|\d{2})$/))) {
    const m = MONTHS.indexOf(r[1].slice(0, 3));
    return m >= 0 && (r[1].length === 3 || MONTHS_FULL[m].startsWith(r[1])) ? ok(year(r[2]), m + 1, 1) : null;
  }
  if ((r = s.match(/^(\d{4})[\s\-/]*([a-z]{3,9})\.?$/))) {
    const m = MONTHS.indexOf(r[2].slice(0, 3));
    return m >= 0 && (r[2].length === 3 || MONTHS_FULL[m].startsWith(r[2])) ? ok(Number(r[1]), m + 1, 1) : null;
  }
  return null;
}

const MONTHS_FULL = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];

function ok(y: number, m: number, d: number): YMD | null {
  return y >= 1000 && y <= 9999 && m >= 1 && m <= 12 && d >= 1 && d <= 31 ? { y, m, d } : null;
}

/** Years between two dates under the 30/360 (US) day count convention. */
export function years360(a: YMD, b: YMD): number {
  let d1 = a.d, d2 = b.d;
  if (d1 === 31) d1 = 30;
  if (d2 === 31 && d1 === 30) d2 = 30;
  return (360 * (b.y - a.y) + 30 * (b.m - a.m) + (d2 - d1)) / 360;
}

/**
 * Number of years a CAGR between two categories spans: the explicit `years` when given, else the
 * 30/360 distance between the dates read from both labels, else the category index distance.
 */
export function cagrYears(categories: string[], from: number, to: number, explicit?: number): { years: number; source: 'explicit' | 'dates' | 'index' } {
  if (explicit && explicit > 0) return { years: explicit, source: 'explicit' };
  const a = parseCategoryDate(categories[from] ?? ''), b = parseCategoryDate(categories[to] ?? '');
  if (a && b) {
    const y = years360(a, b);
    if (y > 0) return { years: y, source: 'dates' };
  }
  return { years: to - from, source: 'index' };
}
