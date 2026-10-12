// Scene-level KPI checks for `npm run quality` (docs/MASTER-PLAN.md §9): each reads the numbers a slide
// shows back from its scene and checks them against the data or each other.
import { cagrYears, LOCALES, type ChartSpec, type LineShape, type RectShape, type Scene, type TextShape } from '../src/core/index.ts';

// ---- Reading numbers back from the scene ------------------------------------------------------

/** The numbers a label shows, in its own units (locale marks, minus sign, parentheses). */
export function numbers(text: string, locale?: string): number[] {
  const loc = LOCALES[locale ?? 'en-US'] ?? LOCALES['en-US'];
  const out: number[] = [];
  // "(12)" is an accounting negative; "(46%)" is not.
  for (const [, open, minus, digits, close] of text.matchAll(/(\()?([−-])?(\d[\d.,\u00a0\u202f]*)(\))?/g)) {
    const neg = !!minus || (!!open && !!close);
    let t = digits.trim().split(loc.group).join('');
    if (loc.decimal !== '.') t = t.replace(loc.decimal, '.');
    t = t.replace(/[^\d.]/g, '').replace(/\.$/, '');
    if (t) out.push((neg ? -1 : 1) * Number(t));
  }
  return out;
}
const pctOf = (text: string, locale?: string) => { const ns = numbers(text.slice(text.lastIndexOf('(') + 1), locale); return text.includes('%') && ns.length ? ns[ns.length - 1] : null; };
const decimalsOf = (text: string, locale?: string) => {
  const dec = (LOCALES[locale ?? 'en-US'] ?? LOCALES['en-US']).decimal;
  const mm = text.match(new RegExp(`\\${dec}(\\d+)`));
  return mm ? mm[1].length : 0;
};

const texts = (sc: Scene, re: RegExp) => sc.shapes.flatMap((s, i) => (s.kind === 'text' && re.test(sc.roles![i]) ? [{ t: s as TextShape, role: sc.roles![i] }] : []));
const col = (role: string) => Number(role.match(/:c(\d+)$/)?.[1] ?? -1);
const ser = (role: string) => Number(role.match(/:s(\d+):/)?.[1] ?? -1);

export interface Check { name: string; ok: boolean; detail?: string }

// 100% sums: percent columns, 100% area categories, percent Mekko columns, pie / doughnut slices.
export function hundredChecks(spec: ChartSpec, sc: Scene): Check[] {
  const o = spec.options ?? {};
  const loc = o.locale;
  const percentCols = spec.type === 'percent' && !o.percentData || spec.type === 'area100' || (spec.type === 'mekko' && (o.mekko ?? 'percent') === 'percent');
  const out: Check[] = [];
  if (percentCols) {
    const nonEmpty = (c: number) => spec.series.filter((s) => typeof s.values[c] === 'number' && (s.values[c] as number) > 0).length;
    spec.categories.forEach((_, c) => {
      const shown = texts(sc, new RegExp(`^label:value:s\\d+:c${c}$`)).map(({ t }) => pctOf(t.text, loc)).filter((v): v is number => v != null);
      if (!shown.length || shown.length < nonEmpty(c)) return; // a hidden share: not verifiable from the slide
      const sum = shown.reduce((a, v) => a + v, 0);
      out.push({ name: `column ${c}`, ok: Math.abs(sum - 100) < 1e-6, detail: `${+sum.toFixed(6)}%` });
    });
  }
  if (spec.type === 'pie' || spec.type === 'doughnut') {
    if ((o.labelContent ?? 'percent') === 'value') return out;
    const by = new Map<number, number>();
    for (const { t, role } of texts(sc, /:c\d+$/)) { const p = pctOf(t.text, loc); if (p != null) by.set(col(role), p); }
    const slices = (spec.series[0]?.values ?? []).filter((v) => typeof v === 'number' && v > 0).length;
    if (by.size === slices) {
      const sum = [...by.values()].reduce((a, v) => a + v, 0);
      out.push({ name: 'pie', ok: Math.abs(sum - 100) < 1e-6, detail: `${+sum.toFixed(6)}%` });
    }
  }
  return out;
}

// Stacked columns: the segment labels add up to the bold total above the column (consistent rounding).
export function stackChecks(spec: ChartSpec, sc: Scene): Check[] {
  if (spec.type !== 'stacked' || spec.options?.format) return [];
  const loc = spec.options?.locale;
  return spec.categories.flatMap((_, c) => {
    const total = texts(sc, new RegExp(`^label:total:c${c}$`))[0];
    const segs = texts(sc, new RegExp(`^label:value:s\\d+:c${c}$`));
    const nonZero = spec.series.filter((s) => typeof s.values[c] === 'number' && s.values[c] !== 0).length;
    if (!total || segs.length < nonZero) return [];
    const sum = segs.reduce((a, { t }) => a + (numbers(t.text, loc)[0] ?? 0), 0);
    const tv = numbers(total.t.text, loc)[0];
    const d = decimalsOf(total.t.text, loc);
    return [{ name: `column ${c}`, ok: Math.abs(sum - tv) < 10 ** -d / 2, detail: `segments ${+sum.toFixed(6)} vs total ${tv}` }];
  });
}

// CAGR bubble = rate recomputed from the totals shown at both ends (±0.05 pp).
export function cagrChecks(spec: ChartSpec, sc: Scene): Check[] {
  const ref = spec.options?.cagr;
  if (!ref || (spec.type !== 'stacked' && spec.type !== 'percent')) return [];
  const loc = spec.options?.locale;
  const bubble = texts(sc, /^arrow:cagr$/)[0];
  const a = texts(sc, new RegExp(`^label:total:c${ref.from}$`))[0], b = texts(sc, new RegExp(`^label:total:c${ref.to}$`))[0];
  if (!bubble || !a || !b) return [];
  const years = cagrYears(spec.categories, ref.from, ref.to, ref.years).years;
  const want = ((numbers(b.t.text, loc)[0] / numbers(a.t.text, loc)[0]) ** (1 / years) - 1) * 100;
  const got = numbers(bubble.t.text, loc)[0];
  return [{ name: 'cagr', ok: Math.abs(got - want) <= 0.05 + 1e-9, detail: `label ${got}% vs recomputed ${want.toFixed(3)}%` }];
}

// Waterfall: every connector meets the bars on both sides; every total = previous total + steps shown.
export function waterfallChecks(spec: ChartSpec, sc: Scene): Check[] {
  if (spec.type !== 'waterfall') return [];
  const loc = spec.options?.locale;
  const out: Check[] = [];
  const bars = sc.shapes.flatMap((s, i) => (s.kind === 'rect' && sc.roles![i].startsWith('bar:') ? [{ r: s as RectShape, c: col(sc.roles![i]) }] : []));
  const touches = (c: number, y: number) => bars.some(({ r, c: k }) => k === c && (Math.abs(r.y - y) < 0.25 || Math.abs(r.y + r.h - y) < 0.25));
  sc.shapes.forEach((s, i) => {
    const mm = sc.roles![i].match(/^connector:c(\d+)$/);
    if (!mm || s.kind !== 'line') return;
    const c = Number(mm[1]), y = (s as LineShape).y1;
    out.push({ name: `connector ${c}→${c + 1}`, ok: touches(c, y) && touches(c + 1, y) });
  });
  // Totals: the bold label of a total column vs the running sum of the labels shown before it.
  const totals = texts(sc, /^label:total:c\d+$/).map(({ t, role }) => ({ c: col(role), v: numbers(t.text, loc)[0], d: decimalsOf(t.text, loc) })).sort((p, q) => p.c - q.c);
  const steps = texts(sc, /^label:value:s\d+:c\d+$/).map(({ t, role }) => ({ c: col(role), v: numbers(t.text, loc)[0] }));
  const deltas = (from: number, to: number) => spec.series.reduce((n, s) => n + s.values.slice(from + 1, to).filter((v) => typeof v === 'number' && v !== 0).length, 0);
  for (let k = 1; k < totals.length; k++) {
    const prev = totals[k - 1], cur = totals[k];
    const shown = steps.filter((st) => st.c > prev.c && st.c < cur.c);
    if (shown.length < deltas(prev.c, cur.c)) continue; // a hidden step: not verifiable
    const sum = prev.v + shown.reduce((a, st) => a + st.v, 0);
    // Each shown figure carries up to half a unit of rounding.
    const tol = (shown.length + 2) * 10 ** -cur.d / 2;
    out.push({ name: `total c${cur.c}`, ok: Math.abs(sum - cur.v) <= tol + 1e-9, detail: `${prev.v} + steps = ${+sum.toFixed(6)} vs ${cur.v}` });
  }
  return out;
}

// Mekko / Mekko XY: column widths add up to the plot width (±0.5 pt), touch, and follow the data.
export function mekkoChecks(spec: ChartSpec, sc: Scene): Check[] {
  if (spec.type !== 'mekko' && spec.type !== 'mekkoxy') return [];
  const cols = new Map<number, { x: number; w: number }>();
  sc.shapes.forEach((s, i) => { if (s.kind === 'rect' && sc.roles![i].startsWith('bar:')) cols.set(col(sc.roles![i]), { x: s.x, w: s.w }); });
  const axis = sc.shapes.find((s, i) => sc.roles![i] === 'axis' && s.kind === 'line' && Math.abs(s.y1 - s.y2) < 0.01) as LineShape | undefined;
  if (!axis || !cols.size) return [{ name: 'mekko widths', ok: false, detail: 'no columns or baseline' }];
  const list = [...cols.entries()].sort((a, b) => a[1].x - b[1].x);
  const sumW = list.reduce((a, [, c]) => a + c.w, 0);
  const plotW = Math.abs(axis.x2 - axis.x1);
  const gaps = Math.max(0, ...list.slice(1).map(([, c], k) => Math.abs(c.x - (list[k][1].x + list[k][1].w))));
  const weight = (c: number) => (spec.type === 'mekko'
    ? spec.series.reduce((a, s) => a + Math.max(0, typeof s.values[c] === 'number' ? (s.values[c] as number) : 0), 0)
    : Math.max(0, typeof spec.series[0].values[c] === 'number' ? (spec.series[0].values[c] as number) : 0));
  const grand = spec.categories.reduce((a, _, c) => a + weight(c), 0);
  const prop = Math.max(...list.map(([c, b]) => Math.abs(b.w / plotW - weight(c) / grand)));
  return [
    { name: 'width sum', ok: Math.abs(sumW - plotW) <= 0.5, detail: `Σ widths ${sumW.toFixed(2)} pt vs plot ${plotW.toFixed(2)} pt` },
    { name: 'columns touch', ok: gaps <= 0.5, detail: `largest gap ${gaps.toFixed(2)} pt` },
    { name: 'widths ∝ data', ok: prop <= 0.001, detail: `largest share error ${(prop * 100).toFixed(3)} pp` },
  ];
}

// Typography: no text below 8 pt; one size per label role (category row, values, totals, series, axis).
export function typeChecks(sc: Scene): Check[] {
  const all = texts(sc, /./);
  const small = all.filter(({ t }) => t.size < 8 - 1e-9);
  const byRole = new Map<string, Set<number>>();
  for (const { t, role } of all) {
    const k = role.replace(/:(s|c)\d+/g, '').replace(/:\d+$/, '');
    if (!byRole.has(k)) byRole.set(k, new Set());
    byRole.get(k)!.add(t.size);
  }
  const mixed = [...byRole].filter(([, s]) => s.size > 1).map(([k, s]) => `${k} ${[...s].join('/')} pt`);
  return [
    { name: 'min size', ok: !small.length, detail: small.map(({ t }) => `"${t.text}" ${t.size} pt`).join(', ') },
    { name: 'one size per role', ok: !mixed.length, detail: mixed.join(', ') },
  ];
}
