// Focused tests for rules in docs/behaviour-spec.md that no other test pinned down. Each test names
// the rule id(s) it covers; the spec lists every rule with its implementation, source and tests.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import type { ChartSpec, Frame, LineShape, PathShape, RectShape, Scene, Shape, TextShape } from '../src/core/index.ts';
import { defaultStyle, layoutChart, layoutGroup, minimumSize, resolveStyle } from '../src/core/index.ts';
import { formatValue, localeOf, niceTicks } from '../src/core/numbers.ts';
import { parseCategoryDate } from '../src/core/dates.ts';
import { breakPieces, measurerFromMetrics, wrap } from '../src/core/text.ts';
import { resolveCollisions } from '../src/core/resolve.ts';
import { makeCtx, valueTicks } from '../src/core/common.ts';
import { placePending } from '../src/core/labels.ts';
import { overlaps, segmentHitsBox } from '../src/core/place.ts';
import { samples } from '../src/core/samples.ts';
import { cachedMetrics, createMeasurer } from '../src/fonts.ts';
import { CHART_FRAME } from '../src/deck.ts';

const m = createMeasurer();
const F = CHART_FRAME;
const lay = (spec: ChartSpec, frame: Frame = F) => layoutChart(spec, frame, m);
const clone = <T>(v: T): T => structuredClone(v);
const byRole = (s: Scene, re: RegExp) => s.shapes.filter((_, i) => re.test(s.roles![i]));
const textsOf = (s: Scene, re: RegExp) => byRole(s, re).filter((x): x is TextShape => x.kind === 'text').map((t) => t.text);
const rectsOf = (s: Scene, re: RegExp) => byRole(s, re).filter((x): x is RectShape => x.kind === 'rect');
const linesOf = (s: Scene, re: RegExp) => byRole(s, re).filter((x): x is LineShape => x.kind === 'line');
const M = '−';
/** The default style as a scene shows it: theme slots resolved ("1F3A5F@accent1"). */
const DS = resolveStyle(defaultStyle);

// ---- Numbers ------------------------------------------------------------------------------------

test('AX-2/3: empty axis range costs more than an extra tick (620 → 0–700, not 0–800); 3–9 ticks', () => {
  assert.deepEqual(niceTicks(0, 620, 5), [0, 100, 200, 300, 400, 500, 600, 700]);
  for (const [a, b, n] of [[0, 1, 3], [0, 999, 5], [-37, 12, 4], [3, 3, 5]] as const) {
    const t = niceTicks(a, b, n);
    assert.ok(t.length >= 3 && t.length <= 9, `${a}–${b}: ${t}`);
    assert.ok(t[0] <= a && t[t.length - 1] >= Math.max(b, a + 1));
  }
});

test('NUM-4: an unknown locale falls back to en-US; a language-only id to its country', () => {
  assert.equal(localeOf('xx-YY'), localeOf('en-US'));
  assert.equal(localeOf('fr'), localeOf('fr-FR'));
  assert.equal(formatValue(1234.5, { decimals: 1, locale: 'pt-BR' }), '1,234.5');
});

test('NUM-5: no sign on a value that rounds to zero; "+" only on signed positives', () => {
  assert.equal(formatValue(-0.004, { decimals: 1 }), '0.0');
  assert.equal(formatValue(0.004, { decimals: 1, signed: true }), '0.0');
  assert.equal(formatValue(-2, { decimals: 0, signed: true }), `${M}2`);
  assert.equal(formatValue(2, { decimals: 0, signed: true }), '+2');
});

test('DATE-2: two-digit years pivot at 70; estimate markers and apostrophes are dropped', () => {
  assert.deepEqual(parseCategoryDate("'24"), { y: 2024, m: 1, d: 1 });
  assert.deepEqual(parseCategoryDate('FY99'), { y: 1999, m: 1, d: 1 });
  assert.deepEqual(parseCategoryDate('2026 fcst'), { y: 2026, m: 1, d: 1 });
  assert.deepEqual(parseCategoryDate('01/2024'), { y: 2024, m: 1, d: 1 });
  assert.deepEqual(parseCategoryDate('2024 Mar'), { y: 2024, m: 3, d: 1 });
  assert.equal(parseCategoryDate('Q1'), null);
});

// ---- Arrows -------------------------------------------------------------------------------------

const growth: ChartSpec = {
  type: 'stacked', categories: ['2021', '2022', '2023'],
  series: [{ name: 'A', values: [60, 70, 90] }, { name: 'B', values: [40, 50, 30] }],
};

test('ARR-3: the difference arrow shows the signed absolute difference of the column totals', () => {
  const s = lay({ ...growth, options: { difference: { from: 0, to: 2 } } });
  assert.deepEqual(textsOf(s, /^arrow:diff$/), ['+20']);
  const down = lay({ ...growth, options: { difference: { from: 1, to: 2 } } });
  assert.deepEqual(textsOf(down, /^arrow:diff$/), ['0'], 'no difference, no sign');
});

test('ARR-4: no CAGR for a non-positive total: the arrow is left out and the reason reported', () => {
  const s = lay({ ...growth, series: [{ name: 'A', values: [-10, 70, 90] }], options: { cagr: { from: 0, to: 2 } } });
  assert.equal(byRole(s, /^arrow:cagr$/).length, 0);
  assert.ok(s.hidden.includes('cagr: needs positive totals'));
});

test('ARR-1: 100% charts compute arrows from the absolute column totals', () => {
  const s = lay({ ...growth, type: 'percent', options: { cagr: { from: 0, to: 2 } } });
  assert.deepEqual(textsOf(s, /^arrow:cagr$/), ['+9.5%']); // 100 → 120 over 2 years
});

// ---- Stacked / clustered / 100% columns --------------------------------------------------------

test('COL-1: negative segments hang below the baseline, positives stack above it', () => {
  const s = lay({ type: 'stacked', categories: ['A', 'B'], series: [{ name: 'P', values: [50, 40] }, { name: 'N', values: [-20, -10] }] });
  const base = linesOf(s, /^axis$/).find((l) => l.y1 === l.y2)!.y1;
  for (const b of rectsOf(s, /^bar:s0:/)) assert.ok(Math.abs(b.y + b.h - base) < 0.01);
  for (const b of rectsOf(s, /^bar:s1:/)) assert.ok(Math.abs(b.y - base) < 0.01);
});

test('COL-5: a 100% chart ignores negative values and says so', () => {
  const s = lay({ type: 'percent', categories: ['A', 'B'], series: [{ name: 'P', values: [50, 40] }, { name: 'N', values: [-20, 60] }] });
  assert.ok(s.hidden.includes('negative values ignored in 100% chart'));
  assert.deepEqual(textsOf(s, /^label:value:s0:c0$/), ['100%']);
});

test('TOT-2: totals wider than their slots are all left out, not every other one', () => {
  const cats = Array.from({ length: 24 }, (_, i) => `${i + 1}`);
  const s = lay({ type: 'stacked', categories: cats, series: [{ name: 'A', values: cats.map(() => 123456) }], options: { seriesLabels: 'none' } }, { x: 0, y: 0, w: 380, h: 300 });
  assert.equal(textsOf(s, /^label:total:/).length, 0);
  assert.ok(s.hidden.includes('column totals wider than their slots: hidden'));
});

test('COL-7: column width is slot / (1 + gap)', () => {
  const spec: ChartSpec = { type: 'clustered', categories: ['A', 'B', 'C'], series: [{ name: 'S', values: [3, 4, 5] }], options: { valueAxis: true, seriesLabels: 'none' } };
  const w = (gap: number) => rectsOf(lay({ ...spec, options: { ...spec.options, gap } }), /^bar:/)[0].w;
  assert.ok(Math.abs(w(0) / w(1) - 2) < 1e-6);
  assert.ok(Math.abs(w(0.5) / w(1) - 4 / 3) < 1e-6);
});

test('SER-1: default series identification: names right for stacked, a legend for clustered and horizontal bars', () => {
  assert.ok(lay(samples.stacked!).roles!.some((r) => r.startsWith('label:series:')));
  assert.ok(lay(samples.clustered!).roles!.some((r) => r.startsWith('legend:')));
  assert.ok(lay({ ...samples.stacked!, options: { direction: 'right' } }).roles!.some((r) => r.startsWith('legend:')));
});

test('SER-3: names beside the plot that would take over a third of the frame become a legend', () => {
  const long: ChartSpec = { type: 'stacked', categories: ['A', 'B'], series: [{ name: 'A very long series name indeed', values: [3, 4] }, { name: 'B', values: [2, 2] }] };
  const narrow = lay(long, { x: 0, y: 0, w: 220, h: 200 });
  assert.ok(narrow.roles!.some((r) => r.startsWith('legend:')));
  assert.ok(!narrow.roles!.some((r) => r.startsWith('label:series:')));
  assert.ok(lay(long).roles!.some((r) => r.startsWith('label:series:')));
});

test('SER-2: a series name moved off its segment gets a leader; one level with it does not', () => {
  const s = lay(samples.stacked!);
  const leaders = linesOf(s, /^label:series:/);
  const names = byRole(s, /^label:series:/).filter((x) => x.kind === 'text').length;
  assert.equal(names, 3);
  for (const l of leaders) assert.ok(Math.abs(l.y2 - l.y1) > 1, 'a leader only where the name moved');
});

test('LBL-11: segment labels of the last column never go right when the series names are there', () => {
  const spec: ChartSpec = { type: 'stacked', categories: ['A', 'B', 'C'], series: [{ name: 'Big', values: [80, 90, 100] }, { name: 'Thin', values: [2, 3, 2] }, { name: 'Top', values: [40, 40, 40] }] };
  const s = lay(spec);
  const lastBar = rectsOf(s, /^bar:s1:c2$/)[0];
  const label = byRole(s, /^label:value:s1:c2$/)[0] as TextShape;
  assert.ok(label, 'thin last segment labelled');
  assert.ok(label.x + label.w <= lastBar.x + lastBar.w + 0.01, 'not to the right of its column');
});

test('AX-1: a short value axis asks for fewer ticks (3–5 by length)', () => {
  const ticks = (h: number, w = 600) => valueTicks(makeCtx(samples.stacked!, defaultStyle, { x: 0, y: 0, w, h }, m));
  assert.equal(ticks(400), 5);
  assert.equal(ticks(130), 3);
  assert.equal(valueTicks(makeCtx(samples.stacked!, defaultStyle, { x: 0, y: 0, w: 200, h: 400 }, m), true), 3);
});

// ---- Label placer -------------------------------------------------------------------------------

test('LBL-9: a boxed label never covers a neighbouring segment whose own label points at it from outside', () => {
  const spec: ChartSpec = { type: 'stacked', categories: ['A'], series: [{ name: 'x', values: [1] }] };
  const ctx = makeCtx(spec, defaultStyle, { x: 0, y: 0, w: 400, h: 300 }, m);
  const segA = { kind: 'rect', x: 180, y: 140, w: 40, h: 4, fill: '3E6A9E' } as const;
  const segB = { kind: 'rect', x: 180, y: 144, w: 40, h: 9, fill: '1F3A5F' } as const;
  ctx.add('bar:s0:c0', segA); ctx.add('bar:s1:c0', segB);
  placePending(ctx, [
    { s: 0, c: 0, text: '4', seg: segA, sides: ['right', 'left'], boxCover: [segA, segB], reason: 'A' },
    { s: 1, c: 0, text: '9', seg: segB, sides: ['right', 'left'], boxCover: [segA, segB], reason: 'B' },
  ]);
  const labelB = ctx.shapes[ctx.roles.indexOf('label:value:s1:c0')] as TextShape;
  assert.ok(labelB, 'B is placed');
  assert.ok(!overlaps(labelB, segA, 0.5), 'B does not sit over A');
});

// ---- Waterfall ----------------------------------------------------------------------------------

test('WF-5: single-series colours: totals palette[0], increases positive, decreases negative', () => {
  const s = lay({ type: 'waterfall', categories: ['S', 'Up', 'Down', 'E'], series: [{ name: 'v', values: [100, 20, -30, 'e'] }] });
  assert.equal(rectsOf(s, /^bar:s0:c0$/)[0].fill, DS.palette[0]);
  assert.equal(rectsOf(s, /^bar:s0:c1$/)[0].fill, defaultStyle.positive);
  assert.equal(rectsOf(s, /^bar:s0:c2$/)[0].fill, defaultStyle.negative);
  assert.equal(rectsOf(s, /^bar:s0:c3$/)[0].fill, DS.palette[0]);
});

test('WF-3: the value of a later total is ignored: totals show the running sum', () => {
  const s = lay({ type: 'waterfall', categories: ['S', 'Up', 'E'], series: [{ name: 'v', values: [100, 20, 999] }] });
  assert.deepEqual(textsOf(s, /^label:total:c2$/), ['120']);
});

// ---- Mekko / Mekko XY ---------------------------------------------------------------------------

test('MK-4: each column\'s share of the grand total sits under its name, rounded to add up to 100%', () => {
  const s = lay(samples.mekko!);
  const shares = textsOf(s, /^label:category:/).filter((t) => /%$/.test(t)).map((t) => parseFloat(t));
  assert.equal(shares.length, samples.mekko!.categories.length);
  assert.equal(shares.reduce((a, b) => a + b, 0), 100);
});

test('MK-8: a Mekko chart without positive totals draws nothing and says why', () => {
  const s = lay({ type: 'mekko', categories: ['A', 'B'], series: [{ name: 'x', values: [0, -5] }] });
  assert.ok(s.hidden.includes('Mekko chart needs positive totals'));
  assert.equal(rectsOf(s, /^bar:/).length, 0);
});

test('MXY-2/3/5: reverse order; one height series: values on top, axis titles from the series names, widths without the unit', () => {
  const spec = clone(samples.mekkoxy!);
  spec.options = { sort: 'reverse', unit: { text: '$', position: 'prefix' } };
  const s = lay(spec);
  const xs = spec.categories.map((_, c) => rectsOf(s, new RegExp(`^bar:s1:c${c}$`))[0].x);
  for (let c = 1; c < xs.length; c++) assert.ok(xs[c] < xs[c - 1], 'reverse sheet order');
  assert.ok(!s.roles!.some((r) => r.startsWith('label:series:') || r.startsWith('legend:')), 'no series names');
  assert.deepEqual(textsOf(s, /^axis$/).sort(), ['Capacity (Mt)', 'Cash cost ($/t)']);
  assert.ok(textsOf(s, /^label:value:s1:/).every((t) => t.startsWith('$')));
  assert.ok(textsOf(s, /^label:category:/).filter((t) => /^\d+$/.test(t)).length === spec.categories.length, 'widths plain');
});

// ---- Line ---------------------------------------------------------------------------------------

const lineOf = (values: (number | null)[][], options: ChartSpec['options'] = {}): ChartSpec => ({
  type: 'line', categories: values[0].map((_, i) => `${2020 + i}`),
  series: values.map((v, i) => ({ name: `S${i}`, values: v })), options,
});

test('LN-1: the value axis includes zero unless the data sits in a narrow band; zeroBased overrides', () => {
  const lowest = (s: Scene) => Math.min(...textsOf(s, /^axis:tick:/).map(Number));
  assert.equal(lowest(lay(lineOf([[10, 30, 50]]))), 0);
  assert.ok(lowest(lay(lineOf([[95, 100, 120]]))) > 0, 'narrow band floats');
  assert.equal(lowest(lay(lineOf([[95, 100, 120]], { zeroBased: true }))), 0);
  assert.ok(lowest(lay(lineOf([[30, 60, 90]], { zeroBased: false }))) > 0);
  assert.equal(lowest(lay(lineOf([[30, 60, 90]]))), 0);
});

test('LN-3: every point is labelled up to 24 points, beyond that only the last category', () => {
  const few = lay(lineOf([[1, 2, 3, 4], [2, 3, 4, 5]]));
  assert.equal(textsOf(few, /^label:value:/).length, 8);
  const many = lay(lineOf([Array.from({ length: 13 }, (_, i) => i + 1), Array.from({ length: 13 }, (_, i) => i + 20)]));
  assert.deepEqual(byRole(many, /^label:value:/).map((_, k) => k).length, 2);
  assert.ok(many.roles!.filter((r) => r.startsWith('label:value:')).every((r) => r.endsWith(':c12')));
});

test('LN-5/6: lines break at empty cells; a dashed zero line when the data crosses zero', () => {
  const s = lay(lineOf([[5, null, 7, -3]]));
  assert.equal(linesOf(s, /^line:s0$/).length, 1, 'only the 2022–2023 segment');
  assert.ok(linesOf(s, /^axis$/).some((l) => l.dash), 'dashed zero line');
  assert.ok(!linesOf(lay(lineOf([[5, 6, 7]])), /^axis$/).some((l) => l.dash));
});

test('LN-4: at a category the top series is labelled above its point, the bottom one below', () => {
  const s = lay(lineOf([[50, 52, 54], [40, 42, 44]]));
  const mk = (sIdx: number, c: number) => rectsOf(s, new RegExp(`^marker:s${sIdx}:c${c}$`))[0];
  const lb = (sIdx: number, c: number) => byRole(s, new RegExp(`^label:value:s${sIdx}:c${c}$`))[0] as TextShape;
  assert.ok(lb(0, 0).y + lb(0, 0).h <= mk(0, 0).y + 0.01, 'top above');
  assert.ok(lb(1, 0).y >= mk(1, 0).y + mk(1, 0).h - 0.01, 'bottom below');
});

// ---- Area ---------------------------------------------------------------------------------------

test('AR-4/5: area charts ignore negatives (reported); totals sit above the top band of a stacked area', () => {
  const s = lay({ type: 'area', categories: ['A', 'B', 'C'], series: [{ name: 'x', values: [10, -5, 12] }, { name: 'y', values: [8, 9, 10] }], options: { totals: true } });
  assert.ok(s.hidden.includes('negative values ignored in area chart'));
  assert.deepEqual(textsOf(s, /^label:total:/), ['18', '9', '22']);
});

test('AR-7: a single category is drawn as a flat band across the plot', () => {
  const s = lay({ type: 'area', categories: ['2024'], series: [{ name: 'x', values: [10] }, { name: 'y', values: [5] }] });
  const band = byRole(s, /^area:s0$/)[0] as PathShape;
  const pts = band.commands.filter((c): c is { op: 'M' | 'L'; x: number; y: number } => c.op === 'M' || c.op === 'L');
  const xs = pts.map((p) => p.x);
  assert.ok(Math.max(...xs) - Math.min(...xs) > F.w * 0.5, 'spans the plot');
  assert.equal(new Set(pts.slice(0, 2).map((p) => p.y.toFixed(3))).size, 1, 'flat top');
});

// ---- Pie / doughnut -----------------------------------------------------------------------------

test('PIE-10: beyond the palette, slice colours repeat as lighter tints: no two slices match', () => {
  const n = defaultStyle.palette.length + 3;
  const cats = Array.from({ length: n }, (_, i) => `C${i}`);
  const s = lay({ type: 'pie', categories: cats, series: [{ name: 'x', values: cats.map(() => 1) }] });
  const fills = (byRole(s, /^wedge:/) as PathShape[]).map((p) => p.fill);
  assert.equal(new Set(fills).size, n);
});

test('PIE-11: on a narrow frame the category names go to a legend on top', () => {
  const s = lay(samples.pie!, { x: 0, y: 0, w: 200, h: 300 });
  assert.ok(s.roles!.some((r) => r.startsWith('legend:c')));
  assert.ok(!s.roles!.some((r) => r.startsWith('label:category:')));
});

test('PIE-12: a doughnut total too wide for the hole is left out and reported', () => {
  const s = lay({ type: 'doughnut', categories: ['A', 'B'], series: [{ name: 'x', values: [1234567890, 987654321] }], options: { totals: true } }, { x: 0, y: 0, w: 200, h: 120 });
  assert.ok(s.hidden.includes('total too wide for the doughnut hole'));
  assert.equal(textsOf(s, /^label:total$/).length, 0);
});

// ---- Harvey / combo / scatter -------------------------------------------------------------------

test('HB-5: balls use the series colour, else the first palette colour', () => {
  const spec = clone(samples.harvey!);
  spec.series[0].color = 'C0504D';
  const s = lay(spec);
  assert.ok(byRole(s, /^ball:s0:/).every((x) => (x as RectShape).stroke === 'C0504D' || x.kind === 'path'));
  assert.ok(byRole(s, /^ball:s1:/).every((x) => (x as RectShape).stroke === DS.palette[0] || x.kind === 'path'));
});

test('CB-3: lines default to dark greys over a white halo; a series colour wins', () => {
  const spec: ChartSpec = { type: 'combo', categories: ['A', 'B', 'C'], series: [{ name: 'Col', values: [10, 12, 14] }, { name: 'L1', values: [5, 6, 7], as: 'line' }, { name: 'L2', values: [3, 4, 5], as: 'line', color: 'C0504D' }] };
  const s = lay(spec);
  const l1 = linesOf(s, /^line:s1$/);
  assert.deepEqual(new Set(l1.map((l) => `${l.color}/${l.width}`)), new Set([`${DS.background}/4`, `${DS.text}/2`]));
  assert.ok(linesOf(s, /^line:s2$/).some((l) => l.color === 'C0504D' && l.width === 2));
});

test('SC-6/7: scatter series: no names with one group, a legend with several; the x label at the origin gives way to the y label', () => {
  const one = lay(samples.scatter!);
  assert.ok(!one.roles!.some((r) => r.startsWith('legend:') || r.startsWith('label:series:')));
  const two = clone(samples.scatter!);
  two.series.push({ name: 'Other', values: two.series[0].values.map((v) => [(v as number[])[0] + 1, (v as number[])[1] + 1]) });
  assert.ok(lay(two).roles!.some((r) => r.startsWith('legend:s1')));
  // Both axes start at 0: the x axis "0" would sit on the y axis "0".
  const s = lay({ type: 'scatter', categories: ['a', 'b'], series: [{ name: 'x', values: [[1, 1], [9, 9]] }], options: { unit: '%' } });
  const zeros = textsOf(s, /^axis:tick:/).filter((t) => t === '0%');
  assert.equal(zeros.length, 1);
});

// ---- Table --------------------------------------------------------------------------------------

test('TBL-2/3: decimals per row under period headers, per column otherwise; text columns left, numbers right', () => {
  const byRowSpec: ChartSpec = { type: 'table', categories: ['2023', '2024'], series: [{ name: 'Revenue', values: [100, 120] }, { name: 'Margin', values: [0.15, 0.2] }] };
  const r = lay(byRowSpec);
  assert.deepEqual(textsOf(r, /^label:value:s0:/), ['100', '120']);
  assert.deepEqual(textsOf(r, /^label:value:s1:/), ['0.15', '0.20']);
  const byCol: ChartSpec = { type: 'table', categories: ['Price', 'Status'], series: [{ name: 'A', values: [1.5, 'Open'] }, { name: 'B', values: [2, 'Closed'] }] };
  const c = lay(byCol);
  assert.deepEqual(textsOf(c, /^label:value:s1:c0$/), ['2.0']);
  const at = (re: RegExp) => (byRole(c, re)[0] as TextShape).align;
  assert.equal(at(/^label:value:s0:c0$/), 'right');
  assert.equal(at(/^label:value:s0:c1$/), 'left');
});

// ---- Gantt --------------------------------------------------------------------------------------

const plan = (start: string, end: string, ms: string, options: ChartSpec['options'] = {}): ChartSpec => ({
  type: 'gantt', categories: ['Work', 'Gate'],
  series: [{ name: 'Start', values: [start, null] }, { name: 'End', values: [end, null] }, { name: 'Milestone', values: [null, ms] }],
  options,
});

test('GNT-1: a date is the start of its day: a bar ending on the 2nd of a month covers one day', () => {
  const s = lay(plan('2026-03-01', '2026-03-31', '2026-03-15'));
  const one = lay({ ...plan('2026-03-01', '2026-03-02', '2026-03-31') });
  const w30 = rectsOf(s, /^bar:s0:c0$/)[0].w, w1 = rectsOf(one, /^bar:s0:c0$/)[0].w;
  assert.ok(Math.abs(w30 / w1 - 30) < 1e-6);
});

test('GNT-9: milestone dates: "Mar 15" within a year, with the year beyond, month and year beyond three; day first outside en-US', () => {
  const label = (s: Scene) => textsOf(s, /^label:value:s2:/)[0];
  assert.equal(label(lay(plan('2026-01-05', '2026-06-01', '2026-03-16'))), 'Mar 16');
  assert.equal(label(lay(plan('2026-01-05', '2027-06-01', '2026-03-16'))), 'Mar 16, 2026');
  assert.equal(label(lay(plan('2026-01-05', '2030-06-01', '2026-03-16'))), 'Mar 2026');
  assert.equal(label(lay(plan('2026-01-05', '2026-06-01', '2026-03-16', { locale: 'en-GB' }))), '16 Mar');
});

test('GNT-5/12/13: upper-tier separators are solid, lower-tier dashed; key only with two milestone kinds; today outside the range is reported', () => {
  const s = lay(plan('2026-01-05', '2026-06-01', '2026-03-16'));
  const grid = linesOf(s, /^gridline:/);
  assert.ok(grid.some((l) => l.dash) && grid.some((l) => !l.dash && l.x1 === l.x2));
  assert.ok(!s.roles!.some((r) => r.startsWith('legend:')));
  const two = plan('2026-01-05', '2026-06-01', '2026-03-16');
  two.series.push({ name: 'Review', values: ['2026-02-02', null] });
  assert.ok(lay(two).roles!.some((r) => r === 'legend:s3'));
  assert.ok(!lay({ ...two, options: { seriesLabels: 'none' } }).roles!.some((r) => r.startsWith('legend:')));
  const late = lay(plan('2026-01-05', '2026-06-01', '2026-03-16', { today: '2030-01-01' }));
  assert.ok(late.hidden.includes('today: outside the chart\'s date range'));
});

// ---- Category labels ----------------------------------------------------------------------------

test('CAT-3: a crowded time axis shows every n-th label instead of hiding at random', () => {
  const years = Array.from({ length: 30 }, (_, i) => `${1995 + i}`);
  const s = lay({ type: 'line', categories: years, series: [{ name: 'x', values: years.map((_, i) => i) }] }, { x: 0, y: 0, w: 480, h: 240 });
  const shown = s.roles!.filter((r) => r.startsWith('label:category:')).map((r) => +r.slice(r.lastIndexOf('c') + 1));
  const step = shown[1] - shown[0];
  assert.ok(step >= 2 && shown.every((c, k) => c === shown[0] + k * step), `every ${step}th: ${shown}`);
  assert.ok(s.hidden.some((h) => /category labels thinned to every/.test(h)));
});

test('CAT-4: names that cannot all show in a frame narrower than a slide make the chart ask for room', () => {
  const cats = ['Telecommunications', 'Pharmaceuticals', 'Infrastructure', 'Semiconductors', 'Transportation', 'Entertainment', 'Manufacturing', 'Agriculture'];
  const s = lay({ type: 'stacked', categories: cats, series: [{ name: 'x', values: cats.map(() => 5) }] }, { x: 0, y: 0, w: 300, h: 240 });
  assert.deepEqual(s.roles, ['placeholder', 'placeholder']);
  assert.match(s.hidden[0], /category labels need wider slots/);
});

// ---- Text, minimum sizes, resolution ------------------------------------------------------------

test('TXT-5: without a bold font, bold text is measured 6% wider', () => {
  const font = ['/System/Library/Fonts/Supplemental/Arial.ttf', 'C:\\Windows\\Fonts\\arial.ttf', '/usr/share/fonts/truetype/msttcorefonts/Arial.ttf',
    '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf', '/usr/share/fonts/liberation-sans/LiberationSans-Regular.ttf'].find(existsSync)!;
  const regular = measurerFromMetrics(cachedMetrics(font));
  assert.ok(Math.abs(regular.width('Revenue', 12, true) / regular.width('Revenue', 12) - 1.06) < 1e-9);
});

test('TXT-3: no line starts with closing punctuation; a dash before a space or a closer does not break', () => {
  assert.deepEqual(breakPieces('東京、大阪').map((p) => p.s), ['東', '京、大', '阪']);
  assert.deepEqual(breakPieces('A-) b').map((p) => p.s), ['A-)', 'b']);
  assert.deepEqual(wrap('one two three four five', m.width('one two', 12) + 1, 12, m, 2), ['one two', 'three four five'], 'the last line keeps the overflow');
});

test('PLC-3: segment–box intersection (Liang–Barsky) counts crossings and containment, not near misses', () => {
  const b = { x: 10, y: 10, w: 10, h: 10 };
  assert.equal(segmentHitsBox(0, 0, 30, 30, b), true);
  assert.equal(segmentHitsBox(12, 12, 14, 14, b), true);
  assert.equal(segmentHitsBox(0, 25, 30, 25, b), false);
  assert.equal(segmentHitsBox(0, 0, 9, 30, b), false);
});

test('LEG-2: legend swatches follow the mark: boxes for columns and areas, short lines for lines, circles for points', () => {
  const kinds = (s: Scene) => byRole(s, /^legend:s0$/).filter((x) => x.kind !== 'text').map((x) => (x.kind === 'rect' ? x.geom ?? 'rect' : x.kind));
  assert.deepEqual(kinds(lay(samples.clustered!)), ['rect']);
  assert.deepEqual(kinds(lay({ ...samples.line!, options: { seriesLabels: 'legend' } })), ['line']);
  assert.deepEqual(kinds(lay({ ...samples.area!, options: { seriesLabels: 'legend' } })), ['rect']);
  const two = clone(samples.scatter!);
  two.series.push({ name: 'Other', values: two.series[0].values.map(() => null) });
  two.series[1].values[0] = [1, 1];
  assert.deepEqual(kinds(lay(two)), ['ellipse']);
});

test('MIN-2: minimum sizes per type grow with the categories', () => {
  const spec = (type: ChartSpec['type'], n: number, options?: ChartSpec['options']) => ({ type, categories: Array.from({ length: n }, (_, i) => `${i}`), series: [{ name: 's', values: [] }], options }) as unknown as ChartSpec;
  assert.deepEqual(minimumSize(spec('stacked', 20)), { w: 320, h: 120 });
  assert.deepEqual(minimumSize(spec('stacked', 20, { direction: 'right' })), { w: 200, h: 270 });
  assert.deepEqual(minimumSize(spec('mekko', 12)), { w: 288, h: 140 });
  assert.deepEqual(minimumSize(spec('pie', 30)), { w: 200, h: 120 });
  assert.deepEqual(minimumSize(spec('scatter', 3)), { w: 200, h: 140 });
  assert.deepEqual(minimumSize(spec('harvey', 5)), { w: 180, h: 80 });
  assert.deepEqual(minimumSize(spec('table', 5)), { w: 120, h: 50 });
  assert.deepEqual(minimumSize(spec('gantt', 10)), { w: 320, h: 180 });
});

test('RES-1/2/3: collisions keep category labels over value labels, drop arrows first, and take a hidden label\'s leader with it', () => {
  const t = (x: number, y: number, text: string): Shape => ({ kind: 'text', x, y, w: 40, h: 14, text, size: 12, color: '000000', align: 'center', valign: 'middle' });
  const shapes: Shape[] = [
    t(10, 10, 'Value'), t(20, 12, 'Cat'),
    { kind: 'line', x1: 0, y1: 50, x2: 30, y2: 30, color: '000000', width: 0.5 }, t(60, 30, 'Seg'), t(62, 32, 'Total'),
    { kind: 'line', x1: 100, y1: 5, x2: 100, y2: 80, color: '000000', width: 0.75 }, t(150, 60, 'Arrow'), t(152, 60, 'Axis'),
  ];
  const roles = ['label:value:s0:c0', 'label:category:c0', 'leader:s1:c1', 'label:value:s1:c1', 'label:total:c2', 'arrow:cagr', 'arrow:cagr', 'axis:tick:0'];
  const why = resolveCollisions(shapes, roles, { x: 0, y: 0, w: 300, h: 100 }, m);
  assert.deepEqual(roles, ['label:category:c0', 'label:total:c2', 'axis:tick:0']);
  assert.equal(why.length, 3);
});

test('RES-4: at the default frame the last-resort pass hides nothing in any sample', () => {
  for (const [type, spec] of Object.entries(samples)) {
    const s = lay(spec!);
    assert.deepEqual(s.hidden.filter((h) => /\) hidden: /.test(h)), [], type);
  }
});

// ---- The spec itself ----------------------------------------------------------------------------

// ---- Theme and same scale -----------------------------------------------------------------------

test('THM-7: a chart with a value axis break keeps its own scale in a same-scale group', () => {
  const broken: ChartSpec = { type: 'stacked', categories: ['A', 'B', 'C'], series: [{ name: 'v', values: [20, 30, 400] }], options: { axisBreak: { from: 40, to: 330 }, scaleGroup: 'g' } };
  const small: ChartSpec = { type: 'stacked', categories: ['A', 'B', 'C'], series: [{ name: 'v', values: [2, 3, 4] }], options: { scaleGroup: 'g' } };
  const half = (i: number) => ({ ...F, x: F.x + i * (F.w / 2), w: F.w / 2 - 12 });
  const [sb, ss] = layoutGroup([{ spec: broken, frame: half(0) }, { spec: small, frame: half(1) }], m);
  assert.equal(sb.valueScale, undefined, 'no linear scale to share');
  assert.equal(ss.valueScale!.ptPerUnit, ss.valueScale!.own, 'the other chart keeps its own scale');
});

test('docs/behaviour-spec.md: unique rule ids, the count is right, and every test it names exists', () => {
  const doc = readFileSync(new URL('../docs/behaviour-spec.md', import.meta.url), 'utf8');
  const ids = [...doc.matchAll(/^\| ([A-Z]+-\d+) \|/gm)].map((x) => x[1]);
  assert.equal(new Set(ids).size, ids.length, 'unique ids');
  assert.match(doc, new RegExp(`\\*\\*Count: ${ids.length} rules\\*\\*`));
  const dir = new URL('./', import.meta.url);
  const src: Record<string, string> = Object.fromEntries(readdirSync(dir).filter((f) => f.endsWith('.test.ts'))
    .map((f) => [f.replace('.test.ts', ''), readFileSync(new URL(f, dir), 'utf8')]));
  src.e2e = readFileSync(new URL('e2e/editor.spec.ts', dir), 'utf8');
  for (const [, file, title] of doc.matchAll(/([a-z0-9-]+) › "([^"]+)"/g)) {
    if (title === 'matrix') continue;
    const prefix = title.split('…')[0].trim();
    const body = src[file];
    assert.ok(body && (body.includes(prefix) || body.includes(prefix.replace(/'/g, "\\'"))), `${file} › "${title}"`);
  }
  // Every focused test in this file is listed under its rule.
  for (const [, id] of src.behaviour.matchAll(/^test\('([^:]+):/gm)) {
    if (id.startsWith('docs/')) continue;
    assert.ok(doc.includes(`behaviour › "${id}"`), id);
  }
});
