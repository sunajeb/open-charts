// Percent axis (100% with percentData), combination, Mekko XY, scatter, bubble, table and Gantt.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import {
  checkScene, layoutChart, validateDeck, validateSpec,
  type ChartSpec, type LineShape, type RectShape, type Scene, type TextShape,
} from '../src/core/index.ts';
import { samples } from '../src/core/samples.ts';
import { createMeasurer } from '../src/fonts.ts';
import { CHART_FRAME } from '../src/deck.ts';

const m = createMeasurer();
const F = CHART_FRAME;
const byRole = (s: Scene, re: RegExp) => s.shapes.filter((_, i) => re.test(s.roles![i]));
const rectsOf = (s: Scene, re: RegExp) => byRole(s, re).filter((x): x is RectShape => x.kind === 'rect');
const textsOf = (s: Scene, re: RegExp) => byRole(s, re).filter((x): x is TextShape => x.kind === 'text');
const strict = (s: Scene) => checkScene(s, { strict: true, measurer: m });
const ROLE = /^(bar:s\d+:c\d+|label:value:s\d+:c\d+|label:total:c\d+|label:category:c\d+|label:series:s\d+|axis|axis:tick:\d+|gridline:\d+|legend:s\d+|marker:s\d+:c\d+|line:s\d+|shade:\d+|leader:c\d+|legend:size|bracket:\d+|today)$/;
const NEW = ['combo', 'mekkoxy', 'scatter', 'bubble', 'table', 'gantt'] as const;
const close = (a: number, b: number, eps = 1e-6) => Math.abs(a - b) < eps;

test('samples: every new type has one, it lays out with roles and passes the strict gate', () => {
  for (const type of NEW) {
    const spec = samples[type];
    assert.ok(spec, `sample for ${type}`);
    assert.deepEqual(validateSpec(spec), [], type);
    const scene = layoutChart(spec!, F, m);
    assert.ok(scene.shapes.length > 5, type);
    for (const r of scene.roles!) assert.match(r, ROLE, `${type}: ${r}`);
    assert.deepEqual(strict(scene), [], type);
  }
});

test('example decks validate and every slide passes the strict gate', () => {
  for (const f of ['combo', 'mekkoxy', 'percent', 'scatter', 'table', 'gantt']) {
    const deck = JSON.parse(readFileSync(`examples/${f}.json`, 'utf8'));
    assert.deepEqual(validateDeck(deck), [], f);
    for (const s of deck.slides) {
      const scene = layoutChart(s.chart, F, m);
      for (const r of scene.roles!) assert.match(r, ROLE, `${f}: ${r}`);
      const issues = strict(scene).filter((i) => !/Services \/ Consumer: segment too small/.test(i.detail));
      assert.deepEqual(issues, [], `${f}: ${s.title}`);
    }
  }
});

// ---- Percent axis -------------------------------------------------------------------------------

test('100% chart with a value axis: ticks and labels in %, every column reaches 100%', () => {
  const scene = layoutChart({ ...samples.percent!, options: { valueAxis: true } }, F, m);
  const ticks = textsOf(scene, /^axis:tick:/).map((t) => t.text);
  assert.equal(ticks[0], '0%');
  assert.equal(ticks.at(-1), '100%');
  for (const t of textsOf(scene, /^label:value:/)) assert.match(t.text, /^\d+%$/);
  const tops = new Set(rectsOf(scene, /^bar:s2:/).map((r) => r.y.toFixed(6)));
  assert.equal(tops.size, 1);
  assert.deepEqual(strict(scene), []);
});

test('percentData: values are shares already, columns keep their own totals on the % axis', () => {
  const spec: ChartSpec = {
    type: 'percent', categories: ['A', 'B'],
    series: [{ name: 'Yes', values: [40, 70] }, { name: 'No', values: [20, 10] }],
    options: { percentData: true, valueAxis: true },
  };
  const scene = layoutChart(spec, F, m);
  const yes = rectsOf(scene, /^bar:s0:/), no = rectsOf(scene, /^bar:s1:/);
  const base = yes[0].y + yes[0].h;
  // Heights proportional to the entered percentages, A tops out at 60%, B at 80%.
  assert.ok(close(yes[0].h / yes[1].h, 40 / 70));
  assert.ok(close((base - no[0].y) / (base - no[1].y), 60 / 80));
  assert.deepEqual(textsOf(scene, /^label:value:/).map((t) => t.text).sort(), ['10%', '20%', '40%', '70%']);
  assert.equal(textsOf(scene, /^axis:tick:/).at(-1)!.text, '100%');
  assert.deepEqual(strict(scene), []);
  assert.deepEqual(validateSpec({ ...spec, type: 'stacked' }).map((e) => e.path), ['/options/percentData']);
});

// ---- Combination ---------------------------------------------------------------------------------

test('combo: stacked columns plus a line series on the same value axis', () => {
  const scene = layoutChart(samples.combo!, F, m);
  assert.equal(rectsOf(scene, /^bar:/).length, 10);
  assert.equal(rectsOf(scene, /^bar:s2:/).length, 0, 'the line series draws no columns');
  assert.equal(rectsOf(scene, /^marker:s2:c\d$/).length, 5);
  const segs = byRole(scene, /^line:s2$/).filter((x): x is LineShape => x.kind === 'line' && x.width === 2);
  assert.equal(segs.length, 4);
  // Same axis: the 2021 plan point (90) sits above the 2021 column top (86).
  const col = rectsOf(scene, /^bar:s1:c0$/)[0];
  const mk = rectsOf(scene, /^marker:s2:c0$/)[0];
  const base = rectsOf(scene, /^bar:s0:c0$/)[0];
  const yOf = (v: number) => base.y + base.h - (v / 62) * base.h;
  assert.ok(close(col.y, yOf(86), 1e-6));
  assert.ok(close(mk.y + mk.h / 2, yOf(90), 1e-6));
  assert.deepEqual(strict(scene), []);
});

test('combo: clustered columns with a line; a line that covers labels hides them and says so', () => {
  const scene = layoutChart({
    type: 'combo', categories: ['A', 'B', 'C'],
    series: [{ name: 'X', values: [30, 32, 34] }, { name: 'Y', values: [28, 30, 33] }, { name: 'Avg', values: [29, 31, 33.5], as: 'line' }],
    options: { combo: 'clustered' },
  }, F, m);
  const bars = rectsOf(scene, /^bar:/);
  assert.equal(bars.length, 6);
  assert.ok(close(bars[0].w, bars[1].w) && close(bars[0].x + bars[0].w, bars[1].x), 'clustered side by side');
  assert.deepEqual(checkScene(scene, { measurer: m }), [], 'never an overlap');
  assert.ok(byRole(scene, /^legend:s2$/).some((x) => x.kind === 'line'), 'line swatch for the line series');
  assert.deepEqual(validateSpec({ ...samples.combo!, series: [{ name: 'a', values: [1, 2, 3, 4, 5], as: 'bar' as never }] }).map((e) => e.path), ['/series/0/as']);
});

// ---- Mekko XY ------------------------------------------------------------------------------------

test('mekkoxy: width from the first series, height from the rest, sorted by height', () => {
  const spec = samples.mekkoxy!;
  const scene = layoutChart(spec, F, m);
  const bars = rectsOf(scene, /^bar:s1:/).sort((a, b) => a.x - b.x);
  const widths = [12, 8, 15, 5, 9], heights = [41, 55, 38, 72, 49];
  const total = widths.reduce((a, b) => a + b, 0);
  const pw = bars.reduce((a, r) => a + r.w, 0);
  // Ascending by height: C, A, E, B, D; columns touch.
  const order = [2, 0, 4, 1, 3];
  bars.forEach((r, k) => {
    const c = order[k];
    assert.equal(byRole(scene, new RegExp(`^bar:s1:c${c}$`))[0], r);
    assert.ok(close(r.w / pw, widths[c] / total));
    assert.ok(close(r.h / bars[0].h, heights[c] / heights[2]));
    if (k) assert.ok(close(r.x, bars[k - 1].x + bars[k - 1].w));
  });
  assert.deepEqual(textsOf(scene, /^label:value:s1:/).map((t) => t.text).sort(), ['$38', '$41', '$49', '$55', '$72']);
  assert.deepEqual(strict(scene), []);
});

test('mekkoxy: value axes on both edges; stacked heights; zero width reported', () => {
  const scene = layoutChart({
    type: 'mekkoxy', categories: ['A', 'B', 'C', 'D'],
    series: [{ name: 'Users (m)', values: [40, 30, 0, 20] }, { name: 'H1', values: [10, 8, 5, 6] }, { name: 'H2', values: [5, 6, 2, 3] }],
    options: { valueAxis: true },
  }, F, m);
  const ticks = textsOf(scene, /^axis:tick:/).map((t) => t.text);
  assert.deepEqual(ticks.slice(-4), ['0', '25', '50', '75'], 'x ticks inside the total width (90)');
  assert.deepEqual(textsOf(scene, /^label:total:/).map((t) => t.text), ['15', '14', '9']);
  assert.ok(scene.hidden.some((h) => /C: zero width/.test(h)));
  assert.deepEqual(validateSpec({ type: 'mekkoxy', categories: ['a'], series: [{ name: 'w', values: [-1] }] }).map((e) => e.path), ['/series/0/values/0', '/series']);
});

// ---- Scatter and bubble -------------------------------------------------------------------------

test('scatter: points map linearly onto two nice value axes, labels never overlap', () => {
  const scene = layoutChart(samples.scatter!, F, m);
  const mk = (c: number) => { const r = rectsOf(scene, new RegExp(`^marker:s0:c${c}$`))[0]; return { x: r.x + r.w / 2, y: r.y + r.h / 2 }; };
  // x: 12.5 vs 4.2 vs 14.1 — ratios of differences are preserved.
  assert.ok(close((mk(0).x - mk(2).x) / (mk(3).x - mk(2).x), (12.5 - 4.2) / (14.1 - 4.2)));
  assert.ok(close((mk(0).y - mk(2).y) / (mk(3).y - mk(2).y), (42 - 31) / (47 - 31)));
  assert.ok(mk(3).y < mk(2).y, 'larger y is higher');
  const ticks = textsOf(scene, /^axis:tick:/).map((t) => Number(t.text));
  assert.ok(ticks.every(Number.isFinite));
  assert.equal(textsOf(scene, /^label:value:/).length, 6);
  assert.ok(textsOf(scene, /^axis$/).some((t) => t.text === 'Seats (thousands)'));
  assert.deepEqual(strict(scene), []);
});

test('scatter: labels that cannot be placed are hidden and reported, never overlapped', () => {
  const pts = [[0, 0], [100, 100], ...Array.from({ length: 30 }, (_, i) => [50 + (i % 3) * 0.2, 50 + Math.floor(i / 3) * 0.2])];
  const scene = layoutChart({ type: 'scatter', categories: pts.map((_, i) => `Point ${i}`), series: [{ name: 'S', values: pts }] }, F, m);
  assert.ok(scene.hidden.some((h) => /no room for point label/.test(h)));
  assert.deepEqual(checkScene(scene, { measurer: m }), []);
});

test('bubble: area is proportional to size; small bubbles in front', () => {
  const scene = layoutChart(samples.bubble!, F, m);
  const r = (c: number) => rectsOf(scene, new RegExp(`^marker:s0:c${c}$`))[0];
  assert.ok(close((r(0).w / r(3).w) ** 2, 420 / 90, 1e-9));
  const idx = (c: number) => scene.shapes.indexOf(r(c));
  assert.ok(idx(3) > idx(0), 'smaller drawn later');
  assert.ok(rectsOf(scene, /^marker:/).every((b) => b.geom === 'ellipse'));
  assert.deepEqual(strict(scene), []);
  const errs = validateSpec({ type: 'bubble', categories: ['a', 'b'], series: [{ name: 's', values: [[1, 2], [1, 2, -3]] }] });
  assert.deepEqual(errs.map((e) => e.path), ['/series/0/values/0', '/series/0/values/1/2']);
});

// ---- Table ----------------------------------------------------------------------------------------

test('table: header, row names and formatted cells in fitted, non-overlapping columns', () => {
  const scene = layoutChart({
    type: 'table', categories: ['2024', '2025'],
    series: [{ name: 'Revenue', values: [1234.5, 2000] }, { name: 'Headcount', values: [3120, 3340] }, { name: 'Rating', values: ['High', null] }],
    options: { locale: 'de-DE' },
  }, F, m);
  assert.deepEqual(textsOf(scene, /^label:category:/).map((t) => t.text), ['2024', '2025']);
  assert.deepEqual(textsOf(scene, /^label:series:/).map((t) => t.text), ['Revenue', 'Headcount', 'Rating']);
  // Decimals per row when the headers are periods; locale from the numbers service.
  assert.deepEqual(textsOf(scene, /^label:value:s0:/).map((t) => t.text), ['1.234,5', '2.000,0']);
  assert.deepEqual(textsOf(scene, /^label:value:s1:/).map((t) => t.text), ['3.120', '3.340']);
  assert.deepEqual(textsOf(scene, /^label:value:s2:/).map((t) => t.text), ['High']);
  const cells = textsOf(scene, /^label:value:s0:/);
  assert.ok(cells[0].x + cells[0].w <= cells[1].x);
  assert.ok(Math.max(...scene.shapes.map((s) => (s.kind === 'line' ? s.x2 : 0))) <= F.x + F.w + 0.01);
  assert.deepEqual(strict(scene), []);
});

test('table: wide content wraps headers, then shrinks, and reports rows that do not fit', () => {
  const cats = Array.from({ length: 9 }, (_, i) => `Very long column heading number ${i + 1}`);
  const rows = (n: number) => Array.from({ length: n }, (_, i) => ({ name: `Row ${i}`, values: cats.map(() => 123456.78) }));
  const wide = layoutChart({ type: 'table', categories: cats, series: rows(3) }, F, m);
  assert.ok(textsOf(wide, /^label:category:/).every((t) => t.text.includes('\n')), 'headers wrap');
  assert.deepEqual(checkScene(wide, { measurer: m }), []);
  const tall = layoutChart({ type: 'table', categories: cats, series: rows(40) }, F, m);
  assert.ok(tall.hidden.some((h) => /shrunk/.test(h)) && tall.hidden.some((h) => /table row does not fit/.test(h)));
  assert.deepEqual(checkScene(tall, { measurer: m }), []);
});

// ---- Gantt ------------------------------------------------------------------------------------------

test('gantt: bars span their dates, milestones are diamonds with date labels', () => {
  const scene = layoutChart(samples.gantt!, F, m);
  const bar = (r: number) => rectsOf(scene, new RegExp(`^bar:s0:c${r}$`))[0];
  const day = (s: string) => Date.parse(s) / 864e5;
  // Pixels per day are the same for every bar.
  const ppd0 = bar(0).w / (day('2026-02-02') - day('2026-01-05'));
  const ppd3 = bar(3).w / (day('2026-11-30') - day('2026-05-04'));
  assert.ok(close(ppd0, ppd3, 1e-9));
  assert.ok(close((bar(2).x - bar(0).x) / ppd0, day('2026-03-02') - day('2026-01-05'), 1e-6));
  const ms = rectsOf(scene, /^marker:s2:/);
  assert.equal(ms.length, 4);
  assert.ok(ms.every((d) => d.geom === 'diamond'));
  assert.ok(close(ms[0].x + ms[0].w / 2, bar(0).x + bar(0).w), 'milestone on the bar end');
  assert.ok(textsOf(scene, /^label:value:s2:c0$/)[0].text === 'Feb 2');
  // One-year range: months under a year tier.
  const scale = textsOf(scene, /^axis:tick:/).map((t) => t.text);
  assert.ok(scale.includes('2026') && scale.includes('Jan') && scale.includes('Nov'), scale.join(' '));
  assert.deepEqual(strict(scene), []);
});

test('gantt: the scale follows the range (weeks, months, quarters)', () => {
  const g = (a: string, b: string): string[] => textsOf(layoutChart({
    type: 'gantt', categories: ['x'], series: [{ name: 'Start', values: [a] }, { name: 'End', values: [b] }],
  }, F, m), /^axis:tick:/).map((t) => t.text);
  const weeks = g('2026-03-02', '2026-04-20');
  assert.ok(weeks.includes('Mar 2026') && weeks.includes('9') && weeks.includes('16'), weeks.join(' '));
  const months = g('2026-01-01', '2026-12-01');
  assert.ok(months.includes('Jun') && months.includes('2026'), months.join(' '));
  const quarters = g('Q1 2026', 'Q3 2030');
  assert.ok(quarters.includes('Q3') && quarters.includes('2028'), quarters.join(' '));
});

test('gantt: validation of dates, bars and shades', () => {
  const base = { type: 'gantt', categories: ['a', 'b'] };
  assert.deepEqual(validateSpec({ ...base, series: [{ name: 'S', values: ['2026-01-01', 'soon'] }, { name: 'E', values: ['2025-12-01', null] }] }).map((e) => e.path),
    ['/series/0/values/1', '/series/1/values/0']);
  assert.deepEqual(validateSpec({ ...base, series: [{ name: 'S', values: ['2026-01-01', null] }] }).map((e) => e.path), ['/series']);
  assert.deepEqual(validateSpec({ ...base, series: [{ name: 'S', values: [null, null] }, { name: 'E', values: [null, null] }], options: { shades: [{ from: '2026-02-01', to: '2026-01-01' }] } }).map((e) => e.path),
    ['/options/shades/0/to']);
  const scene = layoutChart({ ...base, series: [{ name: 'S', values: ['2026-01-05', null] }, { name: 'E', values: ['2026-03-02', null] }, { name: 'M', values: [null, '2026-02-10'] }],
    options: { shades: [{ from: '2026-02-01', to: '2026-02-15', label: 'Freeze' }] } } as ChartSpec, F, m);
  assert.equal(rectsOf(scene, /^shade:0$/).length, 1);
  assert.equal(textsOf(scene, /^shade:0$/)[0].text, 'Freeze');
  assert.deepEqual(strict(scene), []);
});
