// Chart decorations (reference manual: Chart annotations, Axes): difference arrow modes, level
// difference arrows, value lines, the 100% indicator and the 100%= row, category gaps and breaks,
// value axis breaks and the combination chart's secondary axis.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { checkScene, layoutChart, validateSpec, type ChartSpec, type LineShape, type RectShape, type Scene, type TextShape } from '../src/core/index.ts';
import { createMeasurer } from '../src/fonts.ts';

const m = createMeasurer();
const F = { x: 48, y: 96, w: 864, h: 408 };
const lay = (spec: ChartSpec, f = F) => layoutChart(spec, f, m);
const byRole = (s: Scene, re: RegExp) => s.shapes.filter((_, i) => re.test(s.roles![i]));
const texts = (s: Scene, re: RegExp) => (byRole(s, re).filter((x) => x.kind === 'text') as TextShape[]).map((t) => t.text);
const clean = (s: Scene) => assert.deepEqual(checkScene(s, { measurer: m }), []);

const stacked = (options: ChartSpec['options'] = {}): ChartSpec => ({
  type: 'stacked', categories: ['2021', '2022', '2023', '2024'],
  series: [{ name: 'Base', values: [12, 18, 20, 22] }, { name: 'Upsell', values: [5, 2, 4, 6] }, { name: 'New', values: [4, 5, 6, 7] }],
  options,
});

test('total difference arrow: three modes with their labels and heads', () => {
  const heads = (s: Scene) => (byRole(s, /^arrow:diff$/).filter((x) => x.kind === 'line') as LineShape[]).filter((l) => l.arrowEnd).length;
  const abs = lay(stacked({ difference: { from: 0, to: 3, mode: 'absolute' } }));
  assert.deepEqual(texts(abs, /^arrow:diff$/), ['+14']);
  assert.equal(heads(abs), 2, 'absolute is double-ended');
  const rel = lay(stacked({ difference: { from: 0, to: 3 } }));
  assert.deepEqual(texts(rel, /^arrow:diff$/), ['+67%'], 'relative is the default');
  assert.equal(heads(rel), 1);
  const rev = lay(stacked({ difference: { from: 0, to: 3, mode: 'reverse' } }));
  assert.deepEqual(texts(rev, /^arrow:diff$/), ['−40%']);
  for (const s of [abs, rel, rev]) clean(s);
});

test('level arrows: between totals (relative by default) and between segment levels', () => {
  const s = lay(stacked({ levelArrows: [
    { from: { category: 0 }, to: { category: 1 } },
    { from: { category: 2, series: 0 }, to: { category: 3, series: 0 }, mode: 'absolute' },
  ] }));
  assert.deepEqual(texts(s, /^arrow:level:0$/), ['+19%']); // 21 → 25
  assert.deepEqual(texts(s, /^arrow:level:1$/), ['+2']); // Base 20 → 22
  const shaft = (byRole(s, /^arrow:level:0$/) as LineShape[]).find((l) => l.kind === 'line' && l.arrowEnd)!;
  const col1 = byRole(s, /^bar:s2:c1$/)[0] as RectShape;
  assert.ok(shaft.x1 > col1.x + col1.w, 'beside the later column');
  assert.ok(byRole(s, /^arrow:level:0$/).some((x) => x.kind === 'line' && x.dash), 'dotted guide from the earlier column');
  clean(s);
});

test('level arrows: "between" sits in the gap; line charts connect points of two series', () => {
  const s = lay(stacked({ categoryGaps: [0], levelArrows: [{ from: { category: 0 }, to: { category: 1 }, position: 'between' }] }));
  const shaft = (byRole(s, /^arrow:level:0$/) as LineShape[]).find((l) => l.kind === 'line' && l.arrowEnd)!;
  const c0 = byRole(s, /^bar:s0:c0$/)[0] as RectShape, c1 = byRole(s, /^bar:s0:c1$/)[0] as RectShape;
  assert.ok(shaft.x1 > c0.x + c0.w && shaft.x1 < c1.x);
  clean(s);
  const line = lay({ type: 'line', categories: ['a', 'b', 'c'], series: [{ name: 'Us', values: [40, 50, 58] }, { name: 'Market', values: [30, 33, 36] }],
    options: { levelArrows: [{ from: { category: 2, series: 1 }, to: { category: 2, series: 0 }, mode: 'absolute' }] } });
  assert.deepEqual(texts(line, /^arrow:level:0$/), ['+22']);
  clean(line);
});

test('level arrows: validation needs a series on clustered and line charts and a real value line', () => {
  const errs = validateSpec({ type: 'clustered', categories: ['a', 'b'], series: [{ name: 'x', values: [1, 2] }],
    options: { levelArrows: [{ from: { category: 0 }, to: { valueLine: 0 } }] } });
  assert.deepEqual(errs.map((e) => e.path).sort(), ['/options/levelArrows/0/from/series', '/options/levelArrows/0/to/valueLine']);
  assert.match(validateSpec(stacked({ levelArrows: [{ from: { category: 1 }, to: { category: 1 } }] }))[0].message, /different/);
});

test('value lines: mean of the column totals with "Ø", fixed values widen the axis, custom labels', () => {
  const s = lay(stacked({ valueLines: [{ value: 'mean' }, { value: 40, label: 'Target {value}', position: 'left' }] }));
  assert.deepEqual(texts(s, /^valueline:0$/), ['Ø 28']); // (21 + 25 + 30 + 35) / 4 = 27.75
  assert.deepEqual(texts(s, /^valueline:1$/), ['Target 40']);
  const tops = (byRole(s, /^bar:/) as RectShape[]).map((b) => b.y);
  const target = (byRole(s, /^valueline:1$/) as LineShape[]).find((l) => l.kind === 'line')!;
  assert.ok(target.y1 < Math.min(...tops), 'a target above the data is still drawn');
  assert.ok(byRole(s, /^valueline:0$/).every((x) => x.kind !== 'line' || x.dash), 'dashed');
  clean(s);
  const clustered = lay({ type: 'clustered', categories: ['a', 'b'], series: [{ name: 'x', values: [2, 4] }, { name: 'y', values: [6, 8] }], options: { valueLines: [{ value: 'mean' }] } });
  assert.deepEqual(texts(clustered, /^valueline:0$/), ['Ø 5'], 'a clustered chart averages its bars');
  assert.match(validateSpec({ type: 'line', categories: ['a'], series: [{ name: 'x', values: [1] }], options: { valueLines: [{ value: 'mean' }] } })[0].message, /give a number/);
});

test('value lines step around labels they would cross', () => {
  // The mean (≈ 27.75) runs through the 2024 column, whose Upsell label sits there.
  const s = lay(stacked({ valueLines: [{ value: 26 }] }));
  const pieces = (byRole(s, /^valueline:0$/) as LineShape[]).filter((l) => l.kind === 'line');
  assert.ok(pieces.length >= 1);
  clean(s);
});

test('100% indicator: on by default for 100% charts, left of the plot when names are on the right; none removes it', () => {
  const pct = (options: ChartSpec['options'] = {}): ChartSpec => ({ ...stacked(), type: 'percent', options });
  const def = lay(pct());
  assert.deepEqual(texts(def, /^indicator$/), ['100%']);
  const t = byRole(def, /^indicator$/).find((x) => x.kind === 'text') as TextShape;
  const bars = byRole(def, /^bar:/) as RectShape[];
  assert.ok(t.x + t.w <= Math.min(...bars.map((b) => b.x)), 'left (series names take the right)');
  const right = lay(pct({ seriesLabels: 'legend' }));
  const tr = byRole(right, /^indicator$/).find((x) => x.kind === 'text') as TextShape;
  assert.ok(tr.x >= Math.max(...bars.map((b) => b.x + b.w)) - 1, 'right by default without names');
  assert.deepEqual(texts(lay(pct({ percentIndicator: 'none' })), /^indicator$/), []);
  assert.deepEqual(texts(lay(pct({ valueAxis: true })), /^indicator$/), [], 'the axis reads 100% already');
  for (const s of [def, right]) clean(s);
});

test('100%= row: shares of the base; percentages with a base show absolute totals', () => {
  const abs = lay({ type: 'percent', categories: ['a', 'b'], series: [{ name: 'x', values: [20, 30] }, { name: 'y', values: [20, 10] }],
    options: { percentBase: [50, null], seriesLabels: 'none' } });
  assert.deepEqual(texts(abs, /^label:value:s\d+:c0$/), ['40%', '40%']); // of 50, not of 40
  assert.deepEqual(texts(abs, /^label:value:s\d+:c1$/), ['75%', '25%']);
  const pd = lay({ type: 'percent', categories: ['a', 'b'], series: [{ name: 'x', values: [60, 50] }, { name: 'y', values: [40, 50] }],
    options: { percentData: true, percentBase: [1200, 1500], totals: true } });
  assert.deepEqual(texts(pd, /^label:total:c\d+$/), ['1,200', '1,500']);
  assert.match(validateSpec({ ...stacked(), type: 'percent', options: { percentBase: [1, 2] } })[0].message, /2 values but there are 4/);
  clean(abs); clean(pd);
});

test('category gaps widen the space after a category; category breaks draw a symbol on the baseline', () => {
  const s = lay(stacked({ categoryGaps: [1], categoryBreaks: [2] }));
  const x = (c: number) => (byRole(s, new RegExp(`^bar:s0:c${c}$`))[0] as RectShape).x;
  const d01 = x(1) - x(0), d12 = x(2) - x(1), d23 = x(3) - x(2);
  assert.ok(Math.abs(d12 - 2 * d01) < 0.01, 'one extra slot after category 1');
  assert.ok(Math.abs(d23 - d01) < 0.01);
  assert.equal(byRole(s, /^break:cat:2$/).length, 3);
  clean(s);
  assert.match(validateSpec(stacked({ categoryGaps: [3] }))[0].message, /nothing to separate/);
  const wf = lay({ type: 'waterfall', categories: ['a', 'b', 'c'], series: [{ name: 'x', values: [10, 2, 'e'] }], options: { categoryGaps: [1] } });
  clean(wf);
});

test('value axis break: compressed range, symbols on spanning bars and the axis, labels clear of it', () => {
  const spec: ChartSpec = { type: 'stacked', categories: ['2013', '2014', '2015'],
    series: [{ name: 'A', values: [2, 6, 12] }, { name: 'B', values: [5, 7, 2] }, { name: 'C', values: [4, 378, 351] }],
    options: { axisBreak: { from: 40, to: 330 }, valueAxis: true } };
  const s = lay(spec);
  assert.ok(byRole(s, /^break:s2:c1$/).length && byRole(s, /^break:s2:c2$/).length && byRole(s, /^break:axis$/).length);
  assert.equal(byRole(s, /^break:s\d+:c0$/).length, 0, 'the short column is not broken');
  const small = byRole(s, /^bar:s0:c0$/)[0] as RectShape;
  const plain = byRole(lay({ ...spec, options: { valueAxis: true } }), /^bar:s0:c0$/)[0] as RectShape;
  assert.ok(small.h > plain.h * 3, 'small segments become readable');
  clean(s);
  const crossing = lay({ ...spec, options: { axisBreak: { from: 5, to: 330 } } });
  assert.ok(crossing.hidden.some((h) => /segment ends inside it/.test(h)));
  const wiggle = lay({ type: 'clustered', categories: ['a', 'b'], series: [{ name: 'x', values: [15, 3] }, { name: 'y', values: [14, 2] }],
    options: { axisBreak: { from: 5, to: 12, style: 'wiggle' }, direction: 'right' } });
  assert.equal(byRole(wiggle, /^break:s\d+:c0$/).length, 6);
  clean(wiggle);
});

test('combo secondary axis: lines on their own scale, right-hand axis, legend', () => {
  const s = lay({ type: 'combo', categories: ['a', 'b', 'c'], series: [{ name: 'Revenue', values: [120, 150, 180] }, { name: 'Margin', values: [8.5, 10.2, 12.1], as: 'line' }],
    options: { secondaryAxis: [1], unit: 'm' } });
  assert.ok(texts(s, /^axis:secondary:tick:\d+$/).length >= 3);
  assert.deepEqual(texts(s, /^label:value:s1:c\d+$/), ['8.5', '10.2', '12.1'], 'no primary unit on the secondary values');
  const bar = byRole(s, /^bar:s0:c0$/)[0] as RectShape;
  const mk = byRole(s, /^marker:s1:c0$/)[0] as RectShape;
  assert.ok(mk.y < bar.y, 'margin 8.5 of 15 sits higher than revenue 120 of 200');
  assert.ok(byRole(s, /^legend:s1$/).length > 0);
  clean(s);
  assert.match(validateSpec({ type: 'combo', categories: ['a'], series: [{ name: 'x', values: [1] }], options: { secondaryAxis: [0] } })[0].message, /only a line/);
});

test('decorations degrade at small sizes: hidden and reported, never colliding', () => {
  const spec = stacked({ levelArrows: [{ from: { category: 0 }, to: { category: 3 } }], valueLines: [{ value: 'mean' }], categoryGaps: [1] });
  for (const f of [{ ...F, w: 360, h: 200 }, { ...F, w: 240, h: 300 }]) {
    const s = lay(spec, f);
    assert.deepEqual(checkScene(s, { measurer: m }), [], `${f.w}×${f.h}`);
  }
});
