// Engine defects found by reviewing the bench gallery (test/bench), each pinned by a test.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { layoutChart, type ChartSpec, type Scene, type TextShape } from '../src/core/index.ts';
import { createMeasurer } from '../src/fonts.ts';
import { CHART_FRAME } from '../src/deck.ts';

const m = createMeasurer();
const lay = (spec: ChartSpec) => layoutChart(spec, CHART_FRAME, m);
const textsOf = (s: Scene, re: RegExp) => s.shapes.filter((x, i): x is TextShape => x.kind === 'text' && re.test(s.roles![i]));

test('line, floating axis: the lowest point keeps air above the axis (no marker on the category row)', () => {
  const spec: ChartSpec = { type: 'line', categories: ['Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6'], series: [{ name: 'NRR', values: [112, 118, 104, 109, 111, 113] }], options: { unit: '%' } };
  const sc = lay(spec);
  const ticks = textsOf(sc, /^axis:tick:\d+$/).map((t) => Number(t.text.replace('%', '')));
  assert.ok(Math.min(...ticks) < 104, `axis starts at ${Math.min(...ticks)}`);
  assert.equal(textsOf(sc, /^label:category:/).length, 6, 'every category label shown');
  assert.deepEqual(sc.hidden, []);
});

test('clustered: value labels share one size, shrunk together to fit their bars', () => {
  const cats = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'];
  const spec: ChartSpec = { type: 'clustered', categories: cats, series: [{ name: 'x', values: cats.map((_, i) => 100 + i * 37) }, { name: 'y', values: cats.map((_, i) => 90 + i * 41) }] };
  const sc = lay(spec);
  const sizes = new Set(textsOf(sc, /^label:value:/).map((t) => t.size));
  assert.equal(sizes.size, 1);
  assert.equal(textsOf(sc, /^label:value:/).length, 24, 'all or nothing');
});

test('clustered: labels too wide for the bars at the minimum size give way to a value axis, never a partial row', () => {
  const cats = Array.from({ length: 16 }, (_, i) => `M${i + 1}`);
  const spec: ChartSpec = { type: 'clustered', categories: cats, series: [{ name: '2024', values: cats.map((_, i) => 10 + i) }, { name: '2025', values: cats.map((_, i) => 12 + i) }], options: { unit: '%' } };
  const sc = lay(spec);
  assert.equal(textsOf(sc, /^label:value:/).length, 0);
  assert.ok(textsOf(sc, /^axis:tick:/).length >= 3, 'value axis shown');
  assert.ok(sc.hidden.some((h) => h.includes('value axis shows the values')));
});

test('text: CJK characters missing from a Latin font are measured one em wide (PowerPoint falls back to an East Asian font)', () => {
  assert.equal(m.width('中国', 12), 24);
  assert.equal(m.width('サービス', 10), 40);
  assert.equal(m.width('한국', 12), 24);
});
