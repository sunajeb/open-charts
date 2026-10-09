import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import type { ChartSpec, RectShape, TextShape } from '../src/core/index.ts';
import { checkScene, createMeasurer, layoutChart } from '../src/core/index.ts';
import { buildDeck, CHART_FRAME, type DeckSpec } from '../src/deck.ts';

const m = createMeasurer();
const texts = (scene: ReturnType<typeof layoutChart>, role: TextShape['role']) =>
  scene.shapes.filter((s): s is TextShape => s.kind === 'text' && s.role === role);
const rects = (scene: ReturnType<typeof layoutChart>) => scene.shapes.filter((s): s is RectShape => s.kind === 'rect');

const market: ChartSpec = {
  type: 'mekko',
  categories: ['Enterprise', 'Mid-market', 'SMB'],
  series: [
    { name: 'Us', values: [300, 100, 50] },
    { name: 'Them', values: [200, 100, 150] },
    { name: 'Others', values: [100, 50, 100] },
  ],
  options: { unit: 'm' },
};

test('mekko column widths are proportional to category totals and fill the plot', () => {
  const scene = layoutChart(market, CHART_FRAME, m);
  const totals = [600, 250, 300];
  const grand = 1150;
  const cols = totals.map((_, c) => rects(scene).filter((r) => Math.abs(r.x - rects(scene)[c * 3].x) < 1e-6));
  const widths = cols.map((col) => col[0].w);
  const pw = widths.reduce((a, w) => a + w, 0);
  totals.forEach((t, c) => assert.ok(Math.abs(widths[c] / pw - t / grand) < 1e-9, `column ${c}`));
  // Every segment in a column has the column's width; columns touch.
  cols.forEach((col) => col.forEach((r) => assert.equal(r.w, col[0].w)));
  for (let c = 1; c < cols.length; c++) assert.ok(Math.abs(cols[c][0].x - (cols[c - 1][0].x + widths[c - 1])) < 1e-9);
  assert.deepEqual(checkScene(scene), []);
});

test('percent mekko: every column tops out at 100% and its labels sum to 100', () => {
  const scene = layoutChart(market, CHART_FRAME, m);
  const rs = rects(scene);
  const top = Math.min(...rs.map((r) => r.y));
  const bottom = Math.max(...rs.map((r) => r.y + r.h));
  for (const x of new Set(rs.map((r) => r.x))) {
    const col = rs.filter((r) => r.x === x);
    assert.ok(Math.abs(Math.min(...col.map((r) => r.y)) - top) < 1e-6);
    assert.ok(Math.abs(Math.max(...col.map((r) => r.y + r.h)) - bottom) < 1e-6);
    const sum = texts(scene, 'value')
      .filter((t) => t.x >= x && t.x + t.w <= x + col[0].w + 1e-6)
      .reduce((a, t) => a + parseFloat(t.text), 0);
    assert.equal(sum, 100);
  }
  assert.deepEqual(texts(scene, 'total').map((t) => t.text), ['600m', '250m', '300m']);
  // Column shares of the grand total are rounded together and read 100% overall.
  const shares = texts(scene, 'decoration').map((t) => parseFloat(t.text));
  assert.equal(shares.reduce((a, b) => a + b, 0), 100);
  assert.deepEqual(texts(scene, 'series').map((t) => t.text), ['Others', 'Them', 'Us']);
});

test('unit mekko: segment heights are absolute and labelled with values', () => {
  const scene = layoutChart({ ...market, options: { mekko: 'unit', unit: 'm', valueAxis: true } }, CHART_FRAME, m);
  const rs = rects(scene);
  const byX = [...new Set(rs.map((r) => r.x))].sort((a, b) => a - b).map((x) => rs.filter((r) => r.x === x));
  const heights = byX.map((col) => col.reduce((a, r) => a + r.h, 0));
  assert.ok(Math.abs(heights[0] / heights[1] - 600 / 250) < 1e-6);
  assert.ok(Math.abs(heights[2] / heights[1] - 300 / 250) < 1e-6);
  assert.ok(texts(scene, 'value').some((t) => t.text === '300m'));
  assert.ok(texts(scene, 'axis').length >= 3);
  assert.deepEqual(checkScene(scene), []);
});

test('narrow columns hide labels that cannot fit instead of overlapping', () => {
  const scene = layoutChart({
    type: 'mekko',
    categories: ['A very long category name indeed', 'Another rather long category name', 'Tiny'],
    series: [{ name: 'x', values: [500, 500, 2] }, { name: 'y', values: [300, 300, 1] }],
  }, CHART_FRAME, m);
  assert.deepEqual(checkScene(scene), []);
  assert.ok(scene.hidden.some((h) => h.includes('Tiny')));
});

test('negative and zero data are tolerated', () => {
  const scene = layoutChart({
    type: 'mekko', categories: ['A', 'B'],
    series: [{ name: 'x', values: [10, 0] }, { name: 'y', values: [-5, 0] }],
  }, CHART_FRAME, m);
  assert.equal(rects(scene).length, 1);
  assert.ok(scene.hidden.some((h) => h.includes('negative')));
  assert.ok(scene.hidden.some((h) => h.includes('zero total')));
  assert.deepEqual(checkScene(scene), []);
});

test('mekko example deck lays out with no overlaps or out-of-frame labels', () => {
  const deck: DeckSpec = JSON.parse(readFileSync('examples/mekko.json', 'utf8'));
  const { pptx, slides } = buildDeck(deck, m);
  assert.equal(String.fromCharCode(pptx[0], pptx[1]), 'PK');
  assert.equal(slides.length, 3);
  for (const s of slides) assert.deepEqual(s.issues.filter((i) => !i.startsWith('hidden')), [], s.title);
});
