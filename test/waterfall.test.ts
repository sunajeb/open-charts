import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import type { ChartSpec, LineShape, RectShape, TextShape } from '../src/core/index.ts';
import { checkScene, layoutChart } from '../src/core/index.ts';
import { createMeasurer } from '../src/fonts.ts';
import { buildDeck, CHART_FRAME, type DeckSpec } from '../src/deck.ts';

const m = createMeasurer();
const rects = (s: ReturnType<typeof layoutChart>) => s.shapes.filter((x): x is RectShape => x.kind === 'rect');
const lines = (s: ReturnType<typeof layoutChart>) => s.shapes.filter((x): x is LineShape => x.kind === 'line');
const texts = (s: ReturnType<typeof layoutChart>, role: TextShape['role']) =>
  s.shapes.filter((x): x is TextShape => x.kind === 'text' && x.role === role);

const bridge: ChartSpec = {
  type: 'waterfall',
  categories: ['Start', 'Up', 'Down', 'Sub', 'Up2', 'End'],
  series: [{ name: 'v', values: [100, 20, -30, null, 10, null] }],
  options: { totalsAt: [0, 3, 5] },
};

test('waterfall: bars float from the running sum, totals sit on the baseline', () => {
  const scene = layoutChart(bridge, CHART_FRAME, m);
  const bars = rects(scene);
  assert.equal(bars.length, 6);
  const [start, up, down, sub, up2, end] = bars;
  const baseline = lines(scene).find((l) => l.width === 0.75 && l.x1 < start.x)!;
  const y0 = baseline.y1;
  const tol = 0.01;
  // Totals are anchored at the baseline.
  for (const t of [start, sub, end]) assert.ok(Math.abs(t.y + t.h - y0) < tol, 'total bottom on baseline');
  // Increases start where the previous column ended; decreases hang from it.
  assert.ok(Math.abs(up.y + up.h - start.y) < tol);
  assert.ok(Math.abs(down.y - up.y) < tol);
  assert.ok(Math.abs(sub.y - (down.y + down.h)) < tol);
  assert.ok(Math.abs(up2.y + up2.h - sub.y) < tol);
  assert.ok(Math.abs(end.y - up2.y) < tol);
  // Heights are proportional to values: start 100, sub 90, end 100.
  assert.ok(Math.abs(sub.h / start.h - 0.9) < 1e-6);
  assert.ok(Math.abs(end.h - start.h) < tol);
  // Totals use palette[0]; increases / decreases use the semantic colours.
  assert.equal(start.fill, '1F3A5F');
  assert.equal(up.fill, '3E8E5E');
  assert.equal(down.fill, 'C0504D');
  assert.deepEqual(checkScene(scene), []);
  assert.deepEqual(scene.hidden, []);
});

test('waterfall: one connector between each pair of neighbours, at the exit level', () => {
  const scene = layoutChart(bridge, CHART_FRAME, m);
  const bars = rects(scene);
  const connectors = lines(scene).filter((l) => l.width === 0.5);
  assert.equal(connectors.length, bridge.categories.length - 1);
  connectors.forEach((l, i) => {
    assert.equal(l.y1, l.y2);
    assert.ok(Math.abs(l.x1 - (bars[i].x + bars[i].w)) < 0.01 && Math.abs(l.x2 - bars[i + 1].x) < 0.01, 'spans the gap');
    const exit = i === 2 ? bars[i].y + bars[i].h : bars[i].y; // the decrease exits at its bottom
    assert.ok(Math.abs(l.y1 - exit) < 0.01, `connector ${i} at exit level`);
  });
});

test('waterfall: labels are signed deltas and bold totals from the running sum', () => {
  const scene = layoutChart(bridge, CHART_FRAME, m);
  assert.deepEqual(texts(scene, 'total').map((t) => t.text), ['100', '90', '100']);
  assert.ok(texts(scene, 'total').every((t) => t.bold));
  assert.deepEqual(texts(scene, 'value').map((t) => t.text), ['+20', '−30', '+10']);
});

test('waterfall: a bridge crossing below zero keeps the baseline and reserves room below', () => {
  const scene = layoutChart({
    type: 'waterfall', categories: ['Start', 'Loss', 'Sub', 'Gain', 'End'],
    series: [{ name: 'v', values: [10, -25, null, 30, null] }],
    options: { totalsAt: [0, 2, 4], valueAxis: true },
  }, CHART_FRAME, m);
  const [start, loss, sub] = rects(scene);
  const y0 = start.y + start.h;
  assert.ok(loss.y < y0 && loss.y + loss.h > y0, 'decrease crosses the baseline');
  assert.ok(Math.abs(sub.y - y0) < 0.01 && sub.h > 0, 'negative subtotal hangs below the baseline');
  const subLabel = texts(scene, 'total').find((t) => t.text === '−15')!;
  assert.ok(subLabel.y >= sub.y + sub.h, 'negative total labelled below its bar');
  assert.ok(texts(scene, 'axis').length >= 3);
  assert.deepEqual(checkScene(scene), []);
  assert.deepEqual(scene.hidden, []);
});

test('waterfall: multi-series stacks positives up and negatives down from the entry level', () => {
  const scene = layoutChart({
    type: 'waterfall', categories: ['Start', 'Mixed', 'End'],
    series: [{ name: 'a', values: [50, 10, null] }, { name: 'b', values: [0, -4, null] }],
  }, CHART_FRAME, m);
  const [start, a, b, end] = rects(scene).filter((r) => r.w > 20); // skip legend swatches
  assert.ok(Math.abs(a.y + a.h - start.y) < 0.01, 'positive segment rises from entry');
  assert.ok(Math.abs(b.y - start.y) < 0.01, 'negative segment hangs from entry');
  assert.ok(Math.abs(end.h / start.h - 56 / 50) < 1e-6, 'end total = entry + net');
  assert.ok(texts(scene, 'series').length === 2, 'legend for multi-series');
  assert.deepEqual(checkScene(scene), []);
});

test('waterfall example deck lays out with no issues', () => {
  const deck: DeckSpec = JSON.parse(readFileSync('examples/waterfall.json', 'utf8'));
  const { pptx, slides } = buildDeck(deck, m);
  assert.equal(String.fromCharCode(pptx[0], pptx[1]), 'PK');
  assert.equal(slides.length, 3);
  // The stress slide's long driver names need four lines at 12 pt; the engine shrinks the whole
  // category row one step instead of letting the last line overflow (B8), and says so.
  const expected = [[], [], ['hidden: category labels shrunk to 11 pt to fit']];
  slides.forEach((s, i) => assert.deepEqual(s.issues, expected[i], s.title));
});

test('waterfall: a zero delta is labelled "0" and the chain continues', () => {
  const scene = layoutChart({
    type: 'waterfall', categories: ['Start', 'Flat', 'End'],
    series: [{ name: 'v', values: [10, 0, null] }],
  }, CHART_FRAME, m);
  assert.deepEqual(texts(scene, 'value').map((t) => t.text), ['0']);
  assert.deepEqual(texts(scene, 'total').map((t) => t.text), ['10', '10']);
  assert.deepEqual(checkScene(scene), []);
});
