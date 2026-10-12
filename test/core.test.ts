import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { formatNumber, inferDecimals, niceTicks, roundToTotal, textOn } from '../src/core/numbers.ts';
import { spread } from '../src/core/place.ts';
import { checkScene, layoutChart } from '../src/core/index.ts';
import { createMeasurer } from '../src/fonts.ts';
import { buildDeck, CHART_FRAME, type DeckSpec } from '../src/deck.ts';

const m = createMeasurer();

test('nice ticks use 1/2/5 steps and cover the range', () => {
  assert.deepEqual(niceTicks(0, 216), [0, 50, 100, 150, 200, 250]);
  assert.deepEqual(niceTicks(0, 38), [0, 10, 20, 30, 40]);
  const t = niceTicks(-12, 7);
  assert.ok(t[0] <= -12 && t[t.length - 1] >= 7);
});

test('largest-remainder rounding sums to the rounded total', () => {
  const shares = [33.33, 33.33, 33.34];
  assert.equal(roundToTotal(shares).reduce((a, b) => a + b), 100);
  assert.equal(roundToTotal([12.5, 37.5, 50]).reduce((a, b) => a + b), 100);
  const r = roundToTotal([41.6, 31.9, 18.5, 8.0], 0);
  assert.equal(r.reduce((a, b) => a + b), 100);
});

test('number formatting', () => {
  assert.equal(formatNumber(1234.5, 1, 'm'), '1,234.5m');
  assert.equal(formatNumber(-3, 0), '−3');
  assert.equal(formatNumber(4, 0, '%', true), '+4%');
  assert.equal(inferDecimals([1, 2, 3]), 0);
  assert.equal(inferDecimals([1.5, 2]), 1);
  assert.equal(textOn('1F3A5F'), 'FFFFFF');
  assert.equal(textOn('D9D9D9'), '222222');
});

test('spread keeps order and removes overlaps', () => {
  const ys = spread([100, 101, 102, 300], [12, 12, 12, 12], 0, 400, 0);
  for (let i = 1; i < 3; i++) assert.ok(ys[i] - ys[i - 1] >= 12 - 1e-9);
  assert.equal(ys[3], 300);
});

test('stacked column chart: totals and segments line up', () => {
  const scene = layoutChart({
    type: 'stacked', categories: ['A', 'B'],
    series: [{ name: 'x', values: [10, 20] }, { name: 'y', values: [5, 5] }],
  }, CHART_FRAME, m);
  const totals = scene.shapes.filter((s) => s.kind === 'text' && s.role === 'total').map((s) => (s as { text: string }).text);
  assert.deepEqual(totals, ['15', '25']);
  assert.equal(scene.shapes.filter((s) => s.kind === 'rect').length, 4);
  assert.deepEqual(checkScene(scene), []);
});

test('100% chart labels always sum to 100', () => {
  const scene = layoutChart({
    type: 'percent', categories: ['A'],
    series: [{ name: 'a', values: [1] }, { name: 'b', values: [1] }, { name: 'c', values: [1] }],
  }, CHART_FRAME, m);
  const sum = scene.shapes
    .filter((s) => s.kind === 'text' && s.role === 'value')
    .reduce((a, s) => a + parseFloat((s as { text: string }).text), 0);
  assert.equal(sum, 100);
});

test('every example deck lays out with no overlaps and writes a valid pptx', () => {
  for (const file of ['examples/deck.json']) {
    const deck: DeckSpec = JSON.parse(readFileSync(file, 'utf8'));
    const { pptx, slides } = buildDeck(deck, m);
    assert.equal(String.fromCharCode(pptx[0], pptx[1]), 'PK');
    for (const s of slides) assert.deepEqual(s.issues.filter((i) => !i.startsWith('hidden')), [], `${file}: ${s.title}`);
  }
});
