// Performance budgets (the add-in re-lays a chart out on every resize): warm layout p95 < 20 ms for
// every sample, at the default frame and resized; a 100-slide deck builds in < 2 s. Budgets are ~10x
// above what a laptop measures (`npm run bench`), so they catch regressions, not noise.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { layoutChart, type ChartSpec, type Frame } from '../src/core/index.ts';
import { samples } from '../src/core/samples.ts';
import { createMeasurer } from '../src/fonts.ts';
import { buildDeck } from '../src/deck.ts';

const m = createMeasurer();
const FRAMES: Frame[] = [{ x: 48, y: 96, w: 864, h: 408 }, { x: 0, y: 0, w: 360, h: 200 }, { x: 0, y: 0, w: 900, h: 120 }, { x: 0, y: 0, w: 240, h: 500 }];

test('perf: warm layout p95 < 20 ms per sample chart, default and resized frames', () => {
  const slow: string[] = [];
  for (const [type, spec] of Object.entries(samples) as [string, ChartSpec][]) {
    for (let i = 0; i < 5; i++) layoutChart(spec, FRAMES[i % FRAMES.length], m);
    const times = Array.from({ length: 40 }, (_, i) => {
      const t = performance.now();
      layoutChart(spec, FRAMES[i % FRAMES.length], m);
      return performance.now() - t;
    }).sort((a, b) => a - b);
    const p95 = times[Math.floor(times.length * 0.95)];
    if (p95 >= 20) slow.push(`${type}: p95 ${p95.toFixed(1)} ms`);
  }
  assert.deepEqual(slow, []);
});

test('perf: a 100-slide deck builds in < 2 s', () => {
  const all = Object.values(samples) as ChartSpec[];
  const deck = { slides: Array.from({ length: 100 }, (_, i) => ({ title: `Slide ${i + 1}`, chart: all[i % all.length] })) };
  buildDeck({ slides: deck.slides.slice(0, 3) }, m);
  const t = performance.now();
  const { pptx, slides } = buildDeck(deck, m);
  const ms = performance.now() - t;
  assert.equal(slides.length, 100);
  assert.ok(pptx.length > 0);
  assert.ok(ms < 2000, `${ms.toFixed(0)} ms`);
});
