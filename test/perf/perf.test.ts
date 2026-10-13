// Performance budgets (the add-in re-lays a chart out on every resize): warm layout p95 < 20 ms for
// every sample, at the default frame and resized; a 100-slide deck builds in < 2 s; a big deck (100
// slides, 50 charts) builds, reads back and goes through the add-in engine within budget and without
// holding on to memory. Budgets are ~10x above what a laptop measures (`npm run bench`), so they
// catch regressions, not noise (CI runs on slower Windows machines too).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { layoutChart, type ChartSpec, type Frame } from '../../src/core/index.ts';
import { samples } from '../../src/core/samples.ts';
import { createMeasurer } from '../../src/fonts.ts';
import { buildDeck, type DeckSpec } from '../../src/deck.ts';
import { handleRequest, parseRequest } from '../../src/addin/engine.ts';
import { parseResponse } from '../../src/addin/format.ts';
import { importDeck, readCharts } from '../../src/ooxml/pptx-read.ts';

const m = createMeasurer();
const FRAMES: Frame[] = [{ x: 48, y: 96, w: 864, h: 408 }, { x: 0, y: 0, w: 360, h: 200 }, { x: 0, y: 0, w: 900, h: 120 }, { x: 0, y: 0, w: 240, h: 500 }];

test('perf: warm layout p95 < 20 ms per sample chart, default and resized frames', () => {
  const slow: string[] = [];
  for (const [type, spec] of Object.entries(samples) as [string, ChartSpec][]) {
    for (let i = 0; i < 5; i++) layoutChart(spec, FRAMES[i % FRAMES.length], m);
    // Best of three rounds: `npm test` runs test files in parallel, so one round can catch a busy CPU.
    const round = () => {
      const times = Array.from({ length: 40 }, (_, i) => {
        const t = performance.now();
        layoutChart(spec, FRAMES[i % FRAMES.length], m);
        return performance.now() - t;
      }).sort((a, b) => a - b);
      return times[Math.floor(times.length * 0.95)];
    };
    const p95 = Math.min(round(), round(), round());
    if (p95 >= 20) slow.push(`${type}: p95 ${p95.toFixed(1)} ms`);
  }
  assert.deepEqual(slow, []);
});

test('perf: a 100-slide deck builds in < 2 s', () => {
  const all = Object.values(samples) as ChartSpec[];
  const deck = { slides: Array.from({ length: 100 }, (_, i) => ({ title: `Slide ${i + 1}`, chart: all[i % all.length] })) };
  buildDeck({ slides: deck.slides.slice(0, 3) }, m);
  let ms = Infinity, built = buildDeck({ slides: [] }, m);
  for (let k = 0; k < 2 && ms >= 2000; k++) {
    const t = performance.now();
    built = buildDeck(deck, m);
    ms = performance.now() - t;
  }
  const { pptx, slides } = built;
  assert.equal(slides.length, 100);
  assert.ok(pptx.length > 0);
  assert.ok(ms < 2000, `${ms.toFixed(0)} ms`);
});

/** The D5 big deck: 100 slides, a chart on every other one (50 charts, every sample type in turn). */
function bigDeck(): DeckSpec {
  const all = Object.values(samples) as ChartSpec[];
  return { slides: Array.from({ length: 100 }, (_, i) => (i % 2 ? { title: `Slide ${i + 1}`, chart: all[(i >> 1) % all.length] } : { title: `Section ${i / 2 + 1}` })) };
}

/** Best of `n` runs of `f` in ms (a busy CPU during a parallel test run spoils one run, not all). */
function best<T>(n: number, f: () => T): { ms: number; out: T } {
  let ms = Infinity, out!: T;
  for (let k = 0; k < n; k++) { const t = performance.now(); out = f(); ms = Math.min(ms, performance.now() - t); }
  return { ms, out };
}

test('perf: a 100-slide deck with 50 charts builds in < 2 s and reads back (readCharts, importDeck) in < 1 s', () => {
  const deck = bigDeck();
  buildDeck({ slides: deck.slides.slice(0, 4) }, m);
  const heap = process.memoryUsage().heapUsed;
  const built = best(2, () => buildDeck(deck, m));
  assert.equal(built.out.slides.length, 100);
  assert.ok(built.ms < 2000, `build ${built.ms.toFixed(0)} ms`);
  const read = best(2, () => readCharts(built.out.pptx));
  assert.equal(read.out.length, 100);
  assert.equal(read.out.flatMap((s) => s.charts).length, 50);
  assert.ok(read.out.every((s) => s.charts.every((c) => c.spec && c.roles.length > 0)), 'every chart has its spec and roles');
  assert.ok(read.ms < 1000, `readCharts ${read.ms.toFixed(0)} ms`);
  const imported = best(2, () => importDeck(built.out.pptx));
  assert.deepEqual(imported.out.warnings, []);
  assert.equal(imported.out.deck.slides.length, 50);
  assert.ok(imported.ms < 1000, `importDeck ${imported.ms.toFixed(0)} ms`);
  // Memory: what stays after building and reading is the deck itself, not a copy per slide.
  globalThis.gc?.();
  const kept = (process.memoryUsage().heapUsed - heap) / 1e6;
  assert.ok(kept < 200, `heap grew by ${kept.toFixed(0)} MB`);
});

test('perf: the add-in engine inserts and then updates (native edits kept) 50 charts in < 3 s', async () => {
  const F = '48,96,864,408';
  const types = (Object.keys(samples) as ChartSpec['type'][]);
  const tag = (r: ReturnType<typeof parseResponse>, k: string) => r.tags.find(([n]) => n === k)?.[1] ?? '';
  const round = async () => {
    const t = performance.now();
    for (let i = 0; i < 50; i++) {
      const ins = parseResponse(await handleRequest(parseRequest(`insert\n${types[i % types.length]}\n${F}\n`), m));
      assert.deepEqual(ins.errors, []);
      // Update as the add-in sends it: OC_BASE and the shapes as they are now, one recoloured.
      const base = JSON.parse(tag(ins, 'OC_BASE'));
      const native = Object.entries(base).map(([role, props], k) => ({ role, kind: 'mark', props: k ? props : { ...(props as object), fill: 'FF0000' } }));
      const upd = parseResponse(await handleRequest(parseRequest(`update\n\n${tag(ins, 'OC_FRAME')}\n\nbase=${JSON.stringify(base)}\nnative=${JSON.stringify(native)}\n${tag(ins, 'OC_SPEC')}`), m));
      assert.deepEqual(upd.errors, []);
    }
    return performance.now() - t;
  };
  await round();
  const ms = Math.min(await round(), await round());
  assert.ok(ms < 3000, `${ms.toFixed(0)} ms for 50 inserts + 50 updates`);
});
