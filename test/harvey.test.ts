import { test } from 'node:test';
import assert from 'node:assert/strict';
import type { ChartSpec, PathShape, RectShape, Scene, Shape, TextShape } from '../src/core/index.ts';
import { checkScene, layoutChart } from '../src/core/index.ts';
import { samples } from '../src/core/samples.ts';
import { createMeasurer } from '../src/fonts.ts';
import { CHART_FRAME } from '../src/deck.ts';

const m = createMeasurer();
const ball = (scene: Scene, s: number, c: number): Shape[] => scene.shapes.filter((_, i) => scene.roles![i] === `ball:s${s}:c${c}`);
const sweepOf = (shapes: Shape[]) => {
  const p = shapes.find((x): x is PathShape => x.kind === 'path');
  const a = p?.commands.find((x) => x.op === 'A') as { start: number; sweep: number } | undefined;
  return a;
};

const spec: ChartSpec = {
  type: 'harvey',
  categories: ['Price', 'Quality', 'Speed', 'Reach', 'Fit'],
  series: [
    { name: 'Option A', values: [0, 1, 2, 3, 4] },
    { name: 'Option B', values: [4, null, 2, 1, 0] },
  ],
};

test('harvey: one ball per value, filled in quarters clockwise from 12 o\'clock', () => {
  const scene = layoutChart(spec, CHART_FRAME, m);
  // 0: outline only; 1–3: outline + wedge; 4: a filled disc.
  assert.deepEqual(ball(scene, 0, 0).map((s) => s.kind), ['rect']);
  assert.equal((ball(scene, 0, 0)[0] as RectShape).geom, 'ellipse');
  [1, 2, 3].forEach((q) => {
    const a = sweepOf(ball(scene, 0, q))!;
    assert.equal(a.start, -90);
    assert.equal(a.sweep, q * 90);
  });
  const full = ball(scene, 0, 4);
  assert.equal(full.length, 1);
  assert.equal((full[0] as RectShape).fill, (full[0] as RectShape).stroke);
  // Empty cell: no ball.
  assert.equal(ball(scene, 1, 1).length, 0);
});

test('harvey: grid of rows (series) and columns (categories) with labels', () => {
  const scene = layoutChart(spec, CHART_FRAME, m);
  const centre = (s: number, c: number) => { const r = ball(scene, s, c)[0] as RectShape; return { x: r.x + r.w / 2, y: r.y + r.h / 2 }; };
  assert.equal(centre(0, 0).y, centre(0, 4).y, 'a row shares its y');
  assert.equal(centre(0, 2).x, centre(1, 2).x, 'a column shares its x');
  assert.ok(centre(1, 0).y > centre(0, 0).y);
  const texts = (prefix: string) => scene.shapes.filter((s, i): s is TextShape => s.kind === 'text' && scene.roles![i].startsWith(prefix)).map((t) => t.text);
  assert.deepEqual(texts('label:category'), spec.categories);
  assert.deepEqual(texts('label:series'), ['Option A', 'Option B']);
  // Header above the balls, names left of them.
  const head = scene.shapes[scene.roles!.indexOf('label:category:c2')] as TextShape;
  assert.ok(head.y + head.h <= centre(0, 2).y);
  assert.deepEqual(checkScene(scene, { measurer: m }), []);
});

test('harvey: fractions 0–1 fill proportionally; out-of-range values are clamped and reported', () => {
  const frac = layoutChart({ ...spec, series: [{ name: 'Done', values: [0.25, 0.6, 1, 0, 0.1] }] }, CHART_FRAME, m);
  assert.ok(Math.abs(sweepOf(ball(frac, 0, 1))!.sweep - 216) < 1e-9);
  assert.equal(ball(frac, 0, 2).length, 1, '1 = full');
  const clamped = layoutChart({ ...spec, series: [{ name: 'X', values: [5, -1, 2, 2, 2] }] }, CHART_FRAME, m);
  assert.equal(ball(clamped, 0, 0).length, 1, '5 clamps to a full ball');
  assert.ok(clamped.hidden.some((h) => h.includes('clamped')));
});

test('harvey sample passes the quality gate', () => {
  const scene = layoutChart(samples.harvey!, CHART_FRAME, m);
  assert.deepEqual(checkScene(scene, { measurer: m }), []);
  assert.equal(scene.hidden.length, 0);
});
