import { test } from 'node:test';
import assert from 'node:assert/strict';
import type { ChartSpec, LineShape, Scene, TextShape } from '../src/core/index.ts';
import { checkScene, layoutChart, validateSpec } from '../src/core/index.ts';
import { createMeasurer } from '../src/fonts.ts';
import { CHART_FRAME } from '../src/deck.ts';

const m = createMeasurer();
const lines = (scene: Scene, s: number) =>
  scene.shapes.filter((sh, i): sh is LineShape => sh.kind === 'line' && scene.roles![i] === `line:s${s}`);

const spec: ChartSpec = {
  type: 'line',
  categories: ['Q1', 'Q2', 'Q3', 'Q4', 'Q5'],
  series: [
    { name: 'Price', values: [10, 12, 12, 9, null] },
    { name: 'Floor', values: [6, 6, 7, 7, 8] },
  ],
  options: { step: true, unit: '$' },
};

test('step line: every segment is horizontal or vertical, holding the value until the next category', () => {
  const scene = layoutChart(spec, CHART_FRAME, m);
  const segs = lines(scene, 0);
  for (const l of segs) assert.ok(l.x1 === l.x2 || l.y1 === l.y2, 'ortho segment');
  // 10→12: across at 10, then up; 12→12: across only; 12→9: across, then down; 9→null: broken.
  assert.equal(segs.length, 5);
  assert.equal(segs[0].y1, segs[0].y2);
  assert.equal(segs[1].x1, segs[1].x2);
  assert.equal(segs[0].x2, segs[1].x1);
  assert.ok(segs[1].y2 < segs[1].y1, 'rises to the higher value');
  // Markers stay on the data points: the vertical riser ends on the next marker.
  const marker = scene.shapes[scene.roles!.indexOf('marker:s0:c1')] as { x: number; y: number; w: number; h: number };
  assert.ok(Math.abs(marker.x + marker.w / 2 - segs[1].x2) < 1e-9 && Math.abs(marker.y + marker.h / 2 - segs[1].y2) < 1e-9);
  const values = scene.shapes.filter((s): s is TextShape => s.kind === 'text' && s.role === 'value').map((t) => t.text);
  assert.ok(values.includes('12$') || values.includes('$12'));
  assert.deepEqual(checkScene(scene, { measurer: m }), []);
});

test('plain line is unchanged when step is off', () => {
  const scene = layoutChart({ ...spec, options: { unit: '$' } }, CHART_FRAME, m);
  assert.equal(lines(scene, 0).length, 3);
});

test('step is validated: boolean, line charts only', () => {
  assert.deepEqual(validateSpec({ ...spec, options: { step: 'yes' } }).map((e) => e.path), ['/options/step']);
  assert.deepEqual(validateSpec({ ...spec, type: 'stacked', options: { step: true } }).map((e) => e.path), ['/options/step']);
});
