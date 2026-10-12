import { test } from 'node:test';
import assert from 'node:assert/strict';
import type { ChartSpec, PathShape, Scene, TextShape } from '../src/core/index.ts';
import { checkScene, layoutChart } from '../src/core/index.ts';
import { flattenPath } from '../src/core/path.ts';
import { samples } from '../src/core/samples.ts';
import { createMeasurer } from '../src/fonts.ts';
import { CHART_FRAME } from '../src/deck.ts';

const m = createMeasurer();
const at = (scene: Scene, role: string) => scene.shapes[scene.roles!.indexOf(role)];
const textsOf = (scene: Scene, prefix: string) =>
  scene.shapes.filter((s, i): s is TextShape => s.kind === 'text' && scene.roles![i].startsWith(prefix));
const arcs = (p: PathShape) => p.commands.filter((c) => c.op === 'A') as Extract<PathShape['commands'][number], { op: 'A' }>[];

const spec: ChartSpec = {
  type: 'pie',
  categories: ['North', 'South', 'East', 'West', 'Islands'],
  series: [{ name: 'Sales', values: [1, 1, 1, 0.5, 0.05] }],
};

test('pie: slices start at 12 o\'clock, run clockwise and their sweeps are proportional', () => {
  const scene = layoutChart(spec, CHART_FRAME, m);
  const total = 3.55;
  let angle = -90;
  spec.series[0].values.forEach((v, c) => {
    const p = at(scene, `wedge:s0:c${c}`) as PathShape;
    assert.equal(p.kind, 'path');
    const [a] = arcs(p);
    assert.ok(Math.abs(a.start - angle) < 1e-9, `slice ${c} starts where the previous ended`);
    assert.ok(a.sweep > 0, 'clockwise');
    assert.ok(Math.abs(a.sweep - ((v as number) / total) * 360) < 1e-9);
    angle += a.sweep;
    assert.equal(p.commands[0].op, 'M'); // from the centre
  });
  assert.ok(Math.abs(angle - 270) < 1e-9, 'full circle');
});

test('pie: shares add up to exactly 100 (largest remainder) and small slices are labelled outside with a leader', () => {
  const scene = layoutChart(spec, CHART_FRAME, m);
  const shares = [0, 1, 2, 3, 4].map((c) => {
    const inside = scene.roles!.includes(`label:value:s0:c${c}`) ? (at(scene, `label:value:s0:c${c}`) as TextShape).text : null;
    const outside = (at(scene, `label:category:c${c}`) as TextShape).text;
    return parseFloat(inside ?? outside.split(' ').pop()!);
  });
  assert.equal(shares.reduce((a, b) => a + b, 0), 100);
  assert.deepEqual(shares, [28, 28, 28, 14, 2]);
  // 2% does not fit in its sliver: it moves outside after the name, with a leader line.
  assert.equal(scene.roles!.includes('label:value:s0:c4'), false);
  assert.equal((at(scene, 'label:category:c4') as TextShape).text, 'Islands 2%');
  assert.ok(scene.roles!.includes('leader:c4'));
  // Names on the right of the pie are left-aligned, those on the left right-aligned.
  assert.equal((at(scene, 'label:category:c0') as TextShape).align, 'left');
  assert.equal((at(scene, 'label:category:c2') as TextShape).align, 'right');
  assert.deepEqual(checkScene(scene, { measurer: m }), []);
});

test('pie: inside labels sit within their slice', () => {
  const scene = layoutChart(spec, CHART_FRAME, m);
  const w = at(scene, 'wedge:s0:c0') as PathShape;
  const [a] = arcs(w);
  const t = at(scene, 'label:value:s0:c0') as TextShape;
  for (const [x, y] of [[t.x, t.y], [t.x + t.w, t.y], [t.x, t.y + t.h], [t.x + t.w, t.y + t.h]]) {
    const d = Math.hypot(x - a.cx, y - a.cy);
    const ang = (Math.atan2(y - a.cy, x - a.cx) * 180) / Math.PI;
    assert.ok(d < a.r, 'inside the rim');
    assert.ok(ang >= a.start && ang <= a.start + a.sweep, 'inside the angle');
  }
});

test('doughnut: ring slices with a hole, total in the middle, legend variant', () => {
  const scene = layoutChart({ ...spec, type: 'doughnut', options: { totals: true, unit: 'm' } }, CHART_FRAME, m);
  const p = at(scene, 'wedge:s0:c0') as PathShape;
  const [outer, inner] = arcs(p);
  assert.ok(inner.r > 0 && inner.r < outer.r);
  assert.equal(inner.sweep, -outer.sweep, 'inner arc runs back');
  assert.equal(flattenPath(p.commands)[0].closed, true);
  assert.equal((at(scene, 'label:total') as TextShape).text, '3.55m');
  assert.deepEqual(checkScene(scene, { measurer: m }), []);

  const legend = layoutChart({ ...spec, type: 'doughnut', options: { seriesLabels: 'legend' } }, CHART_FRAME, m);
  assert.deepEqual(textsOf(legend, 'legend:c').map((t) => t.text), spec.categories);
  assert.equal(textsOf(legend, 'label:category').length, 0);
  assert.deepEqual(checkScene(legend, { measurer: m }), []);
});

test('pie: a single value is a full disc; zero and negative values get no slice', () => {
  const one = layoutChart({ ...spec, series: [{ name: 'S', values: [5, 0, null, -1, 0] }] }, CHART_FRAME, m);
  const p = at(one, 'wedge:s0:c0') as PathShape;
  assert.equal(arcs(p)[0].sweep, 360);
  assert.equal(one.roles!.filter((r) => r.startsWith('wedge:')).length, 1);
  assert.ok(one.hidden.some((h) => h.includes('negative')));
});

test('pie and doughnut samples pass the quality gate', () => {
  for (const t of ['pie', 'doughnut'] as const) {
    const scene = layoutChart(samples[t]!, CHART_FRAME, m);
    assert.deepEqual(checkScene(scene, { measurer: m }), [], t);
    assert.equal(scene.roles!.filter((r) => r.startsWith('wedge:')).length, samples[t]!.categories.length);
  }
});
