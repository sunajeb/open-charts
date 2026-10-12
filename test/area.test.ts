import { test } from 'node:test';
import assert from 'node:assert/strict';
import type { ChartSpec, PathShape, Scene, TextShape } from '../src/core/index.ts';
import { checkScene, layoutChart } from '../src/core/index.ts';
import { flattenPath } from '../src/core/path.ts';
import { samples } from '../src/core/samples.ts';
import { createMeasurer } from '../src/fonts.ts';
import { CHART_FRAME } from '../src/deck.ts';

const m = createMeasurer();
const band = (scene: Scene, s: number) => scene.shapes[scene.roles!.indexOf(`area:s${s}`)] as PathShape;
const points = (p: PathShape) => flattenPath(p.commands)[0].points;
const labels = (scene: Scene, prefix: string) =>
  scene.shapes.map((sh, i) => ({ sh, role: scene.roles![i] })).filter((x) => x.sh.kind === 'text' && x.role.startsWith(prefix)).map((x) => x.sh as TextShape);

const spec: ChartSpec = {
  type: 'area',
  categories: ['2021', '2022', '2023', '2024'],
  series: [
    { name: 'A', values: [40, 50, 60, 70] },
    { name: 'B', values: [30, null, 34, 36] },
    { name: 'C', values: [20, 22, 24, 26] },
  ],
  options: { unit: 'm', totals: true },
};

test('stacked area: bands are closed paths stacked on the cumulative sums', () => {
  const scene = layoutChart(spec, CHART_FRAME, m);
  const [a, b, c] = [0, 1, 2].map((s) => points(band(scene, s)));
  const n = 4;
  assert.equal(a.length, 2 * n);
  // Upper edge of one band is the lower edge of the next, point for point.
  for (let k = 0; k < n; k++) {
    assert.ok(Math.abs(a[k].y - b[2 * n - 1 - k].y) < 1e-9);
    assert.ok(Math.abs(b[k].y - c[2 * n - 1 - k].y) < 1e-9);
  }
  // Heights are proportional: A at 2021 is 40 units, the whole stack 90.
  const base = a[2 * n - 1].y;
  assert.ok(Math.abs((base - a[0].y) / (base - c[0].y) - 40 / 90) < 1e-9);
  // Points run left to right across the plot.
  assert.ok(a[0].x < a[1].x && a[1].x < a[2].x);
  assert.equal(flattenPath(band(scene, 0).commands)[0].closed, true);
});

test('stacked area: missing values are interpolated and carry no label', () => {
  const scene = layoutChart(spec, CHART_FRAME, m);
  const b = points(band(scene, 1)), a = points(band(scene, 0));
  const base = a[7].y;
  const unit = (base - a[0].y) / 40;
  // B at 2022 is interpolated to 32 between 30 and 34.
  assert.ok(Math.abs((a[1].y - b[1].y) / unit - 32) < 1e-6);
  assert.equal(scene.roles!.includes('label:value:s1:c1'), false);
  assert.ok(scene.roles!.includes('label:value:s1:c2'));
  assert.deepEqual(labels(scene, 'label:total').map((t) => t.text), ['90m', '104m', '118m', '132m']);
  assert.deepEqual(labels(scene, 'label:series').map((t) => t.text).sort(), ['A', 'B', 'C']);
  assert.deepEqual(checkScene(scene, { measurer: m }), []);
});

test('100% area: every category reaches 100% and its labels add up to exactly 100', () => {
  const scene = layoutChart({
    type: 'area100',
    categories: ['a', 'b', 'c'],
    series: [{ name: 'X', values: [1, 1, 2] }, { name: 'Y', values: [1, 1, 2] }, { name: 'Z', values: [1, 1, 2] }],
  }, CHART_FRAME, m);
  const top = points(band(scene, 2));
  assert.ok(top.slice(0, 3).every((p) => Math.abs(p.y - top[0].y) < 1e-9), 'flat top at 100%');
  for (let c = 0; c < 3; c++) {
    const sum = [0, 1, 2].reduce((acc, s) => acc + parseFloat((scene.shapes[scene.roles!.indexOf(`label:value:s${s}:c${c}`)] as TextShape).text), 0);
    assert.equal(sum, 100);
  }
  assert.deepEqual(checkScene(scene, { measurer: m }), []);
});

test('area samples and the legend variant pass the quality gate', () => {
  for (const t of ['area', 'area100'] as const) {
    const scene = layoutChart(samples[t]!, CHART_FRAME, m);
    assert.deepEqual(checkScene(scene, { measurer: m }), [], t);
    assert.ok(scene.roles!.some((r) => r.startsWith('area:s')));
  }
  const scene = layoutChart({ ...spec, options: { seriesLabels: 'legend', valueAxis: true } }, CHART_FRAME, m);
  assert.equal(scene.roles!.filter((r) => r.startsWith('legend:s')).length, 6);
  assert.deepEqual(checkScene(scene, { measurer: m }), []);
});

test('area: every category label is centred under its point, also the first one next to a value axis', () => {
  for (const options of [{ valueAxis: true }, { valueAxis: true, seriesLabels: 'legend' as const }, {}]) {
    const scene = layoutChart({ ...spec, options: { ...spec.options, ...options } }, CHART_FRAME, m);
    const xs = points(band(scene, 0)).slice(0, spec.categories.length).map((p) => p.x);
    labels(scene, 'label:category').forEach((t, c) => assert.ok(Math.abs(t.x + t.w / 2 - xs[c]) < 0.01, `category ${c} centred`));
    // Totals sit over their points too, the first one as well (no tick labels to its left without an axis).
    if (!options.valueAxis) {
      const t = labels(scene, 'label:total:c0')[0];
      assert.equal(t.x, Math.max(xs[0] - t.w / 2, CHART_FRAME.x), 'first total as centred as the frame allows');
    }
    assert.deepEqual(checkScene(scene, { measurer: m }), [], JSON.stringify(options));
  }
});
