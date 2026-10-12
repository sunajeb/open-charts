import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import type { ChartSpec, LineShape, PathShape, Scene, TextShape } from '../src/core/index.ts';
import { checkScene, layoutChart, validateSpec } from '../src/core/index.ts';
import { segmentHitsBox } from '../src/core/place.ts';
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

test('pie: shares add up to exactly 100 (largest remainder); names outside, shares inside, a sliver\'s with its name', () => {
  const scene = layoutChart(spec, CHART_FRAME, m);
  const shares = [0, 1, 2, 3].map((c) => parseFloat((at(scene, `label:value:s0:c${c}`) as TextShape).text));
  // The 2% sliver is narrower than any box holding "2%": the share joins its name outside, on a leader.
  assert.ok(!scene.roles!.includes('label:value:s0:c4'));
  assert.equal((at(scene, 'label:category:c4') as TextShape).text, 'Islands 2%');
  assert.ok(scene.roles!.includes('leader:c4'));
  shares.push(2);
  assert.equal(shares.reduce((a, b) => a + b, 0), 100);
  assert.deepEqual(shares, [28, 28, 28, 14, 2]);
  assert.deepEqual(textsOf(scene, 'label:category').map((t) => t.text).sort(), [...spec.categories.slice(0, 4), 'Islands 2%'].sort());
  // Names on the right of the pie are left-aligned, those on the left right-aligned.
  assert.equal((at(scene, 'label:category:c0') as TextShape).align, 'left');
  assert.equal((at(scene, 'label:category:c2') as TextShape).align, 'right');
  assert.deepEqual(checkScene(scene, { measurer: m }), []);
  assertLeaders(scene);
});

/** Leaders start on the rim, never enter the pie, never cross each other or another label. */
function assertLeaders(scene: Scene) {
  const items = scene.shapes.map((s, i) => ({ s, role: scene.roles![i] }));
  const disc = arcs(items.find((x) => x.role.startsWith('wedge:'))!.s as PathShape)[0];
  const leaders = items.filter((x) => x.role.startsWith('leader:')) as { s: LineShape; role: string }[];
  const texts = items.filter((x) => x.s.kind === 'text') as { s: TextShape; role: string }[];
  for (const { s, role } of leaders) {
    assert.ok(Math.abs(Math.hypot(s.x1 - disc.cx, s.y1 - disc.cy) - disc.r) < 0.5, `${role} starts on the rim`);
    for (let t = 0.02; t <= 1; t += 0.02) {
      const d = Math.hypot(s.x1 + (s.x2 - s.x1) * t - disc.cx, s.y1 + (s.y2 - s.y1) * t - disc.cy);
      assert.ok(d >= disc.r - 0.5, `${role} enters the pie`);
    }
    const c = role.slice('leader:'.length);
    for (const t of texts) {
      if (t.role.endsWith(c) && t.s.align !== 'center') continue; // its own outside label
      assert.ok(!segmentHitsBox(s.x1, s.y1, s.x2, s.y2, t.s), `${role} crosses "${t.s.text}"`);
    }
  }
  const side = (px: number, py: number, qx: number, qy: number, rx: number, ry: number) => Math.sign((qx - px) * (ry - py) - (qy - py) * (rx - px));
  const cross = (a: LineShape, b: LineShape) =>
    side(a.x1, a.y1, a.x2, a.y2, b.x1, b.y1) * side(a.x1, a.y1, a.x2, a.y2, b.x2, b.y2) < 0
    && side(b.x1, b.y1, b.x2, b.y2, a.x1, a.y1) * side(b.x1, b.y1, b.x2, b.y2, a.x2, a.y2) < 0;
  leaders.forEach((a, i) => leaders.slice(i + 1).forEach((b) => assert.ok(!cross(a.s, b.s), `${a.role} crosses ${b.role}`)));
}

test('pie: small slices\' shares stack in a column outside, spaced apart, with clean leaders (example deck)', () => {
  const deck = JSON.parse(readFileSync(new URL('../examples/pie.json', import.meta.url), 'utf8'));
  for (const slide of deck.slides) {
    const scene = layoutChart(slide.chart, CHART_FRAME, m);
    assert.deepEqual(checkScene(scene, { measurer: m, strict: true }), [], slide.title);
    assertLeaders(scene);
  }
  // Survey (legend): 2.6%, 1.3% and 0.8% do not fit inside and stack outside, one line apart, in a
  // column (shared right edge), in slice order from the bottom up, each joined to its slice by a leader.
  const scene = layoutChart(deck.slides[2].chart, CHART_FRAME, m);
  const out = [4, 5, 6].map((c) => at(scene, `label:value:s0:c${c}`) as TextShape);
  assert.deepEqual(out.map((t) => t.text), ['2.6%', '1.3%', '0.8%']);
  for (const c of [4, 5, 6]) assert.ok(scene.roles!.includes(`leader:c${c}`));
  out.forEach((t) => assert.equal(t.align, 'right'));
  assert.ok(out.every((t) => Math.abs(t.x + t.w - (out[0].x + out[0].w)) < 0.01), 'one column');
  for (let k = 1; k < out.length; k++) assert.ok(out[k - 1].y >= out[k].y + out[k].h - 0.01, 'spread apart, smaller slice higher');
});

test('pie: a boxed share is never wider than its slice; too wide, it goes outside with the name', () => {
  // Example deck: "Other" is 3% of revenue, a 9.5° slice — the "3%" box would be wider than it.
  const deck = JSON.parse(readFileSync(new URL('../examples/pie.json', import.meta.url), 'utf8'));
  const scene = layoutChart(deck.slides[0].chart, CHART_FRAME, m);
  assert.ok(!scene.roles!.includes('label:value:s0:c4'));
  assert.equal((at(scene, 'label:category:c4') as TextShape).text, 'Other 3%');
  assert.ok(scene.roles!.includes('leader:c4'));
  assertLeaders(scene);
  // Every slide and sample, at several sizes: a filled share is no wider (across the slice) than the
  // slice's chord at the rim.
  const specs: ChartSpec[] = [...deck.slides.map((x: { chart: ChartSpec }) => x.chart), samples.pie!, samples.doughnut!, spec];
  for (const sp of specs) {
    for (const [w, h] of [[864, 408], [520, 300], [360, 200], [1100, 560]]) {
      const sc = layoutChart(sp, { x: 0, y: 0, w, h }, m);
      sc.shapes.forEach((t, i) => {
        const mm = sc.roles![i].match(/^label:value:s0:c(\d+)$/);
        if (t.kind !== 'text' || !t.fill || !mm) return;
        const [a] = arcs(at(sc, `wedge:s0:c${mm[1]}`) as PathShape);
        const mid = ((a.start + a.sweep / 2) * Math.PI) / 180;
        const across = t.w * Math.abs(Math.sin(mid)) + t.h * Math.abs(Math.cos(mid));
        assert.ok(a.sweep >= 180 || across <= 2 * a.r * Math.sin((a.sweep * Math.PI) / 360) + 0.01, `${sp.type} ${w}x${h}: "${t.text}" wider than its slice`);
      });
    }
  }
});

test('pie: label content shows values, or values and shares', () => {
  const vals = layoutChart({ ...spec, options: { labelContent: 'value', unit: 'm' } }, CHART_FRAME, m);
  assert.equal((at(vals, 'label:value:s0:c0') as TextShape).text, '1.00m');
  const both = layoutChart({ ...spec, series: [{ name: 'S', values: [118, 60, 29, 14, 6] }], options: { labelContent: 'both' } }, CHART_FRAME, m);
  assert.equal((at(both, 'label:value:s0:c0') as TextShape).text, '118 (52%)');
  assert.deepEqual(checkScene(both, { measurer: m }), []);
  assertLeaders(both);
  assert.deepEqual(validateSpec({ ...spec, options: { labelContent: 'share' } } as unknown as ChartSpec).map((e) => e.path), ['/options/labelContent']);
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
