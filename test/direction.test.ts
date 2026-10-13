// options.direction (the gallery's ▲ ▼ ◀ ▶ arrows, manual: Rotate and flip elements): which way a
// chart grows, per type; down / left mirror the up / right layouts; the old `orientation` migrates.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import type { ChartSpec, LineShape, PathShape, RectShape, Scene, TextShape } from '../src/core/index.ts';
import { checkScene, layoutChart, migrate, validateSpec } from '../src/core/index.ts';
import { DIRECTIONS, directionsOf } from '../src/core/spec.ts';
import { samples } from '../src/core/samples.ts';
import { createMeasurer } from '../src/fonts.ts';

const m = createMeasurer();
const F = { x: 48, y: 96, w: 864, h: 408 };
const lay = (spec: ChartSpec, direction?: string) => layoutChart({ ...spec, options: { ...spec.options, ...(direction ? { direction } : {}) } } as ChartSpec, F, m);
const byRole = <T extends Scene['shapes'][number]>(s: Scene, prefix: string) =>
  s.shapes.filter((_, i) => s.roles![i].startsWith(prefix)) as T[];
const sample = (t: keyof typeof samples) => samples[t]!;
const near = (a: number, b: number, tol = 0.01) => Math.abs(a - b) <= tol;

test('direction: the types and directions match the reference gallery', () => {
  assert.deepEqual(Object.keys(DIRECTIONS), ['stacked', 'clustered', 'percent', 'waterfall', 'mekko', 'line', 'area', 'area100', 'combo']);
  for (const t of ['stacked', 'clustered', 'percent', 'waterfall', 'mekko', 'combo']) assert.deepEqual(directionsOf(t), ['up', 'down', 'left', 'right'], t);
  for (const t of ['line', 'area', 'area100']) assert.deepEqual(directionsOf(t), ['up', 'left', 'right'], t);
  for (const t of ['pie', 'doughnut', 'scatter', 'bubble', 'gantt']) assert.deepEqual(directionsOf(t), ['up'], t);
});

test('direction: validated per type', () => {
  assert.deepEqual(validateSpec({ ...sample('waterfall'), options: { direction: 'left' } }), []);
  assert.match(validateSpec({ ...sample('area'), options: { direction: 'down' } })[0].message, /up, left, right for area charts/);
  assert.match(validateSpec({ ...sample('pie'), options: { direction: 'up' } })[0].message, /not used by pie charts/);
  assert.match(validateSpec({ ...sample('stacked'), options: { direction: 'north' } })[0].message, /up, down, left, right/);
});

test('direction: orientation migrates to direction, in v0 and v1 specs, with a warning', () => {
  const warnings: string[] = [];
  const v1 = migrate({ ...sample('stacked'), options: { unit: 'm', orientation: 'horizontal', totals: true } } as ChartSpec, warnings);
  assert.deepEqual(v1.options, { unit: 'm', direction: 'right', totals: true }, 'renamed in place, order kept');
  assert.match(warnings[0], /\/options\/orientation: renamed to "direction" \("right"\)/);
  const v0 = migrate({ type: 'clustered', categories: ['a'], series: [{ name: 's', values: [1] }], options: { orientation: 'vertical' } } as unknown as ChartSpec);
  assert.deepEqual(v0.options, {}, 'vertical is the default direction');
  assert.equal(v0.specVersion, 1);
  // The old spec lays out exactly like the new one.
  const old = layoutChart({ ...sample('stacked'), options: { ...sample('stacked').options, orientation: 'horizontal' } } as unknown as ChartSpec, F, m);
  assert.deepEqual(old, lay(sample('stacked'), 'right'));
});

/** Every bar of `a` mirrored across the frame equals the same bar of `b`. */
function assertMirrored(a: Scene, b: Scene, axis: 'x' | 'y') {
  const bars = a.roles!.flatMap((r, i) => (r.startsWith('bar:') ? [i] : []));
  assert.ok(bars.length > 0);
  for (const i of bars) {
    const ra = a.shapes[i] as RectShape, rb = b.shapes[b.roles!.indexOf(a.roles![i])] as RectShape;
    if (axis === 'y') assert.ok(near(rb.y, 2 * F.y + F.h - ra.y - ra.h) && near(rb.x, ra.x) && near(rb.h, ra.h), a.roles![i]);
    else assert.ok(near(rb.x, 2 * F.x + F.w - ra.x - ra.w) && near(rb.y, ra.y) && near(rb.w, ra.w), a.roles![i]);
  }
}

test('direction down: columns hang from a baseline at the top, categories above them, totals below', () => {
  for (const t of ['stacked', 'clustered', 'percent', 'waterfall', 'mekko', 'combo'] as const) {
    const up = lay(sample(t)), down = lay(sample(t), 'down');
    assertMirrored(up, down, 'y');
    const bars = byRole<RectShape>(down, 'bar:');
    const cats = byRole<TextShape>(down, 'label:category:').filter((x) => x.role === 'category');
    const barTop = Math.min(...bars.map((b) => b.y));
    assert.ok(cats.every((c) => c.y + c.h <= barTop + 0.5), `${t}: category labels above the hanging columns`);
    assert.ok(cats.every((c) => c.valign === 'bottom'), `${t}: category labels hug the baseline`);
  }
  const st = lay(sample('stacked'), 'down');
  sample('stacked').categories.forEach((_, c) => {
    const [total] = byRole<TextShape>(st, `label:total:c${c}`);
    const col = st.shapes.filter((_, i) => new RegExp(`^bar:s\\d+:c${c}$`).test(st.roles![i])) as RectShape[];
    assert.ok(total.y >= Math.max(...col.map((b) => b.y + b.h)) - 0.5, 'a total sits below its hanging column');
  });
});

test('direction left: bars grow leftwards from a baseline on the right, categories right of them', () => {
  for (const t of ['stacked', 'clustered', 'percent', 'waterfall', 'mekko', 'combo', 'line', 'area', 'area100'] as const) {
    const right = lay(sample(t), 'right'), left = lay(sample(t), 'left');
    if (t !== 'line' && t !== 'area' && t !== 'area100') assertMirrored(right, left, 'x');
    const cats = byRole<TextShape>(left, 'label:category:').filter((x) => x.role === 'category');
    const marks = left.shapes.filter((_, i) => /^(bar|marker|area):/.test(left.roles![i]));
    const xMax = Math.max(...marks.map((s) => (s.kind === 'rect' ? s.x + s.w : Math.max(...(s as PathShape).commands.flatMap((c) => ('x' in c ? [c.x] : []))))));
    assert.ok(cats.length > 0 && cats.every((c) => c.x >= xMax - 0.5 && c.align === 'left'), `${t}: category labels right of the bars, left-aligned`);
  }
});

test('direction right: waterfall bars run the bridge across, connectors drop to the next bar', () => {
  const s = lay(sample('waterfall'), 'right');
  const bars = byRole<RectShape>(s, 'bar:');
  assert.ok(bars.every((b) => b.w > 0 && near(b.h, bars[0].h)), 'one bar thickness');
  for (let c = 1; c < bars.length; c++) assert.ok(bars[c].y > bars[c - 1].y, 'categories top to bottom');
  for (const l of byRole<LineShape>(s, 'connector:')) assert.ok(near(l.x1, l.x2) && l.y2 > l.y1, 'connectors are vertical');
  // The start total begins at zero, the step after it continues from its end.
  assert.ok(near(bars[1].x, bars[0].x + bars[0].w, 0.5));
  assert.deepEqual(checkScene(s, { measurer: m }), []);
});

test('direction right: a line chart becomes a profile — categories down, values across', () => {
  const s = lay(sample('line'), 'right');
  const pts = byRole<RectShape>(s, 'marker:s0:');
  for (let c = 1; c < pts.length; c++) assert.ok(pts[c].y > pts[c - 1].y, 'one point per row, top to bottom');
  const vals = (sample('line').series[0].values as number[]);
  const order = (xs: number[]) => xs.map((_, i) => i).sort((a, b) => xs[a] - xs[b]).join();
  assert.equal(order(pts.map((p) => p.x)), order(vals), 'larger values further right');
  const ticks = byRole<TextShape>(s, 'axis:tick:').filter((t) => t.kind === 'text');
  assert.ok(ticks.length && ticks.every((t) => t.y > Math.max(...pts.map((p) => p.y))), 'value axis along the bottom');
});

test('direction right: Mekko rows are as tall as their share of the total; areas stack across', () => {
  const mk = lay(sample('mekko'), 'right');
  const spec = sample('mekko');
  const totals = spec.categories.map((_, c) => spec.series.reduce((a, x) => a + (x.values[c] as number), 0));
  const rowH = spec.categories.map((_, c) => byRole<RectShape>(mk, `bar:s0:c${c}`)[0].h);
  for (let c = 1; c < totals.length; c++) assert.ok(near(rowH[c] / rowH[0], totals[c] / totals[0], 1e-6));
  const ar = lay(sample('area100'), 'right');
  const band = byRole<PathShape>(ar, 'area:s0')[0].commands.filter((c) => c.op === 'L' || c.op === 'M') as { x: number; y: number }[];
  assert.ok(Math.max(...band.map((p) => p.y)) - Math.min(...band.map((p) => p.y)) > F.h * 0.6, 'a band runs down the plot');
});

test('direction: every directional sample is clean in every direction', () => {
  for (const [t, dirs] of Object.entries(DIRECTIONS)) {
    for (const d of dirs) {
      const s = lay(sample(t as keyof typeof samples), d);
      assert.deepEqual(checkScene(s, { measurer: m }), [], `${t} ${d}`);
      // A mirrored layout draws nothing that points a way (a triangle would keep pointing up).
      if (d === 'down' || d === 'left') assert.ok(!s.shapes.some((x) => x.kind === 'rect' && x.geom === 'triangle'), `${t} ${d}`);
    }
  }
});

test('direction: a label pinned by hand stays where it was put in a mirrored chart', () => {
  const spec = { ...sample('stacked'), options: { ...sample('stacked').options, direction: 'down' }, overrides: { 'label:total:c0': { pin: { x: 0.1, y: 0.2 } } } } as unknown as ChartSpec;
  const s = layoutChart(spec, F, m);
  const [t] = byRole<TextShape>(s, 'label:total:c0');
  assert.ok(near(t.x + t.w / 2, F.x + 0.1 * F.w, 0.5) && near(t.y + t.h / 2, F.y + 0.2 * F.h, 0.5), 'at its pin, not mirrored');
});
