// Optional rows: `Series.hidden` leaves a datasheet row out of the chart while its data stays in the spec.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import type { ChartSpec, RectShape, Scene, Series } from '../src/core/index.ts';
import { CHART_TYPES, checkScene, defaultStyle, layoutChart, validateSpec } from '../src/core/index.ts';
import { samples } from '../src/core/samples.ts';
import { createMeasurer } from '../src/fonts.ts';
import { altText, CHART_FRAME } from '../src/deck.ts';

const m = createMeasurer();
const lay = (spec: ChartSpec) => layoutChart(spec, CHART_FRAME, m);
const clone = <T>(v: T): T => structuredClone(v);
const fillOf = (scene: Scene, role: string) => (scene.shapes[scene.roles!.indexOf(role)] as RectShape | undefined)?.fill;
const seriesIn = (scene: Scene) => new Set(scene.roles!.flatMap((r) => [...r.matchAll(/(?:^|:)s(\d+)(?=:|$)/g)].map((x) => +x[1])));

test('hidden: validated (boolean, something visible, positional rows stay)', () => {
  const base = clone(samples.stacked!);
  base.series[1].hidden = 'yes' as unknown as boolean;
  assert.deepEqual(validateSpec(base).map((e) => e.path), ['/series/1/hidden']);
  const all = clone(samples.stacked!);
  all.series.forEach((s) => { s.hidden = true; });
  assert.match(validateSpec(all).map((e) => `${e.path} ${e.message}`).join('\n'), /^\/series every series is hidden/);
  const gantt = clone(samples.gantt!);
  gantt.series[0].hidden = true;
  gantt.series[1].hidden = true;
  assert.deepEqual(validateSpec(gantt).map((e) => e.path), ['/series/0/hidden', '/series/1/hidden']);
  const xy = clone(samples.mekkoxy!);
  xy.series[0].hidden = true;
  assert.deepEqual(validateSpec(xy).map((e) => e.path), ['/series/0/hidden']);
  const heights = clone(samples.mekkoxy!);
  heights.series.slice(1).forEach((s) => { s.hidden = true; });
  assert.match(validateSpec(heights)[0].message, /at least one visible height series/);
});

test('hidden: an extra hidden row changes nothing in any chart type', () => {
  for (const type of CHART_TYPES) {
    const spec = samples[type];
    if (!spec) continue;
    const extra = clone(spec);
    extra.series.push({ ...clone(spec.series[spec.series.length - 1]), name: 'Hidden extra', color: undefined, hidden: true } as Series);
    assert.deepEqual(lay(extra), lay(spec), type);
  }
});

test('hidden: a hidden middle row leaves the chart; the others keep colour and role index', () => {
  for (const type of ['stacked', 'clustered', 'percent', 'mekko', 'line', 'area', 'area100', 'combo', 'waterfall', 'scatter', 'bubble'] as const) {
    const spec = clone(samples[type]!);
    while (spec.series.length < 3) spec.series.push({ ...clone(spec.series[0]), name: `Extra ${spec.series.length}`, as: undefined });
    const shown = lay(spec);
    spec.series[1].hidden = true;
    const scene = lay(spec);
    assert.ok(!seriesIn(scene).has(1), `${type}: no shape of the hidden series`);
    assert.ok(seriesIn(scene).has(2), `${type}: series 2 keeps its index`);
    assert.deepEqual(checkScene(scene, { measurer: m }), [], type);
    // Same colour for the same row, hidden neighbour or not.
    for (const role of scene.roles!.filter((r) => /^(bar|marker):s[02]:/.test(r))) {
      assert.equal(fillOf(scene, role), fillOf(shown, role), `${type} ${role}`);
    }
  }
});

test('hidden: mekkoxy widths and gantt dates are read from their rows, not the first visible ones', () => {
  const xy = clone(samples.mekkoxy!);
  xy.series.push({ name: 'Extra height', values: xy.series[1].values.slice() });
  xy.series[1].hidden = true;
  const scene = lay(xy);
  assert.ok(scene.roles!.some((r) => r.startsWith('bar:s2:')) && !scene.roles!.some((r) => r.startsWith('bar:s1:')));
  const gantt = clone(samples.gantt!);
  const ms = gantt.series.findIndex((s, i) => i >= 2 && s.values.some((v) => v != null));
  assert.ok(ms >= 2, 'sample has a milestone row');
  gantt.series[ms].hidden = true;
  const g = lay(gantt);
  assert.ok(!g.roles!.some((r) => r.startsWith(`marker:s${ms}:`)));
  assert.equal(g.roles!.filter((r) => r.startsWith('bar:s0:')).length, lay(samples.gantt!).roles!.filter((r) => r.startsWith('bar:s0:')).length);
});

test('hidden: waterfall categories marked "e" only in a hidden row stay totals', () => {
  const spec: ChartSpec = {
    type: 'waterfall', categories: ['Start', 'A', 'B', 'End'],
    series: [{ name: 'Marks', values: [100, 10, 5, 'e'] }, { name: 'Deltas', values: [null, -20, 15, null] }],
  };
  const all = lay(spec);
  spec.series[0].hidden = true;
  const scene = lay(spec);
  // End is still a total: a bar from the baseline, labelled with the running sum of the visible deltas.
  const totals = scene.shapes.filter((s, i) => scene.roles![i] === 'label:total:c3');
  assert.equal(totals.length, 1);
  assert.equal((totals[0] as { text: string }).text, '\u22125');
  assert.ok(all.roles!.includes('label:total:c3'));
});

test('hidden: pie and doughnut show the first visible series', () => {
  const pie = clone(samples.pie!);
  pie.series = [{ name: 'Old', values: pie.series[0].values.map(() => 1) }, ...pie.series];
  pie.series[0].hidden = true;
  const scene = lay(pie);
  const plain = lay(samples.pie!);
  assert.deepEqual(scene.shapes, plain.shapes);
  assert.ok(!seriesIn(scene).has(0));
  assert.ok(scene.roles!.some((r) => r.startsWith('wedge:s1:')));
});

test('hidden: table and Harvey rows drop out; alt text names visible series only', () => {
  const table = clone(samples.table!);
  table.series[0].hidden = true;
  const t = lay(table);
  assert.ok(!t.shapes.some((s) => s.kind === 'text' && s.text === table.series[0].name));
  const harvey = clone(samples.harvey!);
  harvey.series[1].hidden = true;
  const h = lay(harvey);
  assert.ok(!h.roles!.some((r) => /^ball:s1:/.test(r)));
  assert.ok(!altText(harvey).includes(harvey.series[1].name));
  assert.equal(defaultStyle.palette.length, 6, 'palette untouched');
});
