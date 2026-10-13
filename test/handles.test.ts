// Drag handles (src/core/handles.ts, src/addin/handles.ts): which handles a chart offers and where,
// what a drag does to the spec, the slide -> screen mapping, and the controller end to end with a
// stand-in helper (selection -> handles, drag -> ghosts -> one apply on release, reset).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createMeasurer } from '../src/fonts.ts';
import { defaultStyle, dragHandles, dragSpec, describeDrag, layoutChart, resetSpec, valueAt, niceStep, type ChartSpec, type Scene } from '../src/core/index.ts';
import { ghostShapes, handlesController, MARKER, screenHandles, toScreen, toSlide, viewFrom, type View } from '../src/addin/handles.ts';
import { parseRequest } from '../src/addin/engine.ts';
import { getSession, type Session } from '../src/server.ts';
import type { PopupHelper } from '../src/addin/popup.ts';

const m = createMeasurer();
const F = { x: 48, y: 96, w: 600, h: 360 };
const lay = (s: ChartSpec) => layoutChart(s, F, m, defaultStyle);
const stacked: ChartSpec = { type: 'stacked', categories: ['A', 'B', 'C', 'D'], series: [{ name: 'x', values: [10, 20, 30, 40] }, { name: 'y', values: [5, 5, 5, 5] }], options: { valueAxis: true } };
const handle = (s: ChartSpec, sc: Scene, id: string, sel: string[] = []) => dragHandles(s, sc, sel).find((h) => h.id === id)!;

test('handles: gap on the first column edge, axis top, selected labels only, none on a placeholder', () => {
  const sc = lay(stacked);
  const hs = dragHandles(stacked, sc);
  assert.deepEqual(hs.map((h) => h.id), ['gap', 'axisMax']);
  const bar = sc.shapes[sc.roles!.indexOf('bar:s0:c0')] as { x: number; w: number };
  const gap = hs[0];
  assert.ok(Math.abs(gap.at.x - (bar.x + bar.w)) < 0.01, 'on the right edge of the first column');
  assert.equal(gap.cursor, 'ew');
  assert.equal(hs[1].cursor, 'ns');
  assert.deepEqual(hs[1].at, { x: sc.valueScale!.at![2], y: sc.valueScale!.at![3] });
  // No value axis drawn: no axis handle.
  assert.deepEqual(dragHandles({ ...stacked, options: {} }, lay({ ...stacked, options: {} })).map((h) => h.id), ['gap']);
  // A selected label is an area handle; a selected bar is not a handle.
  const withLabel = dragHandles(stacked, sc, ['label:value:s0:c1', 'bar:s0:c1']);
  const l = withLabel.find((h) => h.kind === 'label')!;
  assert.equal(l.id, 'label:label:value:s0:c1');
  assert.ok(l.box && l.box.w > 0);
  // Horizontal bars: the gap handle sits on the bottom edge and drags up/down.
  const hz = { ...stacked, options: { direction: 'right' as const } };
  assert.equal(dragHandles(hz, lay(hz))[0].cursor, 'ns');
  // Mirrored directions: the axis handle is still at the value axis's top end.
  for (const direction of ['down', 'left', 'right'] as const) {
    const s = { ...stacked, options: { valueAxis: true, direction } };
    const sc2 = lay(s);
    const ax = dragHandles(s, sc2).find((h) => h.id === 'axisMax')!;
    assert.ok(ax, direction);
    assert.ok(Math.abs(valueAt(sc2, ax.at.x, ax.at.y)! - sc2.valueScale!.max) < 1e-6, direction);
    const axisLine = sc2.shapes.filter((_, i) => sc2.roles![i] === 'axis');
    assert.ok(axisLine.some((l) => l.kind === 'line' && Math.min(Math.hypot(l.x1 - ax.at.x, l.y1 - ax.at.y), Math.hypot(l.x2 - ax.at.x, l.y2 - ax.at.y)) < 40), `${direction}: near an axis end`);
  }
  // Pie: nothing to drag here.
  const pie: ChartSpec = { type: 'pie', categories: ['a', 'b'], series: [{ name: 's', values: [1, 2] }] };
  assert.deepEqual(dragHandles(pie, lay(pie)), []);
});

test('gap drag: the column edge follows the pointer, clamped and rounded', () => {
  const sc = lay(stacked);
  const h = handle(stacked, sc, 'gap');
  assert.equal(dragSpec(stacked, sc, h, 0, 0).options!.gap, 0.5);
  const wider = dragSpec(stacked, sc, h, 10, 0);
  assert.ok(wider.options!.gap! < 0.5);
  // The new layout puts the edge where the pointer is (within the 0.05 rounding).
  const bar0 = sc.shapes[sc.roles!.indexOf('bar:s0:c0')] as { x: number; w: number };
  const after = lay(wider);
  const bar1 = after.shapes[after.roles!.indexOf('bar:s0:c0')] as { x: number; w: number };
  assert.ok(Math.abs(bar1.w - (bar0.w + 20)) < 4, `width ${bar0.w} -> ${bar1.w}`);
  assert.equal(dragSpec(stacked, sc, h, 1000, 0).options!.gap, 0, 'columns touch at most');
  assert.equal(dragSpec(stacked, sc, h, -1000, 0).options!.gap, 10, 'the thinnest column the limits allow');
  assert.equal(describeDrag(wider, h), `Gap ${Math.round(wider.options!.gap! * 100)}%`);
  assert.equal(resetSpec(wider, h).options!.gap, undefined);
  assert.equal(stacked.options!.gap, undefined, 'the input is not changed');
});

test('axis max drag: value under the pointer, rounded; below the data top it goes back to auto', () => {
  const sc = lay(stacked);
  const h = handle(stacked, sc, 'axisMax');
  const vs = sc.valueScale!;
  assert.equal(valueAt(sc, vs.at![0], vs.at![1]), vs.min);
  assert.equal(valueAt(sc, vs.at![2], vs.at![3]), vs.max);
  const up = dragSpec(stacked, sc, h, 0, -(vs.at![1] - vs.at![3]) / 2, lay); // half the axis higher: 1.5 × max
  assert.equal(up.options!.axisMax, vs.max * 1.5);
  assert.equal(lay(up).valueScale!.max, vs.max * 1.5, 're-laid out to the new top');
  const down = dragSpec(up, lay(up), handle(up, lay(up), 'axisMax'), 0, 200, lay);
  assert.equal(down.options!.axisMax, undefined, 'below the data: automatic');
  assert.equal(niceStep(0.3), 0.5);
  assert.equal(niceStep(7), 10);
  assert.equal(resetSpec(up, h).options!.axisMax, undefined);
});

test('label drag pins the centre as frame fractions; reset unpins', () => {
  const sc = lay(stacked);
  const h = handle(stacked, sc, 'label:label:value:s0:c1', ['label:value:s0:c1']);
  const moved = dragSpec(stacked, sc, h, 30, -20);
  const pin = moved.overrides!['label:value:s0:c1'].pin!;
  assert.ok(Math.abs(pin.x - (h.at.x + 30 - F.x) / F.w) < 1e-3);
  assert.ok(Math.abs(pin.y - (h.at.y - 20 - F.y) / F.h) < 1e-3);
  const t = lay(moved).shapes[lay(moved).roles!.indexOf('label:value:s0:c1')] as { x: number; y: number; w: number; h: number; pinned?: boolean };
  assert.ok(t.pinned);
  assert.ok(Math.abs(t.x + t.w / 2 - (h.at.x + 30)) < 0.5, 'the label is where it was dropped');
  assert.equal(dragSpec(stacked, sc, h, -5000, 0).overrides!['label:value:s0:c1'].pin!.x, 0, 'kept inside the frame');
  assert.equal(resetSpec(moved, h).overrides, undefined);
});

test('level arrow: dragged to the nearer of beside / between', () => {
  const spec: ChartSpec = { ...stacked, options: { levelArrows: [{ from: { category: 1 }, to: { category: 3 } }] } };
  const sc = lay(spec);
  const h = handle(spec, sc, 'level:0');
  assert.ok(h, 'the arrow has a handle');
  const between = lay({ ...spec, options: { levelArrows: [{ ...spec.options!.levelArrows![0], position: 'between' }] } });
  const t0 = sc.shapes[sc.roles!.lastIndexOf('arrow:level:0')] as { x: number; w: number };
  const t1 = between.shapes[between.roles!.lastIndexOf('arrow:level:0')] as { x: number; w: number };
  const out = dragSpec(spec, sc, h, t1.x + t1.w / 2 - (t0.x + t0.w / 2), 0, lay);
  assert.equal(out.options!.levelArrows![0].position, 'between');
  assert.equal(dragSpec(spec, sc, h, 0, 0, lay).options!.levelArrows![0].position, 'beside');
  assert.equal(resetSpec(out, h).options!.levelArrows![0].position, undefined);
});

// ---- Mapping ----------------------------------------------------------------------------------------

const screens = [{ frame: { x: 0, y: 0, w: 1512, h: 982 }, visible: { x: 0, y: 38, w: 1512, h: 944 }, scale: 2 }];
const win = [{ x: 100, y: 50, w: 1200, h: 800 }];

test('view: slide -> screen from the group corners, in points or Retina pixels; untrusted otherwise', () => {
  const bounds = { x: 48, y: 96, w: 600, h: 360 };
  // Zoom 80 %, slide origin at (300, 250) on screen.
  const pts: [number, number, number, number] = [300 + 48 * 0.8, 250 + 96 * 0.8, 300 + 648 * 0.8, 250 + 456 * 0.8];
  const v = viewFrom(pts, bounds, win, screens)!;
  assert.ok(Math.abs(v.k - 0.8) < 1e-9 && Math.abs(v.ox - 300) < 1e-9 && Math.abs(v.oy - 250) < 1e-9);
  const px = viewFrom(pts.map((n) => n * 2) as typeof pts, bounds, win, screens)!;
  assert.ok(Math.abs(px.k - 0.8) < 1e-9, 'pixels on a Retina display are halved');
  // Whole-pixel rounding still gives one zoom; a stretched rect or one outside the window does not.
  assert.ok(viewFrom(pts.map(Math.round) as typeof pts, bounds, win, screens));
  assert.equal(viewFrom([pts[0], pts[1], pts[2], pts[3] + 60], bounds, win, screens), undefined);
  assert.equal(viewFrom([5000, 5000, 5100, 5100], bounds, win, screens), undefined);
  const p = toScreen(v, 100, 200), q = toSlide(v, p.x, p.y);
  assert.ok(Math.abs(q.x - 100) < 1e-9 && Math.abs(q.y - 200) < 1e-9);
});

test('screen handles: markers keep their size at any zoom; label boxes scale', () => {
  const sc = lay(stacked);
  const hs = dragHandles(stacked, sc, ['label:value:s0:c1']);
  for (const k of [0.5, 1.5]) {
    const v: View = { k, ox: 10, oy: 20 };
    const s = screenHandles(hs, v);
    const gap = s.find((h) => h.id === 'gap')!;
    assert.deepEqual([gap.rect[2], gap.rect[3]], [MARKER, MARKER]);
    const c = toScreen(v, hs[0].at.x, hs[0].at.y);
    assert.ok(Math.abs(gap.rect[0] + MARKER / 2 - c.x) < 1e-9);
    const label = s.find((h) => !h.marker)!;
    const box = hs.find((h) => h.box)!.box!;
    assert.ok(Math.abs(label.rect[2] - Math.max(8, box.w * k)) < 1e-9);
  }
  const g = ghostShapes(sc, { k: 1, ox: 0, oy: 0 }, sc.frame);
  assert.equal((g[0] as { t: string }).t, 'r', 'a backdrop first');
  assert.ok(g.length > sc.shapes.length);
});

// ---- Controller --------------------------------------------------------------------------------------

function fakeHelper() {
  const sent: Array<Record<string, unknown>> = [];
  const helper: PopupHelper = {
    warm: () => true, show: async () => undefined, hide: () => {}, eval: async () => undefined, stop: () => {}, running: true,
    info: async () => ({ windows: win, screens }),
    send: (msg) => { sent.push(msg as Record<string, unknown>); },
  };
  return { helper, sent };
}

const tick = () => new Promise((r) => setImmediate(r));

test('controller: re-measures only while PowerPoint shows a window (a macro then would start it again)', async () => {
  const { helper, sent } = fakeHelper();
  let windows = win;
  helper.info = async () => ({ windows, screens });
  let measured = 0;
  const c = handlesController({ helper, measurer: () => m, apply: () => {}, now: () => 0, measure: async () => { measured++; return undefined; } });
  c.event({ event: 'remeasure' });
  await tick(); await tick();
  assert.equal(measured, 1, 'a window: measured');
  windows = [];
  c.event({ event: 'remeasure' });
  await tick(); await tick();
  assert.equal(measured, 1, 'no window (PowerPoint closing or quitting): no macro');
  assert.equal(sent.at(-1)?.hide, true, 'the handles are hidden');
});

test('controller: selection shows handles; a drag sends ghosts and applies once on release; reset', async () => {
  const { helper, sent } = fakeHelper();
  const applied: Session[] = [];
  const c = handlesController({ helper, measurer: () => m, apply: (s) => { applied.push(s); }, now: () => 0 });
  const spec = JSON.stringify(stacked);
  const k = 0.8, ox = 300, oy = 250;
  const view = [48 * k + ox, 96 * k + oy, 648 * k + ox, 456 * k + oy].join(',');
  const text = `select\n\n48,96,600,360\n48,96,600,360\nview=${view}\n${spec}`;
  await c.select(parseRequest(text));
  const shown = sent.at(-1)!;
  assert.equal(shown.cmd, 'handles');
  const hs = shown.handles as Array<{ id: string; rect: number[] }>;
  assert.deepEqual(hs.map((h) => h.id), ['gap', 'axisMax']);
  const gap = hs[0];
  const x0 = gap.rect[0] + gap.rect[2] / 2, y0 = gap.rect[1] + gap.rect[3] / 2;
  c.event({ event: 'handle', phase: 'down', id: 'gap', x: x0, y: y0 });
  for (let i = 1; i <= 5; i++) c.event({ event: 'handle', phase: 'move', id: 'gap', x: x0 + i * 2, y: y0 });
  await tick();
  const ghosts = sent.filter((s) => s.cmd === 'ghost' && s.shapes);
  assert.equal(ghosts.length, 1, 'moves are coalesced: the newest wins');
  assert.match(String(ghosts[0].tip), /^Gap \d+%$/);
  c.event({ event: 'handle', phase: 'up', id: 'gap', x: x0 + 10, y: y0 });
  await tick();
  assert.equal(applied.length, 1, 'one Apply on release');
  const s = getSession(applied[0].id)!;
  assert.equal(s.state, 'applied');
  const gapNow = (s.chart as ChartSpec).options!.gap!;
  assert.ok(gapNow < 0.5, `gap ${gapNow}`);
  // Moved 10 screen pt at 80 %: 12.5 slide pt per edge.
  const sc = lay(stacked);
  assert.equal(gapNow, dragSpec(stacked, sc, handle(stacked, sc, 'gap'), 10 / k, 0).options!.gap);
  // A click without a move changes nothing; a double-click resets.
  c.event({ event: 'handle', phase: 'down', id: 'gap', x: x0, y: y0 });
  c.event({ event: 'handle', phase: 'up', id: 'gap', x: x0, y: y0 });
  await tick();
  assert.equal(applied.length, 1);
  c.event({ event: 'handle', phase: 'reset', id: 'gap', x: x0, y: y0 });
  await tick();
  assert.equal(applied.length, 2);
  assert.equal((getSession(applied[1].id)!.chart as ChartSpec).options!.gap, undefined);
  // Nothing selected: hidden.
  await c.select(parseRequest('select\n\n'));
  assert.deepEqual(sent.at(-1), { cmd: 'handles', hide: true });
});
