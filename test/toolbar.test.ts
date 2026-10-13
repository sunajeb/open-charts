// The element mini toolbar (src/addin/toolbar.ts): which controls a selection gets, how a change
// becomes per-element overrides, where the toolbar goes, and the server routes end to end with a
// stand-in helper.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createServer } from 'node:net';
import { getSession, serve, type Session } from '../src/server.ts';
import { specKey } from '../src/addin/engine.ts';
import { controlsFor, currentStyle, placeToolbar, TOOLBAR_SIZE, toolbarRoutes, withOverrides } from '../src/addin/toolbar.ts';
import type { PopupHelper, Rect } from '../src/addin/popup.ts';

const screen = { frame: { x: 0, y: 0, w: 1512, h: 982 }, visible: { x: 0, y: 38, w: 1512, h: 944 }, scale: 2 };

test('controls: the properties every selected element takes', () => {
  assert.deepEqual(controlsFor(['label:value:s0:c1']).map((c) => c.prop), ['size', 'rotate', 'fill']);
  assert.deepEqual(controlsFor(['bar:s0:c1']).map((c) => c.prop), ['fill']);
  assert.deepEqual(controlsFor(['label:value:s0:c1', 'bar:s0:c1']).map((c) => c.prop), ['fill']);
  assert.deepEqual(controlsFor(['gridline:2']), []);
  assert.equal(controlsFor(['label:value:s0:c0'])[1].group, 'rotation');
});

test('a change becomes overrides on each selected role; null removes; empty overrides go away', () => {
  const chart = { type: 'clustered', overrides: { 'bar:s0:c0': { fill: '#000000' } } };
  const a = withOverrides(chart, ['bar:s0:c0', 'bar:s0:c1'], { fill: '#1F3A5F' });
  assert.deepEqual(a.overrides, { 'bar:s0:c0': { fill: '#1F3A5F' }, 'bar:s0:c1': { fill: '#1F3A5F' } });
  assert.deepEqual(chart.overrides, { 'bar:s0:c0': { fill: '#000000' } }, 'the input is not changed');
  const b = withOverrides(a, ['label:value:s0:c0'], { size: 14, rotate: 90 });
  assert.deepEqual(b.overrides!['label:value:s0:c0'], { size: 14, rotate: 90 });
  assert.deepEqual(currentStyle(b, ['bar:s0:c0', 'bar:s0:c1']), { fill: '#1F3A5F' });
  assert.deepEqual(currentStyle(withOverrides(b, ['bar:s0:c1'], { fill: '#FFFFFF' }), ['bar:s0:c0', 'bar:s0:c1']), {});
  const c = withOverrides({ type: 'pie' }, ['wedge:s0:c0'], { fill: null });
  assert.equal('overrides' in c, false);
});

test('toolbar placement: above the selection, below when there is no room, inside the display', () => {
  assert.deepEqual(placeToolbar({ x: 300, y: 400, w: 40, h: 20 }, [], [screen]), { x: 300, y: 400 - 8 - TOOLBAR_SIZE.h, ...TOOLBAR_SIZE });
  assert.deepEqual(placeToolbar({ x: 300, y: 50, w: 40, h: 20 }, [], [screen]), { x: 300, y: 78, ...TOOLBAR_SIZE });
  const r = placeToolbar({ x: 1500, y: 400, w: 10, h: 10 }, [], [screen])!;
  assert.equal(r.x + r.w, 1512);
  assert.equal(placeToolbar(undefined, [], [screen]), undefined);
});

const freePort = () => new Promise<number>((resolve) => {
  const s = createServer();
  s.listen(0, '127.0.0.1', () => { const { port } = s.address() as { port: number }; s.close(() => resolve(port)); });
});

test('routes: a selection shows the toolbar next to it; a change applies overrides; deselect hides it', async () => {
  const port = await freePort();
  const base = `http://127.0.0.1:${port}`;
  const shown: Array<{ panel: string; arg: unknown; frame: Rect }> = [], hidden: string[] = [], applied: Session[] = [];
  const helper = {
    info: async () => ({ windows: [{ x: 0, y: 38, w: 1512, h: 944 }], screens: [screen] }),
    showPanel: (panel: string, _url: string, arg: unknown, frame: Rect) => { shown.push({ panel, arg, frame }); },
    hide: (p: string) => { hidden.push(p); },
  } as unknown as PopupHelper;
  const server = serve(port, {}, { routes: toolbarRoutes({ base, helper, apply: (s) => applied.push(s) }) });
  try {
    await new Promise((r) => server.once('listening', r));
    const spec = JSON.stringify({ type: 'clustered', categories: ['A', 'B'], series: [{ name: 'S', values: [1, 2] }] });
    const id = specKey(spec);
    const select = (body: string) => fetch(`${base}/api/addin/selection`, { method: 'POST', body });
    // The toolbar page.
    assert.match(await (await fetch(`${base}/toolbar`)).text(), /ocPopupOpen/);
    // Selection of two labels (seq 10).
    assert.equal((await select(`seq=10.5\nselect\nlabel:value:s0:c0\tlabel:value:s0:c1\n48,96,864,408\n50,100,860,400\nscreen=300,400,380,420\n${spec}`)).status, 204);
    await new Promise((r) => setTimeout(r, 50));
    assert.equal(shown.length, 1);
    assert.deepEqual(shown[0].arg, { id, roles: ['label:value:s0:c0', 'label:value:s0:c1'] });
    assert.deepEqual(shown[0].frame, { x: 300, y: 400 - 8 - TOOLBAR_SIZE.h, ...TOOLBAR_SIZE });
    assert.equal(getSession(id)?.state, 'open');
    // An older post that arrives late is dropped.
    await select(`seq=9\nselect\n\n`);
    await new Promise((r) => setTimeout(r, 50));
    assert.deepEqual(hidden, []);
    // The properties for the selection, then a change.
    const props = await (await fetch(`${base}/api/addin/properties?id=${id}&roles=label:value:s0:c0,label:value:s0:c1`)).json();
    assert.ok(props.controls.some((c: { prop: string }) => c.prop === 'size'));
    const r = await fetch(`${base}/api/addin/format`, { method: 'POST', body: JSON.stringify({ id, roles: ['label:value:s0:c0', 'label:value:s0:c1'], style: { size: 14 } }) });
    assert.equal(r.status, 200);
    assert.equal(applied.length, 1);
    assert.equal(applied[0].state, 'applied');
    assert.deepEqual((applied[0].chart as { overrides: unknown }).overrides, { 'label:value:s0:c0': { size: 14 }, 'label:value:s0:c1': { size: 14 } });
    const now = await (await fetch(`${base}/api/addin/properties?id=${id}&roles=label:value:s0:c0,label:value:s0:c1`)).json();
    assert.deepEqual(now.style, { size: 14 });
    // Deselect (newer seq): hidden.
    await select(`seq=11\nselect\n\n`);
    await new Promise((r) => setTimeout(r, 50));
    assert.deepEqual(hidden, ['toolbar']);
    // A change for a session that is gone: 404 with what to do.
    assert.equal((await fetch(`${base}/api/addin/format`, { method: 'POST', body: JSON.stringify({ id: 'nope', roles: ['x'], style: { fill: '#000000' } }) })).status, 404);
  } finally {
    server.close();
  }
});

test('the toolbar reads the engine\'s property table (src/core/overrides.ts)', async () => {
  const { propertyTable } = await import('../src/addin/toolbar.ts');
  const { ELEMENT_PROPERTIES } = await import('../src/core/overrides.ts');
  const t = await propertyTable();
  assert.equal(t.info, ELEMENT_PROPERTIES);
  const label = controlsFor(['label:value:s0:c1'], t).map((c) => c.prop);
  for (const p of ['size', 'rotate', 'fill', 'color', 'bold']) assert.ok(label.includes(p), p);
  const bar = controlsFor(['bar:s0:c1'], t).map((c) => c.prop);
  assert.ok(bar.includes('fill') && !bar.includes('size'));
});
