// The Edit data pop-up: placement maths, the request's screen lines, the per-platform openPopup, and
// the helper protocol (src/addin/popup.ts) against a stand-in helper.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { chmodSync, mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { chartRect, clampInto, edgeArgs, openPopup, parseWhere, placePopup, popupHelper, POPUP_SIZE, popupUrl, type PopupHelper, type Rect, type ScreenInfo } from '../src/addin/popup.ts';
import { parseRequest } from '../src/addin/engine.ts';

// A 1512×982 pt MacBook display (menu bar 38 pt) and a second display to its right.
const main: ScreenInfo = { frame: { x: 0, y: 0, w: 1512, h: 982 }, visible: { x: 0, y: 38, w: 1512, h: 944 }, scale: 2 };
const right: ScreenInfo = { frame: { x: 1512, y: 0, w: 1920, h: 1080 }, visible: { x: 1512, y: 25, w: 1920, h: 1055 }, scale: 1 };
const ppt: Rect = { x: 0, y: 38, w: 1512, h: 944 };
const within = (r: Rect, box: Rect) => r.x >= box.x - 0.01 && r.y >= box.y - 0.01 && r.x + r.w <= box.x + box.w + 0.01 && r.y + r.h <= box.y + box.h + 0.01;
const overlaps = (a: Rect, b: Rect) => a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h;

test('placement: right of the chart, top-aligned, when it fits', () => {
  const chart = { x: 300, y: 250, w: 500, h: 300 };
  const r = placePopup(chart, [ppt], [main]);
  assert.deepEqual(r, { x: 812, y: 250, ...POPUP_SIZE });
  assert.ok(!overlaps(r, chart));
});

test('placement: left of the chart when the right side has no room, then below, then above', () => {
  const wideRight = { x: 800, y: 200, w: 600, h: 300 };
  assert.deepEqual(placePopup(wideRight, [ppt], [main]), { x: 800 - 12 - POPUP_SIZE.w, y: 200, ...POPUP_SIZE });
  const wide = { x: 100, y: 100, w: 1300, h: 300 }; // no room left or right: below
  assert.deepEqual(placePopup(wide, [ppt], [main]), { x: 100, y: 412, ...POPUP_SIZE });
  const low = { x: 100, y: 500, w: 1300, h: 400 }; // nor below: above
  assert.deepEqual(placePopup(low, [ppt], [main]), { x: 100, y: 500 - 12 - POPUP_SIZE.h, ...POPUP_SIZE });
});

test('placement: a chart filling the display gets the pop-up inside the visible area, never under the menu bar', () => {
  const huge = { x: 20, y: 60, w: 1480, h: 900 };
  const r = placePopup(huge, [ppt], [main]);
  assert.ok(within(r, main.visible), JSON.stringify(r));
  const nearTop = { x: 700, y: 10, w: 200, h: 100 }; // top-aligned would sit under the menu bar
  assert.ok(placePopup(nearTop, [ppt], [main]).y >= 38);
});

test('placement: on the display the chart is on; shrunk to a small display', () => {
  const chart = { x: 1700, y: 300, w: 400, h: 250 };
  const r = placePopup(chart, [], [main, right]);
  assert.ok(within(r, right.visible));
  assert.equal(r.x, 2112);
  const tiny: ScreenInfo = { frame: { x: 0, y: 0, w: 500, h: 300 }, visible: { x: 0, y: 0, w: 500, h: 300 } };
  const t = placePopup({ x: 10, y: 10, w: 100, h: 50 }, [], [tiny]);
  assert.ok(within(t, tiny.visible));
  assert.deepEqual([t.w, t.h], [500, 300]);
});

test('placement without a chart: inside PowerPoint\'s front window at the right, below the ribbon; else centred', () => {
  const win = { x: 100, y: 60, w: 1200, h: 800 };
  assert.deepEqual(placePopup(undefined, [win], [main]), { x: 100 + 1200 - POPUP_SIZE.w - 24, y: 60 + 170, ...POPUP_SIZE });
  const c = placePopup(undefined, [], [main]);
  assert.deepEqual([c.x, c.y], [(1512 - POPUP_SIZE.w) / 2, 38 + (944 - POPUP_SIZE.h) / 2]);
  assert.ok(within(placePopup(undefined, [], []), { x: 0, y: 0, w: 1440, h: 900 }));
});

test('chartRect: trusts PointsToScreenPixels only inside a PowerPoint window, at scale 1 or the display\'s', () => {
  // Points (scale 1).
  assert.deepEqual(chartRect({ screen: [300, 250, 800, 550] }, [ppt], [main]), { x: 300, y: 250, w: 500, h: 300 });
  // Retina pixels: only half of it lands in the window.
  const win = { x: 0, y: 38, w: 1000, h: 700 };
  assert.deepEqual(chartRect({ screen: [1200, 600, 1800, 1000] }, [win], [main]), { x: 600, y: 300, w: 300, h: 200 });
  // Nonsense (outside every window at any scale), empty, inverted or missing: not trusted.
  assert.equal(chartRect({ screen: [9000, 9000, 9500, 9300] }, [win], [main]), undefined);
  assert.equal(chartRect({ screen: [10, 10, 10, 10] }, [win], [main]), undefined);
  assert.equal(chartRect({ screen: [500, 500, 100, 100] }, [win], [main]), undefined);
  assert.equal(chartRect({ window: [0, 0, 100, 100, 100] }, [win], [main]), undefined);
  assert.equal(chartRect(undefined, [win], [main]), undefined);
  // No window list (the helper could not tell): taken as given.
  assert.deepEqual(chartRect({ screen: [1, 2, 3, 4] }, [], []), { x: 1, y: 2, w: 2, h: 2 });
});

test('clampInto moves a rectangle inside and shrinks it only when it is bigger', () => {
  assert.deepEqual(clampInto({ x: -50, y: 900, w: 100, h: 100 }, { x: 0, y: 0, w: 1000, h: 950 }), { x: 0, y: 850, w: 100, h: 100 });
  assert.deepEqual(clampInto({ x: 0, y: 0, w: 2000, h: 50 }, { x: 10, y: 10, w: 1000, h: 950 }), { x: 10, y: 10, w: 1000, h: 50 });
});

test('Edit data request: screen= and window= are read, and are not part of the layout context', () => {
  const req = parseRequest('edit\n\n48,96,864,408\n50,100,860,400\ntheme=000000,FFFFFF\nscreen=310.5,250,810,550\nwindow=0,38,1512,944,100\n{"type":"clustered"}');
  assert.deepEqual(req.where, { screen: [310.5, 250, 810, 550], window: [0, 38, 1512, 944, 100] });
  assert.equal(req.context, '48,96,864,408\n50,100,860,400\ntheme=000000,FFFFFF\n');
  assert.equal(req.spec, '{"type":"clustered"}');
  assert.equal(parseRequest('edit\n\n1,2,3,4\n1,2,3,4\nscreen=a,b\n{}').where, undefined);
  assert.equal(parseWhere({ window: '1,2,3,4' }), undefined);
});

test('Windows: Edge app mode with size, position, its own profile; browser when Edge is missing', async () => {
  const args = edgeArgs('http://localhost:4174/?session=x&popup=1', { x: 812.4, y: 250, w: 620, h: 360 }, 'C:\\p');
  assert.deepEqual(args.slice(0, 4), ['--app=http://localhost:4174/?session=x&popup=1', '--window-size=620,360', '--window-position=812,250', '--user-data-dir=C:\\p']);
  const launched: string[][] = [], opened: string[] = [];
  const env = { env: { ProgramFiles: 'C:\\Program Files' }, platform: 'win32' as const, launch: (c: string, a: string[]) => { launched.push([c, ...a]); return true; }, browser: (u: string) => { opened.push(u); }, profile: 'C:\\p' };
  const how = await openPopup('http://localhost:4174', 'abc', { screen: [300, 250, 800, 550] }, { ...env, exists: (p) => p.endsWith('msedge.exe') });
  assert.equal(how, 'edge');
  assert.match(launched[0][0], /msedge\.exe$/);
  assert.ok(launched[0].includes('--window-position=812,250'));
  assert.ok(launched[0].includes(`--app=${popupUrl('http://localhost:4174', 'abc')}`));
  assert.equal(await openPopup('http://localhost:4174', 'abc', undefined, { ...env, exists: () => false }), 'browser');
  assert.deepEqual(opened, ['http://localhost:4174/?session=abc']);
});

test('Mac: the helper shows the pop-up; when it cannot, the editor opens in the browser as before', async () => {
  const opened: string[] = [], logged: string[] = [];
  const ok = { show: async () => ({ event: 'shown', ms: 40, ready: true }) } as unknown as PopupHelper;
  assert.equal(await openPopup('http://b', 'i1', undefined, { platform: 'darwin', helper: ok, browser: (u) => opened.push(u), log: (l) => logged.push(l) }), 'helper');
  assert.match(logged[0], /^POPUP shown in \d+ ms \(page ready after 40 ms\)$/);
  const dead = { show: async () => undefined } as unknown as PopupHelper;
  assert.equal(await openPopup('http://b', 'i2', undefined, { platform: 'darwin', helper: dead, browser: (u) => opened.push(u), log: (l) => logged.push(l) }), 'browser');
  assert.deepEqual(opened, ['http://b/?session=i2']);
  const none = popupHelper('http://b', undefined);
  assert.equal(none.warm(), false);
  assert.equal(await none.show('http://b', 'x'), undefined);
});

test('Mac helper protocol: preload, info, show at the placed frame, eval, and exit with the server', async () => {
  // A stand-in helper: answers like addin/popup-mac/main.swift and records what it was sent.
  const dir = mkdtempSync(join(tmpdir(), 'oc-popup-'));
  const exe = join(dir, 'helper');
  writeFileSync(exe, `#!${process.execPath}
const rl = require('node:readline').createInterface({ input: process.stdin });
const out = (m) => process.stdout.write(JSON.stringify(m) + '\\n');
const seen = [];
rl.on('line', (l) => {
  const m = JSON.parse(l); seen.push(m);
  if (m.cmd === 'info') out({ event: 'info', screens: [{ frame: [0, 0, 1512, 982], visible: [0, 38, 1512, 944], scale: 2 }], windows: [[0, 38, 1512, 944]] });
  if (m.cmd === 'show') out({ event: 'shown', panel: m.panel, ms: 12, ready: true });
  if (m.cmd === 'eval') out({ event: 'eval', result: seen });
});
rl.on('close', () => process.exit(0));
`);
  chmodSync(exe, 0o755);
  const h = popupHelper('http://127.0.0.1:4174', exe);
  assert.equal(h.warm(), true);
  const info = await h.info();
  assert.deepEqual(info?.windows, [ppt]);
  const shown = await h.show('http://127.0.0.1:4174', 'sess1', { screen: [300, 250, 800, 550] });
  assert.equal(shown?.ready, true);
  const seen = (await h.eval('x'))?.result as Array<Record<string, unknown>>;
  assert.deepEqual(seen[0], { cmd: 'load', panel: 'datasheet', url: 'http://127.0.0.1:4174/?popup=1' });
  const show = seen.find((m) => m.cmd === 'show')!;
  assert.deepEqual(show, { cmd: 'show', panel: 'datasheet', id: 'sess1', url: 'http://127.0.0.1:4174/?session=sess1&popup=1', frame: [812, 250, POPUP_SIZE.w, POPUP_SIZE.h], focus: true });
  h.stop();
  await new Promise((r) => setTimeout(r, 100));
  assert.equal(h.running, false);
});
