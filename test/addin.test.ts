import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { installLauncher, launcherPath, reinstallCommand, removeLauncher } from '../src/addin/launcher.ts';
import { createMeasurer } from '../src/fonts.ts';
import { flattenPath } from '../src/core/path.ts';
import { escapeField, parseResponse, shapeRecord, SHAPE_COLUMNS, unescapeField } from '../src/addin/format.ts';
import { contextOf, DEFAULT_FRAME, handleRequest, layoutForApply, parseRequest, resolveFrame, sceneBounds, specKey } from '../src/addin/engine.ts';
import { addinRoutes, applyErrorMessage, applyInPowerPoint, editorPage, outlivedPowerPoint, parseApplyAnswer, RESTARTED_MESSAGE } from '../src/addin/server.ts';
import { HELPER_JXA, macroRunner } from '../src/addin/runner.ts';
import { getSession, setSession, type Session } from '../src/server.ts';
import type { Sessions } from '../src/addin/engine.ts';
import { buildAddin, ADDIN_CT, VBA_REL, UI_REL, UI_NS, IMAGE_REL } from '../src/addin/package.ts';
import { galleryItems, insertableItems, sampleFor } from '../src/addin/registry.ts';
import { unzip } from '../src/ooxml/unzip.ts';
import { readCfb } from '../src/addin/cfb.ts';
import { decompress } from '../src/addin/ovba.ts';
import { layoutChart, type Shape } from '../src/core/index.ts';

const m = createMeasurer();
const req = (op: string, arg = '', frame = '', bounds = '', spec = '') => parseRequest([op, arg, frame, bounds, spec].join('\n'));

test('field escaping round-trips tabs, newlines, CR and backslashes', async () => {
  for (const s of ['plain', 'a\tb', 'two\nlines', 'cr\r\n', 'back\\slash', '\\t literal', '']) {
    const e = escapeField(s);
    assert.ok(!/[\t\n\r]/.test(e));
    assert.equal(unescapeField(e), s);
  }
});

test('shape records carry every property VBA needs', async () => {
  const rect = parseResponse(shapeRecord({ kind: 'rect', x: 1.234, y: 2, w: 3, h: 4, fill: '#1f3a5f', stroke: 'FFFFFF', strokeWidth: 0.5 }, 'Arial', 'bar:s0:c0', 'Bar s0 c0')).shapes[0];
  assert.deepEqual([rect.kind, rect.x, rect.fill, rect.lineColor, rect.lineWidth, rect.role, rect.name], ['rect', '1.23', '1F3A5F', 'FFFFFF', '0.5', 'bar:s0:c0', 'Bar s0 c0']);
  const ell = parseResponse(shapeRecord({ kind: 'rect', x: 0, y: 0, w: 4, h: 4, fill: '000000', ellipse: true }, 'Arial')).shapes[0];
  assert.equal(ell.kind, 'ellipse');
  assert.equal(ell.lineColor, '');
  const geom = parseResponse(shapeRecord({ kind: 'rect', x: 0, y: 0, w: 4, h: 4, fill: '000000', geom: 'diamond' } as Shape, 'Arial')).shapes[0];
  assert.equal(geom.kind, 'diamond');
  const line = parseResponse(shapeRecord({ kind: 'line', x1: 10, y1: 20, x2: 5, y2: 40, color: '7F7F7F', width: 1, dash: true, arrowEnd: true }, 'Arial')).shapes[0];
  assert.deepEqual([line.kind, line.x, line.y, line.w, line.h, line.dash, line.arrowEnd], ['line', '10', '20', '5', '40', '1', '1']);
  const text = parseResponse(shapeRecord({ kind: 'text', x: 0, y: 0, w: 50, h: 14, text: 'CAGR\t+12%\nline 2', size: 10, color: '222222', bold: true, align: 'center', valign: 'middle', fill: 'FFFFFF', stroke: '000000', rounded: true }, 'Arial')).shapes[0];
  assert.deepEqual([text.text, text.font, text.size, text.bold, text.align, text.valign, text.wrap, text.autosize, text.textGeom], ['CAGR\t+12%\nline 2', 'Arial', '10', '1', 'c', 'm', '0', '0', 'roundRect']);
  assert.equal(Object.keys(text).length, SHAPE_COLUMNS.length);
  assert.throws(() => shapeRecord({ kind: 'blob' } as unknown as Shape, 'Arial'), /cannot draw shape kind "blob"/);
});

test('path shapes flatten arcs into polygon points', async () => {
  const cmds = [{ op: 'M', x: 100, y: 100 }, { op: 'A', cx: 100, cy: 100, r: 50, start: -90, sweep: 90 }, { op: 'Z' }] as const;
  const subs = flattenPath([...cmds] as never);
  assert.equal(subs.length, 1);
  assert.equal(subs[0].closed, true);
  assert.deepEqual(subs[0].points[0], { x: 100, y: 100 });
  const last = subs[0].points.at(-1)!;
  assert.ok(Math.abs(last.x - 150) < 1e-9 && Math.abs(last.y - 100) < 1e-9, 'arc ends at 3 o\'clock');
  const rec = parseResponse(shapeRecord({ kind: 'path', commands: [...cmds], fill: 'FF0000' } as unknown as Shape, 'Arial')).shapes[0];
  assert.equal(rec.kind, 'path');
  assert.match(rec.points, /^100,100 100,50 .* 150,100 z$/);
  assert.deepEqual([rec.x, rec.y, rec.w, rec.h], ['100', '50', '50', '50']);
  const ring = flattenPath([{ op: 'A', cx: 0, cy: 0, r: 10, start: 0, sweep: 360 }, { op: 'Z' }, { op: 'A', cx: 0, cy: 0, r: 5, start: 0, sweep: -360 }, { op: 'Z' }] as never);
  assert.equal(ring.length, 2);
});

test('insert answers with group, tags and one shape per scene shape, roles and names', async () => {
  const r = parseResponse(await handleRequest(req('insert', 'stacked', '48,96,864,408'), m));
  assert.deepEqual(r.errors, []);
  assert.equal(r.version, '2');
  assert.equal(r.group?.name, 'Stacked column chart');
  const tags = Object.fromEntries(r.tags);
  assert.equal(tags.OC_V, '1');
  assert.equal(tags.OC_FRAME, '48,96,864,408');
  assert.equal(tags.OC_FONT, 'Arial');
  assert.equal(JSON.parse(tags.OC_SPEC).type, 'stacked');
  const sample = sampleFor('stacked');
  assert.ok('spec' in sample);
  const scene = layoutChart(sample.spec, { x: 48, y: 96, w: 864, h: 408 }, m);
  assert.equal(r.shapes.length, scene.shapes.length);
  assert.ok(r.shapes.every((s) => s.role && s.name));
  assert.equal(r.shapes[0].name, 'Bar s0 c0');
});

test('every insertable gallery item lays out; others explain themselves', async () => {
  const items = insertableItems();
  assert.ok(items.length >= 6);
  for (const id of items) assert.deepEqual(parseResponse(await handleRequest(req('insert', id, '0,0,600,400'), m)).errors, [], id);
  const missing = galleryItems().find((g) => !items.includes(g.id));
  if (missing) assert.match(parseResponse(await handleRequest(req('insert', missing.id), m)).errors[0], /not available in this version/);
  assert.match(parseResponse(await handleRequest(req('insert', 'nope'), m)).errors[0], /not available/);
  assert.match(parseResponse(await handleRequest(req('bogus'), m)).errors[0], /do not match/);
});

test('gallery: the reference tiles in order; "<item>@<direction>" inserts that way, a direction the type lacks is ignored', async () => {
  assert.deepEqual(galleryItems().slice(0, 16).map((g) => g.id), ['stacked', 'percent', 'clustered', 'waterfall', 'waterfallDown', 'mekko', 'combo', 'line', 'area', 'area100', 'mekkoUnit', 'pie', 'doughnut', 'scatter', 'bubble', 'gantt']);
  const spec = (item: string) => { const r = sampleFor(item); assert.ok('spec' in r, item); return r.spec; };
  assert.equal(spec('stacked@left').options?.direction, 'left');
  assert.equal(spec('line@right').options?.direction, 'right');
  assert.equal(spec('line@down').options?.direction, undefined, 'lines never hang');
  assert.equal(spec('pie@left').options?.direction, undefined);
  assert.equal(spec('mekkoUnit').options?.mekko, 'unit');
  // Build-down starts with the total and walks down to the result; build-up stacks components to a total.
  const down = spec('waterfallDown').series[0].values, up = spec('waterfall').series[0].values;
  assert.ok((down[0] as number) > 0 && down.slice(1, -1).every((v) => (v as number) < 0) && down.at(-1) === 'e');
  assert.ok(up.slice(0, -1).every((v) => (v as number) > 0) && up.at(-1) === 'e');
  const r = parseResponse(await handleRequest(req('insert', 'waterfallDown@right', '0,0,600,400'), m));
  assert.deepEqual(r.errors, []);
  assert.match(Object.fromEntries(r.tags).OC_SPEC, /"direction":"right"/);
});

test('direction request: the chart is laid out again in its place, growing the new way', async () => {
  const insert = parseResponse(await handleRequest(req('insert', 'stacked', '10,20,600,400'), m));
  const tags = Object.fromEntries(insert.tags);
  const left = parseResponse(await handleRequest(req('direction', 'left', tags.OC_FRAME, '', tags.OC_SPEC), m));
  assert.deepEqual(left.errors, []);
  const lt = Object.fromEntries(left.tags);
  assert.equal(lt.OC_FRAME, tags.OC_FRAME, 'same frame');
  assert.match(lt.OC_SPEC, /"direction":"left"/);
  const bars = left.shapes.filter((s) => s.role.startsWith('bar:'));
  const row = bars.filter((b) => b.role.endsWith(':c0'));
  assert.ok(row.length > 1 && row.every((b) => b.y === row[0].y && b.h === row[0].h), 'one category = one row of segments side by side');
  assert.ok(+row[1].x < +row[0].x, 'stacked leftwards');
  // Back up: the direction option goes away again (up is the default).
  const up = parseResponse(await handleRequest(req('direction', 'up', lt.OC_FRAME, '', lt.OC_SPEC), m));
  assert.doesNotMatch(Object.fromEntries(up.tags).OC_SPEC, /direction/);
  // Old specs (orientation) are migrated first; a direction the type lacks is explained.
  const old = JSON.stringify({ specVersion: 1, type: 'clustered', categories: ['a', 'b'], series: [{ name: 's', values: [1, 2] }], options: { orientation: 'horizontal' } });
  assert.match(Object.fromEntries(parseResponse(await handleRequest(req('direction', 'down', '0,0,500,300', '', old), m)).tags).OC_SPEC, /"direction":"down"/);
  const line = JSON.stringify({ specVersion: 1, type: 'line', categories: ['a', 'b'], series: [{ name: 's', values: [1, 2] }] });
  assert.match(parseResponse(await handleRequest(req('direction', 'down', '0,0,500,300', '', line), m)).errors[0], /Line charts can grow up, left, right, not down/);
  const pie = JSON.stringify({ specVersion: 1, type: 'pie', categories: ['a', 'b'], series: [{ name: 's', values: [1, 2] }] });
  assert.match(parseResponse(await handleRequest(req('direction', 'left', '0,0,500,300', '', pie), m)).errors[0], /no direction/);
  assert.match(parseResponse(await handleRequest(req('direction', 'left'), m)).errors[0], /Select an open-charts chart/);
});

test('large specs are chunked into OC_SPEC_0..n + OC_SPEC_N', async () => {
  const cats = Array.from({ length: 400 }, (_, i) => `Category with a fairly long name, padded so the spec passes the tag chunk size, number ${i}`);
  const spec = { specVersion: 1, type: 'line', categories: cats, series: [{ name: 'A', values: cats.map((_, i) => i) }] };
  const r = parseResponse(await handleRequest(req('update', '', '0,0,900,500', '', JSON.stringify(spec)), m));
  assert.deepEqual(r.errors, []);
  const tags = Object.fromEntries(r.tags);
  assert.ok(Number(tags.OC_SPEC_N) >= 1 && tags.OC_SPEC === undefined);
  const joined = Array.from({ length: Number(tags.OC_SPEC_N) }, (_, i) => tags[`OC_SPEC_${i}`]).join('');
  assert.equal(JSON.parse(joined).categories.length, 400);
});

test('update re-lays out in the stored frame, or follows a move/resize', async () => {
  const s = sampleFor('waterfall');
  assert.ok('spec' in s);
  const f0 = { x: 48, y: 96, w: 864, h: 408 };
  const b0 = sceneBounds(layoutChart(s.spec, f0, m))!;
  assert.deepEqual(resolveFrame(s.spec, undefined, f0, b0, m), f0);
  const moved = { x: b0.x + 100, y: b0.y + 20, w: b0.w / 2, h: b0.h / 2 };
  const f1 = resolveFrame(s.spec, undefined, f0, moved, m);
  assert.ok(Math.abs(f1.w - f0.w / 2) < 1e-6 && Math.abs(f1.h - f0.h / 2) < 1e-6);
  const b1 = sceneBounds(layoutChart(s.spec, f1, m))!;
  assert.ok(Math.abs(b1.x - moved.x) < 20 && Math.abs(b1.y - moved.y) < 20, 'chart lands where the user put it');
  const r = parseResponse(await handleRequest(req('update', '', '48,96,864,408', '', JSON.stringify(s.spec)), m));
  assert.deepEqual(r.errors, []);
  assert.match(parseResponse(await handleRequest(req('update'), m)).errors[0], /Select an open-charts chart/);
  assert.match(parseResponse(await handleRequest(req('update', '', '', '', '{broken'), m)).errors[0], /damaged/);
  assert.match(parseResponse(await handleRequest(req('update', '', '', '', '{"type":"stacked","categories":[],"series":[]}'), m)).errors[0], /categor/i);
  assert.equal(DEFAULT_FRAME.w, 864);
});

/** In-memory stand-in for the add-in server's sessions. */
function fakeSessions() {
  const store = new Map<string, Session>();
  const opened: string[] = [];
  const sessions: Sessions = {
    async get(id) { return store.get(id); },
    async open(id, chart, title) {
      store.set(id, { id, rev: 1, state: 'open', chart, title, updatedAt: '' });
      opened.push(id);
    },
    async settle(id) { store.set(id, { ...store.get(id)!, rev: store.get(id)!.rev + 1, state: 'open' }); },
  };
  const apply = (id: string, chart: object, state: Session['state'] = 'applied') => store.set(id, { ...store.get(id)!, rev: store.get(id)!.rev + 1, state, chart });
  return { sessions, opened, apply };
}

test('edit data opens a session; the applied chart is used by update of that chart', async () => {
  const { sessions, opened, apply } = fakeSessions();
  const s = sampleFor('clustered');
  assert.ok('spec' in s);
  const json = JSON.stringify(s.spec);
  const id = specKey(json);
  const r0 = await handleRequest(req('edit', '', '48,96,864,408', '', json), m, sessions);
  assert.match(r0, new RegExp(`^I\t${id}$`, 'm'), 'answers the session id');
  assert.deepEqual(opened, [id]);

  // Before Apply: update keeps the chart's own data; relayout has nothing to do.
  const r1 = parseResponse(await handleRequest(req('update', id, '48,96,864,408', '', json), m, sessions));
  assert.equal(JSON.parse(Object.fromEntries(r1.tags).OC_SPEC).categories[0], 'North America');
  assert.equal(parseResponse(await handleRequest(req('relayout', '', '48,96,864,408', '48,96,864,408', json), m, sessions)).nothing, true);

  const cats = async (op: string, arg: string) =>
    JSON.parse(Object.fromEntries(parseResponse(await handleRequest(req(op, arg, '48,96,864,408', '48,96,864,408', json), m, sessions)).tags).OC_SPEC ?? '{}').categories;
  for (const arg of [id, '']) { // by session id (ApplyEdit) and by the chart's own spec (Update chart)
    apply(id, { ...s.spec, categories: ['A', 'B', 'C', 'D'] });
    assert.deepEqual(await cats('update', arg), ['A', 'B', 'C', 'D']);
    // Drawn once: another chart that still has the old data (a fresh sample, a copy) keeps it.
    assert.equal((await sessions.get(id))!.state, 'open');
    assert.equal((await cats('update', ''))[0], 'North America');
  }
  // The automatic check also picks up an applied edit.
  apply(id, { ...s.spec, categories: ['A', 'B', 'C', 'D'] });
  assert.equal(parseResponse(await handleRequest(req('relayout', '', '48,96,864,408', '48,96,864,408', json), m, sessions)).nothing, false);
  // An Apply that did not reach the slide: Update chart (clicked on that chart) takes it, but the
  // automatic re-layout of a same-data chart does not.
  apply(id, { ...s.spec, categories: ['E', 'F', 'G', 'H'] }, 'error');
  assert.equal(parseResponse(await handleRequest(req('relayout', '', '48,96,864,408', '48,96,864,408', json), m, sessions)).nothing, true);
  assert.deepEqual(await cats('update', ''), ['E', 'F', 'G', 'H']);
  assert.equal((await sessions.get(id))!.state, 'open');
});

test('relayout: nothing to do in place, a new frame after a resize', async () => {
  const s = sampleFor('line');
  assert.ok('spec' in s);
  const json = JSON.stringify(s.spec);
  const same = parseResponse(await handleRequest(req('relayout', '', '48,96,864,408', '48,96,864,408', json), m));
  assert.equal(same.nothing, true);
  const half = parseResponse(await handleRequest(req('relayout', '', '48,96,864,408', '48,96,432,204', json), m));
  assert.equal(half.nothing, false);
  assert.equal(half.frame, '48,96,432,204', 'the group bounds are the frame (the add-in draws a frame rectangle)');
  assert.equal(Object.fromEntries(half.tags).OC_FRAME, '48,96,432,204');
});

test('the editor page for a session: preloaded only for an editor without session support', () => {
  const html = readFileSync(new URL('../src/web/editor.html', import.meta.url), 'utf8');
  const page = editorPage('0123456789abcdef', { type: 'line', categories: ['a'], series: [{ name: 's', values: [1] }] });
  if (html.includes('/api/session/')) assert.equal(page, undefined);
  else {
    assert.match(page!, /\/api\/session\/' \+ id/);
    assert.ok(page!.indexOf('/api/session/') < page!.indexOf('<script', page!.indexOf('/api/session/')), 'boot script runs first');
  }
});

test('ppam package: add-in content type, VBA and ribbon parts, one callback prefix', async () => {
  const built = buildAddin();
  const e = unzip(built.ppam);
  const ct = e.get('[Content_Types].xml')!.text();
  assert.ok(ct.includes(`PartName="/ppt/presentation.xml" ContentType="${ADDIN_CT}"`));
  assert.match(ct, /Extension="bin" ContentType="application\/vnd\.ms-office\.vbaProject"/);
  assert.ok(e.get('_rels/.rels')!.text().includes(`Type="${UI_REL}" Target="customUI/customUI14.xml"`));
  assert.ok(e.get('ppt/_rels/presentation.xml.rels')!.text().includes(`Type="${VBA_REL}" Target="vbaProject.bin"`));
  assert.ok(!e.get('ppt/presentation.xml')!.text().includes('sldIdLst'), 'an add-in has no slides');

  const ui = e.get('customUI/customUI14.xml')!.text();
  assert.ok(ui.includes(`xmlns="${UI_NS}"`));
  const callbacks = [...ui.matchAll(/(?:onAction|onLoad|getEnabled|getPressed)="([^"]+)"/g)].map((x) => x[1]);
  assert.ok(callbacks.length >= 4 && callbacks.every((c) => c.startsWith(`${built.prefix}_`)));
  for (const g of galleryItems()) assert.ok(ui.includes(`id="chart_${g.id}"`), g.id);
  assert.match(ui, /id="btnEditData"/);
  assert.match(ui, /id="btnUpdate"/);

  const vba = readCfb(e.get('ppt/vbaProject.bin')!.bytes()) as { VBA: Record<string, Uint8Array>; PROJECT: Uint8Array };
  const source = new TextDecoder().decode(decompress(vba.VBA.OpenCharts));
  for (const c of new Set(callbacks)) assert.ok(source.includes(`Sub ${c}(`), `VBA defines ${c}`);
  assert.ok(!source.includes('OCPREFIX'));
  const project = new TextDecoder().decode(vba.PROJECT);
  assert.match(project, /Name="OpenChartsAddIn"/);
  assert.match(project, /Class=AppEvents/);
  const cls = new TextDecoder().decode(decompress(vba.VBA.AppEvents));
  assert.match(cls, /^Attribute VB_Name = "AppEvents"\r\nAttribute VB_Base = "0\{FCFB3D2A/);
  assert.match(cls, /WithEvents App As Application/);
  // Gallery tiles show an embedded icon: every image id has a relationship to a PNG in the package.
  const rels = e.get('customUI/_rels/customUI14.xml.rels')!.text();
  for (const [, id] of ui.matchAll(/ image="([^"]+)"/g)) {
    const target = new RegExp(`Id="${id}" Type="${IMAGE_REL.replace(/\//g, '\\/')}" Target="([^"]+)"`).exec(rels)?.[1];
    assert.ok(target, `relationship for ${id}`);
    const png = e.get(`customUI/${target}`)!.bytes();
    assert.deepEqual([...png.slice(1, 4)], [0x50, 0x4e, 0x47], `${target} is a PNG`);
  }
  assert.match(ct, /Extension="png" ContentType="image\/png"/);
  assert.match(ui, /<gallery id="galCharts"[^>]* columns="5"[^>]* showItemLabel="true"/);
  for (const d of ['dirUp', 'dirDown', 'dirLeft', 'dirRight']) assert.match(ui, new RegExp(`<toggleButton id="${d}"[^>]* onAction="${built.prefix}_onDirection"`));
  assert.match(source, /^Private Const DIRECTIONS As String = "stacked:up,down,left,right;.*line:up,left,right/m);
  assert.deepEqual(buildAddin().ppam, built.ppam, 'reproducible build');
});

test('VBA source is ASCII and every Sub/Function is closed', async () => {
  const src = ['OpenCharts.bas', 'AppEvents.cls'].map((f) => readFileSync(new URL(`../addin/vba/${f}`, import.meta.url), 'utf8')).join('\n');
  assert.ok(/^[\x09\x0a\x0d\x20-\x7e]*$/.test(src));
  const opens = (src.match(/^\s*(Public |Private )?(Sub|Function) /gm) ?? []).length;
  const closes = (src.match(/^\s*End (Sub|Function)\b/gm) ?? []).length;
  assert.equal(opens, closes);
});

test('VBA: Apply is silent, an ungrouped chart is named as such, the engine starts as the add-in loads', () => {
  const src = readFileSync(new URL('../addin/vba/OpenCharts.bas', import.meta.url), 'utf8');
  const body = (name: string) => { const a = src.search(new RegExp(`(Sub|Function) ${name}\\(`)); return src.slice(a, src.slice(a).search(/End (Sub|Function)/) + a); };
  // ApplyEdit runs under the server's AppleScript call: a dialog would block it until it times out.
  assert.match(body('OCPREFIX_ApplyEdit'), /gSilent = True/);
  assert.doesNotMatch(body('OCPREFIX_ApplyEdit'), /MsgBox|Tell /);
  assert.match(body('OCPREFIX_ApplyEdit'), /On Error GoTo Failed/);
  // O3: pieces of an ungrouped chart keep OC_ROLE; say so instead of "select a chart".
  assert.match(src, /This chart was ungrouped, so it can't be edited as a chart\. Undo \(" & ChrW\(&H2318\) & "Z\) the ungroup or insert it again\./);
  assert.match(body('OCPREFIX_getEnabled'), /UngroupedSelected\(\)/, 'the buttons stay enabled so the click can explain');
  // O5: the warm server is started from onLoad through the bridge's startServer handler.
  assert.match(body('OCPREFIX_onLoad'), /StartEngine/);
  assert.doesNotMatch(body('Auto_Open'), /StartEngine/, 'not while PowerPoint launches (breaks register add in)');
  assert.match(src, /AppleScriptTask\(SCRIPT_FILE, "startServer", ""\)/);
  const bridge = readFileSync(new URL('../addin/applescript/open-charts.applescript', import.meta.url), 'utf8');
  assert.match(bridge, /^on startServer\(ignored\)$/m);
  assert.match(bridge, /\{ \/usr\/bin\/curl -sf [^}]*addin\/ping \|\| " & serverCommand & "; \} > \/dev\/null 2>&1 &/, 'backgrounded, output detached');
});

test('request parsing tolerates CRLF and missing fields', async () => {
  const r = parseRequest('update\r\n\r\n1,2,3,4\r\n5,6,7,8\r\n{"a":1}\r\n');
  assert.deepEqual([r.op, r.arg, r.frame, r.bounds, r.spec], ['update', '', { x: 1, y: 2, w: 3, h: 4 }, { x: 5, y: 6, w: 7, h: 8 }, '{"a":1}']);
  assert.deepEqual(parseRequest(''), { op: '', arg: '', frame: undefined, bounds: undefined, spec: '' });
});

test('launcher: installed open-charts command runs the CLI; messages name a command that works', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'oc-launcher-'));
  const path = process.env.PATH;
  try {
    const cli = join(dir, 'fake cli.js');
    writeFileSync(cli, 'console.log(JSON.stringify(process.argv.slice(2)))\n');
    process.env.PATH = ''; // no folder to link into: the install explains how to add one
    const lines = installLauncher(process.execPath, cli, dir);
    assert.match(lines.join('\n'), /add its folder to your PATH/);
    const bin = launcherPath(dir);
    assert.equal(statSync(bin).mode & 0o111, 0o111, 'executable');
    process.env.PATH = path;
    assert.equal(execFileSync(bin, ['addin', 'install', "it's"], { encoding: 'utf8' }).trim(), JSON.stringify(['addin', 'install', "it's"]));
    // The command shown to users is that launcher, quoted for a shell.
    assert.equal(execFileSync('/bin/sh', ['-c', `${reinstallCommand(dir).replace(' addin install', '')} hello`], { encoding: 'utf8' }).trim(), '["hello"]');
    removeLauncher(dir);
    assert.throws(() => statSync(bin));
  } finally {
    process.env.PATH = path;
    rmSync(dir, { recursive: true, force: true });
  }
  // Every message that tells the user to reinstall shows reinstallCommand(), never a bare `open-charts`.
  const mismatch = parseResponse(await handleRequest(req('bogus'), m)).errors[0];
  assert.ok(mismatch.includes(`In Terminal, run: ${reinstallCommand()}`), mismatch);
  const vba = buildAddin().source;
  assert.ok(vba.includes(`Private Const REINSTALL As String = "${reinstallCommand().replace(/"/g, '""')}"`));
  assert.match(vba, /"Open Terminal and run:" & vbLf & REINSTALL/);
  assert.doesNotMatch(vba, /run:\s+open-charts addin/);
});

test('Apply in PowerPoint records the outcome on the session: drawn, or an error that says what to do', async () => {
  const home = mkdtempSync(join(tmpdir(), 'oc-apply-'));
  const chart = { type: 'stacked', categories: ['a'], series: [{ name: 'S', values: [1] }] };
  const macros: string[] = [];
  const argv: string[][] = [];
  try {
    // Not installed: no AppleScript, a clear error.
    let s = setSession('ap1', { chart, state: 'applied' });
    await applyInPowerPoint(s, home, async () => assert.fail('must not run'));
    assert.equal(getSession('ap1')!.state, 'error');
    assert.match(getSession('ap1')!.message!, /not installed.*addin install/);

    writeFileSync(join(home, 'addin.json'), JSON.stringify({ prefix: 'oc123' }));
    // The macro draws the chart: the engine settles the session ("open"), and once the macro has
    // returned the server says "drawn", never earlier, while PowerPoint is still drawing.
    s = setSession('ap1', { chart, state: 'applied' });
    await applyInPowerPoint(s, home, async (macro, args) => {
      macros.push(macro); argv.push(args);
      setSession('ap1', { state: 'open' });
      return { out: '' };
    });
    assert.deepEqual([getSession('ap1')!.state, getSession('ap1')!.message], ['drawn', undefined]);
    assert.deepEqual(macros, ['oc123_ApplyEdit']);
    assert.deepEqual(argv[0], ['ap1', '', ''], 'no layout of its own without prepare');

    // The macro ran but drew nothing (no chart found).
    s = setSession('ap1', { chart, state: 'applied' });
    await applyInPowerPoint(s, home, async () => ({ out: '' }));
    assert.equal(getSession('ap1')!.state, 'error');
    assert.match(getSession('ap1')!.message!, /could not find the chart/);

    // Permission prompt unanswered (killed after the timeout), denied, PowerPoint closed, other.
    const cases: [{ message: string; killed?: boolean }, RegExp][] = [
      [{ message: 'Command failed', killed: true }, /click Allow in that dialog, then Apply again/],
      [{ message: 'execution error: Microsoft PowerPoint got an error: AppleEvent timed out. (-1712)' }, /click Allow/],
      [{ message: 'execution error: Not authorized to send Apple events to Microsoft PowerPoint. (-1743)' }, /System Settings › Privacy & Security › Automation/],
      [{ message: 'execution error: PowerPoint is not running. (-600)' }, /PowerPoint isn't running/],
      [{ message: 'Command failed: osascript\n0:10: execution error: Something odd. (-2700)' }, /could not update the chart \(Something odd\. \(-2700\)\)/],
    ];
    for (const [err, re] of cases) {
      assert.match(applyErrorMessage(err), re);
      s = setSession('ap1', { chart, state: 'applied' });
      await applyInPowerPoint(s, home, async () => ({ error: err }));
      assert.equal(getSession('ap1')!.state, 'error');
      assert.match(getSession('ap1')!.message!, re);
    }

    // Laid out, then the call failed (e.g. timed out while drawing): error, not drawn.
    s = setSession('ap1', { chart, state: 'applied' });
    await applyInPowerPoint(s, home, async () => { setSession('ap1', { state: 'open' }); return { error: { message: 'x', killed: true } }; });
    assert.equal(getSession('ap1')!.state, 'error');

    // A newer Apply (or a Cancel) while this one ran is not overwritten.
    s = setSession('ap1', { chart, state: 'applied' });
    await applyInPowerPoint(s, home, async () => { setSession('ap1', { state: 'cancelled' }); return { error: { message: 'x', killed: true } }; });
    assert.equal(getSession('ap1')!.state, 'cancelled');
    assert.match(readFileSync(join(home, 'addin.log'), 'utf8'), /APPLY ap1 failed after \d+ ms: timed out/);
  } finally {
    rmSync(home, { recursive: true, force: true });
  }
});

test('Apply: the server lays the chart out for its context and ApplyEdit only draws it', async () => {
  const s0 = sampleFor('waterfall');
  assert.ok('spec' in s0);
  const json = JSON.stringify(s0.spec);
  const id = specKey(json);
  const theme = 'theme=000000,FFFFFF,44546A,E7E6E6,C00000,ED7D31,A5A5A5,FFC000,5B9BD5,70AD47,0563C1,954F72\n';
  const context = `48,96,864,408\n48,96,864,408\n${theme}`;
  // Edit data sends the chart's context (frame, bounds, theme) and the engine keeps it for Apply.
  const { sessions } = fakeSessions();
  await handleRequest(parseRequest(`edit\n\n${context}${json}`), m, sessions);
  assert.equal(contextOf(id), context);

  // The server's layout is exactly the update request VBA would send for that context.
  const edited = { ...s0.spec, categories: s0.spec.categories.map((c, i) => (i ? c : 'Start')) };
  const applied: Session = setSession(id, { chart: edited, state: 'applied' });
  const viaVba = await handleRequest(parseRequest(`update\n${id}\n${context}${json}`), m, { get: async () => applied, open: async () => {} });
  assert.equal(await layoutForApply(applied, context, m), viaVba);
  assert.match(viaVba, /\tStart\t/);

  const home = mkdtempSync(join(tmpdir(), 'oc-apply-fast-'));
  try {
    writeFileSync(join(home, 'addin.json'), JSON.stringify({ prefix: 'oc123' }));
    const prepare = (a: Session, c: string) => layoutForApply(a, c, m);
    const calls: string[][] = [];
    const moved = `48,96,864,408\n100,96,432,204\n${theme}`;
    const line = (c: string) => `CONTEXT\t${c.replace(/\n/g, '\t')}`;
    // 1. The context is as expected: one call with the layout; settled while PowerPoint draws, then drawn.
    let s = setSession(id, { chart: edited, state: 'applied' });
    await applyInPowerPoint(s, home, async (_macro, args) => {
      calls.push(args);
      assert.equal(getSession(id)!.state, 'open', 'settled before drawing: a same-data chart keeps its data');
      return { out: `SYNC g 1 of 1 charts\nAPPLIED ${id} OK (200 ms) engine 0 ms, draw 190 ms\n${line(moved)}\n` };
    }, () => false, prepare);
    assert.deepEqual(calls.map((a) => a.slice(0, 2)), [[id, context]]);
    assert.equal(calls[0][2], viaVba);
    assert.equal(getSession(id)!.state, 'drawn');
    assert.equal(contextOf(id), moved, 'the new chart\'s context is kept for the next Apply');
    assert.match(readFileSync(join(home, 'addin.log'), 'utf8'), /SYNC g 1 of 1 charts\n.*APPLIED \w+ OK \(200 ms\)[^\n]*\n.*APPLY \w+ done \(\d+ ms, layout \d+ ms\)/);

    // 2. The chart was moved back meanwhile: ApplyEdit answers its context, the server lays out again.
    calls.length = 0;
    s = setSession(id, { chart: edited, state: 'applied' });
    await applyInPowerPoint(s, home, async (_macro, args) => {
      calls.push(args);
      return { out: args[1] === '*' ? `APPLIED ${id} OK (210 ms)\n${line(context)}` : line(context) };
    }, () => false, prepare);
    assert.deepEqual(calls.map((a) => a.slice(0, 2)), [[id, moved], [id, '*']]);
    assert.equal(calls[1][2], viaVba);
    assert.equal(getSession(id)!.state, 'drawn');

    // 3. No chart, or drawing failed: an error that says what to do.
    for (const answer of [`APPLIED ${id} no chart`, `APPLIED ${id} failed: boom`]) {
      s = setSession(id, { chart: edited, state: 'applied' });
      await applyInPowerPoint(s, home, async () => ({ out: answer }), () => false, prepare);
      assert.equal(getSession(id)!.state, 'error');
      assert.match(getSession(id)!.message!, /could not find the chart/);
    }
  } finally {
    rmSync(home, { recursive: true, force: true });
  }
  assert.deepEqual(parseApplyAnswer('CONTEXT\ta\tb\t\nAPPLIED x OK (1 ms)\n'), { lines: ['APPLIED x OK (1 ms)'], context: 'a\nb\n', applied: 'APPLIED x OK (1 ms)' });
});

test('the AppleScript helper: one request per line, never starts PowerPoint, exits with the server', { skip: process.platform !== 'darwin' }, async () => {
  assert.match(HELPER_JXA, /if \(!ppt\.running\(\)\) reply\(\{ error: 'execution error: PowerPoint is not running\. \(-600\)' \}\)/);
  assert.match(HELPER_JXA, /runVBMacro\(\{ macroName: req\.macro, listOfParameters: req\.args \}\)/);
  const runner = macroRunner();
  assert.equal(await runner.ping(30_000), true, 'answers without an Apple event');
  assert.equal(await runner.ping(30_000), true, 'and again: one helper for many requests');
  runner.stop(); // stdin closes: the helper exits instead of spinning on end of input
  // A fake helper that never answers: the request times out and a fresh helper takes the next one.
  const dir = mkdtempSync(join(tmpdir(), 'oc-runner-'));
  try {
    const n = join(dir, 'n');
    writeFileSync(n, '0');
    const slow = macroRunner('/bin/sh', ['-c', `c=$(cat ${JSON.stringify(n)}); echo $((c+1)) > ${JSON.stringify(n)}; if [ "$c" = 0 ]; then exec sleep 30; else read l; echo '{"out":"pong"}'; fi`]);
    const r = await slow.run('m', [], 200);
    assert.ok('error' in r && r.error.killed);
    assert.equal(await slow.ping(5000), true);
    slow.stop();
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('a server that outlived the PowerPoint that started it never sends Apple events and hands over', async () => {
  // A process id that is gone: a child that has exited.
  const gone = Number(execFileSync('sh', ['-c', 'echo $$'], { encoding: 'utf8' }).trim());
  assert.equal(outlivedPowerPoint(undefined, () => 1), false, 'started without PowerPoint: nothing to outlive');
  assert.equal(outlivedPowerPoint(process.pid, () => 1), false, 'its PowerPoint is still running');
  assert.equal(outlivedPowerPoint(gone, () => undefined), false, 'PowerPoint closed: Apply says it is not running');
  assert.equal(outlivedPowerPoint(gone, () => 1), true, 'another PowerPoint is running');

  // Apply: no osascript (it would freeze the new PowerPoint), an error that says what to do.
  const home = mkdtempSync(join(tmpdir(), 'oc-outlived-'));
  try {
    writeFileSync(join(home, 'addin.json'), JSON.stringify({ prefix: 'oc123' }));
    const s = setSession('ap2', { chart: { type: 'stacked', categories: ['a'], series: [{ name: 'S', values: [1] }] }, state: 'applied' });
    await applyInPowerPoint(s, home, async () => assert.fail('must not send an Apple event'), () => true);
    assert.deepEqual([getSession('ap2')!.state, getSession('ap2')!.message], ['error', RESTARTED_MESSAGE]);
  } finally {
    rmSync(home, { recursive: true, force: true });
  }

  // The new PowerPoint's ping (startServer) and engine calls get 503, so the bridge starts a fresh
  // server; then this one stops. Other routes (the editor's session polls) still answer.
  const call = async (outlived: boolean, method: string, url: string) => {
    let code = 0, retired = false;
    const res = { writeHead: (c: number) => { code = c; return { end: () => undefined }; } };
    const handled = await addinRoutes(() => m, {} as Sessions, () => outlived, () => { retired = true; })(
      Object.assign((async function* () { yield Buffer.from('types'); })(), { method, url }) as never, res as never);
    await new Promise((r) => setImmediate(r));
    return { handled, code, retired };
  };
  assert.deepEqual(await call(true, 'GET', '/api/addin/ping'), { handled: true, code: 503, retired: true });
  assert.deepEqual(await call(true, 'POST', '/api/addin/run'), { handled: true, code: 503, retired: true });
  assert.deepEqual(await call(false, 'GET', '/api/addin/ping'), { handled: true, code: 200, retired: false });
  assert.deepEqual(await call(true, 'GET', '/api/session/x'), { handled: false, code: 0, retired: false });

  // Edit data (and only that) tells the server to open its connection to PowerPoint for the Apply.
  const s = sampleFor('stacked');
  assert.ok('spec' in s);
  let edits = 0;
  const routes = addinRoutes(() => m, fakeSessions().sessions, () => false, () => {}, () => { edits++; });
  const res = { writeHead: () => ({ end: () => undefined }) };
  for (const body of ['types', `edit\n\n48,96,864,408\n48,96,864,408\n${JSON.stringify(s.spec)}`]) {
    await routes(Object.assign((async function* () { yield Buffer.from(body); })(), { method: 'POST', url: '/api/addin/run' }) as never, res as never);
  }
  assert.equal(edits, 1);
});
