import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, writeFileSync, mkdirSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createMeasurer } from '../src/fonts.ts';
import { flattenPath } from '../src/core/path.ts';
import { escapeField, parseResponse, shapeRecord, SHAPE_COLUMNS, unescapeField } from '../src/addin/format.ts';
import { DEFAULT_FRAME, handleRequest, parseRequest, resolveFrame, sceneBounds, specKey } from '../src/addin/engine.ts';
import { fileEdits, editPage, editRoutes } from '../src/addin/edit.ts';
import { buildAddin, ADDIN_CT, VBA_REL, UI_REL, UI_NS } from '../src/addin/package.ts';
import { galleryItems, insertableItems, sampleFor } from '../src/addin/registry.ts';
import { unzip } from '../src/ooxml/unzip.ts';
import { readCfb } from '../src/addin/cfb.ts';
import { decompress } from '../src/addin/ovba.ts';
import { layoutChart, type Shape } from '../src/core/index.ts';

const m = createMeasurer();
const req = (op: string, arg = '', frame = '', bounds = '', spec = '') => parseRequest([op, arg, frame, bounds, spec].join('\n'));

test('field escaping round-trips tabs, newlines, CR and backslashes', () => {
  for (const s of ['plain', 'a\tb', 'two\nlines', 'cr\r\n', 'back\\slash', '\\t literal', '']) {
    const e = escapeField(s);
    assert.ok(!/[\t\n\r]/.test(e));
    assert.equal(unescapeField(e), s);
  }
});

test('shape records carry every property VBA needs', () => {
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

test('path shapes flatten arcs into polygon points', () => {
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

test('insert answers with group, tags and one shape per scene shape, roles and names', () => {
  const r = parseResponse(handleRequest(req('insert', 'stacked', '48,96,864,408'), m));
  assert.deepEqual(r.errors, []);
  assert.equal(r.version, '1');
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

test('every insertable gallery item lays out; others explain themselves', () => {
  const items = insertableItems();
  assert.ok(items.length >= 6);
  for (const id of items) assert.deepEqual(parseResponse(handleRequest(req('insert', id, '0,0,600,400'), m)).errors, [], id);
  const missing = galleryItems().find((g) => !items.includes(g.id));
  if (missing) assert.match(parseResponse(handleRequest(req('insert', missing.id), m)).errors[0], /not available in this version/);
  assert.match(parseResponse(handleRequest(req('insert', 'nope'), m)).errors[0], /not available/);
  assert.match(parseResponse(handleRequest(req('bogus'), m)).errors[0], /Unknown request/);
});

test('large specs are chunked into OC_SPEC_0..n + OC_SPEC_N', () => {
  const cats = Array.from({ length: 400 }, (_, i) => `Category with a fairly long name, padded so the spec passes the tag chunk size, number ${i}`);
  const spec = { specVersion: 1, type: 'line', categories: cats, series: [{ name: 'A', values: cats.map((_, i) => i) }] };
  const r = parseResponse(handleRequest(req('update', '', '0,0,900,500', '', JSON.stringify(spec)), m));
  assert.deepEqual(r.errors, []);
  const tags = Object.fromEntries(r.tags);
  assert.ok(Number(tags.OC_SPEC_N) >= 1 && tags.OC_SPEC === undefined);
  const joined = Array.from({ length: Number(tags.OC_SPEC_N) }, (_, i) => tags[`OC_SPEC_${i}`]).join('');
  assert.equal(JSON.parse(joined).categories.length, 400);
});

test('update re-lays out in the stored frame, or follows a move/resize', () => {
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
  const r = parseResponse(handleRequest(req('update', '', '48,96,864,408', '', JSON.stringify(s.spec)), m));
  assert.deepEqual(r.errors, []);
  assert.match(parseResponse(handleRequest(req('update'), m)).errors[0], /Select an open-charts chart/);
  assert.match(parseResponse(handleRequest(req('update', '', '', '', '{broken'), m)).errors[0], /damaged/);
  assert.match(parseResponse(handleRequest(req('update', '', '', '', '{"type":"stacked","categories":[],"series":[]}'), m)).errors[0], /categor/i);
  assert.equal(DEFAULT_FRAME.w, 864);
});

test('edit data: the edited spec is applied by the next update of that chart only', async () => {
  const home = mkdtempSync(join(tmpdir(), 'oc-addin-'));
  const opened: string[] = [];
  const edits = fileEdits(home, (url) => opened.push(url));
  const s = sampleFor('clustered');
  assert.ok('spec' in s);
  const json = JSON.stringify(s.spec);
  const key = specKey(json);
  const msg = parseResponse(handleRequest(req('edit', '', '', '', json), m, edits));
  assert.match(msg.messages[0], /Update chart/);
  assert.deepEqual(opened, [`http://localhost:4174/edit/${key}`]);

  // The editor page posts the deck back as the user edits.
  const page = editPage(key, readFileSync(join(home, 'edits', `${key}.base.json`), 'utf8'));
  assert.match(page, /\/api\/addin\/save\//);
  assert.ok(page.indexOf('/api/addin/save/') < page.indexOf('<script', page.indexOf('/api/addin/save/')), 'boot script runs before the editor script');
  const route = editRoutes(home);
  const edited = { ...s.spec, categories: ['A', 'B', 'C', 'D'] };
  const body = JSON.stringify({ slides: [{ chart: edited }] });
  let status = 0;
  const res = { writeHead(n: number) { status = n; return this; }, end() {} };
  const reqStream = Object.assign((async function* () { yield Buffer.from(body); })(), { method: 'POST', url: `/api/addin/save/${key}` });
  assert.equal(await route(reqStream as never, res as never), true);
  assert.equal(status, 204);

  const r = parseResponse(handleRequest(req('update', '', '48,96,864,408', '', json), m, edits));
  assert.deepEqual(JSON.parse(Object.fromEntries(r.tags).OC_SPEC).categories, ['A', 'B', 'C', 'D']);
  assert.equal(existsSync(join(home, 'edits', `${key}.json`)), false, 'an edit is applied once');
  const again = parseResponse(handleRequest(req('update', '', '48,96,864,408', '', json), m, edits));
  assert.equal(JSON.parse(Object.fromEntries(again.tags).OC_SPEC).categories[0], 'North America');
});

test('ppam package: add-in content type, VBA and ribbon parts, one callback prefix', () => {
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
  const callbacks = [...ui.matchAll(/(?:onAction|onLoad)="([^"]+)"/g)].map((x) => x[1]);
  assert.ok(callbacks.length >= 4 && callbacks.every((c) => c.startsWith(`${built.prefix}_`)));
  for (const g of galleryItems()) assert.ok(ui.includes(`id="chart_${g.id}"`), g.id);
  assert.match(ui, /id="btnEditData"/);
  assert.match(ui, /id="btnUpdate"/);

  const vba = readCfb(e.get('ppt/vbaProject.bin')!.bytes()) as { VBA: Record<string, Uint8Array>; PROJECT: Uint8Array };
  const source = new TextDecoder().decode(decompress(vba.VBA.OpenCharts));
  for (const c of new Set(callbacks)) assert.ok(source.includes(`Sub ${c}(`), `VBA defines ${c}`);
  assert.ok(!source.includes('OCPREFIX'));
  assert.match(new TextDecoder().decode(vba.PROJECT), /Name="OpenChartsAddIn"/);
  assert.deepEqual(buildAddin().ppam, built.ppam, 'reproducible build');
});

test('VBA source is ASCII and every Sub/Function is closed', () => {
  const src = readFileSync(new URL('../addin/vba/OpenCharts.bas', import.meta.url), 'utf8');
  assert.ok(/^[\x09\x0a\x0d\x20-\x7e]*$/.test(src));
  const opens = (src.match(/^\s*(Public |Private )?(Sub|Function) /gm) ?? []).length;
  const closes = (src.match(/^\s*End (Sub|Function)\b/gm) ?? []).length;
  assert.equal(opens, closes);
});

test('request parsing tolerates CRLF and missing fields', () => {
  const r = parseRequest('update\r\n\r\n1,2,3,4\r\n5,6,7,8\r\n{"a":1}\r\n');
  assert.deepEqual([r.op, r.arg, r.frame, r.bounds, r.spec], ['update', '', { x: 1, y: 2, w: 3, h: 4 }, { x: 5, y: 6, w: 7, h: 8 }, '{"a":1}']);
  assert.deepEqual(parseRequest(''), { op: '', arg: '', frame: undefined, bounds: undefined, spec: '' });
  mkdirSync(join(tmpdir(), 'oc-unused'), { recursive: true });
  writeFileSync(join(tmpdir(), 'oc-unused', 'x'), '');
});
