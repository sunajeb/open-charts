// Office.js add-in, Node side: Path B planning (scene → PowerPoint JS shapes), Path A in place
// (slide patch), the frame rule behind the resize banner, the manifest and the browser build.
// The task pane itself runs against a mocked PowerPoint in test/e2e/officejs.spec.ts.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { inflateRawSync } from 'node:zlib';
import { createMeasurer } from '../src/fonts.ts';
import { samples } from '../src/core/samples.ts';
import type { ChartSpec, Scene } from '../src/core/types.ts';
import { buildChart, chartDeck, CHART_FRAME, resolveFrame, toBase64, fromBase64 } from '../src/officejs/engine.ts';
import { hostCaps, NO_CAPS, planScene, type Op } from '../src/officejs/plan.ts';
import { patchSlide } from '../src/officejs/slide-patch.ts';
import { manifestXml, ADDIN_ID } from '../src/officejs/manifest.ts';
import { writePptx } from '../src/render/pptx.ts';
import { chartContent } from '../src/deck.ts';
import { unzip } from '../src/ooxml/unzip.ts';
import { findAll, parseXml, path } from '../src/ooxml/xml.ts';
import { readCharts } from '../src/ooxml/pptx-read.ts';

const m = createMeasurer();
const ALL = hostCaps(() => true);
const V18 = hostCaps((v) => !['1.9', '1.10'].includes(v));
const V14 = hostCaps((v) => Number(v.split('.')[1]) <= 4);
const scene = (shapes: Scene['shapes'], roles?: string[]): Scene => ({ frame: { x: 0, y: 0, w: 100, h: 100 }, shapes, roles: roles ?? shapes.map(() => 'x'), hidden: [] });
const opsOf = (s: Scene, caps = ALL): Op[] => {
  const p = planScene(s, caps, 'Arial');
  assert.equal(p.path, 'B', p.path === 'A' ? p.reason : '');
  return (p as { ops: Op[] }).ops;
};

test('hostCaps reads the PowerPointApi sets', () => {
  assert.deepEqual(hostCaps(() => true), { insertSlides: true, shapes: true, selection: true, groups: true, rotation: true });
  assert.deepEqual(V14, { insertSlides: true, shapes: true, selection: false, groups: false, rotation: false });
});

test('every sample chart has a plan; Path B where PowerPoint JS can draw it', () => {
  const pathA = new Set<string>();
  for (const [type, spec] of Object.entries(samples)) {
    const b = buildChart(spec as ChartSpec, CHART_FRAME, m);
    const p = planScene(b.scene, ALL, b.font);
    if (p.path === 'A') { pathA.add(type); continue; }
    // One op per shape, plus a triangle per arrowhead; every op keeps its shape's role.
    const heads = b.scene.shapes.filter((s) => s.kind === 'line' && s.arrowEnd).length;
    assert.equal(p.ops.length, b.scene.shapes.length + heads, type);
    for (const op of p.ops) assert.ok(op.role, `${type}: ${op.name} has a role`);
    for (const op of p.ops) {
      const fills = [op.kind !== 'line' ? op.fill : null, op.kind === 'line' || op.line ? op.line?.color : null].filter(Boolean) as string[];
      for (const c of fills) assert.match(c, /^#[0-9A-F]{6}$/, `${type}: colour ${c}`);
    }
  }
  // Only the freeform area bands need the writer's slide.
  assert.deepEqual([...pathA].sort(), ['area', 'area100']);
});

test('without rotation (before 1.10) diagonal lines, arrowheads and wedges fall back to Path A', () => {
  const plan = (type: string, caps = V18) => {
    const b = buildChart(samples[type as keyof typeof samples]!, CHART_FRAME, m);
    return planScene(b.scene, caps, b.font);
  };
  assert.equal(plan('clustered').path, 'B');
  assert.equal(plan('waterfall').path, 'B');
  for (const t of ['line', 'pie', 'doughnut']) assert.equal(plan(t).path, 'A', t);
  // Without groups (before 1.8) nothing is drawn shape by shape.
  const p = plan('clustered', V14);
  assert.equal(p.path, 'A');
  assert.match((p as { reason: string }).reason, /1\.8/);
  assert.equal(planScene(scene([]), NO_CAPS, 'Arial').path, 'A');
});

test('rects, lines and text map onto presets, boxes and text formatting', () => {
  const ops = opsOf(scene([
    { kind: 'rect', x: 10, y: 20, w: 30, h: 40, fill: '1F3A5F@accent1', stroke: 'FFFFFF', strokeWidth: 1 },
    { kind: 'rect', x: 1, y: 2, w: 3, h: 4, fill: 'C0504D', geom: 'diamond' },
    { kind: 'line', x1: 50, y1: 10, x2: 10, y2: 10, color: '7F7F7F', width: 0.75, dash: true },
    { kind: 'text', x: 5, y: 6, w: 40, h: 14, text: 'a\nb', size: 12, color: '222222@tx1', bold: true, align: 'right', valign: 'bottom' },
    { kind: 'text', x: 5, y: 6, w: 40, h: 14, text: '+5%', size: 10, color: 'FFFFFF', align: 'center', valign: 'middle', fill: '1F3A5F', rounded: true },
  ], ['bar:s0:c0', 'marker:s0:c1', 'axis', 'label:value:s0:c0', 'bubble:cagr']));
  assert.deepEqual(ops[0], { kind: 'geom', geom: 'Rectangle', box: { left: 10, top: 20, width: 30, height: 40 }, fill: '#1F3A5F', line: { color: '#FFFFFF', weight: 1 }, role: 'bar:s0:c0', name: 'Bar s0 c0' });
  assert.equal((ops[1] as { geom: string }).geom, 'Diamond');
  assert.deepEqual(ops[2], { kind: 'line', box: { left: 10, top: 10, width: 40, height: 0 }, line: { color: '#7F7F7F', weight: 0.75, dash: 'Dash' }, role: 'axis', name: 'Axis' });
  assert.deepEqual(ops[3], { kind: 'text', box: { left: 5, top: 6, width: 40, height: 14 }, fill: null, line: null, role: 'label:value:s0:c0', name: 'Label value s0 c0',
    text: { text: 'a\nb', font: 'Arial', size: 12, bold: true, color: '#222222', align: 'Right', valign: 'Bottom' } });
  assert.equal(ops[4].kind, 'geom');
  assert.deepEqual((ops[4] as { adjustments: number[] }).adjustments, [0.5]);
  assert.equal((ops[4] as { geom: string }).geom, 'RoundRectangle');
});

test('a rising diagonal line is a rotated horizontal line through its midpoint', () => {
  const [op] = opsOf(scene([{ kind: 'line', x1: 0, y1: 40, x2: 30, y2: 0, color: '000000', width: 1 }]));
  assert.equal(op.kind, 'line');
  assert.deepEqual(op.box, { left: -10, top: 20, width: 50, height: 0 });
  assert.equal((op as { rotation: number }).rotation, 306.87);
  // Falling diagonals need no rotation: the bounding box already runs top-left to bottom-right.
  const [down] = opsOf(scene([{ kind: 'line', x1: 30, y1: 40, x2: 0, y2: 0, color: '000000', width: 1 }]), V18);
  assert.deepEqual(down.box, { left: 0, top: 0, width: 30, height: 40 });
  assert.equal((down as { rotation?: number }).rotation, undefined);
});

test('an arrowhead is a triangle three line widths big, tip on the end point, turned along the line', () => {
  const ops = opsOf(scene([{ kind: 'line', x1: 50, y1: 100, x2: 50, y2: 10, color: '222222', width: 1, arrowEnd: true }], ['arrow:cagr']));
  assert.equal(ops.length, 2);
  // Pointing up: the preset as it is.
  assert.deepEqual(ops[1], { kind: 'geom', geom: 'Triangle', box: { left: 48.5, top: 10, width: 3, height: 3 }, fill: '#222222', line: null, role: 'arrow:cagr', name: 'Arrow cagr head' });
  const right = opsOf(scene([{ kind: 'line', x1: 0, y1: 5, x2: 60, y2: 5, color: '222222', width: 1, arrowEnd: true }]));
  assert.equal((right[1] as { rotation: number }).rotation, 90);
  assert.deepEqual(right[1].box, { left: 57, top: 3.5, width: 3, height: 3 });
  const down = opsOf(scene([{ kind: 'line', x1: 0, y1: 0, x2: 0, y2: 60, color: '222222', width: 1, arrowEnd: true }]));
  assert.equal((down[1] as { rotation: number }).rotation, 180);
});

test('wedges, rings and triangles become Pie, BlockArc, Donut and Triangle presets', () => {
  const c = { cx: 100, cy: 100 };
  const [pie, ring, donut, circle, tri] = opsOf(scene([
    { kind: 'path', fill: '1F3A5F', commands: [{ op: 'M', x: 100, y: 100 }, { op: 'A', ...c, r: 50, start: -90, sweep: 120 }, { op: 'Z' }] },
    { kind: 'path', fill: '3E6A9E', commands: [{ op: 'A', ...c, r: 50, start: 30, sweep: 60 }, { op: 'A', ...c, r: 25, start: 90, sweep: -60 }, { op: 'Z' }] },
    { kind: 'path', fill: '3E6A9E', commands: [{ op: 'A', ...c, r: 50, start: -90, sweep: 360 }, { op: 'A', ...c, r: 30, start: 270, sweep: -360 }, { op: 'Z' }] },
    { kind: 'path', fill: '3E6A9E', commands: [{ op: 'A', ...c, r: 50, start: -90, sweep: 360 }, { op: 'Z' }] },
    { kind: 'path', fill: '222222', commands: [{ op: 'M', x: 10, y: 50 }, { op: 'L', x: 4, y: 47 }, { op: 'L', x: 4, y: 53 }, { op: 'Z' }] },
  ]));
  const circleBox = { left: 50, top: 50, width: 100, height: 100 };
  assert.deepEqual([pie.kind === 'geom' && pie.geom, pie.box, (pie as { adjustments: number[] }).adjustments], ['Pie', circleBox, [270, 30]]);
  assert.deepEqual([ring.kind === 'geom' && ring.geom, (ring as { adjustments: number[] }).adjustments], ['BlockArc', [30, 90, 0.25]]);
  assert.deepEqual([donut.kind === 'geom' && donut.geom, (donut as { adjustments: number[] }).adjustments], ['Donut', [0.2]]);
  assert.equal(circle.kind === 'geom' && circle.geom, 'Ellipse');
  // A triangle pointing right (tip at x=10): base 6 tall, 6 long; turned 90°.
  assert.deepEqual([tri.kind === 'geom' && tri.geom, tri.box, (tri as { rotation: number }).rotation], ['Triangle', { left: 4, top: 47, width: 6, height: 6 }, 90]);
  // A free polygon has no preset.
  const p = planScene(scene([{ kind: 'path', fill: '222222', commands: [{ op: 'M', x: 0, y: 0 }, { op: 'L', x: 9, y: 1 }, { op: 'L', x: 5, y: 9 }, { op: 'L', x: 1, y: 6 }, { op: 'Z' }] }], ['area:s0']), ALL, 'Arial');
  assert.equal(p.path, 'A');
});

test('resolveFrame: unchanged bounds keep the frame; a move keeps the size; a resize re-plans', () => {
  const spec = samples.clustered!;
  const f = CHART_FRAME;
  const b = buildChart(spec, f, m);
  const bounds = { x: Math.min(f.x, ...b.scene.shapes.map(() => f.x)), y: f.y, w: f.w, h: f.h };
  assert.deepEqual(resolveFrame(spec, f, bounds, m), { frame: f, resized: false });
  const moved = resolveFrame(spec, f, { ...bounds, x: bounds.x + 50, y: bounds.y + 10 }, m);
  assert.equal(moved.resized, false);
  assert.ok(Math.abs(moved.frame.x - (f.x + 50)) < 0.01 && Math.abs(moved.frame.w - f.w) < 0.01);
  const resized = resolveFrame(spec, f, { ...bounds, w: bounds.w / 2 }, m);
  assert.equal(resized.resized, true);
  assert.ok(Math.abs(resized.frame.w - f.w / 2) < 0.5);
  // No stored frame (a group without OC_FRAME): lay out where it is.
  assert.deepEqual(resolveFrame(spec, undefined, bounds, m), { frame: bounds, resized: false });
});

const inflate = async (c: Uint8Array, size: number) => new Uint8Array(inflateRawSync(c, { maxOutputLength: Math.max(1, size) }));

test('Path A in place: the chart replaces the old group at its z-order; untouched parts are copied byte for byte', async () => {
  // A user slide: a text box, an old chart, another box on top (the CLI writer stands in for PowerPoint's export).
  const old = buildChart(samples.clustered!, { x: 40, y: 80, w: 400, h: 300 }, m);
  const top = buildChart(samples.waterfall!, { x: 500, y: 80, w: 400, h: 300 }, m);
  const marked = { ...old.content, name: 'open-charts:replace-me' };
  const exported = writePptx([{ title: 'Revenue', scenes: [marked, top.content] }]);
  const fresh = buildChart(samples.area!, { x: 40, y: 80, w: 400, h: 300 }, m);
  const out = await patchSlide(exported, fresh.content, { replaceName: 'open-charts:replace-me', font: 'Arial' }, inflate);

  const before = unzip(exported), after = unzip(out);
  for (const [name, e] of before) {
    if (['ppt/slides/slide1.xml', 'ppt/slides/_rels/slide1.xml.rels', '[Content_Types].xml'].includes(name)) continue;
    assert.deepEqual(after.get(name)?.raw.compressed, e.raw.compressed, name);
  }
  const charts = readCharts(out)[0].charts;
  assert.deepEqual(charts.map((c) => c.name), ['Stacked area chart', 'Waterfall chart'], 'new chart where the old one was, below the other chart');
  assert.deepEqual(JSON.parse(charts[0].spec!), JSON.parse((fresh.content.tags!.find(([k]) => k === 'OC_SPEC'))![1]));
  // Shape ids stay unique on the slide; every child of the new chart has its OC_ROLE.
  const ids = findAll(parseXml(after.get('ppt/slides/slide1.xml')!.text()), 'p:cNvPr').map((n) => n.attrs.id);
  assert.equal(new Set(ids).size, ids.length);
  assert.ok(charts[0].roles.length >= fresh.scene.shapes.length - 1);
});

test('Path A in place, insert: the chart goes on top of the slide; a missing marker is an error', async () => {
  const exported = writePptx([{ title: 'Empty', scenes: [] }]);
  const fresh = buildChart(samples.pie!, CHART_FRAME, m);
  const out = await patchSlide(exported, fresh.content, { font: 'Arial' }, inflate);
  assert.deepEqual(readCharts(out)[0].charts.map((c) => c.name), ['Pie chart']);
  const tree = path(parseXml(unzip(out).get('ppt/slides/slide1.xml')!.text()).children[0], 'p:cSld', 'p:spTree')!;
  assert.equal(tree.children.at(-1)!.name, 'p:grpSp');
  await assert.rejects(patchSlide(exported, fresh.content, { font: 'Arial', replaceName: 'nope' }, inflate), /not found/);
});

test('Path A new slide: the one-slide deck carries tag contract v1 and theme slots', () => {
  const spec = { ...samples.stacked!, style: { theme: { accent1: '112233' } } };
  const b = buildChart(spec, CHART_FRAME, m);
  const deck = chartDeck(b, spec);
  const [slide] = readCharts(deck);
  assert.equal(slide.charts.length, 1);
  assert.deepEqual(JSON.parse(slide.charts[0].spec!), spec);
  assert.match(unzip(deck).get('ppt/theme/theme1.xml')!.text(), /accent1><a:srgbClr val="112233"/);
  assert.deepEqual(fromBase64(toBase64(deck)), deck);
});

test('chart content is the CLI\'s: same group name, alt text and tags as chartContent', () => {
  const spec = samples.line!;
  const b = buildChart(spec, CHART_FRAME, m);
  assert.deepEqual(b.content, chartContent(spec, b.scene, 'Arial', undefined));
});

test('manifest: PowerPoint task pane, one ribbon button, PowerPointApi 1.4 floor, https URLs', () => {
  const xml = manifestXml('https://localhost:4175');
  const doc = parseXml(xml);
  const app = doc.children.find((c) => c.name === 'OfficeApp')!;
  assert.equal(app.attrs['xsi:type'], 'TaskPaneApp');
  assert.equal(path(app, 'Id')!.text, ADDIN_ID);
  assert.match(path(app, 'Version')!.text, /^\d+\.\d+\.\d+\.\d+$/);
  assert.deepEqual(findAll(path(app, 'Hosts')!, 'Host').map((h) => h.attrs.Name), ['Presentation']);
  assert.deepEqual(path(app, 'Requirements', 'Sets', 'Set')!.attrs, { Name: 'PowerPointApi', MinVersion: '1.4' });
  assert.equal(path(app, 'DefaultSettings', 'SourceLocation')!.attrs.DefaultValue, 'https://localhost:4175/taskpane.html?host=office');
  const controls = findAll(app, 'Control');
  assert.equal(controls.length, 1);
  assert.equal(path(controls[0], 'Action')!.attrs['xsi:type'], 'ShowTaskpane');
  const strings = Object.fromEntries(findAll(app, 'bt:String').map((s) => [s.attrs.id, s.attrs.DefaultValue]));
  assert.equal(strings['Button.Label'], 'open-charts');
  for (const u of [...xml.matchAll(/DefaultValue="(http[^"]+)"/g)].map((x) => x[1])) assert.ok(u.startsWith('https://localhost:4175/'), u);
  // Element order the schema requires.
  const order = app.children.map((c) => c.name);
  assert.deepEqual(order, ['Id', 'Version', 'ProviderName', 'DefaultLocale', 'DisplayName', 'Description', 'IconUrl', 'HighResolutionIconUrl', 'SupportUrl', 'AppDomains', 'Hosts', 'Requirements', 'DefaultSettings', 'Permissions', 'VersionOverrides']);
});

test('element overrides: dotted lines, start arrowheads, turned text; patterns and theme fonts need Path A', () => {
  const [dotted, head0, head1] = opsOf(scene([{ kind: 'line', x1: 0, y1: 10, x2: 40, y2: 10, color: '222222', width: 1, dash: 'dot', arrowEnd: true, arrowStart: true }]));
  assert.deepEqual((dotted as { line: object }).line, { color: '#222222', weight: 1, dash: 'SystemDot' });
  assert.equal((head0 as { rotation: number }).rotation, 90); // end: pointing right
  assert.equal((head1 as { rotation: number }).rotation, 270); // start: pointing left
  assert.deepEqual(head1.box, { left: 0, top: 8.5, width: 3, height: 3 });
  // Turned 90° left: the unturned 40 × 12 box, centred on the 12 × 40 footprint, rotated 270°.
  const [turned] = opsOf(scene([{ kind: 'text', x: 10, y: 20, w: 12, h: 40, text: 'Revenue', size: 10, color: '222222', align: 'center', valign: 'middle', rotate: -90 }]));
  assert.deepEqual([turned.box, (turned as { rotation: number }).rotation], [{ left: -4, top: 34, width: 40, height: 12 }, 270]);
  assert.equal(planScene(scene([{ kind: 'text', x: 10, y: 20, w: 12, h: 40, text: 'R', size: 10, color: '222222', align: 'center', valign: 'middle', rotate: 90 }]), V18, 'Arial').path, 'A');
  assert.equal(planScene(scene([{ kind: 'rect', x: 0, y: 0, w: 5, h: 5, fill: '222222', pattern: 'ltUpDiag' }]), ALL, 'Arial').path, 'A');
  assert.equal(planScene(scene([{ kind: 'text', x: 0, y: 0, w: 5, h: 5, text: 'x', size: 10, color: '222222', align: 'left', valign: 'top', font: '+mj-lt' }]), ALL, 'Arial').path, 'A');
  const [own] = opsOf(scene([{ kind: 'text', x: 0, y: 0, w: 5, h: 5, text: 'x', size: 10, color: '222222', align: 'left', valign: 'top', font: 'Georgia', stroke: '000000', strokeWidth: 2 }]));
  assert.deepEqual([(own as { text: { font: string } }).text.font, (own as { line: object }).line], ['Georgia', { color: '#000000', weight: 2 }]);
});

test('Path B groups carry the same tags as CLI charts, OC_BASE included', () => {
  const b = buildChart(samples.stacked!, CHART_FRAME, m);
  const keys = (b.content.tags ?? []).map(([k]) => k);
  for (const k of ['OC_V', 'OC_ENGINE', 'OC_FRAME', 'OC_FONT']) assert.ok(keys.includes(k), k);
  assert.ok(keys.some((k) => k.startsWith('OC_BASE')), 'OC_BASE');
});

test('build:officejs bundles the engine for the browser (no Node built-ins) and the office-mode pane', async () => {
  const { buildOfficeJs } = await import('../scripts/build-officejs.ts');
  const out = mkdtempSync(join(tmpdir(), 'oc-officejs-'));
  try {
    await buildOfficeJs({ out, quiet: true });
    for (const f of ['taskpane.html', 'taskpane.js', 'engine.js', 'commands.html', 'manifest.xml', 'support.html', 'fonts/LiberationSans-Regular.ttf', 'fonts/LiberationSans-Bold.ttf', 'fonts/LiberationSans.LICENSE', 'assets/icon-80.png']) {
      assert.ok(existsSync(join(out, f)), f);
    }
    for (const f of ['taskpane.js', 'engine.js']) assert.doesNotMatch(readFileSync(join(out, f), 'utf8'), /from ?["']node:|require\(["']node:/, f);
    const html = readFileSync(join(out, 'taskpane.html'), 'utf8');
    const office = html.indexOf('appsforoffice.microsoft.com/lib/1/hosted/office.js'), pane = html.indexOf('src="taskpane.js"'), boot = html.indexOf('window.OC_SPEC =');
    assert.ok(office > 0 && office < pane && pane < boot, 'office.js, then taskpane.js, then the editor');
    assert.match(readFileSync(join(out, 'fonts/LiberationSans.LICENSE'), 'utf8'), /SIL OPEN FONT LICENSE Version 1\.1/);
  } finally {
    rmSync(out, { recursive: true, force: true });
  }
});
