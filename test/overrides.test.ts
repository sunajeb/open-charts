// Per-element overrides (core/overrides.ts): spec validation, matching, layout on every chart type,
// renderers (writer, SVG, add-in format), convert to plain text, and OC_BASE native-edit capture.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  checkScene, defaultStyle, layoutChart, validateSpec, compileOverrides, kindOf, propertiesFor, ELEMENT_PROPERTIES, ELEMENT_STYLE_KEYS,
  remapOverrides, detachElement, unpinElement, baseOf, mergeNativeEdits, renderedOf, sceneEntries, resolveElementColor, themeOf, tintHex,
  type ChartSpec, type ChartType, type Scene, type TextShape,
} from '../src/core/index.ts';
import { samples } from '../src/core/samples.ts';
import { createMeasurer } from '../src/fonts.ts';
import { groupXml, Ids, writePptx } from '../src/render/pptx.ts';
import { sceneToSvg } from '../src/render/svg.ts';
import { parseResponse, shapeRecord } from '../src/addin/format.ts';
import { handleRequest } from '../src/addin/engine.ts';
import { chartContent } from '../src/deck.ts';
import { chunkedValue } from '../src/render/tags.ts';
import { readFileSync } from 'node:fs';

const m = createMeasurer();
const F = { x: 48, y: 96, w: 864, h: 408 };
const lay = (spec: ChartSpec, frame = F): Scene => layoutChart(spec, frame, m, defaultStyle);
const texts = (sc: Scene) => sc.shapes.map((s, i) => ({ s, role: sc.roles![i] })).filter((x): x is { s: TextShape; role: string } => x.s.kind === 'text');
const at = (sc: Scene, role: string) => sc.shapes.filter((_, i) => sc.roles![i] === role);
const stacked: ChartSpec = {
  type: 'stacked', categories: ['2023', '2024', '2025'],
  series: [{ name: 'Hardware', values: [40, 44, 47] }, { name: 'Services', values: [20, 26, 33] }],
};
const issues = (sc: Scene) => checkScene(sc, { measurer: m }).filter((i) => i.kind !== 'hidden');

// ---- Spec ----------------------------------------------------------------------------------------

test('overrides validate with path errors, by element kind', () => {
  assert.deepEqual(validateSpec({ ...stacked, overrides: { 'label:value:s1:*': { color: 'accent1-lighter40', size: 14, rotate: -90, content: ['category', { field: 'percent', decimals: 1 }, { text: '!' }] }, 'bar:*:c2': { fill: 'none', outline: 'tx1-darker25', outlineDash: 'dot' } } }), []);
  const errs = validateSpec({ ...stacked, overrides: { 'bar:s0:c0': { size: 12, fill: 'accent9' }, 'foo:s1': {}, 'label:value:s0:c0': { rotate: 45, content: ['val'], bogus: 1 }, 'label::x': {} } });
  const paths = errs.map((e) => e.path);
  assert.ok(paths.includes('/overrides/bar:s0:c0/size'), 'size is not a segment property');
  assert.ok(paths.includes('/overrides/bar:s0:c0/fill'));
  assert.ok(paths.includes('/overrides/foo:s1'));
  assert.ok(paths.includes('/overrides/label:value:s0:c0/rotate'));
  assert.ok(paths.includes('/overrides/label:value:s0:c0/content/0'));
  assert.ok(paths.includes('/overrides/label:value:s0:c0/bogus'));
  assert.ok(paths.includes('/overrides/label::x'));
  // An index past the data is not an error (data edits must not break a chart).
  assert.deepEqual(validateSpec({ ...stacked, overrides: { 'label:value:s9:c9': { bold: true } } }), []);
});

test('the property table drives kinds and the schema', () => {
  assert.equal(kindOf('label:value:s0:c1'), 'label');
  assert.equal(kindOf('wedge:s0:c1'), 'segment');
  assert.equal(kindOf('arrow:cagr'), 'decoration');
  assert.ok(propertiesFor('label').includes('rotate') && !propertiesFor('segment').includes('rotate'));
  assert.ok(propertiesFor('line').includes('arrowheads'));
  const schema = JSON.parse(readFileSync('schema/chart.schema.json', 'utf8'));
  assert.deepEqual(Object.keys(schema.$defs.elementStyle.properties).sort(), [...ELEMENT_STYLE_KEYS].sort());
  const doc = readFileSync('docs/DATA-FORMAT.md', 'utf8');
  for (const k of ELEMENT_STYLE_KEYS) assert.match(doc, new RegExp(`^\\| \`${k}\` \\| ${ELEMENT_PROPERTIES[k].kinds.join(', ')} \\|`, 'm'), `DATA-FORMAT.md row for ${k}`);
});

test('precedence: exact > series / category selector > kind selector, all merged', () => {
  const get = compileOverrides({ 'label:value:s1:c2': { color: 'C0504D' }, 'label:value:*': { size: 10, color: '000000' }, 'label:value:s1:*': { size: 14 } });
  assert.deepEqual(get('label:value:s1:c2'), { size: 14, color: 'C0504D' });
  assert.deepEqual(get('label:value:s0:c2'), { size: 10, color: '000000' });
  assert.equal(get('bar:s1:c2'), undefined);
});

// ---- Layout --------------------------------------------------------------------------------------

test('no overrides (or an empty object) lays out exactly as before', () => {
  for (const [type, spec] of Object.entries(samples)) assert.deepEqual(lay({ ...spec!, overrides: {} }), lay(spec!), type);
});

for (const [type, spec] of Object.entries(samples) as [ChartType, ChartSpec][]) {
  test(`${type}: overrides apply to labels and marks, and the scene stays clean`, () => {
    const base = lay(spec);
    const mark = base.roles!.find((r) => ['segment', 'marker'].includes(kindOf(r) ?? ''));
    const line = base.roles!.find((r, i) => kindOf(r) === 'line' && base.shapes[i].kind === 'line' && !r.startsWith('label:'));
    const sc = lay({ ...spec, overrides: {
      'label:*': { color: 'C0504D', size: 13 },
      ...(mark ? { [mark]: { fill: 'accent2-lighter40', outline: '000000', outlineWidth: 1.5 } } : {}),
      ...(line ? { [line]: { line: 'C0504D', lineWidth: 2, dash: 'dot' } } : {}),
    } });
    const labels = texts(sc).filter((t) => t.role.startsWith('label:'));
    assert.ok(labels.length, 'has labels');
    for (const t of labels) { assert.equal(t.s.color, 'C0504D'); assert.equal(t.s.size, 13); }
    if (mark) {
      const r = at(sc, mark).find((s) => s.kind !== 'text' && s.kind !== 'line') as { fill?: string; stroke?: string; strokeWidth?: number };
      assert.match(r.fill!, /@accent2-lighter40$/);
      assert.equal(r.stroke, '000000');
      assert.equal(r.strokeWidth, 1.5);
    }
    if (line) {
      const l = at(sc, line).find((s) => s.kind === 'line') as { color: string; width: number; dash?: unknown };
      assert.deepEqual([l.color, l.width, l.dash], ['C0504D', 2, 'dot']);
    }
    assert.deepEqual(issues(sc), [], 'bigger labels are placed and collision-checked, never overlapping');
    assert.equal(sc.warnings, undefined);
  });
}

test('turned category labels get a taller axis and stay in the frame', () => {
  const sc = lay({ ...stacked, categories: ['Financial year 2023', 'Financial year 2024', 'Financial year 2025'], overrides: { 'label:category:*': { rotate: -90 } } });
  const cats = texts(sc).filter((t) => t.role.startsWith('label:category'));
  assert.equal(cats.length, 3);
  for (const c of cats) { assert.equal(c.s.rotate, -90); assert.ok(c.s.h > c.s.w * 3, 'footprint is tall'); }
  assert.deepEqual(issues(sc), []);
  assert.deepEqual(sc.hidden, []);
});

test('a bigger segment label moves out of its segment (placed at its drawn size)', () => {
  const f = { x: 0, y: 0, w: 600, h: 260 };
  const spec: ChartSpec = { ...stacked, series: [stacked.series[0], { name: 'Services', values: [6, 26, 33] }] };
  const plain = lay(spec, f);
  const big = lay({ ...spec, overrides: { 'label:value:s1:c0': { size: 18 } } }, f);
  const before = at(plain, 'label:value:s1:c0')[0] as TextShape, after = at(big, 'label:value:s1:c0')[0] as TextShape;
  const seg = at(big, 'bar:s1:c0')[0] as { x: number; y: number; w: number; h: number };
  assert.ok(!at(plain, 'leader:s1:c0').length && !before.fill, 'at 12 pt it sits inside');
  assert.equal(after.size, 18);
  assert.ok(after.h > seg.h || after.w > seg.w, 'too big for its segment');
  assert.ok(at(big, 'leader:s1:c0').length || after.fill, 'outside with a leader, or boxed');
  assert.deepEqual(issues(big), []);
});

test('label content: ordered fields with number formats, custom text, decimals', () => {
  const sc = lay({ ...stacked, categories: ['2023', '2024', '2025'], series: [{ name: 'Hardware', values: [4000, 44, 47] }, { name: 'Services', values: [2000, 26, 33] }], overrides: {
    'label:value:s1:c0': { content: ['series', { field: 'value', thousands: false }, { field: 'percent', decimals: 1 }], separator: ' | ' },
    'label:value:s0:c0': { content: [{ field: 'value', sign: true, decimals: 1 }] },
    'label:total:c1': { decimals: 2 },
    'label:category:c2': { content: [{ text: 'Plan' }, 'category'] },
  } });
  assert.equal((at(sc, 'label:value:s1:c0')[0] as TextShape).text, 'Services | 2000 | 33.3%');
  assert.equal((at(sc, 'label:value:s0:c0')[0] as TextShape).text, '+4,000.0');
  assert.equal((at(sc, 'label:total:c1')[0] as TextShape).text, '70.00');
  assert.equal((at(sc, 'label:category:c2')[0] as TextShape).text, 'Plan 2025');
});

test('font, alignment, fill and outline on a label', () => {
  const sc = lay({ ...stacked, overrides: { 'label:total:c0': { font: '+mj-lt', bold: false, fill: 'FFFF00', outline: 'tx1', outlineWidth: 1, align: 'left' } } });
  const t = at(sc, 'label:total:c0')[0] as TextShape;
  assert.deepEqual([t.font, t.bold, t.fill, t.stroke, t.strokeWidth, t.align], ['+mj-lt', undefined, 'FFFF00', '222222@tx1', 1, 'left']);
});

test('markers: shape and none', () => {
  const sc = lay({ ...samples.line!, overrides: { 'marker:s0:*': { marker: 'diamond' }, 'marker:s1:*': { marker: 'none' } } });
  const ms = sc.shapes.filter((_, i) => sc.roles![i].startsWith('marker:s0')) as { geom?: string }[];
  assert.ok(ms.length && ms.every((s) => s.geom === 'diamond'));
  assert.ok((sc.shapes.filter((_, i) => sc.roles![i].startsWith('marker:s1')) as { fill: string }[]).every((s) => s.fill === ''));
});

test('overrides that match nothing are warnings; hidden series and selectors are not', () => {
  const sc = lay({ ...stacked, overrides: { 'label:value:s0:c7': { bold: true }, 'bar:s5:*': { fill: '000000' }, 'label:value:s1:*': { bold: true } } });
  assert.deepEqual(sc.warnings, [
    'override "label:value:s0:c7" matches nothing: there is no category 7 (the chart has 3)',
    'override "bar:s5:*" matches nothing: there is no series 5 (the chart has 2)',
  ]);
  const hid = lay({ ...stacked, series: [stacked.series[0], { ...stacked.series[1], hidden: true }], overrides: { 'bar:s1:c0': { fill: '000000' } } });
  assert.equal(hid.warnings, undefined);
});

test('overrides name spec indices even when an earlier series is hidden', () => {
  const spec: ChartSpec = { ...stacked, series: [{ ...stacked.series[0], hidden: true }, stacked.series[1], { name: 'Other', values: [1, 2, 3] }], overrides: { 'bar:s1:c0': { fill: '00FF00' } } };
  const sc = lay(spec);
  assert.equal((at(sc, 'bar:s1:c0')[0] as { fill: string }).fill, '00FF00');
  assert.notEqual((at(sc, 'bar:s2:c0')[0] as { fill: string }).fill, '00FF00');
});

test('overrides survive resizing: the resize matrix stays clean with overrides on', () => {
  const spec: ChartSpec = { ...stacked, overrides: { 'label:value:*': { size: 14, bold: true }, 'label:category:*': { rotate: 90 }, 'bar:s0:*': { fill: 'accent3' } } };
  for (const [w, h] of [[864, 408], [600, 300], [420, 260], [320, 220]]) {
    const sc = lay(spec, { x: 10, y: 10, w, h });
    assert.deepEqual(issues(sc), [], `${w}×${h}`);
  }
});

test('remapOverrides keeps overrides on their element through inserts and deletes', () => {
  const o = { 'label:value:s0:c1': { bold: true }, 'label:value:s0:c2': { size: 9 }, 'bar:*:c0': { fill: '000000' }, 'label:series:s1': { bold: true } };
  assert.deepEqual(remapOverrides(o, 'c', (i) => (i === 1 ? -1 : i > 1 ? i - 1 : i)), { 'label:value:s0:c1': { size: 9 }, 'bar:*:c0': { fill: '000000' }, 'label:series:s1': { bold: true } });
  assert.deepEqual(Object.keys(remapOverrides(o, 's', (i) => i + 1)!), ['label:value:s1:c1', 'label:value:s1:c2', 'bar:*:c0', 'label:series:s2']);
});

test('tint / shade colours resolve like DrawingML lumMod / lumOff', () => {
  assert.equal(tintHex('808080', 0), '808080');
  assert.equal(tintHex('000000', 50), '808080');
  assert.equal(tintHex('FFFFFF', -50), '808080');
  assert.equal(resolveElementColor('accent1-darker25', themeOf({})).split('@')[1], 'accent1-darker25');
  assert.equal(resolveElementColor('hlink', themeOf({})), '0563C1@hlink');
});

// ---- Renderers -----------------------------------------------------------------------------------

const styled: ChartSpec = { ...stacked, overrides: {
  'label:category:c0': { rotate: -90, font: '+mj-lt' }, 'label:category:c1': { rotate: 90 },
  'bar:s0:c0': { fill: 'accent1-lighter40', outlineDash: 'dot', outline: '000000' },
} };

test('the .pptx writer draws font, direction, tint and dash per shape', () => {
  const sc = lay(styled);
  const xml = groupXml({ scene: sc, name: 'c', altText: '' }, new Ids(), 'Arial');
  assert.match(xml, /<a:bodyPr vert="vert270"/);
  assert.match(xml, /<a:bodyPr vert="vert"/);
  assert.match(xml, /<a:latin typeface="\+mj-lt"\/>/);
  assert.match(xml, /<a:schemeClr val="accent1"><a:lumMod val="60000"\/><a:lumOff val="40000"\/><\/a:schemeClr>/);
  assert.match(xml, /<a:prstDash val="sysDot"\/>/);
  assert.ok(writePptx([{ scenes: [chartContent(styled, sc, 'Arial')] }]).length > 1000);
});

test('the SVG preview turns text and resolves theme fonts and tints', () => {
  const svg = sceneToSvg([lay(styled)], { w: 960, h: 540 });
  assert.match(svg, /transform="rotate\(-90 /);
  assert.match(svg, /transform="rotate\(90 /);
  assert.ok(!svg.includes('+mj-lt'));
  assert.match(svg, /fill="#[0-9A-F]{6}"/);
});

test('the add-in format carries font, rotation, tint, dash and arrowheads', () => {
  const sc = lay(styled);
  const recs = parseResponse(sc.shapes.map((s, i) => shapeRecord(s, 'Arial', sc.roles![i])).join('\n')).shapes;
  const cat0 = recs.find((r) => r.role === 'label:category:c0')!;
  assert.deepEqual([cat0.rotate, cat0.font], ['-90', '+mj-lt']);
  assert.equal(recs.find((r) => r.role === 'label:category:c1')!.rotate, '90');
  const bar = recs.find((r) => r.role === 'bar:s0:c0')!;
  assert.deepEqual([bar.fill, bar.dash], ['accent1-lighter40', 'dot']);
  const arrow = parseResponse(shapeRecord({ kind: 'line', x1: 0, y1: 0, x2: 9, y2: 9, color: 'hlink', width: 1, arrowStart: true, dash: true }, 'Arial')).shapes[0];
  assert.deepEqual([arrow.arrowStart, arrow.dash, arrow.pattern], ['1', '1', '']);
});

test('VBA reads the new columns and draws them', () => {
  const vba = readFileSync('addin/vba/OpenCharts.bas', 'utf8');
  assert.match(vba, /AddTextbox\(Orientation\(Col\(f, 24\)\)/);
  assert.match(vba, /BeginArrowheadStyle/);
  assert.match(vba, /TintAndShade = shade/);
  assert.match(vba, /SetPattern s, f\(6\), Col\(f, 26\)/);
  assert.match(vba, /Case "dkUpDiag": PatternType = 16/);
  assert.match(vba, /NativeLines\(old\)/);
});

// ---- Convert to plain text -----------------------------------------------------------------------

test('detach (convert to plain text): the label keeps its text through data edits, in the chart', async () => {
  const sc = lay(stacked);
  const spec = detachElement(stacked, sc, 'label:total:c1')!;
  assert.deepEqual(spec.overrides, { 'label:total:c1': { content: [{ text: '70' }], detach: true } });
  assert.deepEqual(validateSpec(spec), []);
  const edited = lay({ ...spec, series: [{ ...spec.series[0], values: [40, 90, 47] }, spec.series[1]] });
  assert.equal((at(edited, 'label:total:c1')[0] as TextShape).text, '70', 'static text');
  assert.equal((at(edited, 'label:total:c2')[0] as TextShape).text, '80', 'other labels follow the data');
  assert.equal(detachElement(stacked, sc, 'label:total:c9'), null);
  assert.ok(validateSpec({ ...stacked, overrides: { 'label:total:c1': { detach: true } } }).some((e) => e.path === '/overrides/label:total:c1/detach'));
  const res = parseResponse(await handleRequest({ op: 'detach', arg: 'label:total:c1', frame: F, spec: JSON.stringify(stacked) }, m));
  assert.deepEqual(res.errors, []);
  assert.ok(res.shapes.some((r) => r.role === 'label:total:c1'), 'stays in the group');
  assert.match(res.tags.find(([k]) => k === 'OC_SPEC')![1], /"label:total:c1":\{"content":\[\{"text":"70"\}\],"detach":true\}/);
  const bad = parseResponse(await handleRequest({ op: 'detach', arg: 'bar:s0:c0', frame: F, spec: JSON.stringify(stacked) }, m));
  assert.equal(bad.errors.length, 1);
});

test('patterns: OOXML presets on segments, in every renderer', () => {
  const spec: ChartSpec = { ...stacked, overrides: { 'bar:s1:*': { pattern: 'dkUpDiag' }, 'bar:s1:c2': { pattern: 'none' } } };
  const sc = lay(spec);
  assert.equal((at(sc, 'bar:s1:c0')[0] as { pattern?: string }).pattern, 'dkUpDiag');
  assert.equal((at(sc, 'bar:s1:c2')[0] as { pattern?: string }).pattern, undefined);
  const xml = groupXml({ scene: sc, name: 'c', altText: '' }, new Ids(), 'Arial');
  assert.match(xml, /<a:pattFill prst="dkUpDiag"><a:fgClr><a:schemeClr val="accent2"\/><\/a:fgClr><a:bgClr><a:schemeClr val="bg1"\/><\/a:bgClr><\/a:pattFill>/);
  assert.match(sceneToSvg([sc], { w: 960, h: 540 }), /<pattern id="p-dkUpDiag-[0-9A-F]{6}"[^>]*>.*fill="url\(#p-dkUpDiag-/s);
  const rec = parseResponse(shapeRecord(at(sc, 'bar:s1:c0')[0], 'Arial', 'bar:s1:c0')).shapes[0];
  assert.equal(rec.pattern, 'dkUpDiag');
  assert.ok(validateSpec({ ...stacked, overrides: { 'bar:s0:c0': { pattern: 'stripes' } } }).length);
  assert.ok(validateSpec({ ...stacked, overrides: { 'label:value:s0:c0': { pattern: 'upDiag' } } }).length, 'labels take no pattern');
});

test('pin: a moved label keeps its place (in the frame) through resizes; unpin resets it', () => {
  const spec: ChartSpec = { ...stacked, overrides: { 'label:total:c0': { pin: { x: 0.5, y: 0.5 } } } };
  for (const f of [F, { x: 0, y: 0, w: 500, h: 300 }]) {
    const t = at(lay(spec, f), 'label:total:c0')[0] as TextShape;
    assert.ok(t.pinned);
    assert.ok(Math.abs(t.x + t.w / 2 - (f.x + f.w / 2)) < 0.01 && Math.abs(t.y + t.h / 2 - (f.y + f.h / 2)) < 0.01, 'centred where pinned');
  }
  const edge = at(lay({ ...stacked, overrides: { 'label:total:c0': { pin: { x: 1, y: 0 } } } }), 'label:total:c0')[0] as TextShape;
  assert.ok(edge.x + edge.w <= F.x + F.w + 0.01 && edge.y >= F.y - 0.01, 'kept inside the frame');
  assert.deepEqual(unpinElement(spec, 'label:total:c0').overrides, {});
  assert.ok(validateSpec({ ...stacked, overrides: { 'label:total:c0': { pin: { x: 2, y: 0 } } } }).some((e) => e.path === '/overrides/label:total:c0/pin/x'));
});

// ---- OC_BASE: native edits ------------------------------------------------------------------------

test('every chart carries OC_BASE: the formatting it drew, per role', () => {
  const sc = lay(stacked);
  const content = chartContent(stacked, sc, 'Arial');
  const base = JSON.parse(chunkedValue(Object.fromEntries(content.tags!), 'OC_BASE')!);
  assert.deepEqual(base['label:value:s1:c0'], { text: '20', font: 'Arial', size: 12, color: 'FFFFFF', bold: false, rotate: 0, fill: 'none', outline: 'none' });
  assert.equal(base['bar:s0:c0'].fill, '1F3A5F@accent1');
  assert.ok(base['label:series:s0'].text);
});

test('mergeNativeEdits: what changed in PowerPoint becomes an override, nothing else', () => {
  const sc = lay(stacked);
  const base = baseOf(sc.shapes, sc.roles!, 'Arial');
  const entries = sceneEntries(sc.shapes, sc.roles!, 'Arial');
  // Unchanged, and the theme alias PowerPoint may report (dk1 for tx1, same hex): no override.
  const same = entries.map((e) => (e.props.color === '222222@tx1' ? { ...e, props: { ...e.props, color: '222222@dk1' } } : e));
  assert.deepEqual(mergeNativeEdits(stacked, base, renderedOf(same)).changes, []);
  const edited = entries.map((e) => {
    if (e.role === 'bar:s0:c1') return { ...e, props: { ...e.props, fill: 'FF0000' } };
    if (e.role === 'label:total:c2') return { ...e, props: { ...e.props, text: 'Big', size: 14, color: '1F3A5F@accent1-darker25' } };
    if (e.role === 'label:category:c0') return { ...e, props: { ...e.props, rotate: 90 as const } };
    if (e.role === 'axis' && e.kind === 'line') return { ...e, props: { ...e.props, lineWidth: 3 } };
    if (e.role === 'label:series:s1' && e.kind === 'text') return { ...e, props: { ...e.props, font: 'Georgia' } };
    return e;
  });
  const { spec, changes } = mergeNativeEdits({ ...stacked, overrides: { 'bar:s0:c1': { outline: '000000' } } }, base, renderedOf(edited));
  assert.deepEqual(changes.sort(), ['axis.lineWidth', 'bar:s0:c1.fill', 'label:category:c0.rotate', 'label:series:s1.font', 'label:total:c2.color', 'label:total:c2.size', 'label:total:c2.text'].sort());
  assert.deepEqual(spec.overrides!['bar:s0:c1'], { outline: '000000', fill: 'FF0000' });
  assert.deepEqual(spec.overrides!['label:total:c2'], { content: [{ text: 'Big' }], size: 14, color: 'accent1-darker25' });
  assert.deepEqual(validateSpec(spec), []);
  // And the next layout draws the user's formatting from the spec.
  const next = lay(spec);
  const t = at(next, 'label:total:c2')[0] as TextShape;
  assert.deepEqual([t.text, t.size], ['Big', 14]);
  assert.equal((at(next, 'label:category:c0')[0] as TextShape).rotate, 90);
});

test('update with base= and native= keeps native edits in OC_SPEC; relayout does not answer "nothing to do"', async () => {
  const sc = lay(stacked);
  const base = baseOf(sc.shapes, sc.roles!, 'Arial');
  const native = sceneEntries(sc.shapes, sc.roles!, 'Arial').map((e) => (e.role === 'bar:s1:c2' ? { ...e, props: { ...e.props, fill: '00B050' } } : e));
  const req = `relayout\n\n${[F.x, F.y, F.w, F.h].join(',')}\n\nbase=${JSON.stringify(base)}\nnative=${JSON.stringify(native)}\n${JSON.stringify(stacked)}`;
  const { parseRequest } = await import('../src/addin/engine.ts');
  const parsed = parseRequest(req);
  assert.ok(parsed.base && parsed.native);
  assert.ok(!parsed.context!.includes('base='), 'base / native are not layout context');
  const res = parseResponse(await handleRequest(parsed, m));
  assert.equal(res.nothing, false);
  assert.match(res.tags.find(([k]) => k === 'OC_SPEC')![1], /"bar:s1:c2":\{"fill":"00B050"\}/);
  assert.equal(res.shapes.find((r) => r.role === 'bar:s1:c2')!.fill, '00B050');
  // Without edits the relayout is still a no-op.
  const quiet = parseRequest(`relayout\n\n${[F.x, F.y, F.w, F.h].join(',')}\n\nbase=${JSON.stringify(base)}\nnative=${JSON.stringify(sceneEntries(sc.shapes, sc.roles!, 'Arial'))}\n${JSON.stringify(stacked)}`);
  assert.equal(parseResponse(await handleRequest(quiet, m)).nothing, true);
});
