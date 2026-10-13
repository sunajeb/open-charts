// Theme colours (style slots → schemeClr / ObjectThemeColor), style validation, number defaults,
// same-scale groups and slides with several charts.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  DEFAULT_THEME, defaultStyle, layoutChart, layoutGroup, resolveColor, validateDeck, validateSpec,
  type ChartSpec, type RectShape, type Scene, type TextShape,
} from '../src/core/index.ts';
import { createMeasurer } from '../src/fonts.ts';
import { buildDeck, CHART_FRAME, slideCharts, type DeckSpec } from '../src/deck.ts';
import { sceneToSvg } from '../src/render/svg.ts';
import { unzip } from '../src/ooxml/unzip.ts';
import { importDeck, readCharts, readThemeColors } from '../src/ooxml/pptx-read.ts';
import { parseResponse, shapeRecord } from '../src/addin/format.ts';
import { handleRequest, parseRequest } from '../src/addin/engine.ts';

const m = createMeasurer();
const F = CHART_FRAME;
const stacked = (values = [20, 30, 47], extra: Partial<ChartSpec> = {}): ChartSpec => ({
  type: 'stacked', categories: ['2022', '2023', '2024'],
  series: [{ name: 'A', values }, { name: 'B', values: values.map((v) => Math.round(v / 5)) }], ...extra,
});
const bars = (s: Scene) => s.shapes.filter((x, i): x is RectShape => x.kind === 'rect' && !!s.roles?.[i].startsWith('bar:'));
const barOf = (s: Scene, role: string) => s.shapes[s.roles!.indexOf(role)] as RectShape;

// ---- Theme colours ---------------------------------------------------------------------------

test('theme: slots resolve to hex and keep the slot; hex passes through', () => {
  assert.equal(resolveColor('accent2', DEFAULT_THEME), '3E6A9E@accent2');
  assert.equal(resolveColor('tx1', DEFAULT_THEME), '222222@tx1', 'tx1 resolves through dk1');
  assert.equal(resolveColor('#c0504d', DEFAULT_THEME), 'c0504d');
  assert.equal(resolveColor('3E6A9E@accent2', DEFAULT_THEME), '3E6A9E@accent2', 'resolving twice is harmless');
});

test('theme: the default style draws series in theme accents, text in tx1, separators in bg1', () => {
  const s = layoutChart(stacked(), F, m);
  assert.equal(barOf(s, 'bar:s0:c0').fill, '1F3A5F@accent1');
  assert.equal(barOf(s, 'bar:s1:c0').fill, '3E6A9E@accent2');
  assert.equal(barOf(s, 'bar:s0:c0').stroke, 'FFFFFF@bg1');
  const total = s.shapes[s.roles!.indexOf('label:total:c0')] as TextShape;
  assert.equal(total.color, '222222@tx1');
  // The SVG preview draws the resolved hex, never the slot.
  const svg = sceneToSvg([s], { w: 960, h: 540 });
  assert.ok(svg.includes('fill="#1F3A5F"') && !svg.includes('@'));
});

test('theme: style.theme sets what the slots stand for, and label contrast follows it', () => {
  const light = layoutChart(stacked(), F, m, { ...defaultStyle, theme: { accent1: 'F2E394' } });
  assert.equal(barOf(light, 'bar:s0:c0').fill, 'F2E394@accent1');
  const label = (s: Scene) => (s.shapes[s.roles!.indexOf('label:value:s0:c2')] as TextShape).color;
  assert.equal(label(layoutChart(stacked(), F, m)), 'FFFFFF', 'white on the default navy');
  assert.equal(label(light), '222222', 'dark on a light accent');
});

test('theme: series colours and the chart\'s own style may name slots', () => {
  const spec = stacked([20, 30, 47], { series: [{ name: 'A', values: [1, 2, 3], color: 'accent4' }, { name: 'B', values: [1, 1, 1], color: '#C0504D' }], style: { palette: ['dk2'] } });
  assert.deepEqual(validateSpec(spec), []);
  const s = layoutChart(spec, F, m);
  assert.equal(barOf(s, 'bar:s0:c0').fill, 'B9CDE5@accent4');
  assert.equal(barOf(s, 'bar:s1:c0').fill, 'C0504D');
  const own = layoutChart({ ...stacked(), style: { palette: ['dk2', 'lt2'] } }, F, m);
  assert.equal(barOf(own, 'bar:s1:c0').fill, 'EEEEEE@lt2');
});

test('theme: the writer emits schemeClr for slots, srgbClr for hex, and the style\'s theme in theme1.xml', () => {
  const deck: DeckSpec = { style: { theme: { accent1: '0B5394' }, negative: 'C00000' }, slides: [{ chart: stacked() }] };
  const files = unzip(buildDeck(deck, m).pptx);
  const slide = files.get('ppt/slides/slide1.xml')!.text();
  assert.match(slide, /<a:solidFill><a:schemeClr val="accent1"\/><\/a:solidFill>/);
  assert.match(slide, /<a:schemeClr val="tx1"\/>/);
  assert.match(slide, /<a:schemeClr val="bg1"\/>/);
  assert.match(slide, /<a:srgbClr val="7F7F7F"\/>/, 'hex colours stay RGB');
  const theme = files.get('ppt/theme/theme1.xml')!.text();
  assert.match(theme, /<a:accent1><a:srgbClr val="0B5394"\/><\/a:accent1>/);
  assert.match(theme, /<a:accent2><a:srgbClr val="3E6A9E"\/><\/a:accent2>/);
  assert.deepEqual(readThemeColors(buildDeck(deck, m).pptx), { ...DEFAULT_THEME, accent1: '0B5394' });
});

test('theme: the add-in line format carries theme colour names for VBA', () => {
  const rec = shapeRecord({ kind: 'rect', x: 0, y: 0, w: 1, h: 1, fill: '1F3A5F@accent1', stroke: 'FFFFFF@bg1' }, 'Arial');
  const r = parseResponse(`OC\t2\n${rec}\n`).shapes[0];
  assert.equal(r.fill, 'accent1');
  assert.equal(r.lineColor, 'bg1');
  assert.equal(parseResponse(`OC\t2\n${shapeRecord({ kind: 'line', x1: 0, y1: 0, x2: 1, y2: 1, color: '#7f7f7f', width: 1 }, 'Arial')}\n`).shapes[0].lineColor, '7F7F7F');
});

test('theme: the add-in resolves slots in the slide\'s theme (request line theme=)', async () => {
  const spec = JSON.stringify(stacked());
  // ThemeColorScheme order: dk1 lt1 dk2 lt2 accent1..6 hlink folHlink. accent1 is a pale yellow here.
  const theme = 'theme=000000,FFFFFF,44546A,E7E6E6,F2E394,ED7D31,A5A5A5,FFC000,5B9BD5,70AD47,0563C1,954F72';
  const req = parseRequest(['update', '', '48,96,864,408', '', theme, spec].join('\n'));
  assert.equal(req.theme?.accent1, 'F2E394');
  assert.equal(req.spec, spec);
  const r = parseResponse(await handleRequest(req, m));
  assert.deepEqual(r.errors, []);
  const label = r.shapes.find((x) => x.role === 'label:value:s0:c2')!;
  assert.equal(label.color, '222222', 'contrast decided on the slide\'s accent1, not the default navy');
  assert.equal(r.shapes.find((x) => x.role === 'bar:s0:c0')!.fill, 'accent1');
  // Old requests (no parameter lines) still read the spec from line 5.
  assert.equal(parseRequest(['update', '', '', '', spec].join('\n')).spec, spec);
});

test('theme: a hidden row leaves the others on their own theme colours', () => {
  const spec: ChartSpec = { type: 'stacked', categories: ['a', 'b'], series: [
    { name: 'A', values: [1, 2] }, { name: 'B', values: [1, 2], hidden: true }, { name: 'C', values: [1, 2] }, { name: 'D', values: [1, 1], color: 'accent6' }] };
  const s = layoutChart(spec, F, m);
  assert.equal(barOf(s, 'bar:s0:c0').fill, '1F3A5F@accent1');
  assert.equal(barOf(s, 'bar:s2:c0').fill, '7FA3CC@accent3', 'row 3 keeps accent3');
  assert.equal(barOf(s, 'bar:s3:c0').fill, 'D9D9D9@accent6');
});

// ---- Validation --------------------------------------------------------------------------------

test('style validation: colours are hex or theme slots; theme and numbers are checked', () => {
  const errs = validateDeck({
    style: { palette: ['accent1', 'accent9'], text: 'tx1', theme: { accent1: 'blue', hlink: '000000' }, numbers: { locale: 'xx-XX', decimals: 11, unit: 'm' } },
    slides: [{ chart: { ...stacked(), style: { axis: 'nope' } } }],
  });
  const paths = errs.map((e) => e.path).sort();
  assert.deepEqual(paths, ['/slides/0/chart/style/axis', '/style/numbers/decimals', '/style/numbers/locale', '/style/numbers/unit', '/style/palette/1', '/style/theme/accent1', '/style/theme/hlink'].sort());
  assert.match(errs.find((e) => e.path === '/style/palette/1')!.message, /theme colour \(accent1, .*tx1/);
});

test('style.numbers: number defaults for every chart; the chart\'s own options win', () => {
  const spec: ChartSpec = { type: 'stacked', categories: ['a', 'b'], series: [{ name: 'A', values: [1.25, 2.5] }] };
  const texts = (s: Scene) => s.shapes.filter((x): x is TextShape => x.kind === 'text').map((t) => t.text);
  const de = layoutChart(spec, F, m, { ...defaultStyle, numbers: { locale: 'de-DE', decimals: 1 } });
  assert.ok(texts(de).includes('1,3') || texts(de).includes('1,2'), texts(de).join(' '));
  const own = layoutChart({ ...spec, options: { locale: 'en-US' } }, F, m, { ...defaultStyle, numbers: { locale: 'de-DE', decimals: 1 } });
  assert.ok(texts(own).some((t) => /^\d\.\d$/.test(t)), texts(own).join(' '));
  // Harvey balls do not read decimals: the default does not apply (and is not an error).
  assert.doesNotThrow(() => layoutChart({ type: 'harvey', categories: ['a'], series: [{ name: 'A', values: [0.5] }] }, F, m, { ...defaultStyle, numbers: { decimals: 2 } }));
});

// ---- Same-scale groups -------------------------------------------------------------------------

const half = (i: number) => ({ ...F, x: F.x + i * (F.w / 2 + 12), w: F.w / 2 - 12 });
const wf: ChartSpec = { type: 'waterfall', categories: ['2023', 'Price', 'Volume', '2024'], series: [{ name: 'Delta', values: [3, 2, 2, 'e'] }], options: { scaleGroup: 'rev' } };

test('same scale: charts in one group draw one unit equally long; the largest scale is kept', () => {
  const a = stacked([20, 30, 47], { options: { scaleGroup: 'rev' } });
  const [sa, sw] = layoutGroup([{ spec: a, frame: half(0) }, { spec: wf, frame: half(1) }], m);
  assert.ok(sa.valueScale && sw.valueScale);
  assert.ok(Math.abs(sa.valueScale.ptPerUnit - sw.valueScale.ptPerUnit) < 1e-6, `${sa.valueScale.ptPerUnit} vs ${sw.valueScale.ptPerUnit}`);
  assert.equal(sa.valueScale.ptPerUnit, sa.valueScale.own, 'the stacked chart keeps its own scale');
  assert.ok(sw.valueScale.own > sw.valueScale.ptPerUnit * 5, 'the waterfall alone would be far taller');
  // Physical heights: 20 in the stacked chart vs the waterfall's 3 → ratio 20/3.
  const ppu = sa.valueScale.ptPerUnit;
  assert.ok(Math.abs(barOf(sa, 'bar:s0:c0').h - 20 * ppu) < 0.05);
  assert.ok(Math.abs(barOf(sw, 'bar:s0:c0').h - 3 * ppu) < 0.05);
  // Alone, nothing changes.
  const alone = layoutChart(a, half(0), m);
  assert.deepEqual(bars(alone).map((b) => b.h), bars(sa).map((b) => b.h));
});

test('same scale: works with value axes on and both orientations; other groups are independent', () => {
  for (const orientation of ['vertical', 'horizontal'] as const) {
    const a = stacked([20, 30, 47], { options: { scaleGroup: 'g', valueAxis: true, orientation } });
    const b = stacked([2, 3, 4], { options: { scaleGroup: 'g', valueAxis: true, orientation } });
    const c = stacked([2, 3, 4], { options: { scaleGroup: 'other', orientation } });
    const [sa, sb, sc] = layoutGroup([{ spec: a, frame: half(0) }, { spec: b, frame: half(1) }, { spec: c, frame: half(1) }], m);
    assert.ok(Math.abs(sa.valueScale!.ptPerUnit - sb.valueScale!.ptPerUnit) < 1e-6, orientation);
    assert.equal(sc.valueScale!.ptPerUnit, sc.valueScale!.own, 'a group of one keeps its own scale');
    const len = (r: RectShape) => (orientation === 'vertical' ? r.h : r.w);
    assert.ok(Math.abs(len(barOf(sb, 'bar:s0:c0')) / len(barOf(sa, 'bar:s0:c0')) - 2 / 20) < 0.01, orientation);
  }
});

test('same scale: line, area and combo charts take part; pie does not read the option', () => {
  const line: ChartSpec = { type: 'line', categories: ['a', 'b', 'c'], series: [{ name: 'L', values: [5, 8, 10] }], options: { scaleGroup: 'g' } };
  const area: ChartSpec = { type: 'area', categories: ['a', 'b', 'c'], series: [{ name: 'L', values: [50, 80, 100] }], options: { scaleGroup: 'g' } };
  const combo: ChartSpec = { type: 'combo', categories: ['a', 'b', 'c'], series: [{ name: 'C', values: [20, 30, 40] }, { name: 'L', values: [10, 12, 14], as: 'line' }], options: { scaleGroup: 'g' } };
  const scenes = layoutGroup([line, area, combo].map((spec, i) => ({ spec, frame: { ...F, x: F.x + i * 290, w: 270 } })), m);
  const ppu = scenes.map((s) => s.valueScale!.ptPerUnit);
  assert.ok(ppu.every((p) => Math.abs(p - ppu[0]) < 1e-6), ppu.join(' '));
  assert.match(validateSpec({ type: 'pie', categories: ['a'], series: [{ name: 'P', values: [1] }], options: { scaleGroup: 'g' } }).map((e) => e.message).join(), /not used by pie/);
});

// ---- Several charts on a slide -------------------------------------------------------------

test('deck: a slide holds several charts; without frames they share the chart area', () => {
  const deck: DeckSpec = {
    slides: [{ title: 'Two', charts: [{ chart: stacked([20, 30, 47], { options: { scaleGroup: 'rev' } }) }, { chart: wf }] }, { chart: stacked() }],
  };
  assert.deepEqual(validateDeck(deck), []);
  const frames = slideCharts(deck.slides[0]).map((c) => c.frame);
  assert.equal(frames[0].x, F.x);
  assert.ok(Math.abs(frames[1].x + frames[1].w - (F.x + F.w)) < 1e-9);
  assert.ok(frames[0].x + frames[0].w < frames[1].x);
  const { pptx, slides } = buildDeck(deck, m);
  assert.equal(slides[0].scenes.length, 2);
  assert.deepEqual(slides[0].issues, []);
  const charts = readCharts(pptx);
  assert.equal(charts[0].charts.length, 2);
  const tags = (i: number) => Object.fromEntries(unzipTags(pptx, i));
  assert.equal(tags(0).OC_SGROUP, 'rev');
  assert.ok(Number(tags(0).OC_SCALE) > 0 && !/e/i.test(tags(0).OC_SCALE));
  // import gives the slide back with its charts and their frames
  const back = importDeck(pptx).deck;
  assert.equal(back.slides[0].charts?.length, 2);
  assert.deepEqual(back.slides[0].charts?.[1].frame, { ...frames[1], w: +frames[1].w.toFixed(2), x: +frames[1].x.toFixed(2) });
  assert.deepEqual(back.slides[1].chart, deck.slides[1].chart);
});

test('deck: chart and charts together, empty charts and bad frames are errors', () => {
  const errs = validateDeck({ slides: [
    { chart: stacked(), charts: [{ chart: stacked() }] },
    { charts: [] },
    { charts: [{ frame: { x: 0, y: 0, w: -1 }, chart: stacked(), extra: 1 }] },
  ] });
  assert.deepEqual(errs.map((e) => e.path), ['/slides/0/chart', '/slides/1/charts', '/slides/2/charts/0/extra', '/slides/2/charts/0/frame/w', '/slides/2/charts/0/frame/h']);
});

/** Group-level tags of the n-th chart in the deck (tag parts in order of writing). */
function unzipTags(pptx: Uint8Array, n: number): [string, string][] {
  const files = unzip(pptx);
  const groups = [...files.keys()].filter((k) => /^ppt\/tags\/tag\d+\.xml$/.test(k)).sort((a, b) => Number(a.match(/\d+/)![0]) - Number(b.match(/\d+/)![0]))
    .map((k) => [...files.get(k)!.text().matchAll(/<p:tag name="([^"]+)" val="([^"]*)"\/>/g)].map((x) => [x[1], x[2]] as [string, string]))
    .filter((t) => t.some(([k]) => k === 'OC_V'));
  return groups[n];
}
