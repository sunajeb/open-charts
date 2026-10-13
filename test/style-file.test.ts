// Style files (--style): validation and sanitising of untrusted files, layering, per-type overrides,
// label colours and line styles by role, and the rule that switching styles changes appearance only.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import {
  STYLE_FILE_KEYS, STYLE_KEYS, defaultStyle, layoutChart, mergeStyles, parseStyleFile, styleFor, validateStyleFile,
  type ChartSpec, type LineShape, type Scene, type Shape, type TextShape,
} from '../src/core/index.ts';
import { createMeasurer } from '../src/fonts.ts';
import { applyStyleFile, buildDeck, CHART_FRAME, type DeckSpec } from '../src/deck.ts';
import { readCharts } from '../src/ooxml/pptx-read.ts';
import { unzip } from '../src/ooxml/unzip.ts';

const m = createMeasurer();
const showcase: DeckSpec = JSON.parse(readFileSync('examples/style-showcase.json', 'utf8'));
const ember = parseStyleFile(readFileSync('examples/styles/ember.json', 'utf8'));
const brand = parseStyleFile(readFileSync('examples/styles/brand.json', 'utf8'));

/** Geometry of a shape without its colours (what a style must never change). */
function geometry(s: Shape) {
  const { color: _c, fill: _f, stroke: _s, width: _w, dash: _d, ...rest } = s as unknown as Record<string, unknown>;
  return rest;
}
const colours = (s: Shape) => JSON.stringify([(s as TextShape).color, (s as TextShape).fill, (s as TextShape).stroke, (s as LineShape).width, (s as LineShape).dash]);

test('style files: the example files are valid', () => {
  for (const f of ['ember', 'brand']) assert.deepEqual(validateStyleFile(JSON.parse(readFileSync(`examples/styles/${f}.json`, 'utf8'))), [], f);
});

test('style files: unknown keys, bad values and unknown chart types are listed with paths', () => {
  const errors = validateStyleFile({
    name: 'x', colour: 'red', palette: [], fontSize: 200, labels: { value: 'red', price: '000000' },
    lines: { axis: { width: 20, dash: 'yes' }, frame: {} }, types: { line: { palette: ['nope'] }, radar: {} },
  });
  const paths = errors.map((e) => e.path).sort();
  assert.deepEqual(paths, ['/colour', '/fontSize', '/labels/price', '/labels/value', '/lines/axis/dash', '/lines/axis/width', '/lines/frame', '/palette', '/types/line/palette/0', '/types/radar'].sort());
});

test('style files: untrusted input is size-limited and sanitised', () => {
  assert.throws(() => parseStyleFile(' '.repeat(300 * 1024)), /larger than 256 KB/);
  assert.throws(() => parseStyleFile('{nope'), /not JSON/);
  assert.throws(() => parseStyleFile('[]'), /must be a JSON object/);
  assert.throws(() => parseStyleFile('{"__proto__": {"font": "x"}}'), /unknown property "__proto__"/);
  assert.throws(() => parseStyleFile(JSON.stringify({ font: 'A'.repeat(65) })), /\/font: must be at most 64/);
  assert.throws(() => parseStyleFile(JSON.stringify({ font: 'Arial"/><a:x' })), /\/font: must be a plain font name/);
  assert.throws(() => parseStyleFile(JSON.stringify({ palette: Array(65).fill('accent1') })), /at most 64 colours/);
  assert.throws(() => parseStyleFile(JSON.stringify({ types: { line: { types: {} } } })), /\/types\/line\/types: unknown property/);
  assert.equal(({} as Record<string, unknown>).font, undefined, 'no prototype pollution');
});

test('style files: layering merges nested keys one level deep; types override the file per chart type', () => {
  const merged = mergeStyles({ lines: { axis: { color: '111111', width: 1 } }, labels: { total: 'accent1' } }, { lines: { axis: { width: 2 } }, labels: { value: '222222' } });
  assert.deepEqual(merged, { lines: { axis: { color: '111111', width: 2 } }, labels: { total: 'accent1', value: '222222' } });
  assert.deepEqual(styleFor(ember, 'line').palette, ['8C2D19', '5B5B5B', 'D9542B', 'BFBFBF']);
  assert.deepEqual(styleFor(ember, 'stacked').palette, ember.palette);
  assert.equal(styleFor(ember, 'line').name, undefined, 'the envelope is not a style key');
  // Deck style, then the chart's own style, win over the file.
  const deck = applyStyleFile({ style: { font: 'Verdana' }, slides: [{ chart: { type: 'line', categories: ['a'], series: [{ name: 'x', values: [1] }], style: { palette: ['00FF00'] } } }] }, ember);
  assert.equal(deck.slides[0].chart!.style!.font, 'Verdana');
  assert.deepEqual(deck.slides[0].chart!.style!.palette, ['00FF00']);
  assert.equal(deck.slides[0].chart!.style!.fontSize, 11);
});

test('style files: schema/style.schema.json lists every style key and the file keys', () => {
  const schema = JSON.parse(readFileSync('schema/style.schema.json', 'utf8'));
  assert.deepEqual(Object.keys(schema.properties).sort(), [...STYLE_KEYS, ...STYLE_FILE_KEYS].sort());
});

test('labels and lines: recolour by role, keep contrast labels, restyle axis / gridline / decoration lines', () => {
  const spec: ChartSpec = {
    type: 'stacked', categories: ['2022', '2023', '2024'],
    series: [{ name: 'A', values: [20, 30, 47] }, { name: 'B', values: [4, 6, 9] }],
    options: { cagr: { from: 0, to: 2 }, valueAxis: true },
  };
  const plain = layoutChart(spec, CHART_FRAME, m);
  const styled = layoutChart({ ...spec, style: { labels: { total: 'C00000', category: '00B050', decoration: '7030A0' }, lines: { axis: { color: '0070C0', width: 2 }, decoration: { color: '7030A0', width: 1.5, dash: true } } } }, CHART_FRAME, m);
  const at = (s: Scene, re: RegExp) => s.shapes.filter((_, i) => re.test(s.roles![i]));
  assert.ok(at(styled, /^label:total/).every((t) => (t as TextShape).color === 'C00000'));
  assert.ok(at(styled, /^label:category/).every((t) => (t as TextShape).color === '00B050'));
  const axisLines = at(styled, /^axis/).filter((s) => s.kind === 'line') as LineShape[];
  assert.ok(axisLines.length && axisLines.every((l) => l.color === '0070C0' && l.width === 2));
  const arrow = at(styled, /^arrow:cagr/);
  assert.ok(arrow.some((s) => s.kind === 'line' && s.color === '7030A0' && s.dash === true && s.width === 1.5));
  assert.ok(arrow.some((s) => s.kind === 'text' && s.color === '7030A0'), 'the CAGR bubble text');
  // Value labels inside dark bars were chosen for contrast: unchanged.
  const values = (s: Scene) => at(s, /^label:value/).map((t) => (t as TextShape).color);
  assert.deepEqual(values(styled), values(plain));
});

test('switching styles changes appearance only: same roles and geometry for every chart type, other colours', () => {
  // Same font and size as the default, other colours, label colours and line styles.
  const recolour = { ...ember, font: defaultStyle.font, fontSize: defaultStyle.fontSize, numbers: undefined };
  const a = buildDeck(showcase, m).slides, b = buildDeck(showcase, m, brand).slides, c = buildDeck(showcase, m, recolour).slides;
  assert.equal(a.length, 17);
  a.forEach((sl, i) => {
    const type = showcase.slides[i].chart!.type;
    for (const other of [b[i], c[i]]) {
      assert.deepEqual(other.scene.roles, sl.scene.roles, `${type}: roles`);
      assert.deepEqual(other.scene.shapes.map(geometry), sl.scene.shapes.map(geometry), `${type}: geometry`);
    }
    assert.notDeepEqual(c[i].scene.shapes.map(colours), sl.scene.shapes.map(colours), `${type}: colours change`);
  });
});

test('switching styles with another font: same roles, the font is written to the deck and its tags', () => {
  const a = buildDeck(showcase, m), b = buildDeck(showcase, m, ember);
  // Labels may hide or wrap differently in another font's widths; marks and lines keep their roles.
  const marks = (r: string[] = []) => r.filter((x) => /^(bar|wedge|area|ball|marker|line):/.test(x));
  a.slides.forEach((sl, i) => assert.deepEqual(marks(b.slides[i].scene.roles), marks(sl.scene.roles)));
  const charts = readCharts(b.pptx);
  assert.equal(charts.length, 17);
  for (const s of charts) {
    assert.equal(s.charts[0].font, 'Georgia');
    const spec = JSON.parse(s.charts[0].spec!);
    assert.equal(spec.style.fontSize, 11, 'the stored spec carries the full style, so the add-in re-lays out the same');
  }
  const xml = unzip(b.pptx).get('ppt/slides/slide1.xml')!.text();
  assert.match(xml, /typeface="Georgia"/);
  assert.doesNotMatch(xml, /typeface="Arial"/);
});

test('CLI: --style builds with the file; an invalid file exits 1 listing every problem', () => {
  const run = (...args: string[]) => spawnSync(process.execPath, ['--experimental-strip-types', '--no-warnings', 'src/cli.ts', ...args], { encoding: 'utf8' });
  const ok = run('examples/deck.json', '--style', 'examples/styles/ember.json', '-o', 'out/test-style/deck');
  assert.equal(ok.status, 0, ok.stderr);
  const bad = run('examples/deck.json', '--style', 'schema/deck.schema.json', '-o', 'out/test-style/bad');
  assert.equal(bad.status, 1);
  assert.match(bad.stderr, /Invalid schema\/deck\.schema\.json:\n/);
  assert.match(bad.stderr, /^ {2}\/title: unknown property "title"$/m);
});
