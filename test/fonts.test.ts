// Any font (tasks B1–B4): the font file reader (core/sfnt.ts), the resolver and index (src/fonts.ts),
// substitutes and the wide fallback (core/fonts.ts), theme fonts, and the wiring into the add-in
// engine and the CLI. Synthetic fonts (test/helpers/make-font.ts) make the resolver's cases exact;
// the fonts installed on this computer are measured where present (skipped, not failed, elsewhere).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { makeFont, type FaceSpec } from './helpers/make-font.ts';
import { bytesReader, extractFace, faceMetrics, readFaces } from '../src/core/sfnt.ts';
import { fontMetrics, measurerFromMetrics } from '../src/core/text.ts';
import { isThemeFont, resolveFont, substitutesFor, WIDE_FALLBACK } from '../src/core/fonts.ts';
import { checkScene, defaultStyle, layoutChart, type ChartSpec, type Style, type TextShape } from '../src/core/index.ts';
import { samples } from '../src/core/samples.ts';
import { cachedMetrics, createMeasurer, familyMetrics, findFont, fontDirs, fontIndex, reportOnce } from '../src/fonts.ts';
import { handleRequest, parseRequest } from '../src/addin/engine.ts';
import { loadMeasurer } from '../src/officejs/measure.ts';
import { unzip } from '../src/ooxml/unzip.ts';
import { parse } from '../vendor/opentype.module.js';

const tmp = mkdtempSync(join(tmpdir(), 'oc-fonts-'));
process.on('exit', () => rmSync(tmp, { recursive: true, force: true }));
const fixtures = join(tmp, 'fonts');
const cache = join(tmp, 'cache');
mkdirSync(join(fixtures, 'CloudFonts', 'Fixture Cloud'), { recursive: true });
process.env.OPEN_CHARTS_CACHE = cache;

const sans = (o: Partial<FaceSpec>): FaceSpec => ({ family: 'Fixture Sans', advances: { a: 400, b: 450 }, ...o });
writeFileSync(join(fixtures, 'fixture-sans.ttc'), makeFont(
  sans({}), sans({ bold: true, advances: { a: 520, b: 560 } }), sans({ italic: true, advances: { a: 410 } }), sans({ bold: true, italic: true, advances: { a: 530 } }),
));
writeFileSync(join(fixtures, 'FixtureVar.ttf'), makeFont({ family: 'Fixture Var', variable: { min: 100, def: 400, max: 900, delta: 100 }, advances: { a: 500 } }));
writeFileSync(join(fixtures, 'FixtureSerif-Book.otf'), makeFont({ family: 'Fixture Serif Book', typoFamily: 'Fixture Serif', subfamily: 'Regular', fullName: 'Fixture Serif Book', cff: true, ascent: 800, descent: 200, lineGap: 300 }));
writeFileSync(join(fixtures, 'CloudFonts', 'Fixture Cloud', '8f3a2c1d.ttf'), makeFont({ family: 'Fixture Cloud', advances: { a: 700 } }));

/** Runs `f` with the font folders set to the fixtures only. */
function withFixtures<T>(f: () => T): T {
  const before = process.env.OPEN_CHARTS_FONT_DIRS;
  const restore = () => { if (before === undefined) delete process.env.OPEN_CHARTS_FONT_DIRS; else process.env.OPEN_CHARTS_FONT_DIRS = before; };
  process.env.OPEN_CHARTS_FONT_DIRS = fixtures;
  let out: T;
  try { out = f(); } catch (e) { restore(); throw e; }
  if (out instanceof Promise) return out.finally(restore) as T;
  restore();
  return out;
}

const ab = (u8: Uint8Array) => u8.buffer.slice(u8.byteOffset, u8.byteOffset + u8.byteLength) as ArrayBuffer;
const file = (p: string) => ab(readFileSync(p));

// ---- B2: the font file reader -------------------------------------------------------------------

test('sfnt: names, weights and styles of every face of a collection; one face extracts as a plain font', () => {
  const bytes = file(join(fixtures, 'fixture-sans.ttc'));
  const faces = readFaces(bytesReader(bytes));
  assert.deepEqual(faces.map((f) => [f.index, f.families[0], f.subfamily, f.weight, f.bold, f.italic]), [
    [0, 'Fixture Sans', 'Regular', 400, false, false], [1, 'Fixture Sans', 'Bold', 700, true, false],
    [2, 'Fixture Sans', 'Italic', 400, false, true], [3, 'Fixture Sans', 'Bold Italic', 700, true, true],
  ]);
  const one = extractFace(bytes, 1);
  assert.deepEqual(readFaces(bytesReader(one)).map((f) => f.subfamily), ['Bold']);
  assert.deepEqual(fontMetrics(one), fontMetrics(bytes, 1));
  assert.equal(fontMetrics(bytes, 1).advances[fontMetrics(bytes, 1).codes.indexOf(97)], 520);
  const otf = readFaces(bytesReader(file(join(fixtures, 'FixtureSerif-Book.otf'))))[0];
  assert.equal(otf.outlines, 'cff');
  assert.deepEqual(otf.typoFamilies, ['Fixture Serif']);
});

test('sfnt: a variable font measures its default instance, and any weight from HVAR (avar-free normalisation)', () => {
  const bytes = file(join(fixtures, 'FixtureVar.ttf'));
  assert.deepEqual(readFaces(bytesReader(bytes))[0].wght, { min: 100, def: 400, max: 900 });
  const at = (wght?: number) => { const m = fontMetrics(bytes, 0, wght); return m.advances[m.codes.indexOf(97)]; };
  assert.equal(at(), 500);
  assert.equal(at(400), 500);
  assert.equal(at(700), 560); // (700 − 400) / (900 − 400) = 0.6 of the +100 delta
  assert.equal(at(900), 600);
  assert.equal(at(100), 500); // the region only covers 0 → 1
});

test('sfnt: line height follows PowerPoint\'s rule (win metrics plus the extra hhea gap)', () => {
  const m = fontMetrics(file(join(fixtures, 'FixtureSerif-Book.otf')));
  assert.equal(m.lineHeight, (800 + 200 + 300) / 1000);
});

test('sfnt reads exactly what opentype.js reads (advances, cmap, line height) for real fonts', () => {
  const paths = [
    'vendor/fonts/LiberationSans-Regular.ttf', 'vendor/fonts/LiberationSans-Bold.ttf',
    '/System/Library/Fonts/Supplemental/Arial.ttf', '/System/Library/Fonts/Supplemental/Georgia.ttf',
    '/System/Library/Fonts/Supplemental/Verdana.ttf', '/System/Library/Fonts/Supplemental/Times New Roman.ttf',
    '/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts/Calibri.ttf',
    '/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts/Aptos.ttf',
    'C:\\Windows\\Fonts\\arial.ttf', 'C:\\Windows\\Fonts\\segoeui.ttf', 'C:\\Windows\\Fonts\\calibri.ttf',
  ].filter(existsSync);
  assert.ok(paths.length >= 2);
  for (const p of paths) {
    const bytes = file(p);
    const font = parse(bytes) as { unitsPerEm: number; glyphs: { get(i: number): { advanceWidth?: number } }; tables: { cmap: { glyphIndexMap: Record<string, number> } } };
    const m = faceMetrics(bytesReader(bytes));
    const map = font.tables.cmap.glyphIndexMap;
    assert.deepEqual(m.codes, Object.keys(map).map(Number).sort((a, b) => a - b), p);
    assert.deepEqual(m.advances, m.codes.map((c) => font.glyphs.get(map[c])?.advanceWidth ?? 0), p);
    assert.equal(m.unitsPerEm, font.unitsPerEm);
  }
});

// ---- B1: the resolver and its index ------------------------------------------------------------

test('resolver: family names (any case, typographic family, full name) pick the right face; files may have any name', () => withFixtures(() => {
  const faces = fontIndex();
  const pick = (family: string, o: { bold?: boolean; italic?: boolean } = {}) => {
    const f = findFont(family, o, faces);
    return f && `${f.path.split(/[\\/]/).pop()}#${f.index}${f.wght ? `@${f.wght}` : ''}`;
  };
  assert.equal(pick('Fixture Sans'), 'fixture-sans.ttc#0');
  assert.equal(pick('fixture  SANS'), 'fixture-sans.ttc#0');
  assert.equal(pick('Fixture Sans', { bold: true }), 'fixture-sans.ttc#1');
  assert.equal(pick('Fixture Sans', { italic: true }), 'fixture-sans.ttc#2');
  assert.equal(pick('Fixture Sans', { bold: true, italic: true }), 'fixture-sans.ttc#3');
  assert.equal(pick('Fixture Var'), 'FixtureVar.ttf#0');
  assert.equal(pick('Fixture Var', { bold: true }), 'FixtureVar.ttf#0@700', 'bold from the weight axis');
  assert.equal(pick('Fixture Serif'), 'FixtureSerif-Book.otf#0', 'typographic family');
  assert.equal(pick('Fixture Serif Book'), 'FixtureSerif-Book.otf#0');
  assert.equal(pick('Fixture Serif', { bold: true }), null, 'no bold: the measurer makes it');
  assert.equal(pick('Fixture Cloud'), '8f3a2c1d.ttf#0', 'a cloud font in a nested folder with a hashed name');
  assert.equal(pick('Not Installed Sans'), null);
}));

test('resolver: the index is cached on disk and rebuilt when a font folder changes', () => withFixtures(() => {
  fontIndex();
  const saved = readdirSync(cache).filter((f) => f.startsWith('font-index-')).map((f) => JSON.parse(readFileSync(join(cache, f), 'utf8')))
    .find((x) => x.roots.join() === fixtures);
  assert.ok(saved, 'index file in the cache directory');
  assert.ok(Object.keys(saved.files).some((p: string) => p.endsWith('fixture-sans.ttc')));
  assert.equal(findFont('Fixture Late'), null);
  writeFileSync(join(fixtures, 'late.ttf'), makeFont({ family: 'Fixture Late' }));
  assert.ok(findFont('Fixture Late'), 'a font installed while running is found');
  rmSync(join(fixtures, 'late.ttf'));
  assert.equal(findFont('Fixture Late'), null);
}));

test('resolver: the platforms\' font folders, Office\'s own and cloud fonts included', () => {
  const env = { LOCALAPPDATA: 'C:\\Users\\a\\AppData\\Local', WINDIR: 'C:\\Windows' };
  assert.deepEqual(fontDirs('win32', env, 'C:\\Users\\a').slice(0, 3), ['C:\\Windows\\Fonts', 'C:\\Users\\a\\AppData\\Local\\Microsoft\\Windows\\Fonts', 'C:\\Users\\a\\AppData\\Local\\Microsoft\\FontCache\\4\\CloudFonts']);
  const mac = fontDirs('darwin', {}, '/Users/a');
  for (const d of ['/System/Library/Fonts', '/Library/Fonts', '/Users/a/Library/Fonts', '/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts', '/Users/a/Library/Group Containers/UBF8T346G9.Office/FontCache']) assert.ok(mac.includes(d), d);
  assert.ok(mac.at(-1)!.replace(/\\/g, '/').endsWith('vendor/fonts'), 'open-charts\' own Liberation Sans last');
  assert.deepEqual(fontDirs('linux', { OPEN_CHARTS_FONT_DIRS: ['/a', '/b'].join(process.platform === 'win32' ? ';' : ':') }, '/h'), ['/a', '/b']);
});

// ---- B2: measuring, substitutes, the wide fallback ---------------------------------------------

test('measurer: each family measures with its own faces; bold from the bold face or the weight axis', () => withFixtures(() => {
  const m = createMeasurer();
  const s = m.forFont!('Fixture Sans');
  assert.equal(s.width('ab', 10), (400 + 450) / 100);
  assert.equal(s.width('ab', 10, true), (520 + 560) / 100);
  assert.equal(s.lineHeight(10), 10 * 1.15); // 900 + 250
  const v = m.forFont!('Fixture Var');
  assert.equal(v.width('a', 10), 5);
  assert.ok(Math.abs(v.width('a', 10, true) - 5.6) < 1e-9);
  const serif = m.forFont!('Fixture Serif');
  assert.equal(serif.width('a', 10, true), serif.width('a', 10) * 1.06, 'no bold face, no axis: slightly wider');
  assert.equal(m.forFont!('Arial'), m.forFont!('arial'), 'one measurer per family');
}));

test('measurer: a missing font uses a metric-compatible substitute, else the wide fallback, reported once per scene', () => withFixtures(() => {
  const m = createMeasurer();
  assert.deepEqual(substitutesFor('Calibri'), ['Carlito']);
  assert.deepEqual(substitutesFor('helvetica'), ['Arial', 'Liberation Sans', 'Arimo']);
  // Liberation Sans is not in the fixture folders: Arial (the base) stands in, exactly.
  const lib = m.forFont!('Liberation Sans');
  assert.equal(lib.fallback, undefined);
  assert.equal(lib.width('Revenue', 12), m.width('Revenue', 12));
  const wide = m.forFont!('Corporate Sans');
  assert.deepEqual(wide.fallback, { requested: 'Corporate Sans', used: m.family, wide: true });
  assert.equal(wide.width('Revenue', 12), m.width('Revenue', 12) * WIDE_FALLBACK.width);
  assert.equal(wide.lineHeight(10), 10 * WIDE_FALLBACK.lineHeight);
  const spec: ChartSpec = { ...samples.stacked!, overrides: { 'label:total:*': { font: 'Corporate Sans' } } };
  const scene = layoutChart(spec, { x: 48, y: 96, w: 864, h: 408 }, m, { ...defaultStyle, font: 'Corporate Sans' });
  const notes = (scene.warnings ?? []).filter((w) => w.startsWith('font "'));
  assert.equal(notes.length, 1, 'one warning for the font, however many labels use it');
  assert.match(notes[0], /^font "Corporate Sans" is not installed here/);
  assert.deepEqual(checkScene(scene, { measurer: wide }).filter((i) => i.kind !== 'hidden'), []);
  assert.equal(layoutChart(samples.stacked!, { x: 48, y: 96, w: 864, h: 408 }, m).warnings, undefined, 'Arial: no warning');
}));

test('measurer: one font file given by path measures every font, as before (--font file.ttf)', () => {
  const m = createMeasurer('vendor/fonts/LiberationSans-Regular.ttf', 'vendor/fonts/LiberationSans-Bold.ttf');
  assert.equal(m.forFont, undefined);
  const a = layoutChart(samples.line!, { x: 0, y: 0, w: 500, h: 300 }, m, { ...defaultStyle, font: 'Corporate Sans' });
  assert.equal(a.warnings, undefined);
});

// ---- B1: theme fonts ---------------------------------------------------------------------------

test('theme fonts: +mn-lt / +mj-lt resolve to the deck\'s fonts, else the style font, else Arial', () => {
  assert.ok(isThemeFont('+mn-lt') && isThemeFont('+mj-lt') && !isThemeFont('Calibri'));
  assert.equal(resolveFont('+mn-lt', { themeFonts: { major: 'Aptos Display', minor: 'Aptos' } }), 'Aptos');
  assert.equal(resolveFont('+mj-lt', { themeFonts: { major: 'Aptos Display', minor: 'Aptos' } }), 'Aptos Display');
  assert.equal(resolveFont('+mj-lt', { themeFonts: { minor: 'Aptos' } }), 'Aptos');
  assert.equal(resolveFont('+mn-lt', { font: 'Georgia' }), 'Georgia', 'a CLI deck\'s theme uses the style font');
  assert.equal(resolveFont('+mn-lt', { font: '+mj-lt' }), 'Arial');
  assert.equal(resolveFont(undefined, { font: 'Verdana' }), 'Verdana');
  assert.equal(resolveFont('Calibri', { themeFonts: { minor: 'Aptos' } }), 'Calibri');
});

test('theme fonts: a chart in +mn-lt lays out exactly as in the deck\'s body font; overrides in +mj-lt measure in the heading font', () => withFixtures(() => {
  const m = createMeasurer();
  const frame = { x: 48, y: 96, w: 864, h: 408 };
  const themeFonts = { major: 'Fixture Cloud', minor: 'Fixture Sans' };
  const viaTheme = layoutChart(samples.clustered!, frame, m, { ...defaultStyle, font: '+mn-lt', themeFonts });
  const direct = layoutChart(samples.clustered!, frame, m, { ...defaultStyle, font: 'Fixture Sans' });
  assert.deepEqual(viaTheme, direct);
  assert.notDeepEqual(direct, layoutChart(samples.clustered!, frame, m), 'the font changes the layout');
  // An element in the heading font is measured in it: its box fits its own text.
  const spec: ChartSpec = { ...samples.clustered!, overrides: { 'label:category:c0': { font: '+mj-lt' } } };
  const scene = layoutChart(spec, frame, m, { ...defaultStyle, font: '+mn-lt', themeFonts });
  const i = scene.roles!.indexOf('label:category:c0');
  const t = scene.shapes[i] as TextShape;
  assert.equal(t.font, '+mj-lt');
  assert.equal(t.lineHeight, m.forFont!('Fixture Cloud').lineHeight(t.size));
  assert.deepEqual(checkScene(scene, { measurer: m.forFont!('Fixture Sans') }).filter((x) => x.kind === 'overlap'), []);
}));

test('style: themeFonts is validated, and a CLI deck\'s theme names real fonts', async () => {
  const { validateStyle } = await import('../src/core/index.ts');
  assert.deepEqual(validateStyle({ font: '+mn-lt', themeFonts: { major: 'Aptos Display', minor: 'Aptos' } }), []);
  assert.ok(validateStyle({ themeFonts: { minor: '+mn-lt' } }).length > 0);
  assert.ok(validateStyle({ themeFonts: { body: 'Aptos' } }).length > 0);
  const { writePptx } = await import('../src/render/pptx.ts');
  const theme = (font: string, fonts?: { major?: string; minor?: string }) => unzip(writePptx([], font, undefined, fonts)).get('ppt/theme/theme1.xml')!.text();
  assert.match(theme('+mn-lt'), /<a:majorFont><a:latin typeface="Arial"\/>.*<a:minorFont><a:latin typeface="Arial"\/>/);
  assert.match(theme('+mn-lt', { major: 'Aptos Display', minor: 'Aptos' }), /<a:majorFont><a:latin typeface="Aptos Display"\/>.*<a:minorFont><a:latin typeface="Aptos"\/>/);
  assert.match(theme('Georgia'), /<a:minorFont><a:latin typeface="Georgia"\/>/);
});

// ---- B3: the add-in, the CLI, Office.js ---------------------------------------------------------

test('add-in: majorfont= / minorfont= carry the deck\'s theme fonts; a missing font is told once, never on re-layout', async () => withFixtures(async () => {
  const m = createMeasurer();
  const ctx = 'majorfont=Fixture Cloud\nminorfont=Fixture Sans\n';
  const req = parseRequest(`update\n\n48,96,864,408\n\n${ctx}{"type":"line"}`);
  assert.deepEqual(req.themeFonts, { major: 'Fixture Cloud', minor: 'Fixture Sans' });
  assert.match(req.context!, /minorfont=Fixture Sans\n/, 'the theme fonts are part of the layout\'s context');
  const spec = (font: string) => JSON.stringify({ ...samples.line, style: { font } });
  const viaTheme = await handleRequest(parseRequest(`update\n\n48,96,864,408\n\n${ctx}${spec('+mn-lt')}`), m);
  const direct = await handleRequest(parseRequest(`update\n\n48,96,864,408\n\n${spec('Fixture Sans')}`), m);
  const shapes = (r: string) => r.split('\n').filter((l) => l.startsWith('S\t')).map((l) => l.split('\t').filter((_, k) => k !== 12));
  assert.deepEqual(shapes(viaTheme), shapes(direct), 'same geometry; only the font name differs');
  assert.match(viaTheme, /\t\+mn-lt\t/);
  const told: string[] = [];
  const notify = (w: string) => { if (told.includes(w)) return false; told.push(w); return true; };
  const missing = `update\n\n48,96,864,408\n\n${spec('Corporate Sans')}`;
  assert.match(await handleRequest(parseRequest(missing), m, undefined, notify), /^M\tfont "Corporate Sans" is not installed here/m);
  assert.doesNotMatch(await handleRequest(parseRequest(missing), m, undefined, notify), /^M\t/m, 'once');
  told.length = 0;
  assert.doesNotMatch(await handleRequest(parseRequest(missing.replace('update', 'relayout')), m, undefined, notify), /^M\t/m, 'not on the automatic re-layout');
}));

test('add-in: a new chart takes the deck\'s body font (+mn-lt), Arial when the deck sends none', async () => withFixtures(async () => {
  const m = createMeasurer();
  const withTheme = await handleRequest(parseRequest('insert\nstacked\n48,96,864,408\n\nmajorfont=Fixture Cloud\nminorfont=Fixture Sans\n'), m);
  assert.match(withTheme, /\t\+mn-lt\t/, 'labels drawn in the deck body font');
  assert.match(withTheme, /"font":"\+mn-lt"/, 'stored with the chart, so it follows theme changes');
  const plain = await handleRequest(parseRequest('insert\nstacked\n48,96,864,408\n\n'), m);
  assert.doesNotMatch(plain, /\+mn-lt/);
}));

test('add-in: the VBA sends the theme fonts with every theme line', () => {
  const vba = readFileSync('addin/vba/OpenCharts.bas', 'utf8');
  assert.match(vba, /ThemeLine = ThemeColorsLine\(sld\) & ThemeFontsLine\(sld\)/);
  assert.match(vba, /ThemeFontScheme/);
  assert.match(vba, /"majorfont=" & major/);
  assert.match(vba, /"minorfont=" & minor/);
});

test('reportOnce: a message is reported the first time only, and remembered in the cache', () => {
  const msg = `font "Once ${Date.now()}" is not installed here`;
  assert.equal(reportOnce(msg), true);
  assert.equal(reportOnce(msg), false);
  assert.ok(JSON.parse(readFileSync(join(cache, 'reported.json'), 'utf8')).includes(msg));
});

test('CLI: --font <family> sets the charts\' font and measures with it; a missing one is reported', () => {
  const deck = join(tmp, 'deck.json');
  writeFileSync(deck, JSON.stringify({ slides: [{ title: 'T', chart: samples.stacked }] }));
  const run = (font: string) => {
    const out = join(tmp, `out-${font.replace(/\W/g, '')}`);
    const r = execFileSync(process.execPath, ['--experimental-strip-types', '--no-warnings', 'src/cli.ts', deck, '-o', out, '--font', font], {
      env: { ...process.env, OPEN_CHARTS_FONT_DIRS: fixtures }, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'],
    });
    return { slide: unzip(new Uint8Array(readFileSync(`${out}.pptx`))).get('ppt/slides/slide1.xml')!.text(), r };
  };
  assert.match(run('Fixture Sans').slide, /typeface="Fixture Sans"/);
  let err = '';
  try { execFileSync(process.execPath, ['--experimental-strip-types', '--no-warnings', 'src/cli.ts', deck, '-o', join(tmp, 'x'), '--font', 'Corporate Sans', '--strict'], { env: { ...process.env, OPEN_CHARTS_FONT_DIRS: fixtures }, stdio: ['ignore', 'pipe', 'pipe'] }); } catch (e) { err = String((e as { stderr: Buffer }).stderr); }
  assert.match(err, /warning: font "Corporate Sans" is not installed here/);
});

test('Office.js measurer: Arial from the bundled Liberation Sans, other fonts from the server\'s exact metrics, else the fallback', async () => withFixtures(async () => {
  const base = 'https://oc.test/taskpane.html';
  const real = globalThis.fetch;
  const asked: string[] = [];
  globalThis.fetch = (async (u: URL | string) => {
    const url = new URL(String(u));
    if (url.pathname.startsWith('/fonts/Liberation')) return new Response(readFileSync(`vendor${url.pathname}`));
    if (url.pathname === '/fonts/metrics') {
      const family = url.searchParams.get('family')!;
      asked.push(family);
      const found = familyMetrics(family);
      return new Response(JSON.stringify(found), { status: 200 });
    }
    return new Response('', { status: 404 });
  }) as typeof fetch;
  try {
    const bm = await loadMeasurer(base);
    const node = createMeasurer();
    await bm.load(['Arial', 'Fixture Sans', 'Corporate Sans', '+mn-lt']);
    assert.deepEqual(asked.sort(), ['Corporate Sans', 'Fixture Sans'], 'Arial is bundled; theme fonts are never asked for');
    const frame = { x: 48, y: 96, w: 864, h: 408 };
    for (const font of ['Arial', 'Fixture Sans', 'Corporate Sans']) {
      const style: Style = { ...defaultStyle, font };
      const a = layoutChart(samples.waterfall!, frame, bm, style), b = layoutChart(samples.waterfall!, frame, node, style);
      if (font === 'Arial') assert.deepEqual(a, b, 'Arial: the same scene as Node');
      if (font === 'Fixture Sans') assert.deepEqual(a, b, 'exact metrics from the server: the same scene as Node');
      if (font === 'Corporate Sans') assert.match(a.warnings![0], /^font "Corporate Sans"/);
    }
  } finally {
    globalThis.fetch = real;
  }
}));

test('Office.js server: GET /fonts/metrics answers an installed family\'s metrics, null otherwise', async () => withFixtures(async () => {
  const { serveOfficeJs } = await import('../src/officejs/server.ts');
  const log = console.log;
  console.log = () => {};
  const server = serveOfficeJs({ root: tmp, port: 0 });
  try {
    await new Promise((r) => server.once('listening', r));
    const at = `http://127.0.0.1:${(server.address() as { port: number }).port}/fonts/metrics?family=`;
    const ok = await fetch(`${at}Fixture%20Sans`);
    assert.equal(ok.status, 200);
    const j = await ok.json() as { family: string; regular: { v: number }; bold: object | null };
    assert.equal(j.family, 'Fixture Sans');
    assert.equal(j.regular.v, 1);
    assert.ok(j.bold);
    assert.equal(await (await fetch(`${at}Corporate%20Sans`)).json(), null);
    assert.equal(await (await fetch(`${at}${'x'.repeat(65)}`)).json(), null);
  } finally {
    server.close();
    console.log = log;
  }
}));

// ---- B4: the fonts on this computer ---------------------------------------------------------------

const FONTS = ['Arial', 'Helvetica', 'Helvetica Neue', 'Georgia', 'Times New Roman', 'Verdana', 'Calibri', 'Aptos', 'Segoe UI', 'Cambria', 'Courier New'];
const installed = FONTS.filter((f) => findFont(f));

test('installed fonts measure with their own widths and line heights', (t) => {
  if (installed.length < 2) { t.skip(`only ${installed.join(', ') || 'no'} test fonts installed`); return; }
  t.diagnostic(`fonts measured: ${installed.join(', ')}; not installed: ${FONTS.filter((f) => !installed.includes(f)).join(', ') || 'none'}`);
  const m = createMeasurer();
  const text = 'Revenue 2024 (EUR m)';
  for (const f of installed) {
    const fm = m.forFont!(f);
    assert.equal(fm.fallback, undefined, f);
    assert.ok(fm.width(text, 12, true) >= fm.width(text, 12), `${f}: bold is not narrower`);
    assert.ok(fm.lineHeight(10) > 9 && fm.lineHeight(10) < 16, `${f}: line height ${fm.lineHeight(10)}`);
    const face = findFont(f)!;
    assert.deepEqual({ ...cachedMetrics(face.path, face.index, face.wght), family: undefined }, { ...fontMetrics(file(face.path), face.index, face.wght), family: undefined }, `${f}: cached metrics are current`);
  }
  if (installed.includes('Verdana') && installed.includes('Arial')) assert.ok(m.forFont!('Verdana').width(text, 12) > m.width(text, 12) * 1.1, 'Verdana is wider than Arial');
  if (installed.includes('Helvetica') && installed.includes('Arial')) assert.equal(m.forFont!('Helvetica').width(text, 12), m.width(text, 12), 'Helvetica and Arial share their widths');
  if (installed.includes('Calibri') && installed.includes('Arial')) assert.ok(m.forFont!('Calibri').width(text, 12) < m.width(text, 12), 'Calibri is narrower than Arial');
});

test('quality gate: every chart type laid out in wide and narrow fonts, in large and small frames, has no overlaps or overflow', (t) => {
  if (!installed.length) { t.skip('no test fonts installed'); return; }
  const m = createMeasurer();
  const frames = [{ x: 48, y: 96, w: 864, h: 408 }, { x: 20, y: 20, w: 420, h: 300 }, { x: 10, y: 10, w: 300, h: 210 }];
  const fonts = [...installed, 'Corporate Sans'];
  for (const font of fonts) {
    const fm = m.forFont!(font);
    for (const [type, spec] of Object.entries(samples)) for (const frame of frames) {
      const scene = layoutChart(spec!, frame, m, { ...defaultStyle, font });
      const bad = checkScene(scene, { measurer: fm }).filter((i) => i.kind !== 'hidden');
      assert.deepEqual(bad, [], `${type} in ${font} at ${frame.w}×${frame.h}`);
    }
  }
  // The reason for all this: a layout measured in Arial but drawn in a wider font overflows.
  const verdana = installed.includes('Verdana') ? m.forFont!('Verdana') : null;
  if (verdana) {
    const arial = layoutChart(samples.mekko!, frames[2], m);
    assert.ok(checkScene(arial, { measurer: verdana }).some((i) => i.kind === 'overflow' || i.kind === 'overlap'), 'Arial-measured mekko overflows in Verdana');
  }
});
