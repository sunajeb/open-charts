// Template fill (open-charts fill): named chart slots, placeholders, {{fields}} across runs, notes,
// slide plans with each, sections, theme colours, refill, dry-run / strict, and — above all —
// untouched parts byte-identical.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { createMeasurer } from '../src/fonts.ts';
import { fillTemplate, type FillReport } from '../src/template/fill.ts';
import type { FillData } from '../src/template/data.ts';
import { validateFillData } from '../src/template/data.ts';
import { makeTemplate } from '../scripts/make-template.ts';
import { unzip } from '../src/ooxml/unzip.ts';
import { readCharts, slideParts } from '../src/ooxml/pptx-read.ts';
import { findAll, parseXml, textOf } from '../src/ooxml/xml.ts';
import { parseStyleFile } from '../src/core/style-file.ts';

const m = createMeasurer();
const template = makeTemplate();
const data: FillData = JSON.parse(readFileSync('examples/template/data.json', 'utf8'));
const fill = (d: FillData, o: Partial<Parameters<typeof fillTemplate>[2]> = {}) => fillTemplate(template, d, { measurer: m, ...o });
const texts = (pptx: Uint8Array, part: string) => findAll(parseXml(unzip(pptx).get(part)!.text()), 'a:p').map((p) => findAll(p, 'a:t').map(textOf).join(''));
const slideTexts = (pptx: Uint8Array) => { const e = unzip(pptx); return slideParts(e).map((p) => texts(pptx, p).join(' | ')); };

test('the committed example template is the one scripts/make-template.ts writes', () => {
  assert.deepEqual(new Uint8Array(readFileSync('examples/template/template.pptx')), template);
});

test('fill: the example data fills every slot: 10 slides, 7 charts, every field, no problems', () => {
  const { pptx, report } = fill(data);
  assert.equal(report.ok, true, JSON.stringify(report, null, 1));
  assert.deepEqual(report.unused, []);
  assert.equal(report.slides.length, 10);
  const charts = readCharts(pptx!);
  assert.deepEqual(charts.flatMap((s) => s.charts.map((c) => `${s.index}:${c.name}`)), ['3:chart:revenue', '4:chart:mix_left', '4:chart:mix_right', '5:chart:region_sales', '6:chart:region_sales', '7:chart:region_sales']);
  for (const s of slideTexts(pptx!)) assert.doesNotMatch(s, /\{\{/, s);
});

test('fill: untouched parts are byte-identical (compressed bytes, CRC, method) and keep their order', () => {
  const { pptx, report } = fill(data);
  const before = unzip(template), after = unzip(pptx!);
  const changed = new Set([...report.changedParts, ...report.removedParts]);
  let same = 0;
  for (const [name, e] of before) {
    if (changed.has(name)) continue;
    const a = after.get(name);
    assert.ok(a, `${name} kept`);
    assert.equal(a.raw.method, e.raw.method, name);
    assert.equal(a.raw.crc, e.raw.crc, name);
    assert.deepEqual(a.raw.compressed, e.raw.compressed, `${name}: same compressed bytes`);
    same++;
  }
  // What a fill may touch: slides with slots, their rels and notes, the slide list, content types, new parts.
  for (const p of ['ppt/slides/slide2.xml', 'ppt/slides/slide7.xml', 'ppt/slides/slide8.xml', 'ppt/slideMasters/slideMaster1.xml', 'ppt/slideLayouts/slideLayout3.xml', 'ppt/theme/theme1.xml', 'ppt/media/image1.png', 'docProps/app.xml', 'ppt/notesMasters/notesMaster1.xml']) {
    assert.ok(!changed.has(p), `${p} untouched`);
  }
  assert.ok(same >= 20, `${same} parts copied verbatim`);
  const order = [...after.keys()].filter((k) => before.has(k));
  assert.deepEqual(order, [...before.keys()].filter((k) => after.has(k)), 'original parts keep their order');
});

test('fill: a values-only fill edits only the slides with slots; presentation.xml stays byte-identical', () => {
  const { pptx, report } = fill({ values: { revenue: data.values!.revenue, revenue_title: 'Revenue', source: 'Filings', revenue_note: 'Note' } });
  assert.ok(pptx);
  const touched = report.changedParts.filter((p) => !p.startsWith('ppt/tags/'));
  assert.deepEqual(touched, ['[Content_Types].xml', 'ppt/notesSlides/notesSlide1.xml', 'ppt/slides/_rels/slide3.xml.rels', 'ppt/slides/slide3.xml']);
  assert.deepEqual(report.removedParts, []);
  // Every other slide reports its unfilled fields.
  assert.equal(report.ok, false);
  assert.ok(report.slides[0].fields.some((f) => f.name === 'client' && f.status === 'missing'));
});

test('fill: a placeholder without its own frame takes the layout placeholder\'s; charts carry tags and the slot name', () => {
  const { pptx, report } = fill(data);
  assert.deepEqual(report.slides[3].charts.map((c) => c.frame), [{ x: 36, y: 96, w: 432, h: 384 }, { x: 492, y: 96, w: 432, h: 384 }]);
  const c = readCharts(pptx!)[3].charts;
  assert.deepEqual(c.map((x) => x.frame), [{ x: 36, y: 96, w: 432, h: 384 }, { x: 492, y: 96, w: 432, h: 384 }]);
  assert.ok(c.every((x) => x.contract === '1' && x.roles.length > 5 && JSON.parse(x.spec!).type));
});

test('fill: charts take the template theme: accent slots as schemeClr, OC_SPEC theme = template colours', () => {
  const { pptx } = fill(data);
  const xml = unzip(pptx!).get('ppt/slides/slide3.xml')!.text();
  assert.match(xml, /<a:schemeClr val="accent1"\/>/);
  assert.doesNotMatch(xml, /<a:srgbClr val="1F3A5F"\/>/, 'not the open-charts default navy');
  const spec = JSON.parse(readCharts(pptx!)[2].charts[0].spec!);
  assert.equal(spec.style.theme.accent1, '0F6E7A');
  assert.equal(spec.style.font, 'Arial', 'the theme body font');
});

test('fill: {{fields}} split across runs are replaced, keeping each run\'s formatting; table cells and notes too', () => {
  const { pptx } = fill(data);
  const xml = unzip(pptx!).get('ppt/slides/slide1.xml')!.text();
  assert.deepEqual(texts(pptx!, 'ppt/slides/slide1.xml'), ['Northwind Traders: Growth review', 'Board update · October 2026']);
  // The value went into the run where the field started; the run with lang="en-GB" kept its rPr.
  assert.match(xml, /<a:r><a:rPr lang="en-US" dirty="0"\/><a:t>Northwind Traders<\/a:t><\/a:r><a:r><a:rPr lang="en-GB" dirty="0"\/><a:t><\/a:t><\/a:r>/);
  const table = slideTexts(pptx!)[7];
  assert.match(table, /Revenue \| \$163m \| \$188m \| EBITDA margin \| 18\.2% \| 19\.5%/);
  assert.deepEqual(texts(pptx!, 'ppt/notesSlides/notesSlide1.xml').filter(Boolean), ['Talk track: Subscriptions passed licences in 2024; services stayed flat.']);
});

test('fill: a line break in a value becomes a:br with the run\'s formatting', () => {
  const { pptx } = fill({ values: { ...data.values, project: 'Growth review\nPhase 2' } });
  const xml = unzip(pptx!).get('ppt/slides/slide1.xml')!.text();
  assert.match(xml, /<a:t>: Growth review<\/a:t><\/a:r><a:br><a:rPr lang="en-US" dirty="0"\/><\/a:br><a:r><a:rPr lang="en-US" dirty="0"\/><a:t>Phase 2<\/a:t>/);
});

test('fill: each repeats a slide per item (own slide, rels, tags; unique ids), sections follow', () => {
  const { pptx } = fill(data);
  const e = unzip(pptx!);
  const parts = slideParts(e);
  assert.equal(parts.length, 10);
  assert.equal(new Set(parts).size, 10);
  assert.deepEqual(slideTexts(pptx!).slice(4, 7).map((t) => t.split(' | ')[0]), ['Americas: subscriptions doubled in two years', 'EMEA: growth slowed as licences declined', 'Asia Pacific: smallest region, fastest growth']);
  const pres = parseXml(e.get('ppt/presentation.xml')!.text());
  const ids = findAll(pres, 'p:sldId').map((s) => s.attrs.id);
  assert.equal(new Set(ids).size, 10);
  const sections = findAll(pres, 'p14:section').map((s) => [s.attrs.name, findAll(s, 'p14:sldId').map((x) => x.attrs.id)] as const);
  assert.deepEqual(sections.map(([n, xs]) => `${n}:${xs.length}`), ['Introduction:2', 'Analysis:7', 'Close:1']);
  assert.deepEqual(sections.flatMap(([, xs]) => xs), ids, 'sections cover the slides in order');
  // Copies have their own tag parts (no part shared between slides).
  const ct = e.get('[Content_Types].xml')!.text();
  for (const p of parts) assert.ok(ct.includes(`PartName="/${p}"`), `${p} registered`);
  const tagOwners = new Map<string, string>();
  for (const p of parts) {
    const rels = e.get(p.replace('slides/', 'slides/_rels/') + '.rels')!.text();
    for (const t of rels.matchAll(/Target="\.\.\/tags\/(tag\d+\.xml)"/g)) {
      assert.ok(!tagOwners.has(t[1]), `${t[1]} used by ${tagOwners.get(t[1])} and ${p}`);
      tagOwners.set(t[1], p);
      assert.ok(e.has(`ppt/tags/${t[1]}`));
    }
  }
});

test('fill: a slide plan can drop and reorder slides; parts only dropped slides used are removed', () => {
  const { pptx, report } = fill({ values: data.values, slides: [{ slide: 3 }, { slide: 1 }] });
  assert.deepEqual(slideTexts(pptx!).map((t) => t.split(' | ')[0]), ['Revenue grew 2.4x in four years, led by subscriptions', 'Northwind Traders: Growth review']);
  const e = unzip(pptx!);
  assert.ok(!e.has('ppt/media/image1.png'), 'the picture only slide 7 used');
  assert.ok(!e.has('ppt/slides/slide2.xml'));
  assert.ok(e.has('ppt/notesSlides/notesSlide1.xml'), 'slide 3 keeps its notes');
  assert.ok(e.has('ppt/slideLayouts/slideLayout3.xml'), 'layouts stay (the master uses them)');
  assert.ok(report.removedParts.includes('ppt/slides/_rels/slide2.xml.rels'));
  const ct = e.get('[Content_Types].xml')!.text();
  assert.doesNotMatch(ct, /slide2\.xml/);
  const rels = e.get('ppt/_rels/presentation.xml.rels')!.text();
  assert.doesNotMatch(rels, /slides\/slide2\.xml/);
  for (const r of rels.matchAll(/Target="([^"]+)"/g)) assert.ok(e.has(`ppt/${r[1]}`), `${r[1]} exists`);
  // The new order splits no section: Analysis, then Introduction (Close now empty).
  const pres = parseXml(e.get('ppt/presentation.xml')!.text());
  assert.deepEqual(findAll(pres, 'p14:section').map((s) => findAll(s, 'p14:sldId').length), [1, 1, 0]);
});

test('fill: an order that would split a section drops the sections with a warning', () => {
  const { pptx, report } = fill({ values: data.values, slides: [{ slide: 1 }, { slide: 3 }, { slide: 2 }] });
  assert.ok(report.warnings.some((w) => /sections removed/.test(w)));
  assert.doesNotMatch(unzip(pptx!).get('ppt/presentation.xml')!.text(), /sectionLst/);
});

test('fill: filling a filled deck again replaces the charts in place and drops their old tag parts', () => {
  const values = { ...data.values, ...data.slides![4].each![0] };
  const once = fill({ values }).pptx!;
  const v2 = { ...values, revenue: { ...(values.revenue as object), series: [{ name: 'Total', values: [1, 2, 3, 4, 5] }] } } as FillData['values'];
  const twice = fillTemplate(once, { values: v2 }, { measurer: m });
  const charts = readCharts(twice.pptx!);
  assert.equal(charts[2].charts.length, 1);
  assert.equal(JSON.parse(charts[2].charts[0].spec!).series[0].name, 'Total');
  const tagsOnce = [...unzip(once).keys()].filter((k) => k.startsWith('ppt/tags/')).length;
  const tagsTwice = [...unzip(twice.pptx!).keys()].filter((k) => k.startsWith('ppt/tags/')).length;
  assert.ok(Math.abs(tagsTwice - tagsOnce) <= 50 && tagsTwice < tagsOnce * 1.5, `old tags collected (${tagsOnce} → ${tagsTwice})`);
  assert.ok(twice.report.removedParts.some((p) => p.startsWith('ppt/tags/')));
});

test('fill: --style and data style apply; the template theme still wins for theme colours', () => {
  const style = parseStyleFile(readFileSync('examples/styles/ember.json', 'utf8'));
  const { pptx } = fill(data, { style });
  const spec = JSON.parse(readCharts(pptx!)[2].charts[0].spec!);
  assert.equal(spec.style.font, 'Georgia');
  assert.equal(spec.style.palette[0], '8C2D19');
  assert.equal(spec.style.theme.accent1, '0F6E7A');
  const d2 = fill({ ...data, style: { font: 'Verdana' } }, { style }).pptx!;
  assert.equal(JSON.parse(readCharts(d2)[2].charts[0].spec!).style.font, 'Verdana', 'data style over the file');
});

test('fill: dry run reports slots and frames and writes nothing; problems are counted', () => {
  const { pptx, report } = fill({ values: { client: 'X', revenue: 'not a chart', mix_left: data.values!.mix_left } }, { dryRun: true });
  assert.equal(pptx, undefined);
  const charts = report.slides.flatMap((s) => s.charts);
  assert.deepEqual(charts.map((c) => `${c.name}:${c.status}`), ['revenue:error', 'mix_left:filled', 'mix_right:missing', 'region_sales:missing']);
  assert.ok(report.problems > 5);
});

test('fill data: validated with paths; chart specs inside values too', () => {
  const errors = validateFillData({ values: { 'bad name': 'x', c: { type: 'pie', categories: ['a'], series: [{ name: 's', values: ['x'] }] }, n: [] }, slides: [{ slide: 0, each: [] }], extra: 1 });
  assert.deepEqual(errors.map((e) => e.path).sort(), ['/extra', '/slides/0/each', '/slides/0/slide', '/values/bad name', '/values/c/series/0/values/0', '/values/n'].sort());
  assert.throws(() => fill({ slides: [{ slide: 9 }] }), /template has 8 slides/);
});

test('fill: a slot inside a group is reported, not filled', async () => {
  const { Pkg } = await import('../src/template/package.ts');
  const pkg = new Pkg(template);
  const xml = pkg.text('ppt/slides/slide3.xml');
  const sp = xml.slice(xml.indexOf('<p:sp><p:nvSpPr><p:cNvPr id="4" name="chart:revenue"/>'), xml.indexOf('</p:sp>', xml.indexOf('name="chart:revenue"')) + 7);
  pkg.set('ppt/slides/slide3.xml', xml.replace(sp, `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="40" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${sp}</p:grpSp>`));
  const r: FillReport = fillTemplate(pkg.toBytes(), { values: data.values }, { measurer: m, dryRun: true }).report;
  assert.match(r.slides[2].charts[0].detail ?? '', /inside a group/);
});

test('CLI fill: writes the deck; --dry-run writes nothing; --strict exits 1 on a missing value; --each writes one deck per item', () => {
  const run = (...args: string[]) => spawnSync(process.execPath, ['--experimental-strip-types', '--no-warnings', 'src/cli.ts', ...args], { encoding: 'utf8' });
  rmSync('out/test-fill', { recursive: true, force: true });
  const ok = run('fill', 'examples/template/template.pptx', 'examples/template/data.json', '-o', 'out/test-fill/deck.pptx', '--strict');
  assert.equal(ok.status, 0, ok.stderr);
  assert.match(ok.stdout, /Wrote out\/test-fill\/deck\.pptx \(10 slides\)/);
  const dry = run('fill', 'examples/template/template.pptx', 'examples/template/data.json', '-o', 'out/test-fill/dry.pptx', '--dry-run');
  assert.equal(dry.status, 0);
  assert.match(dry.stdout, /chart:mix_left at 36,96 432x384: would fill/);
  assert.ok(!existsSync('out/test-fill/dry.pptx'));
  writeFileSync('out/test-fill/partial.json', JSON.stringify({ values: { client: 'X' } }));
  const strict = run('fill', 'examples/template/template.pptx', 'out/test-fill/partial.json', '-o', 'out/test-fill/s.pptx', '--strict');
  assert.equal(strict.status, 1);
  assert.match(strict.stderr, /--strict: \d+ problems/);
  const each = run('fill', 'examples/template/template.pptx', 'examples/template/data.json', '--each', 'examples/template/clients.json', '-o', 'out/test-fill/{client}.pptx');
  assert.equal(each.status, 0, each.stderr);
  assert.ok(existsSync('out/test-fill/Northwind Traders.pptx') && existsSync('out/test-fill/Contoso.pptx'));
  assert.match(slideTexts(new Uint8Array(readFileSync('out/test-fill/Contoso.pptx')))[0], /^Contoso: Growth review/);
  const bad = run('fill', 'examples/template/template.pptx', 'examples/template/data.json', '--each', 'examples/template/clients.json', '-o', 'out/test-fill/same.pptx');
  assert.equal(bad.status, 1);
  assert.match(bad.stderr, /put a field in the output name/);
});

test('schema/fill.schema.json lists the fill data keys', async () => {
  const { FILL_KEYS, FILL_SLIDE_KEYS } = await import('../src/template/data.ts');
  const schema = JSON.parse(readFileSync('schema/fill.schema.json', 'utf8'));
  assert.deepEqual(Object.keys(schema.properties).sort(), [...FILL_KEYS].sort());
  assert.deepEqual(Object.keys(schema.properties.slides.items.properties).sort(), [...FILL_SLIDE_KEYS].sort());
});
