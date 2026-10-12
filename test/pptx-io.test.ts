import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { deflateRawSync } from 'node:zlib';
import { layoutChart } from '../src/core/index.ts';
import { createMeasurer } from '../src/fonts.ts';
import type { Scene } from '../src/core/types.ts';
import { buildDeck, CHART_FRAME, chartContent, type DeckSpec } from '../src/deck.ts';
import { writePptx } from '../src/render/pptx.ts';
import { chartTags, specJson } from '../src/render/tags.ts';
import { crc32, zip } from '../src/render/zip.ts';
import { unzip } from '../src/ooxml/unzip.ts';
import { findAll, parseXml, textOf } from '../src/ooxml/xml.ts';
import { importDeck, readCharts } from '../src/ooxml/pptx-read.ts';
import { ENGINE_VERSION } from '../src/version.ts';

const m = createMeasurer();
const deck: DeckSpec = JSON.parse(readFileSync(new URL('../examples/deck.json', import.meta.url), 'utf8'));

test('ENGINE_VERSION matches package.json', () => {
  const pkg = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8'));
  assert.equal(ENGINE_VERSION, pkg.version);
});

test('xml parser: entities, CDATA, offsets, no DOCTYPE expansion', () => {
  const src = '<?xml version="1.0"?><!DOCTYPE x [<!ENTITY e "boom">]><a x="1 &amp; 2"><b>t&lt;&#65;&#x42;</b><![CDATA[<raw>]]><c/>&e;</a>';
  const root = parseXml(src);
  const a = root.children[0];
  assert.equal(a.attrs.x, '1 & 2');
  assert.equal(textOf(a.children[0]), 't<AB');
  assert.ok(a.text.includes('<raw>') && a.text.includes('&e;'));
  assert.equal(src.slice(a.children[1].start, a.children[1].end), '<c/>');
  assert.equal(src.slice(a.children[0].innerStart, a.children[0].innerEnd), 't&lt;&#65;&#x42;');
  assert.throws(() => parseXml('<a><b></a>'));
});

test('unzip reads stored and deflated entries and passes raw entries through unchanged', () => {
  const text = 'hello '.repeat(200);
  const data = new TextEncoder().encode(text);
  const compressed = new Uint8Array(deflateRawSync(data));
  const z = zip([{ name: 'a.txt', data: 'plain' }, { name: 'b.txt', raw: { method: 8, crc: crc32(data), size: data.length, compressed } }]);
  const e = unzip(z);
  assert.equal(e.get('a.txt')!.text(), 'plain');
  assert.equal(e.get('b.txt')!.text(), text);
  assert.deepEqual(e.get('b.txt')!.raw.compressed, compressed);
  assert.throws(() => unzip(zip([{ name: '../evil', data: 'x' }])), /Unsafe/);
  assert.throws(() => unzip(new Uint8Array(10)), /Not a zip/);
});

test('OC_SPEC is chunked above the limit and reassembles exactly, never splitting surrogate pairs', () => {
  const spec = { type: 'stacked', categories: ['😀'.repeat(50)], series: [{ name: 'a', values: [1] }] };
  const tags = chartTags(spec, CHART_FRAME, 'Arial', 7);
  const t = Object.fromEntries(tags);
  assert.equal(t.OC_SPEC, undefined);
  assert.ok(Number(t.OC_SPEC_N) > 10);
  for (let i = 0; i < Number(t.OC_SPEC_N); i++) assert.doesNotMatch(t[`OC_SPEC_${i}`], /^[\uDC00-\uDFFF]|[\uD800-\uDBFF]$/);
  assert.equal(specJson(t), JSON.stringify(spec));
  const small = Object.fromEntries(chartTags(spec, CHART_FRAME, 'Arial'));
  assert.equal(small.OC_SPEC, JSON.stringify(spec));
  assert.equal(small.OC_FRAME, '48,96,864,408');
});

test('writer emits tag contract v1 as real PowerPoint tag parts', () => {
  const { pptx } = buildDeck(deck, m);
  const e = unzip(pptx);
  const ct = e.get('[Content_Types].xml')!.text();
  const slideRels = e.get('ppt/slides/_rels/slide1.xml.rels')!.text();
  assert.match(slideRels, /relationships\/tags" Target="\.\.\/tags\/tag1\.xml"/);
  assert.match(ct, /PartName="\/ppt\/tags\/tag1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.presentationml\.tags\+xml"/);
  const slide = parseXml(e.get('ppt/slides/slide1.xml')!.text());
  const grp = findAll(slide, 'p:grpSp')[0];
  assert.equal(findAll(grp, 'p:cNvPr')[0].attrs.name, 'Stacked column chart');
  const rid = findAll(grp, 'p:tags')[0].attrs['r:id'];
  assert.ok(slideRels.includes(`Id="${rid}"`));
  const tagXml = parseXml(e.get('ppt/tags/tag1.xml')!.text());
  const tags = Object.fromEntries(findAll(tagXml, 'p:tag').map((t) => [t.attrs.name, t.attrs.val]));
  assert.equal(tags.OC_V, '1');
  assert.equal(tags.OC_ENGINE, ENGINE_VERSION);
  assert.equal(tags.OC_FONT, 'Arial');
  assert.equal(tags.OC_FRAME, '48,96,864,408');
  assert.deepEqual(JSON.parse(tags.OC_SPEC), deck.slides[0].chart);
  // every tag part is referenced and registered
  const parts = [...e.keys()].filter((k) => k.startsWith('ppt/tags/'));
  for (const p of parts) assert.ok(ct.includes(`/${p}"`), p);
});

test('each child gets OC_ROLE from scene.roles and a meaningful name', () => {
  const scene = layoutChart(deck.slides[0].chart, CHART_FRAME, m) as Scene & { roles?: string[] };
  scene.roles = scene.shapes.map((s, i) => (s.kind === 'rect' ? `bar:s0:c${i}` : s.kind === 'text' ? `label:value:s0:c${i}` : 'axis'));
  const pptx = writePptx([{ title: 'T', scenes: [chartContent(deck.slides[0].chart, scene, 'Arial')] }]);
  const [info] = readCharts(pptx);
  assert.equal(info.title, 'T');
  assert.equal(info.charts.length, 1);
  assert.deepEqual(info.charts[0].roles, scene.roles);
  const slide = unzip(pptx).get('ppt/slides/slide1.xml')!.text();
  assert.match(slide, /name="Bar s0 c\d+"/);
  assert.match(slide, /name="Label value s0 c\d+"/);
});

test('round trip: build → import gives the identical deck spec', () => {
  const styled: DeckSpec = { ...deck, style: { font: 'Arial', palette: ['112233', '445566', '778899', 'AABBCC'] } };
  for (const d of [deck, styled]) {
    const { deck: back, warnings } = importDeck(buildDeck(d, m).pptx);
    assert.deepEqual(warnings, []);
    assert.deepEqual(back, d);
  }
});

test('inspect reports frames, contract version and tag sizes', () => {
  const info = readCharts(buildDeck(deck, m).pptx);
  assert.equal(info.length, deck.slides.length);
  const c = info[0].charts[0];
  assert.equal(c.contract, '1');
  assert.deepEqual(c.frame, CHART_FRAME);
  assert.deepEqual(c.bounds, CHART_FRAME);
  assert.ok(c.tagBytes > JSON.stringify(deck.slides[0].chart).length);
  assert.ok(c.children > 5);
});
