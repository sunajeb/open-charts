// D6 robustness, offline: decks from older versions, charts edited by hand in PowerPoint without the
// add-in, and charts pasted into another deck. Requests are built the way the VBA sends them
// (addin/vba/OpenCharts.bas Relayout: ChartContext, NativeLines = base= + native=, then SpecOf).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { handleRequest, NEWER_SPEC, parseRequest, sceneBounds, sceneFor } from '../src/addin/engine.ts';
import { parseResponse } from '../src/addin/format.ts';
import { galleryItems, sampleFor } from '../src/addin/registry.ts';
import { sceneEntries, type ChartSpec, type Measurer, type RenderedEntry } from '../src/core/index.ts';
import { samples } from '../src/core/samples.ts';
import { buildDeck } from '../src/deck.ts';
import { createMeasurer } from '../src/fonts.ts';
import { importDeck, readCharts } from '../src/ooxml/pptx-read.ts';
import { unzip } from '../src/ooxml/unzip.ts';
import { zip } from '../src/render/zip.ts';

const m = createMeasurer();
const F = '48,96,864,408';
type Res = ReturnType<typeof parseResponse>;
const tag = (r: Res, k: string) => r.tags.find(([n]) => n === k)?.[1] ?? '';
const ask = async (text: string, mm: Measurer = m) => parseResponse(await handleRequest(parseRequest(text), mm, { get: async () => undefined, open: async () => {} }));
const insert = (item: string, ctx = '') => ask(`insert\n${item}\n${F}\n\n${ctx}`);
const items = galleryItems().map((g) => g.id).filter((id) => 'spec' in sampleFor(id));

/** The chart's shapes as the VBA reads them back (ShapeEntry), laid out from its own tags. */
function nativeOf(r: Res, ctx = ''): RenderedEntry[] {
  const { scene } = sceneFor(JSON.parse(tag(r, 'OC_SPEC')), `${tag(r, 'OC_FRAME')}\n\n${ctx}`, m);
  return sceneEntries(scene.shapes, scene.roles!, tag(r, 'OC_FONT')).filter((e) => e.role && e.role !== 'frame');
}
/** An update request with the chart's OC_BASE and its shapes as they are now (NativeLines). */
const update = (r: Res, native: unknown, { op = 'update', base = tag(r, 'OC_BASE'), frame = tag(r, 'OC_FRAME'), bounds = '', ctx = '' } = {}) =>
  ask(`${op}\n\n${frame}\n${bounds}\n${ctx}base=${base}\nnative=${JSON.stringify(native)}\n${tag(r, 'OC_SPEC')}`);
const own = (r: Res) => JSON.parse(tag(r, 'OC_SPEC')).overrides ?? {};

// ---- Decks from older versions ------------------------------------------------------------------

/** A spec from before spec v1 (no specVersion): `orientation` and options its type does not use. */
const V0 = { type: 'clustered', categories: ['2023', '2024'], series: [{ name: 'Sales', values: [3, 5] }], options: { orientation: 'horizontal', gap: 0.5, totals: true } };

test('older spec (v0, no specVersion): every add-in request lays it out and stores it upgraded', async () => {
  const spec = JSON.stringify(V0);
  for (const op of ['update', 'relayout', 'direction\nright', 'detach\nlabel:value:s0:c0', 'command\nscaleGroup']) {
    const [o, arg = ''] = op.split('\n');
    const r = await ask(`${o}\n${arg}\n${F}\n\n${spec}`);
    assert.deepEqual(r.errors, [], op);
    if (o === 'relayout') { assert.equal(r.nothing, true, 'relayout: nothing to do'); continue; }
    const stored = JSON.parse(tag(r, 'OC_SPEC'));
    assert.equal(stored.specVersion, 1, `${op}: stored as the current version`);
    assert.equal(stored.options.direction, 'right', `${op}: orientation became direction`);
    assert.ok(!('orientation' in stored.options) && !('totals' in stored.options), `${op}: ${JSON.stringify(stored.options)}`);
  }
  const edit = await ask(`edit\n\n${F}\n\n${spec}`);
  assert.deepEqual(edit.errors, []);
  // A chart drawn before OC_BASE existed: the VBA sends no base= / native= lines.
  assert.deepEqual((await ask(`update\n\n${F}\n${F}\n${spec}`)).errors, []);
});

test('a chart from a newer open-charts, or damaged tags, gives a plain message, never a crash', async () => {
  const newer = JSON.stringify({ ...V0, specVersion: 2, options: {} });
  for (const op of ['update', 'relayout', 'edit', 'command\ntotals', 'direction\nright', 'detach\nlabel:value:s0:c0']) {
    const [o, arg = ''] = op.split('\n');
    assert.deepEqual((await ask(`${o}\n${arg}\n${F}\n\n${newer}`)).errors, [NEWER_SPEC], op);
    // OC_SPEC chunk missing (VBA SpecOf joins what is there), hand-edited into something else.
    for (const bad of [JSON.stringify(V0).slice(0, 40), 'null', '[]', '{}', '"x"', '42']) {
      const r = await ask(`${o}\n${arg}\n${F}\n\n${bad}`);
      assert.deepEqual(r.errors, ['This chart\'s data (OC_SPEC tag) is damaged and cannot be read.'], `${op} ${bad}`);
    }
  }
  // Valid JSON that is not a valid chart: where the problem is, in words.
  const r = await ask(`update\n\n${F}\n\n${JSON.stringify({ ...V0, options: {}, series: [{ name: 'Sales', values: ['x', 5] }] })}`);
  assert.match(r.errors[0], /^This chart's data \(OC_SPEC tag\) is not valid, so it cannot be laid out:\n\/series\/0\/values\/0: /);
});

/** `bytes` with every part's text passed through `edit` (a deck changed outside open-charts). */
function rewrite(bytes: Uint8Array, edit: (name: string, text: string) => string): Uint8Array {
  return zip([...unzip(bytes).values()].map((e) => (/\.(xml|rels)$/.test(e.name) ? { name: e.name, data: edit(e.name, e.text()) } : { name: e.name, raw: e.raw })));
}

test('reading older and damaged decks: v0 charts import and rebuild; a damaged chart is reported, the rest still read', () => {
  const ok = samples.stacked as ChartSpec;
  const built = buildDeck({ slides: [{ title: 'Old', chart: V0 as unknown as ChartSpec }, { title: 'Chunk lost', chart: ok }, { title: 'Broken', chart: ok }, { title: 'Fine', chart: ok }] }, m).pptx;
  // v0 as stored by an older version: imports as-is and builds again (migrated on the way).
  const old = importDeck(built);
  assert.deepEqual(old.warnings, []);
  assert.equal((old.deck.slides[0].chart as unknown as typeof V0).options.orientation, 'horizontal');
  assert.equal(buildDeck(old.deck, m).slides.length, 4);
  // Slide 2: OC_SPEC split in two chunks, the second one lost; slide 3: OC_SPEC hand-edited.
  let seen = 0;
  const damaged = rewrite(built, (name, text) => {
    if (!name.startsWith('ppt/tags/') || !text.includes('name="OC_SPEC"')) return text;
    seen++;
    if (seen === 2) return text.replace('name="OC_SPEC"', 'name="OC_SPEC_0"').replace('</p:tagLst>', '<p:tag name="OC_SPEC_N" val="2"/></p:tagLst>');
    if (seen === 3) return text.replace(/name="OC_SPEC" val="[^"]*"/, 'name="OC_SPEC" val="{not json"');
    return text;
  });
  assert.equal(seen, 4);
  const charts = readCharts(damaged).flatMap((s) => s.charts);
  assert.equal(charts.length, 4);
  assert.equal(charts[1].spec, undefined);
  assert.match(charts[1].error!, /OC_SPEC chunk 1 of 2 is missing/);
  const { deck, warnings } = importDeck(damaged);
  assert.deepEqual(deck.slides.map((s) => s.title), ['Old', 'Fine']);
  assert.equal(warnings.length, 2);
  assert.match(warnings[0], /^slide 2: .* is damaged \(OC_SPEC chunk 1 of 2 is missing\); skipped$/);
  assert.match(warnings[1], /^slide 3: .* has a damaged OC_SPEC; skipped$/);
});

// ---- Charts edited by hand in PowerPoint -----------------------------------------------------------

test('hand edits on every chart type: deleted, recoloured, retyped, oversized and untagged pieces update cleanly', async () => {
  for (const item of items) {
    const r = await insert(item);
    assert.deepEqual(r.errors, [], item);
    // Every third piece deleted, or its OC_ROLE lost (the VBA skips those).
    const native = nativeOf(r).filter((_, i) => i % 3 !== 2);
    const mark = native.find((e) => e.kind === 'mark' && e.props.fill && e.props.fill !== 'none');
    const text = native.find((e) => e.kind === 'text' && /^label:/.test(e.role));
    const edited = native.map((e) => {
      if (e === mark) return { ...e, props: { ...e.props, fill: 'C00000' } };
      // Text retyped and set to 96 pt (more than a spec allows): kept as 72 pt.
      if (e === text) return { ...e, props: { ...e.props, text: 'Typed', size: 96 } };
      // Text in several sizes / fonts (PowerPoint reports no single value), lines far too wide.
      if (e.kind === 'text') return { ...e, props: { ...e.props, size: -2, font: '' } };
      if (e.kind === 'line') return { ...e, props: { ...e.props, lineWidth: 40 } };
      return e;
    });
    // A shape the user drew inside the group (no OC_ROLE), a copied piece (same role twice), junk.
    edited.push({ role: '', kind: 'mark', props: { fill: '00FF00' } }, ...(mark ? [{ ...mark, props: { fill: '0000FF' } }] : []));
    const junk = [...edited, null, { role: 'nothing:here', kind: 'mark', props: { fill: '123456' } }, { role: mark?.role, kind: 'text', props: { size: '12', text: 5, bold: 'yes' } }];
    const u = await update(r, junk);
    assert.deepEqual(u.errors, [], item);
    assert.equal(u.frame, tag(r, 'OC_FRAME'), `${item}: frame kept`);
    const o = own(u);
    if (mark) assert.equal(o[mark.role]?.fill, 'C00000', `${item}: recolour kept`);
    if (text) {
      assert.deepEqual(o[text.role]?.content, [{ text: 'Typed' }], `${item}: retyped text kept`);
      assert.equal(o[text.role]?.size, 72, `${item}: 96 pt kept as 72 pt`);
    }
    for (const [role, st] of Object.entries(o) as [string, Record<string, unknown>][]) {
      assert.ok(st.font === undefined, `${item} ${role}: no font from mixed text`);
      assert.ok(st.size === undefined || st.size === 72, `${item} ${role}: size ${st.size}`);
      assert.ok(st.lineWidth === undefined || st.lineWidth === 20, `${item} ${role}: lineWidth ${st.lineWidth}`);
    }
    // Redrawn from the spec, OC_BASE is what is drawn: the next automatic check has nothing to do,
    // also when the 72 pt label is set past 72 pt again (kept as 72 pt: no change).
    const redrawn = nativeOf(u).map((e) => (e.role === text?.role && e.kind === 'text' ? { ...e, props: { ...e.props, size: 96 } } : e));
    const again = await update(u, redrawn, { op: 'relayout' });
    assert.deepEqual(again.errors, [], item);
    assert.equal(again.nothing, true, `${item}: relayout after the edits is a no-op`);
  }
});

test('hand edits: damaged OC_BASE / OC_FRAME and a moved group still update', async () => {
  const r = await insert('stacked');
  const native = nativeOf(r);
  // OC_BASE chunk lost (truncated) or edited: native edits cannot be told apart, the chart still updates.
  for (const base of [tag(r, 'OC_BASE').slice(0, 50), '[]', 'null']) {
    const u = await update(r, native, { base });
    assert.deepEqual(u.errors, []);
    assert.deepEqual(own(u), {});
  }
  // OC_FRAME edited into nonsense: the group's bounds are the frame.
  const { scene } = sceneFor(JSON.parse(tag(r, 'OC_SPEC')), `${F}\n\n`, m);
  const b = sceneBounds(scene)!;
  const u = await update(r, native, { frame: 'x,y', bounds: [b.x, b.y, b.w, b.h].join(',') });
  assert.deepEqual(u.errors, []);
  assert.equal(u.frame, F);
  // Moved by hand (without the add-in seeing it): laid out where the group now is, same size.
  const moved = await update(r, native, { bounds: [b.x + 100, b.y + 20, b.w, b.h].join(',') });
  assert.equal(moved.frame, '148,116,864,408');
});

// ---- Copy / paste into another deck -----------------------------------------------------------------

test('a chart pasted into a deck with other theme colours, fonts and slide size re-lays out in that deck, frame kept', async () => {
  // Deck B's body font is wider than deck A's.
  const wide: Measurer = { ...m, width: (t, s, b) => m.width(t, s, b) * 1.4, family: 'Wide Sans', forFont: (f) => fonts.forFont!(f) };
  const fonts: Measurer = { ...m, forFont: (f) => (f === 'Wide Sans' ? wide : m.forFont!(f)) };
  const deckA = 'majorfont=Arial\nminorfont=Arial\n';
  // Deck B: another colour scheme (accent1 a pale yellow) and theme fonts.
  const deckB = 'theme=000000,FFFFFF,44546A,E7E6E6,F2E394,ED7D31,A5A5A5,FFC000,5B9BD5,70AD47,0563C1,954F72\nmajorfont=Wide Sans\nminorfont=Wide Sans\n';
  for (const item of items) {
    const r = await insert(item, deckA);
    assert.deepEqual(r.errors, [], item);
    const frame = tag(r, 'OC_FRAME');
    const { scene } = sceneFor(JSON.parse(tag(r, 'OC_SPEC')), `${frame}\n\n${deckA}`, m);
    const b = sceneBounds(scene)!;
    // Pasted: the group keeps its tags and its size; on a 4:3 slide (720 pt wide) it may also land elsewhere.
    for (const dx of [0, -60]) {
      const bounds = [b.x + dx, b.y, b.w, b.h].join(',');
      const same = await ask(`update\n\n${frame}\n${bounds}\n${deckA}${tag(r, 'OC_SPEC')}`, fonts);
      const pasted = await ask(`update\n\n${frame}\n${bounds}\n${deckB}${tag(r, 'OC_SPEC')}`, fonts);
      assert.deepEqual(pasted.errors, [], item);
      const f = frame.split(',').map(Number);
      assert.equal(pasted.frame, [f[0] + dx, ...f.slice(1)].join(','), `${item}: frame kept (moved ${dx})`);
      assert.equal(tag(pasted, 'OC_SPEC'), tag(same, 'OC_SPEC'), `${item}: the stored chart does not take deck B's theme`);
      assert.ok(!tag(pasted, 'OC_SPEC').includes('Wide Sans') && !tag(pasted, 'OC_SPEC').includes('F2E394'));
      if (dx) continue;
      // Text in the deck's body font, measured as deck B's: the layout changes with the font.
      assert.ok(pasted.shapes.filter((s) => s.kind === 'text').every((s) => s.font === '+mn-lt'), `${item}: text in the deck's body font`);
      assert.notDeepEqual(pasted.shapes.map((s) => [s.x, s.w]), same.shapes.map((s) => [s.x, s.w]), `${item}: laid out with deck B's font`);
    }
  }
  // Colours: theme slots stay slots (PowerPoint shows deck B's), label contrast decided on deck B's accent1.
  const r = await insert('stacked', deckA);
  const inB = await ask(`update\n\n${F}\n\n${deckB}${tag(r, 'OC_SPEC')}`, fonts);
  assert.equal(inB.shapes.find((s) => s.role === 'bar:s0:c0')!.fill, 'accent1');
  const label = (x: Res) => x.shapes.find((s) => s.role === 'label:value:s0:c0')!.color;
  assert.notEqual(label(inB), label(r), 'label colour follows deck B\'s accent1');
});
