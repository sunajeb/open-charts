import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import {
  checkScene, layoutChart, migrate, migrateDeck, validateDeck, validateSpec, InvalidSpecError,
  CHART_KEYS, OPTION_KEYS, SERIES_KEYS, STYLE_KEYS, CHART_TYPES, OPTIONS, LIMITS, appliesTo, optionDefaults,
  type ChartSpec, type RectShape, type Scene, type TextShape, type LineShape,
} from '../src/core/index.ts';
import { parseCategoryDate, years360, cagrYears } from '../src/core/dates.ts';
import { samples } from '../src/core/samples.ts';
import { createMeasurer } from '../src/fonts.ts';
import { buildDeck, CHART_FRAME, type DeckSpec } from '../src/deck.ts';

const m = createMeasurer();
const F = CHART_FRAME;
const byRole = (s: Scene, re: RegExp) => s.shapes.filter((_, i) => re.test(s.roles![i]));
const textsOf = (s: Scene, re: RegExp) => byRole(s, re).filter((x): x is TextShape => x.kind === 'text').map((t) => t.text);
const ROLE = /^(bar:s\d+:c\d+|label:value:s\d+:c\d+|label:total:c\d+|label:category:c\d+|label:series:s\d+|axis|axis:tick:\d+|gridline:\d+|arrow:(cagr|diff)|arrow:level:\d+|valueline:\d+|indicator|break:(axis|cat:\d+|s\d+:c\d+)|axis:secondary|axis:secondary:tick:\d+|connector:c\d+|legend:s\d+|marker:s\d+:c\d+|line:s\d+|area:s\d+|wedge:s\d+:c\d+|leader:c\d+|leader:s\d+:c\d+|legend:c\d+|ball:s\d+:c\d+|label:total|shade:\d+|legend:size|bracket:\d+|today)$/;

const base = (type: ChartSpec['type'], options: ChartSpec['options'] = {}): ChartSpec => ({
  type, categories: ['2021', '2022', '2023', '2024'],
  series: [{ name: 'A', values: [10, 12, 15, 18] }, { name: 'B', values: [5, -2, 6, 7] }],
  options,
});

// ---- Roles --------------------------------------------------------------------------------------

test('every shape of every chart type has a role in the shared format', () => {
  const specs: ChartSpec[] = [
    base('stacked', { cagr: { from: 0, to: 3 }, difference: { from: 1, to: 3 }, valueAxis: true }),
    base('clustered', { valueAxis: true }),
    base('percent', { seriesLabels: 'legend' }),
    base('stacked', { orientation: 'horizontal', cagr: { from: 0, to: 3 }, valueAxis: true }),
    base('clustered', { orientation: 'horizontal' }),
    base('percent', { orientation: 'horizontal' }),
    base('line', { seriesLabels: 'legend' }),
    base('line'),
    { ...base('waterfall', { valueAxis: true }), series: [{ name: 'A', values: [10, 3, -2, null] }] },
    base('mekko', { valueAxis: true }),
    base('mekko', { mekko: 'unit' }),
  ];
  for (const spec of specs) {
    const scene = layoutChart(spec, F, m);
    assert.equal(scene.roles!.length, scene.shapes.length, spec.type);
    for (const r of scene.roles!) assert.match(r, ROLE, `${spec.type} ${spec.options?.orientation ?? ''}`);
  }
  const s = layoutChart(base('stacked', { cagr: { from: 0, to: 3 } }), F, m);
  assert.ok(s.roles!.includes('bar:s1:c3') && s.roles!.includes('label:total:c0') && s.roles!.includes('arrow:cagr'));
  assert.deepEqual(textsOf(s, /^label:category:/), ['2021', '2022', '2023', '2024']);
});

test('every example deck chart has roles for every shape', () => {
  for (const file of ['examples/deck.json', 'examples/line.json', 'examples/mekko.json', 'examples/waterfall.json', 'examples/step.json', 'examples/area.json', 'examples/pie.json', 'examples/harvey.json', 'examples/percent.json', 'examples/combo.json', 'examples/mekkoxy.json', 'examples/scatter.json', 'examples/table.json', 'examples/gantt.json']) {
    const deck: DeckSpec = JSON.parse(readFileSync(file, 'utf8'));
    for (const sl of deck.slides) {
      const scene = layoutChart(sl.chart, F, m);
      scene.roles!.forEach((r) => assert.match(r, ROLE, `${file}: ${sl.title}`));
    }
  }
});

// ---- B3: CAGR from dates ------------------------------------------------------------------------

test('B3: category labels parse as dates', () => {
  assert.deepEqual(parseCategoryDate('2021'), { y: 2021, m: 1, d: 1 });
  assert.deepEqual(parseCategoryDate('2025E'), { y: 2025, m: 1, d: 1 });
  assert.deepEqual(parseCategoryDate('FY24'), { y: 2024, m: 1, d: 1 });
  assert.deepEqual(parseCategoryDate('Q3 2023'), { y: 2023, m: 7, d: 1 });
  assert.deepEqual(parseCategoryDate('Q1 FY24'), { y: 2024, m: 1, d: 1 });
  assert.deepEqual(parseCategoryDate('2023-Q4'), { y: 2023, m: 10, d: 1 });
  assert.deepEqual(parseCategoryDate('1Q23'), { y: 2023, m: 1, d: 1 });
  assert.deepEqual(parseCategoryDate('H2 2022'), { y: 2022, m: 7, d: 1 });
  assert.deepEqual(parseCategoryDate('Jan 2023'), { y: 2023, m: 1, d: 1 });
  assert.deepEqual(parseCategoryDate('September 2023'), { y: 2023, m: 9, d: 1 });
  assert.deepEqual(parseCategoryDate('Mar-24'), { y: 2024, m: 3, d: 1 });
  assert.deepEqual(parseCategoryDate('2023-02'), { y: 2023, m: 2, d: 1 });
  assert.deepEqual(parseCategoryDate('2023-02-15'), { y: 2023, m: 2, d: 15 });
  assert.equal(parseCategoryDate('North America'), null);
  assert.equal(parseCategoryDate('Q5 2023'), null);
  assert.equal(parseCategoryDate('Marketing 24'), null);
});

test('B3: 30/360 day count', () => {
  assert.equal(years360({ y: 2020, m: 1, d: 1 }, { y: 2024, m: 1, d: 1 }), 4);
  assert.equal(years360({ y: 2023, m: 1, d: 31 }, { y: 2023, m: 7, d: 31 }), 0.5);
  assert.equal(years360({ y: 2023, m: 2, d: 15 }, { y: 2024, m: 5, d: 15 }), 1.25);
});

test('B3: quarterly CAGR uses dates, explicit years still win, unparsable labels fall back to index', () => {
  const q = ['Q1 2023', 'Q2 2023', 'Q3 2023', 'Q4 2023', 'Q1 2024'];
  assert.deepEqual(cagrYears(q, 0, 4), { years: 1, source: 'dates' });
  assert.deepEqual(cagrYears(q, 0, 4, 2), { years: 2, source: 'explicit' });
  assert.deepEqual(cagrYears(['A', 'B', 'C'], 0, 2), { years: 2, source: 'index' });
  assert.deepEqual(cagrYears(['Jan 2023', 'Jul 2023'], 0, 1), { years: 0.5, source: 'dates' });

  const spec: ChartSpec = { type: 'stacked', categories: q, series: [{ name: 'x', values: [100, 103, 106, 108, 110] }], options: { cagr: { from: 0, to: 4 } } };
  assert.deepEqual(textsOf(layoutChart(spec, F, m), /^arrow:cagr$/), ['+10.0%']); // not +2.4% (4 "years")
  spec.options = { cagr: { from: 0, to: 4, years: 2 } };
  assert.deepEqual(textsOf(layoutChart(spec, F, m), /^arrow:cagr$/), ['+4.9%']);
});

// ---- B5 -----------------------------------------------------------------------------------------

test('B5: null in a clustered chart draws no bar and no label', () => {
  const scene = layoutChart({ type: 'clustered', categories: ['a', 'b'], series: [{ name: 'x', values: [5, null] }, { name: 'y', values: [3, 4] }] }, F, m);
  assert.equal(byRole(scene, /^bar:s0:c1$/).length, 0);
  assert.equal(byRole(scene, /^label:value:s0:c1$/).length, 0);
  assert.equal(byRole(scene, /^bar:/).length, 3);
  assert.ok(!textsOf(scene, /^label:value:/).includes('0'));
  const h = layoutChart({ type: 'clustered', categories: ['a', 'b'], series: [{ name: 'x', values: [5, null] }], options: { orientation: 'horizontal' } }, F, m);
  assert.equal(byRole(h, /^bar:/).length, 1);
});

// ---- B8 -----------------------------------------------------------------------------------------

test('B8: a word wider than its slot shrinks the category row, then hides, and is reported', () => {
  const cats = Array.from({ length: 12 }, (_, i) => (i === 3 ? 'Internationalisation' : `C${i}`));
  const scene = layoutChart({ type: 'stacked', categories: cats, series: [{ name: 'x', values: cats.map(() => 5) }] }, F, m);
  const issues = checkScene(scene, { measurer: m, strict: true });
  assert.equal(issues.filter((i) => i.kind === 'overflow').length, 0, 'never silent overflow');
  assert.ok(scene.hidden.some((h) => /shrunk to \d+ pt|too wide for its slot/.test(h)), scene.hidden.join('; '));
  assert.ok(issues.some((i) => i.kind === 'hidden'));
  // Every category label the scene does draw fits its box.
  const cat = byRole(scene, /^label:category:/) as TextShape[];
  for (const t of cat) assert.ok(m.width(t.text, t.size) <= t.w + 0.01, t.text);
  // CJK wraps between ideographs instead of overflowing.
  const cjk = layoutChart({ type: 'stacked', categories: ['東京都港区六本木一丁目', '大阪', '名古屋', '福岡', '札幌', '仙台', '広島', '神戸', '京都', '横浜', '川崎', '千葉'], series: [{ name: 'x', values: Array(12).fill(1) }] }, F, m);
  assert.deepEqual(checkScene(cjk, { measurer: m }).filter((i) => i.kind === 'overflow'), []);
  assert.ok((byRole(cjk, /^label:category:c0$/)[0] as TextShape).text.includes('\n'));
});

test('B8: the gate reports glyph overflow when given a measurer', () => {
  const scene: Scene = { frame: F, hidden: [], roles: ['axis'], shapes: [{ kind: 'text', x: 60, y: 100, w: 10, h: 14, text: 'Overflowing', size: 12, color: '000000', align: 'left', valign: 'top' }] };
  assert.equal(checkScene(scene, { measurer: m })[0].kind, 'overflow');
  assert.deepEqual(checkScene(scene), []);
});

// ---- B9 -----------------------------------------------------------------------------------------

test('B9: the legend wraps onto more rows instead of leaving the frame', () => {
  const series = Array.from({ length: 8 }, (_, i) => ({ name: `Business unit number ${i + 1} (incl. services)`, values: [i + 1, i + 2] }));
  const scene = layoutChart({ type: 'clustered', categories: ['2023', '2024'], series }, F, m);
  const legend = byRole(scene, /^legend:/).filter((s): s is TextShape => s.kind === 'text');
  assert.ok(new Set(legend.map((t) => Math.round(t.y))).size > 1, 'more than one row');
  assert.deepEqual(checkScene(scene, { measurer: m }).filter((i) => i.kind === 'out-of-frame' || i.kind === 'overlap'), []);
});

// ---- B10 / B11 ----------------------------------------------------------------------------------

test('B10: line labels avoid data lines found by role, not stroke width', () => {
  const scene = layoutChart({ type: 'line', categories: ['a', 'b', 'c', 'd'], series: [{ name: 'x', values: [10, 30, 10, 30] }, { name: 'y', values: [30, 10, 30, 10] }] }, F, m);
  const lines = byRole(scene, /^line:/) as LineShape[];
  assert.ok(lines.length === 6 && lines.every((l) => l.width === 2));
  assert.deepEqual(checkScene(scene, { measurer: m }), []);
});

test('B11: right-hand series names never collide with the last point\'s labels', () => {
  const scene = layoutChart({
    type: 'line', categories: ['2019', '2020', '2021', '2022', '2023', '2024'],
    series: [{ name: 'Revenue', values: [100, 104, 108, 112, 116, 120.5] }, { name: 'Costs', values: [99, 103, 107, 111, 115, 119.8] }, { name: 'Cash', values: [98, 102, 106, 110, 114, 119.2] }],
  }, F, m);
  const issues = checkScene(scene, { measurer: m });
  assert.deepEqual(issues.filter((i) => i.kind !== 'hidden'), []);
  const names = byRole(scene, /^label:series:/).filter((s): s is TextShape => s.kind === 'text');
  const last = byRole(scene, /^label:value:s\d+:c5$/) as TextShape[];
  for (const n of names) for (const l of last) assert.ok(n.x >= l.x + l.w - 0.01, `${n.text} vs ${l.text}`);
});

// ---- B12 / B13 ----------------------------------------------------------------------------------

test('B12: gate v2 catches text over marks, shapes outside the frame and arrows through labels', () => {
  const t = (x: number, y: number, text: string): TextShape => ({ kind: 'text', x, y, w: 30, h: 14, text, size: 12, color: '000000', align: 'center', valign: 'middle' });
  const scene: Scene = {
    frame: { x: 0, y: 0, w: 400, h: 300 }, hidden: [],
    shapes: [
      { kind: 'rect', x: 100, y: 100, w: 40, h: 100, fill: '000000' },
      t(105, 120, '5'),                         // own bar: fine
      t(105, 150, 'other'),                     // another series' label on this bar
      { kind: 'rect', x: 200, y: 250, w: 40, h: 80, fill: '000000' }, // bar below the frame
      { kind: 'line', x1: 0, y1: 50, x2: 400, y2: 50, color: '000000', width: 0.75 },
      t(50, 43, 'total'),                       // the arrow runs through it
      { kind: 'line', x1: 0, y1: 270, x2: 380, y2: 10, color: '000000', width: 2 },
      t(300, 60, 'hit'),                        // data line of another series crosses it
    ],
    roles: ['bar:s0:c0', 'label:value:s0:c0', 'label:value:s1:c0', 'bar:s1:c1', 'arrow:cagr', 'label:total:c1', 'line:s0', 'label:value:s1:c3'],
  };
  const kinds = checkScene(scene).map((i) => `${i.kind} ${i.detail}`);
  assert.ok(kinds.includes('text-mark "other" × bar:s0:c0'), kinds.join('\n'));
  assert.ok(!kinds.some((k) => k.startsWith('text-mark "5"')));
  assert.ok(kinds.includes('out-of-frame rect bar:s1:c1'));
  assert.ok(kinds.includes('arrow-label arrow:cagr × "total"'));
  assert.ok(kinds.includes('text-mark "hit" × line:s0'));
});

test('B13: hidden labels are issues in strict mode, with an allow-list', () => {
  // A thin middle segment between touching columns (gap 0.05): no room inside, beside or above it.
  const big = [90, 95, 100, 105, 110, 115];
  const scene = layoutChart({
    type: 'stacked', categories: ['a', 'b', 'c', 'd', 'e', 'f'],
    series: [{ name: 'Big', values: big }, { name: 'Thin', values: big.map(() => 1) }, { name: 'Big2', values: big }], options: { gap: 0.05 },
  }, F, m);
  assert.ok(scene.hidden.length > 0);
  assert.equal(checkScene(scene).filter((i) => i.kind === 'hidden').length, 0);
  assert.equal(checkScene(scene, { strict: true }).filter((i) => i.kind === 'hidden').length, scene.hidden.length);
  assert.deepEqual(checkScene(scene, { strict: true, allowHidden: ['segment too small'] }), []);
});

test('B13: example decks pass the strict gate; declared stress slides list exactly what they hide', () => {
  // Thin segments get their labels boxed or outside with a leader; what still cannot be placed is
  // declared here, so any new hidden label fails.
  const allowed: Record<string, RegExp[]> = {
    'examples/waterfall.json#2': [/^category labels shrunk to 11 pt to fit$/],
  };
  for (const file of ['examples/deck.json', 'examples/line.json', 'examples/mekko.json', 'examples/waterfall.json', 'examples/step.json', 'examples/area.json', 'examples/pie.json', 'examples/harvey.json', 'examples/percent.json', 'examples/combo.json', 'examples/mekkoxy.json', 'examples/scatter.json', 'examples/table.json', 'examples/gantt.json']) {
    const deck: DeckSpec = JSON.parse(readFileSync(file, 'utf8'));
    deck.slides.forEach((sl, i) => {
      const scene = layoutChart(sl.chart, F, m);
      const issues = checkScene(scene, { strict: true, measurer: m, allowHidden: allowed[`${file}#${i}`] ?? [] });
      assert.deepEqual(issues, [], `${file} #${i}: ${sl.title}`);
    });
  }
});

// ---- B14 in charts ------------------------------------------------------------------------------

test('B14: stacked segment labels add up to the total shown', () => {
  const spec: ChartSpec = { type: 'stacked', categories: ['a'], series: [{ name: 'x', values: [1.4] }, { name: 'y', values: [1.4] }, { name: 'z', values: [1.4] }], options: { decimals: 0 } };
  const scene = layoutChart(spec, F, m);
  const segs = textsOf(scene, /^label:value:/).map(Number);
  assert.equal(segs.reduce((a, b) => a + b, 0), Number(textsOf(scene, /^label:total:/)[0]));
});

test('B14: locale, unit prefix, magnitude and format reach the labels and the axis', () => {
  const spec: ChartSpec = {
    type: 'clustered', categories: ['a', 'b'],
    series: [{ name: 'x', values: [1_250_000, 2_500_000] }],
    options: { scale: 'auto', unit: { text: '$', position: 'prefix' }, valueAxis: true },
  };
  assert.deepEqual(textsOf(layoutChart(spec, F, m), /^label:value:/), ['$1.25m', '$2.50m']);
  assert.ok(textsOf(layoutChart(spec, F, m), /^axis:tick:/).includes('$2.5m'));
  const de = layoutChart({ type: 'clustered', categories: ['a'], series: [{ name: 'x', values: [1234.5] }], options: { locale: 'de-DE' } }, F, m);
  assert.deepEqual(textsOf(de, /^label:value:/), ['1.234,5']);
  const neg = layoutChart({ type: 'clustered', categories: ['a', 'b'], series: [{ name: 'x', values: [-1234, 99] }], options: { format: '#,##0;(#,##0)' } }, F, m);
  assert.deepEqual(textsOf(neg, /^label:value:/), ['(1,234)', '99']);
  const pct = layoutChart({ type: 'percent', categories: ['a'], series: [{ name: 'x', values: [1] }, { name: 'y', values: [3] }], options: { locale: 'fr-FR' } }, F, m);
  assert.deepEqual(textsOf(pct, /^label:value:/), ['25 %', '75 %']);
});

// ---- B15 + spec v1 ------------------------------------------------------------------------------

test('B15: validation is path-addressed and catches bad numbers, names, colours and options', () => {
  const errs = validateSpec({
    type: 'stacked', categories: ['a', 'b', 3],
    series: [
      { name: 'x', values: [1, NaN, 2], color: 'blue' },
      { name: 'x', values: [1, '2', Infinity, 4] },
      { name: 'y', values: [1, 2, 'e'] },
    ],
    options: { cagr: { from: 1, to: 1 }, decimals: 1.5, bogus: true, orientation: 'sideways', unit: { text: '$', position: 'left' } },
  });
  const paths = Object.fromEntries(errs.map((e) => [e.path, e.message]));
  assert.match(paths['/categories/2'], /string/);
  assert.match(paths['/series/0/values/1'], /finite number, got NaN/);
  assert.match(paths['/series/0/color'], /hex colour/);
  assert.match(paths['/series/1/name'], /duplicate series name "x"/);
  assert.match(paths['/series/1/values'], /4 values but there are 3 categories/);
  assert.match(paths['/series/1/values/1'], /number or null, got string "2"/);
  assert.match(paths['/series/1/values/2'], /finite number, got Infinity/);
  assert.match(paths['/series/2/values/2'], /number or null, got string "e"/, '"e" is only valid in a waterfall');
  assert.match(paths['/options/cagr/to'], /different/);
  assert.match(paths['/options/decimals'], /whole number/);
  assert.match(paths['/options/bogus'], /unknown property/);
  assert.match(paths['/options/orientation'], /vertical, horizontal/);
  assert.match(paths['/options/unit/position'], /prefix/);
  assert.deepEqual(validateSpec(base('stacked')), []);
  assert.deepEqual(validateSpec({ ...base('waterfall'), series: [{ name: 'a', values: [1, 2, -1, 'e'] }] }), []);
  assert.deepEqual(validateSpec({ ...base('line', { orientation: 'horizontal' }) }), [{ path: '/options/orientation', message: 'not used by line charts (only stacked, clustered, percent)' }]);
});

test('B15: layoutChart throws one error listing every "/path: message"', () => {
  try {
    layoutChart({ type: 'stacked', categories: ['a', 'b'], series: [{ name: 'x', values: [1, 'e' as never] }, { name: 'x', values: [1, 2] }] }, F, m);
    assert.fail('should throw');
  } catch (err) {
    assert.ok(err instanceof InvalidSpecError);
    assert.match((err as Error).message, /^Invalid chart spec:\n {2}\/series\/0\/values\/1: .*\n {2}\/series\/1\/name: duplicate/);
    assert.equal((err as InvalidSpecError).errors.length, 2);
  }
});

test('waterfall "e" cells mark computed totals and do not change inferred decimals', () => {
  const scene = layoutChart({ type: 'waterfall', categories: ['Start', 'Price', 'Cost', 'End'], series: [{ name: 'v', values: [100, 20, -5, 'e'] }], options: { unit: 'm' } }, F, m);
  assert.deepEqual(textsOf(scene, /^label:total:/), ['100m', '115m']);
  assert.deepEqual(textsOf(scene, /^label:value:/).sort(), ['+20m', '−5m']);
  const mid = layoutChart({ type: 'waterfall', categories: ['S', 'a', 'Sub', 'b', 'E'], series: [{ name: 'v', values: [10, 5, 'e', -3, 'e'] }] }, F, m);
  assert.deepEqual(textsOf(mid, /^label:total:/), ['10', '15', '12']);
});

test('multi-series waterfall: legend swatches match the delta bars', () => {
  const scene = layoutChart({ type: 'waterfall', categories: ['S', 'a', 'b', 'E'], series: [{ name: 'Organic', values: [10, 2, -1, null] }, { name: 'Acquired', values: [0, 3, 1, null] }] }, F, m);
  for (const s of [0, 1]) {
    const swatch = byRole(scene, new RegExp(`^legend:s${s}$`)).find((x): x is RectShape => x.kind === 'rect')!;
    const bar = byRole(scene, new RegExp(`^bar:s${s}:c1$`))[0] as RectShape;
    assert.equal(swatch.fill, bar.fill);
  }
});

test('spec v1: migrate stamps v0 losslessly; decks too; newer versions are rejected', () => {
  const v0 = base('stacked', { cagr: { from: 0, to: 3 }, unit: 'm' });
  const v1 = migrate(v0);
  assert.equal(v1.specVersion, 1);
  const { specVersion, ...rest } = v1;
  assert.deepEqual(rest, v0);
  assert.equal((v0 as { specVersion?: number }).specVersion, undefined, 'input not mutated');
  assert.equal(migrate(v1), v1);
  const deck = migrateDeck({ slides: [{ title: 't', chart: v0 }] });
  assert.equal(deck.specVersion, 1);
  assert.equal((deck.slides[0].chart as ChartSpec).specVersion, 1);
  assert.match(validateSpec({ ...v0, specVersion: 2 })[0].message, /unsupported specVersion 2/);
  assert.throws(() => layoutChart({ ...v0, specVersion: 2 as 1 }, F, m), /specVersion/);
  assert.equal(layoutChart({ ...v0, specVersion: 1 }, F, m).shapes.length, layoutChart(v0, F, m).shapes.length);
});

// ---- Options: one applicability table, one set of defaults, one set of limits --------------------

/** A valid value for each option on a chart whose categories are dates, for applicability tests. */
const SAMPLE_OPTION: Record<string, unknown> = {
  unit: 'm', decimals: 1, format: '#,##0', locale: 'de-DE', scale: 'k', orientation: 'horizontal',
  valueAxis: true, zeroBased: true, gap: 0.3, seriesLabels: 'legend', totals: true,
  cagr: { from: 0, to: 1 }, difference: { from: 0, to: 1 }, totalsAt: [0], step: true, percentData: true,
  combo: 'clustered', mekko: 'unit', sort: 'reverse', axisTitles: { x: 'x' }, sizeLegend: true,
  labelContent: 'value', shades: [{ from: '2026-01-01', to: '2026-02-01' }],
  brackets: [{ from: '2026-01-01', to: '2026-02-01' }], today: '2026-01-15',
  levelArrows: [{ from: { category: 0, series: 0 }, to: { category: 1, series: 0 } }], valueLines: [{ value: 1 }],
  percentIndicator: 'left', percentBase: [1, 1], categoryGaps: [0], categoryBreaks: [0], axisBreak: { from: 1, to: 2 }, secondaryAxis: [0],
};

test('options: every option on a type that does not read it is a path-addressed error', () => {
  assert.deepEqual(Object.keys(SAMPLE_OPTION).sort(), [...OPTION_KEYS].sort());
  for (const type of CHART_TYPES) {
    const spec = { ...samples[type]!, options: {} };
    for (const k of OPTION_KEYS) {
      const errs = validateSpec({ ...spec, options: { [k]: SAMPLE_OPTION[k] } }).filter((e) => /not used by/.test(e.message));
      if (appliesTo(k, type)) assert.deepEqual(errs, [], `${k} on ${type}`);
      else assert.deepEqual(errs, [{ path: `/options/${k}`, message: `not used by ${type} charts (only ${OPTIONS[k].types.join(', ')})` }], `${k} on ${type}`);
    }
  }
  assert.equal(validateSpec(base('line', { cagr: null })).length, 0, 'null means "none" anywhere');
  assert.equal(validateSpec({ ...samples.pie!, options: { cagr: { from: 0, to: 1 } } })[0].message, 'not used by pie charts (only stacked, percent)');
});

test('options: migrate() strips options a v0 chart type does not use, with a warning; v1 is rejected', () => {
  const v0 = { ...samples.pie!, specVersion: undefined, options: { cagr: { from: 0, to: 1 }, totals: true, labelContent: 'value' as const } };
  delete v0.specVersion;
  const warnings: string[] = [];
  const up = migrate(v0, warnings);
  assert.deepEqual(up.options, { labelContent: 'value' });
  assert.deepEqual(warnings, ['/options/cagr: not used by pie charts; removed', '/options/totals: not used by pie charts; removed']);
  assert.deepEqual(validateSpec(up), []);
  assert.deepEqual(v0.options.totals, true, 'input not mutated');
  assert.ok(layoutChart(v0, F, m).shapes.length > 0, 'old specs still lay out');
  assert.throws(() => layoutChart({ ...v0, specVersion: 1 }, F, m), /\/options\/cagr: not used by pie charts/);
  const deckWarnings: string[] = [];
  migrateDeck({ slides: [{ chart: base('stacked') }, { chart: v0 }] }, deckWarnings);
  assert.deepEqual(deckWarnings, ['/slides/1/chart/options/cagr: not used by pie charts; removed', '/slides/1/chart/options/totals: not used by pie charts; removed']);
  const built = buildDeck({ slides: [{ chart: v0 }] }, m);
  assert.ok(built.slides[0].issues.includes('warning: /options/cagr: not used by pie charts; removed'));
});

test('options: optionDefaults() is what the engine draws when an option is left out', () => {
  for (const type of CHART_TYPES) {
    const spec = samples[type]!;
    const defaults = optionDefaults(type);
    for (const k of Object.keys(defaults)) assert.ok(appliesTo(k, type), `${k} default on ${type}`);
    const without = layoutChart({ ...spec, options: {} }, F, m);
    const explicit = layoutChart({ ...spec, options: defaults }, F, m);
    assert.deepEqual(explicit, without, `${type}: explicit defaults draw the same chart`);
  }
  // The defaults that differ by type, as documented.
  assert.equal(optionDefaults('mekko').totals, true);
  assert.equal(optionDefaults('mekkoxy').totals, true);
  assert.equal(optionDefaults('stacked').totals, true);
  assert.equal(optionDefaults('percent').totals, false);
  assert.equal(optionDefaults('line').valueAxis, true);
  assert.equal(optionDefaults('stacked').valueAxis, false);
  assert.equal(optionDefaults('pie').totals, undefined, 'not read by pie');
  const mekko = layoutChart({ ...samples.mekko!, options: {} }, F, m);
  assert.ok(textsOf(mekko, /^label:total:c\d+$/).length > 0, 'Mekko draws totals by default');
  const mekkoOff = layoutChart({ ...samples.mekko!, options: { totals: false } }, F, m);
  assert.equal(textsOf(mekkoOff, /^label:total:c\d+$/).length, 0);
});

test('options: limits are defined once and used by the validator and the schema', () => {
  const chart = JSON.parse(readFileSync('schema/chart.schema.json', 'utf8'));
  for (const [k, { min, max }] of Object.entries(LIMITS)) {
    const prop = chart.$defs.options.properties[k];
    assert.equal(prop.minimum, min, `${k} minimum`);
    assert.equal(prop.maximum, max, `${k} maximum`);
    assert.deepEqual(validateSpec(base('stacked', { [k]: max })), [], `${k} = max`);
    assert.match(validateSpec(base('stacked', { [k]: max + 1 }))[0].message, new RegExp(`${min}–${max}`));
    assert.match(validateSpec(base('stacked', { [k]: min - 1 }))[0].message, new RegExp(`${min}–${max}`));
  }
});

test('validateDeck addresses errors from the deck root', () => {
  const errs = validateDeck({ style: { palette: ['zzz'] }, slides: [{ title: 'a', chart: base('stacked') }, { chart: { ...base('stacked'), series: [] } }], extra: 1 });
  assert.deepEqual(errs.map((e) => e.path).sort(), ['/extra', '/slides/1/chart/series', '/style/palette/0']);
  for (const f of ['examples/deck.json', 'examples/line.json', 'examples/mekko.json', 'examples/waterfall.json']) {
    assert.deepEqual(validateDeck(JSON.parse(readFileSync(f, 'utf8'))), [], f);
  }
});

test('JSON Schema files stay in sync with the validator', () => {
  const chart = JSON.parse(readFileSync('schema/chart.schema.json', 'utf8'));
  const deck = JSON.parse(readFileSync('schema/deck.schema.json', 'utf8'));
  assert.equal(chart.$schema, 'https://json-schema.org/draft/2020-12/schema');
  assert.deepEqual(Object.keys(chart.properties).sort(), [...CHART_KEYS].sort());
  assert.deepEqual(Object.keys(chart.$defs.options.properties).sort(), [...OPTION_KEYS].sort());
  for (const k of OPTION_KEYS) {
    const d: string = chart.$defs.options.properties[k].description ?? '';
    assert.ok(d.startsWith(`Chart types: ${OPTIONS[k].types.join(', ')}.`), `schema description of ${k} starts with its chart types`);
  }
  assert.deepEqual(Object.keys(chart.$defs.series.properties).sort(), [...SERIES_KEYS].sort());
  assert.deepEqual(chart.properties.type.enum, [...CHART_TYPES]);
  assert.deepEqual(Object.keys(deck.$defs.style.properties).sort(), [...STYLE_KEYS].sort());
  assert.equal(deck.properties.slides.items.properties.chart.$ref, chart.$id);
});

// ---- Horizontal bars ----------------------------------------------------------------------------

test('horizontal stacked bars: categories on the left, bars grow right, totals beyond the ends', () => {
  const scene = layoutChart({
    type: 'stacked', categories: ['North America', 'Europe', 'Asia Pacific'],
    series: [{ name: 'Enterprise', values: [40, 30, 20] }, { name: 'SMB', values: [10, 12, 8] }],
    options: { orientation: 'horizontal', difference: { from: 0, to: 2 } },
  }, F, m);
  const bars = byRole(scene, /^bar:/) as RectShape[];
  const cats = byRole(scene, /^label:category:/) as TextShape[];
  const totals = byRole(scene, /^label:total:/) as TextShape[];
  assert.equal(bars.length, 6);
  assert.deepEqual(totals.map((t) => t.text), ['50', '42', '28']);
  for (const c of cats) assert.ok(c.x + c.w <= Math.min(...bars.map((b) => b.x)) + 0.01);
  assert.ok((byRole(scene, /^bar:s0:c0$/)[0] as RectShape).y < (byRole(scene, /^bar:s0:c2$/)[0] as RectShape).y, 'first category on top');
  assert.ok(Math.abs(bars[0].h - bars[1].h) < 1e-9 && bars[0].w > bars[0].h, 'bars are horizontal');
  for (const t of totals) assert.ok(bars.every((b) => b.x + b.w <= t.x + 0.01 || b.y + b.h <= t.y || b.y >= t.y + t.h));
  assert.equal(textsOf(scene, /^arrow:diff$/)[0], '−22');
  assert.deepEqual(checkScene(scene, { strict: true, measurer: m }), []);
});

test('horizontal clustered and percent bars pass the strict gate', () => {
  const clustered = layoutChart({
    type: 'clustered', categories: ['North America', 'Europe', 'Asia Pacific', 'Latin America'],
    series: [{ name: 'Enterprise', values: [38, 31, -5, 22] }, { name: 'Mid-market', values: [29, 26, 24, null] }],
    options: { orientation: 'horizontal', unit: '%', valueAxis: true },
  }, F, m);
  assert.deepEqual(checkScene(clustered, { strict: true, measurer: m }), []);
  const neg = byRole(clustered, /^bar:s0:c2$/)[0] as RectShape;
  const negLabel = byRole(clustered, /^label:value:s0:c2$/)[0] as TextShape;
  assert.ok(negLabel.x + negLabel.w <= neg.x + 0.01, 'negative label sits left of its bar');
  const pct = layoutChart({ ...base('percent', { orientation: 'horizontal' }) }, F, m);
  const shares = textsOf(pct, /^label:value:s\d+:c0$/).map((t) => parseFloat(t));
  assert.equal(shares.reduce((a, b) => a + b, 0), 100);
  assert.deepEqual(checkScene(pct, { strict: true, measurer: m }).filter((i) => i.kind !== 'hidden'), []);
});

test('examples render through buildDeck with lineHeight on every text', () => {
  const deck: DeckSpec = JSON.parse(readFileSync('examples/deck.json', 'utf8'));
  const { slides } = buildDeck(deck, m);
  for (const s of slides) for (const sh of s.scene.shapes) if (sh.kind === 'text') assert.ok(Math.abs(sh.lineHeight! - m.lineHeight(sh.size)) < 1e-9);
});
