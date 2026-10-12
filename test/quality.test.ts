// The KPI checks behind `npm run quality` (scripts/kpis.ts) must pass good layouts and catch broken ones.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { layoutChart, type ChartSpec, type Scene, type TextShape } from '../src/core/index.ts';
import { createMeasurer } from '../src/fonts.ts';
import { CHART_FRAME } from '../src/deck.ts';
import { cagrChecks, hundredChecks, mekkoChecks, numbers, stackChecks, typeChecks, waterfallChecks } from '../scripts/kpis.ts';

const m = createMeasurer();
const lay = (spec: ChartSpec) => layoutChart(spec, CHART_FRAME, m);
const failed = (cs: { ok: boolean }[]) => cs.filter((c) => !c.ok).length;
/** Change the text of the first label whose role matches. */
const retext = (sc: Scene, re: RegExp, text: string) => {
  const i = sc.roles!.findIndex((r, k) => re.test(r) && sc.shapes[k].kind === 'text');
  (sc.shapes[i] as TextShape).text = text;
  return sc;
};

test('quality numbers(): locale marks, minus signs, parentheses, units', () => {
  assert.deepEqual(numbers('1,234.5m'), [1234.5]);
  assert.deepEqual(numbers('−3.4pp'), [-3.4]);
  assert.deepEqual(numbers('(12)'), [-12]);
  assert.deepEqual(numbers('1.234,5 €', 'de-DE'), [1234.5]);
  assert.deepEqual(numbers('412 (46%)'), [412, 46]);
});

test('quality: 100% sums pass for the engine and fail for a broken label', () => {
  const spec: ChartSpec = { type: 'percent', categories: ['A', 'B'], series: [{ name: 'x', values: [1, 2] }, { name: 'y', values: [1, 1] }, { name: 'z', values: [1, 3] }] };
  const sc = lay(spec);
  const ok = hundredChecks(spec, sc);
  assert.equal(ok.length, 2);
  assert.equal(failed(ok), 0);
  assert.equal(failed(hundredChecks(spec, retext(sc, /^label:value:s0:c0$/, '35%'))), 1);
});

test('quality: stacked segments must add up to the total shown', () => {
  const spec: ChartSpec = { type: 'stacked', categories: ['A'], series: [{ name: 'x', values: [1.4] }, { name: 'y', values: [1.4] }, { name: 'z', values: [1.4] }], options: { decimals: 0 } };
  const sc = lay(spec);
  assert.equal(failed(stackChecks(spec, sc)), 0); // consistent rounding: 1 + 1 + 2 = 4
  assert.equal(failed(stackChecks(spec, retext(sc, /^label:total:c0$/, '5'))), 1);
});

test('quality: CAGR label is recomputed from the shown totals', () => {
  const spec: ChartSpec = { type: 'stacked', categories: ['2020', '2024'], series: [{ name: 'x', values: [100, 146] }], options: { cagr: { from: 0, to: 1 } } };
  const sc = lay(spec);
  assert.equal(cagrChecks(spec, sc).length, 1);
  assert.equal(failed(cagrChecks(spec, sc)), 0);
  assert.equal(failed(cagrChecks(spec, retext(sc, /^arrow:cagr$/, '+9.0%'))), 1);
});

test('quality: waterfall totals and connectors balance', () => {
  const spec: ChartSpec = { type: 'waterfall', categories: ['Start', 'Up', 'Down', 'End'], series: [{ name: 'v', values: [100, 20, -5, 'e'] }] };
  const sc = lay(spec);
  const ok = waterfallChecks(spec, sc);
  assert.ok(ok.length >= 4, `${ok.length} checks`);
  assert.equal(failed(ok), 0);
  assert.equal(failed(waterfallChecks(spec, retext(sc, /^label:total:c3$/, '120'))), 1);
  const moved = lay(spec);
  const k = moved.roles!.indexOf('connector:c1');
  (moved.shapes[k] as { y1: number }).y1 += 3;
  assert.equal(failed(waterfallChecks(spec, moved)), 1);
});

test('quality: Mekko widths sum to the plot and follow the totals', () => {
  const spec: ChartSpec = { type: 'mekko', categories: ['A', 'B', 'C'], series: [{ name: 'x', values: [50, 30, 20] }, { name: 'y', values: [10, 30, 60] }] };
  const sc = lay(spec);
  assert.equal(failed(mekkoChecks(spec, sc)), 0);
  for (let i = 0; i < sc.shapes.length; i++) if (sc.roles![i].endsWith(':c1') && sc.roles![i].startsWith('bar:')) (sc.shapes[i] as { w: number }).w += 4;
  assert.ok(failed(mekkoChecks(spec, sc)) >= 2);
});

test('quality: typography flags text below 8 pt and mixed sizes in one role', () => {
  const sc = lay({ type: 'stacked', categories: ['A', 'B'], series: [{ name: 'x', values: [3, 4] }] });
  assert.equal(failed(typeChecks(sc)), 0);
  const i = sc.roles!.indexOf('label:category:c0');
  (sc.shapes[i] as TextShape).size = 7;
  assert.equal(failed(typeChecks(sc)), 2);
});
