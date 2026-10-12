// Example titles state facts about their data ("grew 2.2x", "overtook in 2022"): every such claim is
// recomputed here, so a title and its data cannot drift apart.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

type Slide = { title: string; chart: { categories: string[]; series: { name: string; values: (number | null | string)[] }[] } };
const slides = (f: string): Slide[] => JSON.parse(readFileSync(`examples/${f}.json`, 'utf8')).slides;
const slide = (f: string, title: RegExp) => {
  const s = slides(f).find((x) => title.test(x.title));
  assert.ok(s, `${f}: a slide titled ${title}`);
  return s!;
};
const col = (s: Slide, c: number) => s.chart.series.reduce((a, x) => a + (typeof x.values[c] === 'number' ? (x.values[c] as number) : 0), 0);
const series = (s: Slide, name: string) => s.chart.series.find((x) => x.name === name)!.values as number[];

test('deck: "grew N.Nx in four years" is the first-to-last total ratio', () => {
  const s = slide('deck', /grew [\d.]+x in four years/);
  const n = s.chart.categories.length;
  assert.equal(n - 1, 4);
  const claim = +s.title.match(/grew ([\d.]+)x/)![1];
  assert.equal(claim, +(col(s, n - 1) / col(s, 0)).toFixed(1));
  // "led by Enterprise": the largest absolute gain.
  const gain = (x: { values: unknown[] }) => (x.values[n - 1] as number) - (x.values[0] as number);
  const top = [...s.chart.series].sort((a, b) => gain(b) - gain(a))[0];
  assert.ok(s.title.includes(`led by ${top.name}`));
});

test('deck: "Enterprise is now more than half of the mix"', () => {
  const s = slide('deck', /half of the mix/);
  const n = s.chart.categories.length;
  const share = series(s, 'Enterprise')[n - 1] / col(s, n - 1);
  assert.ok(share > 0.5 && /more than half/.test(s.title), `share ${share}`);
});

test('area: online overtook stores in the year it first exceeds them', () => {
  const s = slide('area', /online overtook stores in (\d{4})/);
  const year = s.title.match(/in (\d{4})/)![1];
  const on = series(s, 'Online'), st = series(s, 'Stores');
  assert.equal(s.chart.categories[on.findIndex((v, i) => v > st[i])], year);
});

test('percent: shares quoted in titles match the data', () => {
  const a = slide('percent', /rose from (\d+)% to (\d+)%/);
  const [, from, to] = a.title.match(/from (\d+)% to (\d+)%/)!.map(Number);
  const n = a.chart.categories.length;
  assert.equal(Math.round((series(a, 'Enterprise')[0] / col(a, 0)) * 100), from);
  assert.equal(Math.round((series(a, 'Enterprise')[n - 1] / col(a, n - 1)) * 100), to);

  const b = slide('percent', /a third of buyers are still considering; one in (\w+) gave no answer/);
  const k = b.chart.categories.length;
  const considering = series(b, 'Considering').reduce((x, y) => x + y, 0) / k / 100;
  assert.ok(Math.abs(considering - 1 / 3) < 0.03, `considering ${considering}`);
  const none = b.chart.categories.reduce((x, _, c) => x + (100 - col(b, c)), 0) / k / 100;
  const words: Record<string, number> = { five: 5, six: 6, seven: 7, eight: 8 };
  const said = words[b.title.match(/one in (\w+)/)![1]];
  const nearest = Object.values(words).sort((p, q) => Math.abs(1 / p - none) - Math.abs(1 / q - none))[0];
  assert.equal(said, nearest, `no answer ${none}`);
});

test('pie, mekkoxy, line, waterfall: stated facts hold', () => {
  const pie = slide('pie', /more than half of 2025 revenue/);
  const rev = pie.chart.series[0].values as number[];
  assert.ok(rev[0] / rev.reduce((a, v) => a + v, 0) > 0.5);

  const xy = slide('mekkoxy', /three quarters of capacity produces below \$60\/t/);
  const cap = series(xy, 'Capacity (Mt)'), cost = series(xy, 'Cash cost ($/t)');
  const below = cap.filter((_, i) => cost[i] < 60).reduce((a, v) => a + v, 0) / cap.reduce((a, v) => a + v, 0);
  assert.ok(Math.abs(below - 0.75) < 0.02, `below ${below}`);

  const line = slide('line', /turned positive in (\d{4})/);
  const om = series(line, 'Operating margin');
  assert.equal(line.chart.categories[om.findIndex((v) => v > 0)], line.title.match(/in (\d{4})/)![1]);

  const wf = slide('waterfall', /EBITDA slipped (\d+)m/);
  const deltas = (wf.chart.series[0].values as (number | null)[]);
  const end = deltas.reduce((a: number, v) => a + (v ?? 0), 0);
  assert.equal(deltas[0]! - end, +wf.title.match(/slipped (\d+)m/)![1]);
});
