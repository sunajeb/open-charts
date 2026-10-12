// Label placement and layout polish for combo, Mekko XY, scatter / bubble, table and Gantt: each test
// pins a defect seen in the rendered goldens.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import {
  checkScene, layoutChart,
  type ChartSpec, type LineShape, type RectShape, type Scene, type TextShape,
} from '../src/core/index.ts';
import { createMeasurer } from '../src/fonts.ts';
import { overlaps, segmentHitsBox } from '../src/core/place.ts';
import { CHART_FRAME } from '../src/deck.ts';

const m = createMeasurer();
const F = CHART_FRAME;
const byRole = (s: Scene, re: RegExp) => s.shapes.filter((_, i) => re.test(s.roles![i]));
const rectsOf = (s: Scene, re: RegExp) => byRole(s, re).filter((x): x is RectShape => x.kind === 'rect');
const textsOf = (s: Scene, re: RegExp) => byRole(s, re).filter((x): x is TextShape => x.kind === 'text');
const lines = (s: Scene, re: RegExp) => byRole(s, re).filter((x): x is LineShape => x.kind === 'line');
const roleOf = (s: Scene, sh: object) => s.roles![s.shapes.indexOf(sh as never)];
const strict = (s: Scene) => checkScene(s, { strict: true, measurer: m });
const close = (a: number, b: number, eps = 1e-6) => Math.abs(a - b) < eps;
const slides = (f: string): { title: string; chart: ChartSpec }[] => JSON.parse(readFileSync(`examples/${f}.json`, 'utf8')).slides;
const segHits = (l: { x1: number; y1: number; x2: number; y2: number }, b: { x: number; y: number; w: number; h: number }) => segmentHitsBox(l.x1, l.y1, l.x2, l.y2, b);
/** Distance between two boxes along the axis that separates them (negative when they overlap). */
const gapBetween = (a: TextShape, b: TextShape) => Math.max(b.x - (a.x + a.w), a.x - (b.x + b.w), b.y - (a.y + a.h), a.y - (b.y + b.h));

// ---- Combination -----------------------------------------------------------------------------------

test('combo: line labels read as their point\'s, segment labels stay inside, no two labels touch', () => {
  for (const { title, chart: spec } of slides('combo')) {
    const scene = layoutChart(spec, F, m);
    const lh = m.lineHeight(12);
    const line = spec.series.findIndex((s) => s.as === 'line');
    const lineLabels = textsOf(scene, new RegExp(`^label:value:s${line}:`));
    assert.equal(lineLabels.length, spec.categories.length, `${title}: every point labelled`);
    for (const t of lineLabels) {
      const c = /c(\d+)$/.exec(roleOf(scene, t))![1];
      const mk = rectsOf(scene, new RegExp(`^marker:s${line}:c${c}$`))[0];
      const px = mk.x + mk.w / 2, py = mk.y + mk.h / 2;
      // Directly above or below the point, or touching it from the side; never out in the gap between columns.
      const vertical = close(t.x + t.w / 2, px, 0.5) && (t.y + t.h <= py || t.y >= py);
      const beside = Math.min(Math.abs(t.x - px), Math.abs(t.x + t.w - px)) <= mk.w;
      assert.ok(vertical || beside, `${title}: ${t.text} reads as its point's label`);
      // Close to its point; or, for a point inside its columns, straight above them (manual: line labels
      // go above / below the point, not onto the column segments, where they read as segment values).
      const bars = rectsOf(scene, new RegExp(`^bar:s\\d+:c${c}$`));
      const top = Math.min(...bars.map((b) => b.y));
      const overStack = vertical && py > top && t.y + t.h <= top && top - (t.y + t.h) <= lh * 0.8;
      assert.ok(overStack || Math.min(Math.abs(t.y + t.h - py), Math.abs(t.y - py)) <= lh * 1.5, `${title}: ${t.text} close to its point`);
      for (const b of rectsOf(scene, /^bar:/)) assert.ok(!overlaps(t, b, 0), `${title}: ${t.text} sits on a column`);
    }
    if (spec.options?.combo !== 'clustered') {
      for (const t of textsOf(scene, /^label:value:s[01]:/)) {
        const bar = rectsOf(scene, new RegExp(`^${roleOf(scene, t).replace('label:value', 'bar')}$`))[0];
        assert.ok(t.y >= bar.y + 1 && t.y + t.h <= bar.y + bar.h - 1, `${title}: ${t.text} inside its segment`);
      }
    }
    const labels = textsOf(scene, /^label:(value|total)/);
    for (let i = 0; i < labels.length; i++) {
      for (let j = i + 1; j < labels.length; j++) assert.ok(gapBetween(labels[i], labels[j]) >= 1, `${title}: "${labels[i].text}" touches "${labels[j].text}"`);
    }
    assert.deepEqual(strict(scene), [], title);
  }
});

test('combo: a line point inside its stack is labelled above the stack; on a segment only as a last resort', () => {
  const spec = (top: number): ChartSpec => ({
    type: 'combo', categories: ['A', 'B', 'C'],
    series: [{ name: 'P', values: [50, 50, 50] }, { name: 'Q', values: [50, 50, 50] },
      { name: 'In', values: [60, 60, 60], as: 'line' }, { name: 'Over', values: [top, top, top], as: 'line' }],
    options: { totals: true },
  });
  // Room above the stack: the "60" sits there, off the columns, under the total.
  const free = layoutChart(spec(106), F, m);
  assert.deepEqual(free.hidden, []);
  const bar = rectsOf(free, /^bar:s1:c1$/)[0];
  const t = textsOf(free, /^label:value:s2:c1$/)[0];
  assert.ok(t.y + t.h <= bar.y && !rectsOf(free, /^bar:/).some((b) => overlaps(t, b, 0)), 'above the stack');
  assert.ok(textsOf(free, /^label:total:c1$/)[0].y + lh() <= t.y, 'under the total');
  // A second line runs just above the stack, through every spot there: the "60" falls back onto its
  // segment, wholly inside it, in a contrasting colour.
  const tight = layoutChart(spec(103), F, m);
  assert.deepEqual(tight.hidden, []);
  const seg = rectsOf(tight, /^bar:s1:c1$/)[0];
  const u = textsOf(tight, /^label:value:s2:c1$/)[0];
  assert.ok(u.y >= seg.y && u.y + u.h <= seg.y + seg.h && u.x >= seg.x && u.x + u.w <= seg.x + seg.w, 'inside one segment');
  assert.equal(u.color, 'FFFFFF');
  assert.deepEqual(strict(tight), []);
});
const lh = () => m.lineHeight(12);

// ---- Mekko XY ----------------------------------------------------------------------------------------

test('mekkoxy: a segment too thin for its label gets it outside, beside the last column, with a leader', () => {
  const scene = layoutChart(slides('mekkoxy')[1].chart, F, m);
  assert.deepEqual(scene.hidden, [], 'no label hidden');
  const seg = rectsOf(scene, /^bar:s3:c4$/)[0]; // Services / Consumer, $120
  const label = textsOf(scene, /^label:value:s3:c4$/)[0];
  assert.equal(label.text, '$120');
  assert.ok(label.x > seg.x + seg.w, 'outside, to the right');
  const leader = lines(scene, /^leader:c4$/)[0];
  assert.ok(close(leader.x1, seg.x + seg.w) && close(leader.y1, seg.y + seg.h / 2), 'leader starts at the segment');
  assert.ok(leader.x2 <= label.x && label.x - leader.x2 < 2 && leader.y2 >= label.y && leader.y2 <= label.y + label.h, 'and ends at the label');
  const name = textsOf(scene, /^label:series:s3$/)[0];
  assert.ok(name.x >= label.x + label.w, 'series name further right');
  // The value axis is drawn over the first column, not hidden behind it.
  const axisIdx = scene.shapes.findIndex((sh, i) => scene.roles![i] === 'axis' && sh.kind === 'line' && sh.x1 === sh.x2);
  assert.ok(axisIdx > scene.roles!.findLastIndex((r) => r.startsWith('bar:')));
  assert.deepEqual(strict(scene), []);
});

// ---- Gantt -------------------------------------------------------------------------------------------


test('gantt: rows share the frame height up to a comfortable maximum; the chart keeps its top edge', () => {
  const lh = m.lineHeight(12);
  const rowSeps = (s: Scene) => lines(s, /^gridline:/).filter((l) => l.y1 === l.y2).map((l) => l.y1).sort((a, b) => a - b);
  // Seven activities: the rows fill the frame (minus header, bracket band and key).
  const long = layoutChart(slides('gantt')[0].chart, F, m);
  const ys = rowSeps(long);
  const rowH = ys[1] - ys[0];
  assert.ok(rowH > lh * 3, `rows use the height (${rowH.toFixed(1)} pt)`);
  const key = textsOf(long, /^legend:s\d+$/)[0];
  assert.ok(F.y + F.h - (key.y + key.h) < lh, 'nothing but the key below the band');
  // A short plan: rows stop growing at the maximum, the header stays at the top of the frame.
  const short = layoutChart({ type: 'gantt', categories: ['A', 'B'], series: [{ name: 'S', values: ['2026-01-05', '2026-02-02'] }, { name: 'E', values: ['2026-03-02', '2026-04-06'] }] }, F, m);
  const sep = rowSeps(short)[0];
  const head = lines(short, /^axis$/).find((l) => l.y1 === l.y2 && l.x1 === F.x)!;
  assert.ok(close(sep - head.y1, lh * 4.5, 0.01), 'row height capped');
  assert.ok(textsOf(short, /^axis:tick:/).every((t) => t.y < F.y + lh * 3), 'header at the top');
  assert.deepEqual(strict(short), []);
});

test('gantt: a week scale spans whole weeks even when a milestone label widens the range', () => {
  const scene = layoutChart(slides('gantt')[1].chart, F, m);
  const fine = lines(scene, /^gridline:/).filter((l) => l.dash).map((l) => l.x1).sort((a, b) => a - b);
  const step = fine[1] - fine[0];
  const bars = rectsOf(scene, /^bar:/);
  const right = Math.max(...lines(scene, /^axis$/).map((l) => l.x2));
  const left = Math.min(...bars.map((b) => b.x));
  assert.ok(close((right - left) / step, Math.round((right - left) / step), 1e-6), `${((right - left) / step).toFixed(3)} weeks`);
});

test('gantt: brackets label periods under the rows; a labelled shade gets one; the today line crosses every row', () => {
  const spec = slides('gantt')[0].chart;
  const scene = layoutChart(spec, F, m);
  const day = (s: string) => Date.parse(s) / 864e5;
  const bar0 = rectsOf(scene, /^bar:s0:c0$/)[0]; // Diagnostic 2026-01-05 … 2026-02-02
  const ppd = bar0.w / (day('2026-02-02') - day('2026-01-05'));
  const xOf = (s: string) => bar0.x + (day(s) - day('2026-01-05')) * ppd;
  const rowsBottom = Math.max(...rectsOf(scene, /^shade:0$/).map((r) => r.y + r.h));
  for (const [k, b] of spec.options!.brackets!.entries()) {
    const ls = lines(scene, new RegExp(`^bracket:${k}$`));
    assert.equal(ls.length, 3, 'two ends and a span');
    const span = ls.find((l) => l.y1 === l.y2)!;
    assert.ok(close(span.x1, xOf(b.from), 1e-6) && close(span.x2, xOf(b.to), 1e-6), b.label);
    assert.ok(ls.every((l) => Math.min(l.y1, l.y2) >= rowsBottom), 'under the rows');
    const label = textsOf(scene, new RegExp(`^bracket:${k}$`))[0];
    assert.equal(label.text, b.label);
    assert.ok(label.y >= span.y1, 'label under its bracket');
  }
  assert.equal(lines(scene, /^shade:0$/).length, 3, 'the shade label hangs from a bracket');
  const today = lines(scene, /^today$/)[0];
  assert.ok(close(today.x1, xOf('2026-04-13'), 1e-6) && today.x1 === today.x2);
  assert.ok(today.y1 <= bar0.y && today.y2 >= rowsBottom, 'through every row');
  assert.equal(today.color, 'C0504D');
  assert.equal(textsOf(scene, /^today$/)[0].text, 'Today');
  assert.deepEqual(strict(scene), []);
  const off = layoutChart({ ...spec, options: { ...spec.options, today: '2030-01-01' } }, F, m);
  assert.ok(off.hidden.some((h) => /today: outside/.test(h)));
});

// ---- Table -------------------------------------------------------------------------------------------

test('table: rows share the spare frame height (at most two lines more each), content centred, top edge kept', () => {
  const lh = m.lineHeight(12);
  const scene = layoutChart(slides('table')[0].chart, F, m);
  const rules = lines(scene, /^(axis|gridline:\d+)$/).map((l) => l.y1).sort((a, b) => a - b);
  const rowH = rules[1] - rules[0];
  assert.ok(rowH > lh * 2.5, `rows grew (${rowH.toFixed(1)} pt)`);
  assert.ok(rules.at(-1)! <= F.y + F.h + 0.01);
  const cell = textsOf(scene, /^label:value:s0:c0$/)[0];
  assert.ok(close(cell.y + cell.h / 2, (rules[0] + rules[1]) / 2, 0.6), 'centred in its row');
  assert.ok(textsOf(scene, /^label:category:/).every((t) => t.y < F.y + lh * 2), 'header at the top');
  // A long table that fills the frame does not grow at all.
  const long = layoutChart({ type: 'table', categories: ['A'], series: Array.from({ length: 30 }, (_, i) => ({ name: `R${i}`, values: [i] })) }, F, m);
  assert.ok(long.hidden.length === 0 || long.hidden.every((h) => /does not fit|shrunk/.test(h)));
  assert.deepEqual(checkScene(long, { measurer: m }), []);
});

// ---- Scatter and bubble -----------------------------------------------------------------------------

test('scatter and bubble: legend swatches are circles, like the markers they stand for', () => {
  for (const { title, chart } of slides('scatter')) {
    const sw = rectsOf(layoutChart(chart, F, m), /^legend:s\d+$/);
    assert.equal(sw.length, chart.series.length, title);
    assert.ok(sw.every((r) => r.geom === 'ellipse'), title);
  }
});

test('scatter: a crowded point gets its label further out with a leader line that crosses nothing', () => {
  const grid = [...Array(9).keys()].map((i) => [5 + (i % 3) * 0.15, 5 + Math.floor(i / 3) * 0.375]);
  const pts = [[0, 0], [10, 10], ...grid];
  const scene = layoutChart({ type: 'scatter', categories: pts.map((_, i) => `Pt ${i}`), series: [{ name: 'S', values: pts }] }, F, m);
  const leaders = byRole(scene, /^leader:c\d+$/).filter((x) => x.kind === 'line') as LineShape[];
  assert.ok(leaders.length > 0, 'some labels moved out with a leader');
  assert.equal(textsOf(scene, /^label:value:/).length + scene.hidden.length, pts.length, 'labelled, or hidden and reported');
  assert.ok(scene.hidden.length <= 2, scene.hidden.join('; '));
  const labels = textsOf(scene, /^label:value:/);
  for (const l of leaders) {
    const c = /c(\d+)$/.exec(roleOf(scene, l))![1];
    const mk = rectsOf(scene, new RegExp(`^marker:s0:c${c}$`))[0];
    const lb = textsOf(scene, new RegExp(`^label:value:s0:c${c}$`))[0];
    const r = mk.w / 2, mx = mk.x + r, my = mk.y + r;
    assert.ok(Math.abs(Math.hypot(l.x1 - mx, l.y1 - my) - (r + 1)) < 0.01, 'starts at its marker');
    assert.ok(l.x2 >= lb.x - 0.01 && l.x2 <= lb.x + lb.w + 0.01 && l.y2 >= lb.y - 0.01 && l.y2 <= lb.y + lb.h + 0.01, 'ends on its label');
    for (const t of labels) if (t !== lb) assert.ok(!segHits(l, t), `leader of ${lb.text} crosses ${t.text}`);
  }
  assert.deepEqual(checkScene(scene, { measurer: m }), []);
});

test('bubble: the size legend bubble is drawn to the chart\'s scale and labelled with its round size', () => {
  const spec = { ...slides('scatter')[1].chart };
  const scene = layoutChart(spec, F, m);
  const leg = rectsOf(scene, /^legend:size$/)[0];
  const big = rectsOf(scene, /^marker:s0:c0$/)[0]; // Payments, size 420
  assert.ok(close((leg.w / big.w) ** 2, 200 / 420, 1e-9));
  assert.equal(textsOf(scene, /^legend:size$/)[0].text, 'Revenue ($m): 200');
  assert.deepEqual(strict(scene), []);
  const off = layoutChart({ ...spec, options: { ...spec.options, sizeLegend: undefined } }, F, m);
  assert.equal(rectsOf(off, /^legend:size$/).length, 0, 'off unless asked for, as in the reference');
});
