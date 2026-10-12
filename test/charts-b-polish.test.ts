// Label placement and layout polish for combo, Mekko XY, scatter / bubble, table and Gantt: each test
// pins a defect seen in the rendered goldens.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import {
  checkScene, layoutChart,
  type ChartSpec, type RectShape, type Scene, type TextShape,
} from '../src/core/index.ts';
import { createMeasurer } from '../src/fonts.ts';
import { segmentHitsBox } from '../src/core/place.ts';
import { CHART_FRAME } from '../src/deck.ts';

const m = createMeasurer();
const F = CHART_FRAME;
const byRole = (s: Scene, re: RegExp) => s.shapes.filter((_, i) => re.test(s.roles![i]));
const rectsOf = (s: Scene, re: RegExp) => byRole(s, re).filter((x): x is RectShape => x.kind === 'rect');
const textsOf = (s: Scene, re: RegExp) => byRole(s, re).filter((x): x is TextShape => x.kind === 'text');
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
    const lh = m.lineHeight(10);
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
      assert.ok(Math.min(Math.abs(t.y + t.h - py), Math.abs(t.y - py)) <= lh * 1.5, `${title}: ${t.text} close to its point`);
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
  const leaders = byRole(scene, /^leader:c\d+$/).filter((x) => x.kind === 'line') as import('../src/core/index.ts').LineShape[];
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
