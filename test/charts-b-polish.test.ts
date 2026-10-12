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
