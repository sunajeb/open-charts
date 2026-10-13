// Resize matrix: the add-in re-lays a chart out whenever the user resizes it, so every chart type's
// sample and every example chart must come out clean at any size from 160×90 to 1400×800 pt, or —
// below the type's documented minimum — as a "too small" placeholder.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { checkScene, defaultStyle, layoutChart, minimumSize, type Scene, type TextShape } from '../src/core/index.ts';
import { samples } from '../src/core/samples.ts';
import { createMeasurer } from '../src/fonts.ts';
import { buildReport } from '../src/report.ts';
import { allCases, allFrames, chartMeasurer, isPlaceholder, MIN_TEXT, resizeProblems } from './resize-matrix.ts';

const m = createMeasurer();
const DEFAULT = { x: 48, y: 96, w: 864, h: 408 };
const frames = allFrames();

for (const c of allCases()) {
  test(`resize: ${c.name}`, () => {
    const reference = layoutChart(c.spec, DEFAULT, m);
    assert.ok(!isPlaceholder(reference), 'the default frame is never too small');
    // The last-resort resolver judges like the gate, so at the default frame (where the gate passes)
    // it never has to step in.
    assert.deepEqual(reference.hidden.filter((h) => /\) hidden: /.test(h)), []);
    const min = minimumSize(c.spec);
    const failures: string[] = [];
    for (const f of frames) {
      const at = `${f.w}×${f.h}`;
      let scene: Scene;
      try {
        scene = layoutChart(c.spec, f, m);
      } catch (e) {
        failures.push(`${at}: threw ${(e as Error).message}`);
        continue;
      }
      for (const p of resizeProblems(scene, chartMeasurer(c.spec, m), reference)) failures.push(`${at}: ${p}`);
      const small = f.w < min.w || f.h < min.h;
      if (small && !isPlaceholder(scene)) failures.push(`${at}: below the minimum ${min.w}×${min.h} but no placeholder`);
      if (isPlaceholder(scene)) {
        // The size the placeholder asks for is enough: there the chart lays out.
        const need = scene.hidden[0]?.match(/^chart too small: \w+ needs at least (\d+) × (\d+) pt/);
        if (!need) failures.push(`${at}: placeholder not reported`);
        else if (isPlaceholder(layoutChart(c.spec, { ...f, w: Math.max(f.w, +need[1]), h: Math.max(f.h, +need[2]) }, m))) failures.push(`${at}: still too small at the size it asks for (${need[1]}×${need[2]})`);
      }
    }
    assert.deepEqual(failures, []);
  });
}

test('resize: layout is deterministic at every size', () => {
  for (const c of allCases()) {
    for (const f of [frames[0], frames[9], frames[30], frames[frames.length - 3], DEFAULT]) {
      assert.equal(JSON.stringify(layoutChart(c.spec, f, m)), JSON.stringify(layoutChart(c.spec, f, m)), `${c.name} ${f.w}×${f.h}`);
    }
  }
});

test('resize: below the minimum the chart is an outlined placeholder with a readable message', () => {
  const spec = samples.waterfall!;
  const scene = layoutChart(spec, { x: 10, y: 20, w: 150, h: 80 }, m);
  assert.ok(isPlaceholder(scene));
  assert.ok(scene.roles!.every((r) => r === 'placeholder'));
  const [outline, msg] = scene.shapes;
  assert.deepEqual([outline.kind, msg.kind], ['rect', 'text']);
  assert.match((msg as TextShape).text.replace(/\n/g, ' '), /too small/i);
  assert.ok((msg as TextShape).size >= MIN_TEXT);
  assert.match(scene.hidden[0], /^chart too small: waterfall needs at least \d+ × 120 pt, frame is 150 × 80 pt/);
  // Too small for even the shortest message: the outline alone, still reported.
  const tiny = layoutChart(spec, { x: 0, y: 0, w: 20, h: 8 }, m);
  assert.deepEqual(tiny.shapes.map((s) => s.kind), ['rect']);
  assert.equal(tiny.hidden.length, 1);
});

test('resize: a table whose columns cannot fit at the minimum font size becomes a placeholder', () => {
  const spec = samples.table!;
  const scene = layoutChart(spec, { x: 0, y: 0, w: 130, h: 300 }, m);
  assert.ok(isPlaceholder(scene));
  assert.match(scene.hidden[0], /columns need \d+ pt at 8 pt/);
});

test('resize: the strict gate reports every hidden label of a resized chart, and nothing else', () => {
  for (const type of ['stacked', 'waterfall', 'mekko', 'pie', 'gantt', 'combo'] as const) {
    for (const f of [{ x: 0, y: 0, w: 240, h: 130 }, { x: 0, y: 0, w: 900, h: 120 }, { x: 0, y: 0, w: 100, h: 60 }]) {
      const scene = layoutChart(samples[type]!, f, m);
      const strict = checkScene(scene, { strict: true, measurer: m });
      assert.ok(strict.every((i) => i.kind === 'hidden'), `${type} ${f.w}×${f.h}: ${JSON.stringify(strict)}`);
      assert.equal(strict.length, scene.hidden.length);
      // What `--strict` and `--report` see (buildDeck lists hidden labels as issues).
      const report = buildReport([{ issues: scene.hidden.map((h) => `hidden: ${h}`) }]);
      assert.equal(report.ok, scene.hidden.length === 0);
    }
  }
});

test('resize: every text keeps the theme minimum size, also for a large theme font', () => {
  const style = { fontSize: 18 };
  for (const c of allCases().filter((_, i) => i % 3 === 0)) {
    for (const f of [frames[10], frames[25], frames[45]]) {
      const scene = layoutChart(c.spec, f, m, { ...defaultStyle, ...style });
      for (const s of scene.shapes) if (s.kind === 'text') assert.ok(s.size >= MIN_TEXT, `${c.name} ${f.w}×${f.h}: "${s.text}" at ${s.size}`);
      assert.deepEqual(checkScene(scene, { measurer: chartMeasurer(c.spec, m) }), [], `${c.name} ${f.w}×${f.h}`);
    }
  }
});
