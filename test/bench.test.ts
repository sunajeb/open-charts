// Bench fixtures (docs/plan/02-engine-and-quality.md §4): every reference slide must lay out exactly as its
// committed scene (test/bench/scenes/*.scene.json), pass the strict quality gate (hidden labels only where
// the fixture declares them in test/bench/manifest.json) and be deterministic. To accept a layout change:
// `npm run bench:accept`, then review `git diff test/bench/scenes` and the bench-*.png goldens.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { CHART_TYPES, checkScene } from '../src/core/index.ts';
import { createMeasurer } from '../src/fonts.ts';
import { allowedHidden, benchDir, benchSlides, diffLines, fingerprint, layoutBench, sceneDir, sceneSnapshot } from './bench/harness.ts';

const m = createMeasurer();
const slides = benchSlides();
const manifest = JSON.parse(readFileSync(join(benchDir, 'manifest.json'), 'utf8')) as Record<string, unknown>;

test('bench: at least 22 reference slides, every chart type covered, every slide described in the manifest', () => {
  assert.ok(slides.length >= 22, `${slides.length} bench slides`);
  const types = new Set(slides.map((b) => b.spec.type));
  assert.deepEqual(CHART_TYPES.filter((t) => !types.has(t)), []);
  assert.deepEqual(slides.map((b) => b.name).filter((n) => !manifest[n]), []);
  assert.deepEqual(Object.keys(manifest).filter((n) => !slides.some((b) => b.name === n)), []);
});

for (const b of slides) {
  test(`bench ${b.name}: matches its expected scene, passes the strict gate, deterministic`, (t) => {
    const scene = layoutBench(b, m);
    const snap = sceneSnapshot(scene, m);
    assert.equal(sceneSnapshot(layoutBench(b, m), m), snap, 'layout is not deterministic');

    const issues = checkScene(scene, { measurer: m, strict: true, allowHidden: allowedHidden(b.name) });
    assert.deepEqual(issues.map((i) => `${i.kind}: ${i.detail}`), []);

    const file = join(sceneDir, `${b.name}.scene.json`);
    assert.ok(existsSync(file), `no expected scene for ${b.name}: run \`npm run bench:accept\``);
    const expected = readFileSync(file, 'utf8');
    const fonts = (JSON.parse(expected) as { fonts: string }).fonts;
    if (fonts !== fingerprint(m)) {
      t.skip(`expected scene made under other font metrics (${fonts}; here ${fingerprint(m)})`);
      return;
    }
    if (expected !== snap) {
      assert.fail(`${b.name}: layout differs from test/bench/scenes/${b.name}.scene.json — if intended, run \`npm run bench:accept\` and review\n${diffLines(expected, snap).join('\n')}`);
    }
  });
}
