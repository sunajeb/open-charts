// `npm run bench:accept`: rewrite the expected scenes in test/bench/scenes/ from the current engine and
// list what changed. Review the change before committing: `git diff test/bench/scenes` for numbers,
// `npm run goldens` (bench decks render as bench-*.png) to see it.
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { createMeasurer } from '../../src/fonts.ts';
import { benchSlides, diffLines, layoutBench, sceneDir, sceneSnapshot } from './harness.ts';

const m = createMeasurer();
mkdirSync(sceneDir, { recursive: true });
const slides = benchSlides();
const keep = new Set(slides.map((b) => `${b.name}.scene.json`));
let changed = 0;
for (const b of slides) {
  const file = join(sceneDir, `${b.name}.scene.json`);
  const next = sceneSnapshot(layoutBench(b, m), m);
  const prev = existsSync(file) ? readFileSync(file, 'utf8') : null;
  if (prev === next) continue;
  changed++;
  writeFileSync(file, next);
  const d = prev ? diffLines(prev, next, 4) : [];
  console.log(`${prev ? 'updated' : 'created'} ${b.name}${d.length ? `\n  ${d.join('\n  ')}` : ''}`);
}
for (const f of readdirSync(sceneDir)) {
  if (!keep.has(f)) { rmSync(join(sceneDir, f)); changed++; console.log(`removed ${f}`); }
}
console.log(`\n${changed ? `${changed} scene snapshot(s) written` : 'All scene snapshots current'} (${slides.length} bench slides).`);
