// Layout benchmark: `npm run bench` (add --json for machine-readable output).
// - cold start: a fresh `node` running the CLI on a one-chart deck (node start + font parse + layout + .pptx)
// - font parse: createMeasurer() in this process
// - warm layout: p50 / p95 per sample chart at the default frame and at resized frames
// - deck build: buildDeck() of a 100-slide deck (layout + quality gate + SVG + .pptx)
import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { layoutChart, type ChartSpec, type Frame } from '../src/core/index.ts';
import { samples } from '../src/core/samples.ts';
import { createMeasurer } from '../src/fonts.ts';
import { buildDeck } from '../src/deck.ts';

const root = fileURLToPath(new URL('..', import.meta.url));
const json = process.argv.includes('--json');
const quick = process.argv.includes('--quick');

const pct = (xs: number[], p: number) => { const s = [...xs].sort((a, b) => a - b); return s[Math.min(s.length - 1, Math.floor(p * s.length))]; };
const time = (f: () => void) => { const t = performance.now(); f(); return performance.now() - t; };

// Cold start: median of a few fresh processes.
const tmp = mkdtempSync(join(tmpdir(), 'oc-bench-'));
const one = join(tmp, 'one.json');
writeFileSync(one, JSON.stringify({ slides: [{ chart: samples.stacked }] }));
const coldRuns: number[] = [];
for (let i = 0; i < (quick ? 2 : 5); i++) {
  coldRuns.push(time(() => execFileSync(process.execPath, ['--experimental-strip-types', '--no-warnings', join(root, 'src/cli.ts'), one, '-o', join(tmp, 'one')], { stdio: 'ignore' })));
}
rmSync(tmp, { recursive: true, force: true });

const fontMs = time(() => createMeasurer());
const m = createMeasurer();

const DEFAULT: Frame = { x: 48, y: 96, w: 864, h: 408 };
const RESIZED: Frame[] = [DEFAULT, { x: 0, y: 0, w: 360, h: 200 }, { x: 0, y: 0, w: 900, h: 120 }, { x: 0, y: 0, w: 1400, h: 800 }, { x: 0, y: 0, w: 240, h: 500 }];
const N = quick ? 30 : 200;
const rows: { type: string; firstMs: number; p50: number; p95: number; p95Resized: number }[] = [];
for (const [type, spec] of Object.entries(samples) as [string, ChartSpec][]) {
  const firstMs = time(() => layoutChart(spec, DEFAULT, m));
  for (let i = 0; i < 20; i++) layoutChart(spec, DEFAULT, m);
  const at = Array.from({ length: N }, () => time(() => layoutChart(spec, DEFAULT, m)));
  const resized = Array.from({ length: N }, (_, i) => time(() => layoutChart(spec, RESIZED[i % RESIZED.length], m)));
  rows.push({ type, firstMs, p50: pct(at, 0.5), p95: pct(at, 0.95), p95Resized: pct(resized, 0.95) });
}

const all = Object.values(samples) as ChartSpec[];
const deck = { slides: Array.from({ length: 100 }, (_, i) => ({ title: `Slide ${i + 1}`, chart: all[i % all.length] })) };
buildDeck({ slides: deck.slides.slice(0, 5) }, m); // warm
const deckRuns = Array.from({ length: quick ? 2 : 3 }, () => time(() => buildDeck(deck, m)));

const result = {
  coldStartMs: pct(coldRuns, 0.5),
  fontParseMs: fontMs,
  layout: rows,
  worstP95Ms: Math.max(...rows.map((r) => Math.max(r.p95, r.p95Resized))),
  deck100Ms: Math.min(...deckRuns),
};
if (json) {
  console.log(JSON.stringify(result, null, 2));
} else {
  const f = (v: number) => v.toFixed(2).padStart(8);
  console.log(`cold start (CLI, 1 chart)  ${f(result.coldStartMs)} ms (median of ${coldRuns.length})`);
  console.log(`font parse                 ${f(result.fontParseMs)} ms`);
  console.log(`100-slide deck build       ${f(result.deck100Ms)} ms`);
  console.log(`\nlayout (ms)      first      p50      p95  p95 resized`);
  for (const r of rows) console.log(`${r.type.padEnd(10)} ${f(r.firstMs)} ${f(r.p50)} ${f(r.p95)} ${f(r.p95Resized)}`);
  console.log(`\nworst p95 ${result.worstP95Ms.toFixed(2)} ms`);
}
