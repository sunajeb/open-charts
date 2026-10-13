// Resize report: runs the resize matrix (test/resize-matrix.ts) and groups the problems per chart type.
//   node --experimental-strip-types --no-warnings scripts/resize-report.ts [name-regex] [-v]
import { layoutChart } from '../src/core/index.ts';
import { createMeasurer } from '../src/fonts.ts';
import { allCases, allFrames, chartMeasurer, resizeProblems } from '../test/resize-matrix.ts';

const m = createMeasurer();
const args = process.argv.slice(2);
const verbose = args.includes('-v');
const pattern = args.find((a) => !a.startsWith('-'));
const filter = pattern ? new RegExp(pattern) : null;
const D = { x: 48, y: 96, w: 864, h: 408 };
const byType: Record<string, Record<string, number>> = {};
const examples: Record<string, string[]> = {};
const sizes: Record<string, Set<string>> = {};
let total = 0, bad = 0;
const t0 = performance.now();
for (const c of allCases()) {
  if (filter && !filter.test(c.name)) continue;
  const ref = layoutChart(c.spec, D, m);
  for (const f of allFrames()) {
    total++;
    let probs: string[];
    try { probs = resizeProblems(layoutChart(c.spec, f, m), chartMeasurer(c.spec, m), ref); } catch (e) { probs = [`exception: ${(e as Error).message.split('\n')[0]}`]; }
    if (!probs.length) continue;
    bad++;
    const t = c.spec.type;
    (sizes[t] ??= new Set()).add(`${f.w}x${f.h}`);
    byType[t] ??= {};
    for (const p of probs) {
      const k = p.replace(/".*?"/g, '"…"').replace(/[\d.]+/g, '#').slice(0, 70);
      byType[t][k] = (byType[t][k] ?? 0) + 1;
      const ex = (examples[`${t}|${k}`] ??= []);
      if (ex.length < 3) ex.push(`${c.name} ${f.w}x${f.h}: ${p}`);
    }
  }
}
console.log(`${bad}/${total} frames with problems, ${(performance.now() - t0).toFixed(0)} ms`);
for (const [t, ks] of Object.entries(byType)) {
  console.log(`== ${t}`);
  if (args.includes('-s')) console.log(`  sizes: ${[...sizes[t]].join(' ')}`);
  for (const [k, n] of Object.entries(ks).sort((a, b) => b[1] - a[1])) {
    console.log(`  ${n}\t${k}`);
    if (verbose) for (const e of examples[`${t}|${k}`]) console.log(`      ${e.slice(0, 220)}`);
  }
}
