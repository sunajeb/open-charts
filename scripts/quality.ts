// `npm run quality`: the locally measurable engine KPIs of docs/MASTER-PLAN.md §9 over every bench slide
// (test/bench/*.json) and every example slide (examples/*.json), laid out as the CLI does.
//
//   npm run quality              table + verdict; writes out/quality.json
//   npm run quality -- --json    the JSON report only (stdout)
//
// Hard gate (exit code 1 when broken): 0 overlaps, 0 out-of-frame, 0 glyph overflows, 0 undeclared hidden
// labels on bench slides (test/bench/manifest.json declares the accepted ones), 100% columns and pies read
// exactly 100, stacked segments add up to the total shown, CAGR labels match the shown endpoints
// (±0.05 pp), waterfall columns balance (connectors meet the bars; totals = previous total + steps
// shown), Mekko widths sum to the plot width (±0.5 pt) and follow the totals, bench slides clean at four
// resized frames, no text below 8 pt, one size per label role, deterministic layout, layout p95 < 20 ms
// per chart.
import { mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { basename, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { checkScene, defaultStyle, layoutChart, migrate, type ChartSpec } from '../src/core/index.ts';
import { createMeasurer } from '../src/fonts.ts';
import { CHART_FRAME, type DeckSpec } from '../src/deck.ts';
import { allowedHidden } from '../test/bench/harness.ts';
import { cagrChecks, hundredChecks, mekkoChecks, stackChecks, typeChecks, waterfallChecks, type Check } from './kpis.ts';

const root = fileURLToPath(new URL('..', import.meta.url));
const jsonOnly = process.argv.includes('--json');
const RUNS = process.argv.includes('--quick') ? 10 : 40;
const m = createMeasurer();

interface Slide { source: string; name: string; spec: ChartSpec; style: typeof defaultStyle; allow: string[]; bench: boolean }

function decks(dir: string, bench: boolean): Slide[] {
  return readdirSync(join(root, dir)).filter((f) => (bench ? /^\d+-.*\.json$/ : /\.json$/).test(f)).sort().flatMap((f) => {
    const deck = JSON.parse(readFileSync(join(root, dir, f), 'utf8')) as DeckSpec;
    const style = { ...defaultStyle, ...deck.style };
    const name = basename(f, '.json');
    return deck.slides.map((s, i) => ({
      source: `${dir}/${f}#${i + 1}`, name: deck.slides.length > 1 ? `${name}-${i + 1}` : name,
      spec: migrate(s.chart), style, bench, allow: bench ? allowedHidden(name) : [],
    }));
  });
}

// ---- Run --------------------------------------------------------------------------------------

const RESIZED = [{ x: 0, y: 0, w: 420, h: 260 }, { x: 0, y: 0, w: 300, h: 400 }, { x: 0, y: 0, w: 1200, h: 300 }, { x: 0, y: 0, w: 640, h: 360 }];
const pct = (xs: number[], p: number) => { const s = [...xs].sort((a, b) => a - b); return s[Math.min(s.length - 1, Math.floor(p * s.length))]; };
const slides = [...decks('test/bench', true), ...decks('examples', false)];
const rows = slides.map((sl) => {
  const lay = () => layoutChart(sl.spec, CHART_FRAME, m, sl.style);
  const scene = lay();
  const deterministic = JSON.stringify(lay()) === JSON.stringify(scene);
  for (let i = 0; i < 5; i++) lay();
  const times = Array.from({ length: RUNS }, () => { const t = performance.now(); lay(); return performance.now() - t; });
  const issues = checkScene(scene, { measurer: m });
  // Resized (the add-in re-lays a chart out at any size): bench slides at four other frames stay clean.
  const resized = sl.bench ? RESIZED.flatMap((f) => checkScene(layoutChart(sl.spec, f, m, sl.style), { measurer: m }).map((i) => `${f.w}×${f.h} ${i.kind}: ${i.detail}`)) : [];
  const strict = checkScene(scene, { measurer: m, strict: true, allowHidden: sl.allow }).filter((i) => i.kind === 'hidden');
  const count = (k: string) => issues.filter((i) => i.kind === k).length;
  const checks: Record<string, Check[]> = {
    hundred: hundredChecks(sl.spec, scene), stack: stackChecks(sl.spec, scene), cagr: cagrChecks(sl.spec, scene),
    waterfall: waterfallChecks(sl.spec, scene), mekko: mekkoChecks(sl.spec, scene), type: typeChecks(scene),
  };
  return {
    slide: sl.name, source: sl.source, type: sl.spec.type, bench: sl.bench,
    overlaps: count('overlap') + count('text-mark') + count('arrow-label'), outOfFrame: count('out-of-frame'), overflow: count('overflow'),
    hidden: scene.hidden.length, hiddenUndeclared: strict.length, hiddenDetail: strict.map((i) => i.detail),
    deterministic, p50Ms: pct(times, 0.5), p95Ms: pct(times, 0.95),
    checks: Object.fromEntries(Object.entries(checks).map(([k, cs]) => [k, { n: cs.length, failed: cs.filter((c) => !c.ok).map((c) => `${c.name}: ${c.detail ?? ''}`) }])),
    issues: [...issues.map((i) => `${i.kind}: ${i.detail}`), ...resized],
    resizedIssues: resized.length,
  };
});

const sum = (f: (r: typeof rows[number]) => number) => rows.reduce((a, r) => a + f(r), 0);
const checkKpi = (k: string) => ({ checked: sum((r) => r.checks[k].n), failed: sum((r) => r.checks[k].failed.length) });
const allP95 = rows.map((r) => r.p95Ms);
const kpis = {
  slides: rows.length,
  benchSlides: rows.filter((r) => r.bench).length,
  chartTypes: new Set(rows.map((r) => r.type)).size,
  overlaps: sum((r) => r.overlaps),
  resizedFrames: rows.filter((r) => r.bench).length * RESIZED.length,
  resizedIssues: sum((r) => r.resizedIssues),
  outOfFrame: sum((r) => r.outOfFrame),
  glyphOverflow: sum((r) => r.overflow),
  hiddenLabels: sum((r) => r.hidden),
  hiddenUndeclaredBench: sum((r) => (r.bench ? r.hiddenUndeclared : 0)),
  hiddenExamples: sum((r) => (r.bench ? 0 : r.hidden)),
  hundredPercentSums: checkKpi('hundred'),
  stackedSegmentsEqualTotal: checkKpi('stack'),
  cagrRecomputed: checkKpi('cagr'),
  waterfallBalance: checkKpi('waterfall'),
  mekkoWidths: checkKpi('mekko'),
  typography: checkKpi('type'),
  deterministic: rows.every((r) => r.deterministic),
  layoutP95Ms: +pct(allP95, 0.95).toFixed(3),
  layoutWorstP95Ms: +Math.max(...allP95).toFixed(3),
};
const gate: [string, boolean][] = [
  ['0 overlaps', kpis.overlaps === 0],
  ['0 out-of-frame', kpis.outOfFrame === 0],
  ['bench slides clean at resized frames', kpis.resizedIssues === 0],
  ['0 glyph overflows', kpis.glyphOverflow === 0],
  ['0 undeclared hidden labels on bench slides', kpis.hiddenUndeclaredBench === 0],
  ['100% columns / pies read exactly 100', kpis.hundredPercentSums.failed === 0],
  ['stacked segments = total shown', kpis.stackedSegmentsEqualTotal.failed === 0],
  ['CAGR label = recomputed ± 0.05 pp', kpis.cagrRecomputed.failed === 0],
  ['waterfall columns balance', kpis.waterfallBalance.failed === 0],
  ['Mekko widths = plot width ± 0.5 pt', kpis.mekkoWidths.failed === 0],
  ['no text < 8 pt, one size per role', kpis.typography.failed === 0],
  ['deterministic', kpis.deterministic],
  ['layout p95 < 20 ms per chart', kpis.layoutWorstP95Ms < 20],
];
const report = { ok: gate.every(([, ok]) => ok), kpis, gate: Object.fromEntries(gate), slides: rows };

mkdirSync(join(root, 'out'), { recursive: true });
writeFileSync(join(root, 'out/quality.json'), `${JSON.stringify(report, null, 2)}\n`);
if (jsonOnly) {
  console.log(JSON.stringify(report, null, 2));
} else {
  const c = (k: { checked: number; failed: number }) => `${k.checked - k.failed}/${k.checked}`;
  console.log(`Quality report: ${kpis.slides} slides (${kpis.benchSlides} bench), ${kpis.chartTypes} chart types\n`);
  console.log(`  overlaps                       ${kpis.overlaps}`);
  console.log(`  out-of-frame                   ${kpis.outOfFrame}`);
  console.log(`  resized-frame issues           ${kpis.resizedIssues} (${kpis.resizedFrames} layouts)`);
  console.log(`  glyph overflows                ${kpis.glyphOverflow}`);
  console.log(`  hidden labels                  ${kpis.hiddenLabels} (bench undeclared ${kpis.hiddenUndeclaredBench}; examples ${kpis.hiddenExamples})`);
  console.log(`  100% sums exact                ${c(kpis.hundredPercentSums)}`);
  console.log(`  stacked segments = total       ${c(kpis.stackedSegmentsEqualTotal)}`);
  console.log(`  CAGR = recomputed ±0.05 pp     ${c(kpis.cagrRecomputed)}`);
  console.log(`  waterfall balance              ${c(kpis.waterfallBalance)}`);
  console.log(`  Mekko widths                   ${c(kpis.mekkoWidths)}`);
  console.log(`  typography                     ${c(kpis.typography)}`);
  console.log(`  deterministic                  ${kpis.deterministic ? 'yes' : 'NO'}`);
  console.log(`  layout p95 / worst p95         ${kpis.layoutP95Ms} / ${kpis.layoutWorstP95Ms} ms`);
  const bad = rows.filter((r) => r.resizedIssues || r.overlaps || r.outOfFrame || r.overflow || (r.bench && r.hiddenUndeclared) || Object.values(r.checks).some((k) => k.failed.length));
  for (const r of bad) {
    console.log(`\n  ${r.slide} (${r.source})`);
    for (const i of r.issues.filter((x) => !x.startsWith('hidden'))) console.log(`    ${i}`);
    if (r.bench) for (const h of r.hiddenDetail) console.log(`    hidden (undeclared): ${h}`);
    for (const [k, v] of Object.entries(r.checks)) for (const f of v.failed) console.log(`    ${k}: ${f}`);
  }
  console.log(`\n${gate.map(([n, ok]) => `  ${ok ? 'pass' : 'FAIL'}  ${n}`).join('\n')}`);
  console.log(`\n${report.ok ? 'Quality gate passed' : 'Quality gate FAILED'} — JSON report in out/quality.json`);
}
process.exit(report.ok ? 0 : 1);
