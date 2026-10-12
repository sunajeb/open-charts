// Bench harness shared by test/bench.test.ts, `npm run bench:accept` and `npm run quality`.
//
// test/bench/<NN-name>.json are reference slides (docs/plan/02-engine-and-quality.md §4.1): realistic
// consulting decks covering every chart type and the hard cases. Each has a committed expected scene,
// test/bench/scenes/<NN-name>.scene.json: the laid-out scene with coordinates rounded to 0.1 pt, one
// shape per line, so a layout change shows up as a readable diff. `npm run bench:accept` rewrites them
// (review the change in the goldens gallery: `npm run goldens` renders the bench decks too).
import { createHash } from 'node:crypto';
import { readdirSync, readFileSync } from 'node:fs';
import { basename, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { defaultStyle, layoutChart, migrate, type Scene, type Shape } from '../../src/core/index.ts';
import type { Measurer } from '../../src/core/text.ts';
import { CHART_FRAME, type DeckSpec } from '../../src/deck.ts';

export const benchDir = fileURLToPath(new URL('.', import.meta.url));
export const sceneDir = join(benchDir, 'scenes');

export interface BenchSlide { file: string; name: string; slide: number; title?: string; spec: DeckSpec['slides'][number]['chart']; style: typeof defaultStyle }

/** Every slide of every bench deck, in file order. */
export function benchSlides(dir = benchDir): BenchSlide[] {
  return readdirSync(dir).filter((f) => /^\d+-.*\.json$/.test(f)).sort().flatMap((f) => {
    const deck = JSON.parse(readFileSync(join(dir, f), 'utf8')) as DeckSpec;
    const style = { ...defaultStyle, ...deck.style };
    return deck.slides.map((s, i) => ({ file: f, name: basename(f, '.json') + (deck.slides.length > 1 ? `-${i + 1}` : ''), slide: i + 1, title: s.title, spec: s.chart, style }));
  });
}

/** Lay a bench slide out exactly as the CLI does (buildDeck): migrated spec, default chart frame. */
export const layoutBench = (b: BenchSlide, m: Measurer): Scene => layoutChart(migrate(b.spec), CHART_FRAME, m, b.style);

/** Characters whose widths (regular and bold) identify the font metrics. */
const SAMPLE = [...Array.from({ length: 95 }, (_, i) => String.fromCharCode(32 + i)), ...'−€£¥%‰–—…·äöüÄÖÜßéèêçñåøÉ中円한'];

/**
 * The measurer's fingerprint: snapshots are only comparable under the same font metrics (Arial or a
 * metric-compatible substitute). A snapshot made under other metrics is skipped, not failed.
 */
export const fingerprint = (m: Measurer) => {
  const widths = SAMPLE.flatMap((ch) => [m.width(ch, 12), m.width(ch, 12, true)].map((w) => w.toFixed(3))).join(',');
  return `sha1:${createHash('sha1').update(widths).digest('hex').slice(0, 12)} lh=${m.lineHeight(12).toFixed(2)}`;
};

const r = (v: number) => Math.round(v * 10) / 10 + 0; // + 0 turns -0 into 0

function roundShape(s: Shape): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(s)) {
    if (v === undefined) continue;
    if (typeof v === 'number') out[k] = r(v);
    else if (k === 'commands') out[k] = (v as Record<string, unknown>[]).map((c) => Object.fromEntries(Object.entries(c).map(([ck, cv]) => [ck, typeof cv === 'number' ? r(cv) : cv])));
    else out[k] = v;
  }
  return out;
}

/** Stable, reviewable JSON of a scene: rounded coordinates, role first, one shape per line. */
export function sceneSnapshot(scene: Scene, m: Measurer): string {
  // The scene role first; a text's coarse `role` field is kept as `textRole`.
  const shapes = scene.shapes.map((s, i) => {
    const { role: textRole, ...rest } = roundShape(s);
    return JSON.stringify({ role: scene.roles?.[i] ?? '', ...rest, ...(textRole ? { textRole } : {}) });
  });
  return [
    '{',
    `  "fonts": ${JSON.stringify(fingerprint(m))},`,
    `  "frame": ${JSON.stringify(roundShape(scene.frame as unknown as Shape))},`,
    `  "hidden": ${JSON.stringify(scene.hidden)},`,
    '  "shapes": [',
    shapes.map((s) => `    ${s}`).join(',\n'),
    '  ]',
    '}',
    '',
  ].join('\n');
}

/** Line-level differences between two snapshots, for a readable test failure. */
export function diffLines(expected: string, actual: string, max = 12): string[] {
  const a = expected.split('\n'), b = actual.split('\n');
  const out: string[] = [];
  const n = Math.max(a.length, b.length);
  for (let i = 0; i < n && out.length < max * 2; i++) {
    if (a[i] === b[i]) continue;
    if (a[i] !== undefined) out.push(`- ${a[i].trim()}`);
    if (b[i] !== undefined) out.push(`+ ${b[i].trim()}`);
  }
  return out;
}

/** Hidden labels a fixture declares acceptable (test/bench/manifest.json), as strict-gate patterns. */
export function allowedHidden(name: string, dir = benchDir): string[] {
  const manifest = JSON.parse(readFileSync(join(dir, 'manifest.json'), 'utf8')) as Record<string, { stresses: string; allowHidden?: string[] }>;
  return manifest[name]?.allowHidden ?? [];
}
