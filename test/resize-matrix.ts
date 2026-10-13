// Resize matrix shared by test/resize.test.ts and scripts/bench.ts: every chart type's sample and every
// example chart, laid out at a grid of frames, checked against the invariants a resized chart must keep.
import { readdirSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { checkScene, measurerFor, resolveFont, type ChartSpec, type Frame, type Scene, type Shape } from '../src/core/index.ts';
import { samples } from '../src/core/samples.ts';
import { directionsOf } from '../src/core/spec.ts';
import type { Measurer } from '../src/core/text.ts';

const root = fileURLToPath(new URL('..', import.meta.url));

export interface Case { name: string; spec: ChartSpec }

/** Every sample (one per chart type, and in each other direction it can grow in) and every chart of every example deck. */
export function allCases(): Case[] {
  const out: Case[] = Object.entries(samples).map(([type, spec]) => ({ name: `sample:${type}`, spec: spec! }));
  for (const [type, spec] of Object.entries(samples)) {
    for (const direction of directionsOf(type).filter((d) => d !== 'up')) {
      out.push({ name: `sample:${type}:${direction}`, spec: { ...spec!, options: { ...spec!.options, direction } } });
    }
  }
  for (const f of readdirSync(`${root}examples`).filter((f) => f.endsWith('.json')).sort()) {
    const deck = JSON.parse(readFileSync(`${root}examples/${f}`, 'utf8'));
    const slides: { chart?: ChartSpec; charts?: { chart: ChartSpec }[] }[] = Array.isArray(deck.slides) ? deck.slides : [{ chart: deck }];
    slides.forEach((s, i) => (s.charts?.map((c) => c.chart) ?? (s.chart ? [s.chart] : [])).forEach((chart, k) =>
      out.push({ name: `${f.replace('.json', '')}#${i + 1}${s.charts ? `.${k + 1}` : ''}:${chart.type}`, spec: chart })));
  }
  return out;
}

export const WIDTHS = [160, 240, 360, 520, 700, 864, 1100, 1400];
export const HEIGHTS = [90, 130, 200, 300, 408, 560, 800];
/** Grid plus extreme aspect ratios and the default slide frame. */
export function allFrames(): Frame[] {
  const sizes: [number, number][] = [];
  for (const w of WIDTHS) for (const h of HEIGHTS) sizes.push([w, h]);
  sizes.push([900, 120], [1400, 90], [200, 500], [160, 800], [864, 408]);
  return sizes.map(([w, h]) => ({ x: 48, y: 96, w, h }));
}

/** Theme minimum: no text is ever drawn smaller; labels that would need less are hidden and reported. */
export const MIN_TEXT = 8;

const numbers = (s: Shape): number[] => {
  if (s.kind === 'line') return [s.x1, s.y1, s.x2, s.y2, s.width];
  if (s.kind === 'path') return s.commands.flatMap((c) => (c.op === 'Z' ? [] : c.op === 'A' ? [c.cx, c.cy, c.r, c.start, c.sweep] : [c.x, c.y]));
  if (s.kind === 'text') return [s.x, s.y, s.w, s.h, s.size, ...(s.lineHeight != null ? [s.lineHeight] : [])];
  return [s.x, s.y, s.w, s.h];
};

/** Roles of the text shapes in a scene. */
const textRoles = (s: Scene) => new Set(s.shapes.flatMap((sh, i) => (sh.kind === 'text' ? [s.roles![i]] : [])));

/**
 * Problems of one resized scene. `reference` is the same chart at the default frame: a text role it
 * has that the resized scene lacks must be accounted for in `scene.hidden` (hidden AND reported).
 */
/** The measurer of the chart's own font (an example chart may set `style.font`). */
export const chartMeasurer = (spec: ChartSpec, m: Measurer): Measurer => measurerFor(m, resolveFont(spec.style?.font, spec.style));

export function resizeProblems(scene: Scene, m: Measurer, reference?: Scene): string[] {
  const out: string[] = [];
  scene.shapes.forEach((s, i) => {
    if (numbers(s).some((v) => !Number.isFinite(v))) out.push(`non-finite ${s.kind} ${scene.roles![i]}`);
    if (s.kind === 'text' && s.size < MIN_TEXT - 1e-9) out.push(`text "${s.text}" at ${s.size} pt`);
    if ((s.kind === 'rect' || s.kind === 'text') && (s.w < -1e-9 || s.h < -1e-9)) out.push(`negative size ${s.kind} ${scene.roles![i]}`);
  });
  for (const issue of checkScene(scene, { tol: 1, measurer: m })) out.push(`${issue.kind}: ${issue.detail}`);
  if (reference && !isPlaceholder(scene)) {
    // A label may move (a pie share inside its slice joins the name outside it, a name outside goes to
    // the legend and its share inside): lost means its role is gone and so is one of its words.
    // Tick labels follow the axis length and are not compared.
    const have = textRoles(scene);
    const words = (t: string) => t.split(/\s+/).filter(Boolean);
    const shown = new Set(scene.shapes.flatMap((s) => (s.kind === 'text' ? words(s.text) : [])));
    const lost = reference.shapes.flatMap((s, i) => {
      const r = reference.roles![i];
      return s.kind === 'text' && !have.has(r) && !r.startsWith('axis:tick:') && !words(s.text).every((w) => shown.has(w)) ? [r] : [];
    });
    if (lost.length && scene.hidden.length === 0) out.push(`silently dropped: ${lost.join(', ')}`);
  }
  return out;
}

export const isPlaceholder = (s: Scene) => s.roles?.[0] === 'placeholder';
