// The add-in's native shapes look like the CLI writer's: the same charts (same spec, same frame,
// laid out in the add-in deck's theme) written by the CLI's .pptx writer and rendered by the same
// PowerPoint, pixel-compared slide by slide. Uses this run's insert-all deck when there is one,
// otherwise inserts every type itself.
import { existsSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { PPAM_DIR } from '../../../src/addin/cli.ts';
import { DEFAULT_THEME, defaultStyle, layoutChart } from '../../../src/core/index.ts';
import { chartContent } from '../../../src/deck.ts';
import { createMeasurer } from '../../../src/fonts.ts';
import { readThemeColors } from '../../../src/ooxml/pptx-read.ts';
import { writePptx } from '../../../src/render/pptx.ts';
import { decodePng, diffImages, encodePng, type Image } from '../../../test/goldens/png.ts';
import type { Ctx } from '../scenarios.ts';
import { insertAll, type InsertAll } from './insert-all.ts';

export async function run(ctx: Ctx) {
  const { pp } = ctx;
  let ins = ctx.shared.get('insert-all') as InsertAll | undefined;
  if (ins && existsSync(ins.deck)) ctx.note(`add-in deck from this run's insert-all: ${ins.deck}`);
  else {
    ins = insertAll(ctx);
    pp.closeOwned();
  }
  // Charts use theme colours, which PowerPoint draws in the deck's theme (the new presentation's
  // Office theme): the CLI copy is laid out with, and written in, that same theme.
  const deckTheme = { ...DEFAULT_THEME, ...readThemeColors(new Uint8Array(readFileSync(ins.deck))) };
  const style = defaultStyle;
  const cliSlides = ins.slides.map((sl) => {
    const c = sl.charts[0];
    // A slide without a chart stays empty so slide numbers line up.
    if (!c?.spec || !c.frame) return { scenes: [] };
    const { style: overrides, ...chart } = JSON.parse(c.spec);
    const scene = layoutChart(chart, c.frame, createMeasurer(), { ...style, ...overrides, theme: { ...overrides?.theme, ...deckTheme } });
    return { scenes: [chartContent(chart, scene, style.font, overrides)] };
  });
  const compareDeck = join(PPAM_DIR, 'live-compare.pptx'); // a folder PowerPoint's sandbox may read
  const cliPng = join(ctx.out, 'cli-png');
  writeFileSync(compareDeck, writePptx(cliSlides, style.font, deckTheme));
  try {
    pp.open(compareDeck);
    pp.exportCopy(cliPng, 'png');
  } finally {
    rmSync(compareDeck, { force: true });
  }
  for (const sl of ins.slides) {
    if (!sl.charts[0]) continue;
    const a = join(ins.pngDir, `slide${sl.index}.png`), b = join(cliPng, `slide${sl.index}.png`);
    if (!ctx.check(existsSync(a) && existsSync(b), `slide ${sl.index}: both PNGs exported`)) continue;
    const ia = decodePng(readFileSync(a)), ib = decodePng(readFileSync(b));
    const d = diffImages(ia, ib);
    const off = beyondShift(ia, ib);
    if (d.diff && !d.ok) writeFileSync(join(ctx.out, `diff-${sl.index}.png`), encodePng(d.diff));
    ctx.check(off <= ia.width * ia.height * 2.5e-5, `slide ${sl.index} (${JSON.parse(sl.charts[0].spec ?? '{}').type}): add-in vs CLI writer ${d.ok ? 'identical' : `${(d.fraction * 100).toFixed(3)}% pixels differ, ${off} beyond a 3-px shift`}`);
  }
}

/**
 * Pixels of either image with no match (all channels within 48) in the other within ±R px, and no
 * blend of its neighbours: glyph placement that differs by a pixel or two, or an anti-aliasing step,
 * is not a difference.
 */
export function beyondShift(a: Image, b: Image, R = 3): number {
  if (a.width !== b.width || a.height !== b.height) return a.width * a.height;
  const near = (p: Image, q: Image, x: number, y: number) => {
    const i = (y * p.width + x) * 4;
    for (let dy = -R; dy <= R; dy++) for (let dx = -R; dx <= R; dx++) {
      const xx = x + dx, yy = y + dy;
      if (xx < 0 || yy < 0 || xx >= q.width || yy >= q.height) continue;
      const j = (yy * q.width + xx) * 4;
      if (Math.abs(p.data[i] - q.data[j]) <= 48 && Math.abs(p.data[i + 1] - q.data[j + 1]) <= 48 && Math.abs(p.data[i + 2] - q.data[j + 2]) <= 48) return true;
    }
    // An anti-aliasing step on a high-contrast edge (white separators on saturated theme colours): the
    // pixel is a blend that lies between the colours of the other image's 3x3 neighbourhood.
    const lo = [255, 255, 255], hi = [0, 0, 0];
    for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
      const xx = x + dx, yy = y + dy;
      if (xx < 0 || yy < 0 || xx >= q.width || yy >= q.height) continue;
      const j = (yy * q.width + xx) * 4;
      for (let c = 0; c < 3; c++) { lo[c] = Math.min(lo[c], q.data[j + c]); hi[c] = Math.max(hi[c], q.data[j + c]); }
    }
    return [0, 1, 2].every((c) => p.data[i + c] >= lo[c] - 8 && p.data[i + c] <= hi[c] + 8);
  };
  let n = 0;
  for (let y = 0; y < a.height; y++) for (let x = 0; x < a.width; x++) if (!near(a, b, x, y) || !near(b, a, x, y)) n++;
  return n;
}
