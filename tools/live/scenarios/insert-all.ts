// Every gallery item on its own slide through the add-in's own Dispatch (what the ribbon runs), one
// chart in an empty content placeholder, a type the engine does not support yet; then the saved deck
// is read back (tags, OC_SPEC, OC_ROLE on every shape) and every slide exported as PNG.
import { existsSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { galleryItems, insertableItems } from '../../../src/addin/registry.ts';
import { formatInspect } from '../../../src/inspect.ts';
import { readCharts, type SlideInfo } from '../../../src/ooxml/pptx-read.ts';
import type { Ctx } from '../scenarios.ts';

export interface InsertAll { deck: string; pngDir: string; slides: SlideInfo[] }

/** Inserts every available type (one per slide) into a fresh presentation and exports it. */
export function insertAll(ctx: Ctx, opts: { placeholder?: boolean } = {}) {
  const { pp } = ctx;
  pp.newPresentation();
  const items = insertableItems();
  const slideOf: Record<string, number> = {};
  for (const id of items) {
    slideOf[id] = pp.newSlide();
    const r = pp.run(`chart_${id}`);
    ctx.time('insert', r.ms);
    ctx.check(/ OK \(\d+ ms\)/.test(r.result), `insert ${id}: ${r.result || r.log.split('\n').pop() || 'no RESULT line'}`);
  }
  let placeholderSlide = 0, placeholderBox: number[] = [];
  if (opts.placeholder) {
    // On a title-and-content slide the chart takes the empty content placeholder's place.
    placeholderSlide = pp.newSlide('text slide');
    placeholderBox = pp.tell(`get {left position, top, width, height} of shape 2 of slide ${placeholderSlide} of active presentation`).split(', ').map(Number);
    const r = pp.run('chart_waterfall');
    ctx.check(/ OK \(/.test(r.result), `insert into a content placeholder: ${r.result || 'no RESULT line'}`);
  }
  const deck = join(ctx.out, 'addin.pptx'), pngDir = join(ctx.out, 'png');
  pp.exportCopy(deck, 'pptx');
  pp.exportCopy(pngDir, 'png');
  if (!existsSync(deck)) throw new Error(`the deck was not exported to ${deck}`);
  const slides = readCharts(new Uint8Array(readFileSync(deck)));
  const result: InsertAll = { deck, pngDir, slides };
  return { ...result, items, slideOf, placeholderSlide, placeholderBox };
}

export async function run(ctx: Ctx) {
  const { pp } = ctx;
  const r = insertAll(ctx, { placeholder: true });

  // A type the engine does not support yet answers with a message, not a crash.
  const unsupported = galleryItems().find((g) => !r.items.includes(g.id));
  if (unsupported) {
    pp.newSlide();
    const u = pp.run(`chart_${unsupported.id}`);
    ctx.check(/not available/.test(u.log), `${unsupported.id} (not supported yet) answers "not available" (${u.result})`);
  }

  writeFileSync(join(ctx.out, 'inspect.txt'), formatInspect(r.deck, r.slides));
  for (const id of r.items) {
    const c = r.slides.find((x) => x.index === r.slideOf[id])?.charts[0];
    if (!ctx.check(!!c, `${id}: a chart on slide ${r.slideOf[id]} of the saved deck`) || !c) continue;
    ctx.check(!!c.spec && c.contract === '1' && !!c.frame && !!c.font, `${id}: group tags complete (OC_SPEC, OC_V 1, OC_FRAME, OC_FONT)`);
    ctx.check(!!c.children && c.roles.length === c.children, `${id}: ${c.roles.length} of ${c.children} shapes carry OC_ROLE`);
  }
  const ph = r.slides.find((x) => x.index === r.placeholderSlide)?.charts[0];
  ctx.check(!!ph?.frame && [ph.frame.x, ph.frame.y, ph.frame.w, ph.frame.h].every((v, i) => Math.abs(v - r.placeholderBox[i]) <= 1),
    `content placeholder: chart frame ${JSON.stringify(ph?.frame)} is the placeholder ${r.placeholderBox.join(',')}`);
  const pngs = existsSync(r.pngDir) ? readdirSync(r.pngDir).filter((f) => f.toLowerCase().endsWith('.png')).length : 0;
  ctx.check(pngs === r.slides.length, `${pngs} PNGs exported for ${r.slides.length} slides (${r.pngDir})`);
  ctx.shared.set('insert-all', { deck: r.deck, pngDir: r.pngDir, slides: r.slides } satisfies InsertAll);
}
