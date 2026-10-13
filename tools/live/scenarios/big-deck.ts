// A big deck: 100 slides with 50 charts (every sample type in turn, a chart on every other slide),
// built by the CLI deck builder, opens in PowerPoint without a repair prompt; Update chart on the
// first, middle and last chart stays quick; exported and read back, every chart keeps its tags.
import { rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { PPAM_DIR } from '../../../src/addin/cli.ts';
import type { ChartSpec } from '../../../src/core/index.ts';
import { samples } from '../../../src/core/samples.ts';
import { buildDeck } from '../../../src/deck.ts';
import { createMeasurer } from '../../../src/fonts.ts';
import type { Ctx } from '../scenarios.ts';

export async function run(ctx: Ctx) {
  const { pp, check } = ctx;
  const all = Object.values(samples) as ChartSpec[];
  const slides = Array.from({ length: 100 }, (_, i) => (i % 2 ? { title: `Slide ${i + 1}`, chart: all[(i >> 1) % all.length] } : { title: `Section ${i / 2 + 1}` }));
  const deck = join(PPAM_DIR, 'live-big-deck.pptx'); // a folder PowerPoint's sandbox may read
  writeFileSync(deck, buildDeck({ slides }, createMeasurer()).pptx);
  try {
    let t = Date.now();
    pp.open(deck);
    ctx.time('open 100 slides, 50 charts', Date.now() - t);
    check(Number(pp.tell('count of slides of active presentation')) === 100, 'opened without a repair prompt, 100 slides');
    for (const n of [2, 50, 100]) {
      pp.goTo(n);
      pp.macro('SelectChart', '1');
      const u = pp.run('btnUpdate');
      ctx.time(`update chart on slide ${n}`, u.ms);
      check(/RESULT btnUpdate OK/.test(u.result), `Update chart on slide ${n}: ${u.result || 'no RESULT line'}`);
    }
    t = Date.now();
    const read = pp.charts(join(ctx.out, 'big-deck.pptx'));
    ctx.time('export + read back', Date.now() - t);
    const charts = read.flatMap((s) => s.charts);
    check(charts.length === 50, `${charts.length} charts read back, expected 50`);
    check(charts.every((c) => !!c.spec && !!c.frame && c.roles.length > 0), 'every chart keeps OC_SPEC, OC_FRAME and its roles');
    pp.closeOwned();
  } finally {
    rmSync(deck, { force: true });
  }
}
