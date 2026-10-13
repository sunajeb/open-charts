// Resizing without breaking: a chart shrunk to half and doubled back re-lays out when the user clicks
// elsewhere (a selection change); a moved and shrunk chart re-lays out in its new place on Update;
// a chart made smaller than its type's minimum shows the "too small" placeholder, and is a chart
// again once it is large enough.
import { join } from 'node:path';
import type { Ctx } from '../scenarios.ts';

export async function run(ctx: Ctx) {
  const { pp } = ctx;
  pp.newPresentation();
  const slideOf: Record<string, number> = {};
  for (const id of ['line', 'clustered', 'stacked', 'small']) {
    slideOf[id] = pp.newSlide();
    pp.run(`chart_${id === 'small' ? 'clustered' : id}`);
  }
  /** Resize chart 1 of the slide by `f`, click elsewhere, wait for the automatic re-layout. */
  const resize = (id: string, f: string) => {
    pp.goTo(slideOf[id]);
    pp.macro('SelectChart', '1');
    pp.macro('ResizeChart', '1', f);
    const before = pp.logLines().length, t = Date.now();
    pp.macro('SelectChart', '0'); // the user clicks elsewhere: a selection change
    const done = pp.waitForLog(/RELAYOUT/, before);
    ctx.time('resize (selection change incl. re-layout)', Date.now() - t);
    return done;
  };
  for (const [id, f] of [['line', '0.5'], ['line', '2'], ['clustered', '0.5']]) {
    const done = resize(id, f);
    ctx.check(/RELAYOUT OK/.test(done), `${id} x${f}: ${done || 'no automatic re-layout within 20 s (selection change not delivered?)'}`);
  }

  // Moved and shrunk by the user, then Update chart: re-laid out in the new place and size.
  pp.tell(`set g to shape 1 of slide ${slideOf.stacked} of active presentation
set left position of g to 300
set top of g to 150
set width of g to 400
set height of g to 300`);
  pp.goTo(slideOf.stacked);
  const before = pp.logLines().length;
  pp.macro('SelectChart', '1');
  pp.waitForLog(/RELAYOUT/, before, 5000); // the selection change re-lays the moved chart out by itself
  pp.macro('SelectChart', '1');
  const u = pp.run('btnUpdate');
  ctx.time('update', u.ms);
  ctx.check(/RESULT btnUpdate OK/.test(u.result), `Update chart after a move: ${u.result || 'no RESULT line'}`);

  // Too small, then large enough again.
  const tiny = resize('small', '0.15');
  ctx.check(/RELAYOUT/.test(tiny), `clustered x0.15: ${tiny || 'no automatic re-layout'}`);
  let slides = pp.charts(join(ctx.out, 'small.pptx'));
  const small = slides.find((s) => s.index === slideOf.small)?.charts[0];
  ctx.check(!!small?.roles.includes('placeholder'), `below its minimum size the chart shows the "too small" placeholder (roles ${[...new Set(small?.roles)].join(', ')})`);
  const big = resize('small', '6');
  ctx.check(/RELAYOUT OK/.test(big), `clustered x6 again: ${big || 'no automatic re-layout'}`);

  slides = pp.charts(join(ctx.out, 'resize.pptx'));
  const chart = (id: string) => slides.find((s) => s.index === slideOf[id])?.charts[0];
  for (const [id, w] of [['line', 864], ['clustered', 432]] as const) {
    const c = chart(id);
    ctx.check(!!c?.frame && Math.abs(c.frame.w - w) <= 2, `${id}: OC_FRAME width ${c?.frame?.w} after resizing, expected ${w} (re-laid out)`);
    ctx.check(!!c?.frame && !!c.bounds && Math.abs(c.bounds.w - c.frame.w) <= 1 && Math.abs(c.bounds.h - c.frame.h) <= 1, `${id}: group bounds ${JSON.stringify(c?.bounds)} match OC_FRAME ${JSON.stringify(c?.frame)}`);
  }
  const st = chart('stacked');
  ctx.check(!!st?.bounds && Math.abs(st.bounds.x - 300) <= 40 && st.bounds.w <= 420, `stacked: re-laid out in the moved frame (bounds ${JSON.stringify(st?.bounds)})`);
  const again = chart('small');
  ctx.check(!!again && !again.roles.includes('placeholder'), 'large enough again: a chart, no placeholder');
}
