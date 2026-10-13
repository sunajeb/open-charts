// The Edit data pop-up: time from the macro call to the server's "POPUP shown" line; Escape closes it
// unchanged; a value typed into the pop-up (System Events keystrokes: the terminal needs
// Accessibility) reaches the slide live (the server's APPLY line and the session), and Escape puts
// the chart back. No screen capture.
import { sleep } from '../powerpoint.ts';
import type { Ctx } from '../scenarios.ts';

export async function run(ctx: Ctx) {
  const { pp } = ctx;
  pp.newPresentation();
  pp.newSlide();
  pp.run('chart_clustered');
  sleep(1500); // the server (and with it the helper) starts as the ribbon loads; let it preload
  const series = async (id: string) => JSON.stringify((await pp.session(id))?.chart?.series);

  for (let round = 1; round <= 2; round++) {
    pp.macro('SelectChart', '1');
    const from = pp.logLines().length, t = Date.now();
    const r = pp.run('btnEditData');
    const shown = pp.waitForLog(/POPUP (shown|helper did not answer)/, from, 15_000);
    ctx.time('pop-up shown', Date.now() - t);
    ctx.check(/POPUP shown/.test(shown), `round ${round}: pop-up shown (macro ${r.ms} ms, ${shown ? `"${shown.slice(20)}" after ${Date.now() - t} ms` : 'no POPUP line within 15 s'})`);
    if (!shown) return;
    if (round === 1) { pp.key('key code 53'); sleep(800); continue; } // Escape: close without changes
    // Type a value into the first cell (focused), wait for the live apply.
    const typedAt = pp.logLines().length;
    pp.key('keystroke "77"');
    const apply = pp.waitForLog(/APPLY \w+ (done|ran|failed)/, typedAt, 10_000);
    ctx.check(/APPLY \w+ done/.test(apply), `typed 77: applied live (${apply ? apply.slice(20) : 'no APPLY line'})`);
    const id = /APPLY (\w+)/.exec(apply)?.[1] ?? '';
    if (!id) return;
    const edited = await series(id);
    ctx.check(edited.includes('77'), `the session has the typed value: ${edited}`);
    const escAt = pp.logLines().length;
    pp.key('key code 53');
    const revert = pp.waitForLog(/APPLY \w+ (done|ran|failed)/, escAt, 10_000);
    ctx.check(/APPLY \w+ done/.test(revert), `Escape put the chart back (${revert ? revert.slice(20) : 'no APPLY line'})`);
    const back = await pp.session(id);
    ctx.check(!!back?.chart && JSON.stringify(back.chart.series) !== edited,`session after Escape: ${back?.state} ${JSON.stringify(back?.chart?.series)}`);
  }
}
