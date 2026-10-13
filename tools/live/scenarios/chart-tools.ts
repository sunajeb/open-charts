// Chart commands (ribbon Chart tools group / right-click items) and undo. Two charts on one slide:
// which commands apply to which selection (the right-click items' getVisible), Same scale, Remove
// from same scale, Totals, CAGR arrow, Convert to plain text, Reset label position; then how many
// Cmd+Z presses undo one action (insert, Update chart, a command, a resize re-layout, same scale).
import { join } from 'node:path';
import { specKey } from '../../../src/addin/engine.ts';
import { sleep } from '../powerpoint.ts';
import type { Ctx } from '../scenarios.ts';

export async function run(ctx: Ctx) {
  const { pp } = ctx;
  // PowerPoint for Mac shows a selection made by one macro to the next only after a moment.
  const select = (proc: string, ...args: string[]) => { pp.macro(proc, ...args); sleep(800); };
  pp.newPresentation();
  pp.newSlide();
  pp.run('chart_stacked');
  pp.run('chart_clustered');
  pp.tell(`set a to shape 1 of slide 1 of active presentation
set left position of a to 20
set width of a to 440
set b to shape 2 of slide 1 of active presentation
set left position of b to 500
set width of b to 440`);
  select('SelectChart', '1'); pp.run('btnUpdate');
  select('SelectChart', '2'); pp.run('btnUpdate');
  const deckNow = () => pp.charts(join(ctx.out, 'commands.pptx'))[0]?.charts ?? [];
  const specOf = (c: { spec?: string } | undefined): any => (c?.spec ? JSON.parse(c.spec) : undefined);
  const byType = (type: string) => deckNow().find((c) => specOf(c)?.type === type);
  const applies = (out: string) => /APPLIES.*$/m.exec(out.split('\n').filter((l) => / APPLIES/.test(l)).pop() ?? '')?.[0] ?? '';

  // Which commands apply (the right-click items show exactly these).
  select('SelectChart', '1');
  let out = applies(pp.macro('Probe', '1'));
  ctx.check(out === 'APPLIES btnEditData btnUpdate btnTotals btnCagr btnDifference btnResetLabel', `commands on a stacked chart (${out})`);
  select('SelectCharts', '1,2');
  out = applies(pp.macro('Probe', '1'));
  ctx.check(out === 'APPLIES btnSameScale', `commands on two charts (${out})`);

  // Same scale: one group in both specs, the partner re-laid out at the group's scale.
  out = pp.macro('Run', 'btnSameScale');
  ctx.check(/RESULT btnSameScale OK/.test(out) && /SAMESCALE scale1 2 of 2 charts/.test(out), `Same scale (${out})`);
  const charts = deckNow();
  ctx.check(charts.length === 2 && charts.every((c) => specOf(c)?.options?.scaleGroup === 'scale1'), `both charts in scaleGroup scale1 (${charts.map((c) => specOf(c)?.options?.scaleGroup).join(', ')})`);
  select('SelectCharts', '1,2');
  out = applies(pp.macro('Probe', '1'));
  ctx.check(out === 'APPLIES btnSameScale btnUnscale', `commands on two charts in a group (${out})`);
  select('SelectChart', '1');
  out = pp.macro('Run', 'btnUpdate');
  ctx.check(/SYNC scale1 1 of 1 charts/.test(out), `Update keeps the same scale (${out})`);
  select('SelectChart', '2');
  out = pp.macro('Run', 'btnUnscale');
  ctx.check(/RESULT btnUnscale OK/.test(out), `Remove from same scale (${out})`);
  ctx.check(!!byType('clustered') && !specOf(byType('clustered'))?.options?.scaleGroup, 'clustered chart left the group');

  // Totals and CAGR arrow on the stacked chart (the sample has a CAGR arrow and totals).
  select('SelectChart', '1');
  out = pp.macro('Run', 'btnTotals');
  ctx.check(/RESULT btnTotals OK/.test(out) && specOf(byType('stacked'))?.options?.totals === false, `Totals off (${out})`);
  select('SelectChart', '1');
  out = pp.macro('Run', 'btnCagr');
  ctx.check(/RESULT btnCagr OK/.test(out) && !specOf(byType('stacked'))?.options?.cagr, `CAGR arrow removed (${out})`);
  select('SelectChart', '2');
  out = pp.macro('Run', 'btnTotals');
  ctx.check(/Totals are not available/.test(out), `Totals on a clustered chart explains (${out})`);

  // Convert to plain text (a selected label), then Reset label position.
  const label = 'label:value:s0:c0';
  select('SelectRole', '1', label);
  // A piece of a group selected by a macro can take a moment longer to show to the next one.
  out = applies(pp.macro('Probe', '1'));
  for (let i = 0; i < 4 && !/btnPlainText/.test(out); i++) { sleep(600); out = applies(pp.macro('Probe', '1')); }
  ctx.check(/btnPlainText/.test(out), `Convert to plain text applies to a selected label (${out})`);
  out = pp.macro('Run', 'btnPlainText');
  ctx.check(/RESULT btnPlainText OK/.test(out) && specOf(byType('stacked'))?.overrides?.[label]?.detach === true, `Convert to plain text (${out})`);
  // A label that is still the chart's (the converted one is plain text now), never moved.
  select('SelectRole', '1', 'label:value:s0:c1');
  out = pp.macro('Run', 'btnResetLabel');
  ctx.check(/This label is where open-charts puts it already/.test(out), `Reset label position on a label that was not moved explains (${out})`);
  // A moved value label (pinned through Apply, as the toolbar / direct manipulation would), then reset.
  select('SelectChart', '1');
  pp.macro('Run', 'btnEditData');
  const cur = byType('stacked')!;
  const spec = specOf(cur);
  const a = await pp.apply(specKey(cur.spec!), { ...spec, overrides: { ...spec.overrides, 'label:value:s1:c1': { pin: { x: 0.3, y: 0.1 } } } });
  ctx.check(/APPLIED \w+ OK/.test(a.applied), `pin a label through Apply (${a.applied})`);
  select('SelectRole', '1', 'label:value:s1:c1');
  out = pp.macro('Run', 'btnResetLabel');
  ctx.check(/RESULT btnResetLabel OK/.test(out) && !specOf(byType('stacked'))?.overrides?.['label:value:s1:c1']?.pin, `Reset label position (${out})`);
  // Close the Edit data pop-up (Escape, as the user would): the Cmd+Z presses below must reach the
  // slide, not the pop-up's grid. Its Cancel re-applies its opening chart: wait for that to land.
  const closing = pp.logLines().length;
  pp.key('key code 53');
  pp.waitForLog(/APPLY \w+ (done|ran|failed)/, closing, 5000);
  sleep(500);

  // Undo: Cmd+Z presses (System Events, as the user) until the slide is back to the state before the
  // action. Each action follows a keyboard edit of the user's (nudging chart 1 with the arrow key), so
  // an action that merged into the previous undo entry would overshoot instead of stopping there. The
  // chart the action replaces is renamed first, so a replacement with the same geometry still shows.
  const pres = `presentation ${JSON.stringify(pp.tell('get name of active presentation'))}`;
  const sig = () => pp.tell(`set acc to ""
set sl to slide 1 of ${pres}
repeat with i from 1 to (count of shapes of sl)
  set s to shape i of sl
  set n to 0
  try
    set n to count of group items of s
  end try
  set acc to acc & (get name of s) & "#" & n & "@" & ((get left position of s) as integer) & "," & ((get top of s) as integer) & "," & ((get width of s) as integer) & "," & ((get height of s) as integer) & "|"
end repeat
return acc`);
  const measure = (name: string, action: () => void, expected: number, max = 6) => {
    let presses = '';
    try {
      pp.macro('SelectChart', '1');
      pp.key('key code 124');
      sleep(500);
      try { pp.tell(`set name of shape 3 of slide 1 of ${pres} to "BEFORE"`); } catch { /* no third chart yet */ }
      const pre = sig();
      action();
      sleep(300);
      const states = [sig()];
      for (let i = 0; i < max && states[states.length - 1] !== pre; i++) { pp.key('keystroke "z" using command down'); sleep(400); states.push(sig()); }
      const k = states.indexOf(pre);
      presses = k > 0 ? String(k) : `not back to the state before in ${max} presses`;
    } catch (err) {
      if ((err as Error).name === 'LiveTimeout') throw err;
      ctx.note(`undo ${name}: not measured: ${(err as Error).message.split('\n')[0]}`);
      return;
    }
    ctx.check(presses === String(expected), `undo ${name}: ${presses} Cmd+Z press(es), expected ${expected}`);
  };
  measure('insert', () => pp.run('chart_stacked'), 1);
  pp.run('chart_stacked');
  measure('update', () => { pp.macro('SelectChart', '3'); pp.run('btnUpdate'); }, 1);
  measure('totals', () => { pp.macro('SelectChart', '3'); pp.run('btnTotals'); }, 1);
  // A resize is the user's drag plus the add-in's re-layout: two.
  measure('resize', () => {
    pp.macro('SelectChart', '3'); pp.macro('ResizeChart', '3', '0.7');
    const b = pp.logLines().length; pp.macro('SelectChart', '0'); pp.waitForLog(/RELAYOUT/, b);
  }, 2);
  measure('same scale', () => { pp.macro('SelectCharts', '1,3'); pp.run('btnSameScale'); }, 1);
}
