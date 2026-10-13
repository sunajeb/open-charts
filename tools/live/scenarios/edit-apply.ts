// Edit data -> the session on the add-in's server -> Apply in the editor updates the slide by
// itself -> Cancel puts the opening chart back. Edit data on an ungrouped chart explains. When the
// run started PowerPoint: Apply with PowerPoint closed says so (and does not start it), and Apply
// works in a restarted PowerPoint (the server must not drive the PowerPoint it started with).
import { join } from 'node:path';
import { specKey } from '../../../src/addin/engine.ts';
import { sampleFor } from '../../../src/addin/registry.ts';
import { sleep } from '../powerpoint.ts';
import type { Ctx } from '../scenarios.ts';

export async function run(ctx: Ctx) {
  const { pp } = ctx;
  const wf = sampleFor('waterfall');
  if (!('spec' in wf)) throw new Error('no waterfall sample');
  const specOf = (file: string) => pp.charts(join(ctx.out, file))[0]?.charts[0]?.spec ?? '';

  // 1. Edit data opens a session; Apply reaches the slide.
  pp.newPresentation();
  pp.newSlide();
  ctx.check(/ OK \(/.test(pp.run('chart_waterfall').result), 'insert a waterfall');
  pp.macro('SelectChart', '1');
  const edit = pp.run('btnEditData');
  ctx.time('edit data', edit.ms);
  const id = specKey(JSON.stringify(wf.spec));
  ctx.check(!!(await pp.session(id)), `Edit data opened session ${id} on the add-in server (${edit.result})`);
  const edited = structuredClone(wf.spec);
  edited.categories.splice(5, 0, 'One-offs');
  edited.series[0].values.splice(5, 0, 6);
  const a = await pp.apply(id, edited);
  ctx.time('apply', a.ms);
  ctx.check(/APPLIED \w+ OK/.test(a.applied) && a.outcome?.state === 'drawn', `Apply updated the chart (${a.applied || 'no APPLIED line'}, session ${a.outcome?.state}${a.outcome?.message ? `: ${a.outcome.message}` : ''})`);
  ctx.check(specOf('applied.pptx').includes('One-offs'), 'the chart on the slide has the edited data');

  // 2. Cancel: the editor puts the opening chart back, then closes the session.
  const back = await pp.apply(id, wf.spec);
  await pp.post(id, { state: 'cancelled' });
  ctx.check(/APPLIED \w+ OK/.test(back.applied), `Cancel put the opening chart back (${back.applied || 'no APPLIED line'})`);
  ctx.check((await pp.session(id))?.state === 'cancelled', 'the session is cancelled');
  ctx.check(!specOf('cancelled.pptx').includes('One-offs'), 'the chart on the slide has the opening data again');

  // 3. An ungrouped chart: its pieces keep OC_ROLE, the data is gone; Edit data says so.
  const n = pp.newSlide();
  pp.run('chart_clustered');
  pp.tell(`ungroup shape 1 of slide ${n} of active presentation\nselect shape 2 of slide ${n} of active presentation`);
  const u = pp.run('btnEditData');
  ctx.check(/This chart was ungrouped, so it can't be edited as a chart/.test(u.log), `Edit data on an ungrouped chart explains (${u.result})`);

  // 4. The restart case: only with a PowerPoint this run started (never quit the owner's).
  if (pp.wasRunning) {
    ctx.note('restart case not run: PowerPoint was running before the run, and the runner never quits a PowerPoint it did not start');
    return;
  }
  pp.quit();
  const { rev } = await pp.post(id, { chart: wf.spec, state: 'applied' });
  let s: Awaited<ReturnType<typeof pp.session>> = undefined;
  for (const end = Math.min(Date.now() + 30_000, pp.deadline); Date.now() < end; sleep(100)) {
    s = await pp.session(id);
    if (s && s.rev > rev && ['drawn', 'error'].includes(s.state)) break;
  }
  ctx.check(s?.state === 'error' && /PowerPoint isn't running/.test(s.message ?? ''), `Apply with PowerPoint closed says so (${s?.state}: ${s?.message})`);
  ctx.check(!pp.running(), 'Apply did not start PowerPoint');

  pp.newPresentation(); // PowerPoint starts again: the add-in compiles again
  pp.newSlide();
  ctx.check(/ OK \(/.test(pp.run('chart_stacked').result), 'insert a stacked chart in the restarted PowerPoint');
  const spec = pp.charts(join(ctx.out, 'restarted.pptx'))[0]?.charts[0]?.spec;
  pp.macro('SelectChart', '1');
  const e2 = pp.run('btnEditData');
  if (!ctx.check(!!spec && /RESULT btnEditData OK/.test(e2.log), `Edit data in the restarted PowerPoint (${e2.result})`) || !spec) return;
  const changed = JSON.parse(spec);
  changed.series[0].values[2] = 45;
  const r = await pp.apply(specKey(spec), changed);
  ctx.time('apply after restart', r.ms);
  ctx.check(/APPLIED \w+ OK/.test(r.applied) && r.outcome?.state === 'drawn', `Apply in a restarted PowerPoint (${r.applied || 'no APPLIED line'}, session ${r.outcome?.state}${r.outcome?.message ? `: ${r.outcome.message}` : ''})`);
}
