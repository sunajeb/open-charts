// The chart gallery's directions: insert charts and turn the selected one with the direction
// buttons, pick a direction before an insert, a direction the type cannot take is refused; the
// ribbon's Charts group listed through the accessibility API (no screen capture).
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import type { Ctx } from '../scenarios.ts';

export async function run(ctx: Ctx) {
  const { pp } = ctx;
  pp.newPresentation();
  const steps: string[] = [];
  const run = (step: string, id: string) => {
    const r = pp.run(id);
    steps.push(`${step}: ${r.log.split('\n').filter((l) => /RESULT|MESSAGE|Tell|can grow/.test(l)).join(' | ') || r.log.split('\n').pop()}`);
    return r.log;
  };
  const turn = (id: string, label: string, dir: string) => {
    pp.newSlide(); run(`insert ${label}`, `chart_${id}`); pp.macro('SelectChart', '1'); return run(`${label} ${dir}`, `dir${dir}`);
  };
  // 1. Insert up, then turn the selected chart with each button its type allows.
  turn('stacked', 'stacked', 'Left');
  turn('clustered', 'clustered', 'Down');
  turn('waterfallDown', 'build-down waterfall', 'Right');
  turn('line', 'line', 'Right');
  turn('mekko', 'Mekko', 'Right');
  turn('combo', 'combination', 'Down');
  // 2. A direction the type cannot take: refused with an explanation, the chart unchanged.
  const refused = turn('area', 'area', 'Down');
  // 3. Nothing selected: the arrow picks the next insert's direction (that insert only).
  pp.newSlide(); pp.macro('SelectChart', '0'); run('pick Left with nothing selected', 'dirLeft'); run('insert 100% after Left', 'chart_percent');
  // 4. Turned and turned back up: the direction option goes away again.
  turn('stacked', 'stacked', 'Down'); pp.macro('SelectChart', '1'); run('stacked Up', 'dirUp');

  // 5. The ribbon: the open-charts tab's Charts group, read through the accessibility API.
  let ribbon = '', queried = true;
  try {
    ribbon = pp.osa(`tell application "System Events" to tell process "Microsoft PowerPoint"
  set frontmost to true
  set rb to (first UI element of window 1 whose description is "ribbon")
  click (first radio button of rb whose name is "open-charts")
  delay 1
  set grp to first UI element of (first scroll area of rb)
  set out to ""
  repeat with e in UI elements of grp
    set out to out & (role of e) & " " & (name of e as text) & linefeed
  end repeat
  click (first UI element of grp whose name is "Insert chart")
  delay 1.5
  set g to UI element 1 of UI element 1 of window "Insert chart"
  repeat with c in UI elements of g
    set p to position of c
    set z to size of c
    set out to out & "  tile " & (description of c) & " @" & (item 1 of p as integer) & "," & (item 2 of p as integer) & " " & (item 1 of z as integer) & "x" & (item 2 of z as integer) & linefeed
  end repeat
  key code 53
  return out
end tell`, 30, 'read the ribbon (accessibility)');
  } catch (e) {
    if ((e as Error).name === 'LiveTimeout') throw e;
    ribbon = `accessibility query failed: ${(e as Error).message.split('\n')[0]}`;
    queried = false;
    ctx.note(`ribbon not checked (the terminal needs Accessibility): ${ribbon}`);
  }
  if (queried) ctx.check(/Insert chart/.test(ribbon) && /tile /.test(ribbon), 'the ribbon lists the Charts group and the gallery\'s tiles');

  const slides = pp.charts(join(ctx.out, 'directions.pptx'));
  pp.exportCopy(join(ctx.out, 'png'), 'png');
  const specs = slides.flatMap((s) => s.charts.map((c) => { const j = JSON.parse(c.spec ?? '{}'); return `${j.type} ${j.options?.direction ?? 'up'}`; }));
  writeFileSync(join(ctx.out, 'report.txt'), [...steps, '', 'Saved deck, chart per slide (type direction):', ...specs.map((s, i) => `  slide ${i + 1}: ${s}`), '', 'Ribbon (accessibility):', ribbon].join('\n'));
  const want = ['stacked left', 'clustered down', 'waterfall right', 'line right', 'mekko right', 'combo down', 'area up', 'percent left', 'stacked up'];
  want.forEach((w, i) => ctx.check(specs[i] === w, `slide ${i + 1}: ${w} (got ${specs[i] ?? 'no chart'})`));
  ctx.check(/can grow up, left, right, not down/.test(refused), 'area Down is refused with an explanation');
}
