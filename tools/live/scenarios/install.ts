// Install the add-in from this folder (engine copy, ppam, AppleScript bridge, pop-up helper),
// register and load it in PowerPoint (starts PowerPoint if needed).
import { addinMain } from '../../../src/addin/cli.ts';
import { installedPrefix } from '../powerpoint.ts';
import type { Ctx } from '../scenarios.ts';

export async function run(ctx: Ctx) {
  const t = Date.now();
  await addinMain(['install'], { register: true });
  ctx.time('install', Date.now() - t);
  ctx.pp.ready = false; // the add-in (re)loaded: its VBA compiles before the first macro answers
  ctx.check(installedPrefix() === ctx.pp.prefix, `installed add-in is this build (prefix ${ctx.pp.prefix})`);
  ctx.pp.loaded = ctx.check(ctx.pp.tell('get loaded of add in "open-charts"') === 'true', 'PowerPoint loaded open-charts.ppam (a macro prompt? click "Enable Macros" and run again)');
}
