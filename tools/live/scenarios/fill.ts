// `open-charts fill`: the example template filled (plain and with the Ember style), or the decks
// given with --deck, open in PowerPoint without a repair prompt (a repair prompt blocks the open,
// which times out and is reported with the dialog's text) and export every slide as PNG.
import { execFileSync } from 'node:child_process';
import { copyFileSync, existsSync, readdirSync, rmSync } from 'node:fs';
import { basename, join, resolve } from 'node:path';
import { PPAM_DIR } from '../../../src/addin/cli.ts';
import { sleep } from '../powerpoint.ts';
import type { Ctx } from '../scenarios.ts';

export async function run(ctx: Ctx) {
  const { pp } = ctx;
  let decks = ctx.opts.decks.map((d) => resolve(d));
  if (!decks.length) {
    const cli = (...args: string[]) => execFileSync(process.execPath, ['--experimental-strip-types', '--no-warnings', 'src/cli.ts', 'fill', 'examples/template/template.pptx', 'examples/template/data.json', ...args], { cwd: ctx.root, stdio: 'ignore', timeout: 60_000 });
    cli('-o', join(ctx.out, 'filled.pptx'));
    cli('--style', 'examples/styles/ember.json', '-o', join(ctx.out, 'filled-ember.pptx'));
    decks = [join(ctx.out, 'filled.pptx'), join(ctx.out, 'filled-ember.pptx')];
  }
  for (const deck of decks) {
    const name = basename(deck, '.pptx');
    const copy = join(PPAM_DIR, `fill-check-${name}.pptx`); // a folder PowerPoint's sandbox may read
    copyFileSync(deck, copy);
    try {
      pp.open(copy);
      const slides = Number(pp.tell('count of slides of active presentation'));
      ctx.check(slides > 0, `${name}: opened without a repair prompt, ${slides} slides`);
      const dest = join(ctx.out, name);
      pp.exportCopy(dest, 'png');
      const pngs = () => (existsSync(dest) ? readdirSync(dest).filter((f) => f.toLowerCase().endsWith('.png')).length : 0);
      for (const end = Math.min(Date.now() + 30_000, pp.deadline); Date.now() < end && pngs() < slides; sleep(250));
      ctx.check(pngs() === slides, `${name}: ${pngs()} PNGs for ${slides} slides (${dest})`);
      pp.closeOwned();
    } finally {
      rmSync(copy, { force: true });
    }
  }
}
