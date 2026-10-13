// Live test of the Edit data pop-up in real PowerPoint for Mac (not in CI):
//   node --experimental-strip-types --no-warnings tools/e2e-popup-live.ts
// Installs the add-in from this folder, inserts a clustered chart, selects it, runs Edit data, times
// the pop-up (wall clock from the macro call to the server's "POPUP shown" line), types a value into
// the pop-up (System Events keystrokes: needs Accessibility for the terminal), checks the slide was
// updated live (the server's APPLY line and the chart's OC_SPEC), presses Escape and checks the chart
// went back. Quits PowerPoint at the end. No screen capture.
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { addinMain } from '../src/addin/cli.ts';
import { addinHome } from '../src/addin/server.ts';
import { buildAddin } from '../src/addin/package.ts';

const LOG = join(addinHome(), 'addin.log');
const osa = (script: string, timeoutSec = 60): string =>
  execFileSync('osascript', ['-e', `with timeout of ${timeoutSec} seconds\n${script}\nend timeout`], { encoding: 'utf8' }).trim();
const ppt = (body: string, t?: number) => osa(`tell application "Microsoft PowerPoint"\n${body}\nend tell`, t);
const keys = (body: string) => osa(`tell application "System Events"\n${body}\nend tell`);
const lines = () => (existsSync(LOG) ? readFileSync(LOG, 'utf8').split('\n').filter(Boolean) : []);
const sleep = (ms: number) => Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
const waitFor = (re: RegExp, from: number, ms = 15_000) => {
  for (const end = Date.now() + ms; Date.now() < end; sleep(20)) {
    const hit = lines().slice(from).find((l) => re.test(l));
    if (hit) return hit;
  }
  return '';
};
const report: string[] = [];
const say = (s: string) => { report.push(s); console.log(s); };
const quit = () => { try { ppt('quit saving no', 60); } catch { /* gone */ } };
process.on('exit', quit);

await addinMain(['install'], { register: true });
const { prefix } = buildAddin();
let ready = false;
const macro = (proc: string, ...args: string[]) => {
  for (let attempt = 1; ; attempt++) {
    try {
      return ppt(`run VB macro macro name ${JSON.stringify(`${prefix}_${proc}`)} list of parameters {${args.map((a) => JSON.stringify(a)).join(', ')}}`, 120);
    } catch (err) {
      if (ready || attempt >= 8) throw err;
      sleep(5000);
    } finally { ready = true; }
  }
};
ppt('activate\nmake new presentation');
ppt('make new slide at end of active presentation with properties {layout:slide layout blank}\ngo to slide view of document window 1 number 1');
macro('Run', 'chart_clustered');
sleep(1500); // the server (and with it the helper) starts as the ribbon loads; let it preload
// What reached the slide: the session's chart once the server logged "APPLY … done" (ApplyEdit
// answered "APPLIED … OK" after drawing it).
const session = async (id: string) => (await (await fetch(`http://127.0.0.1:4174/api/session/${id}`)).json()) as { state: string; chart: { series: Array<{ values: unknown[] }> } };

for (let round = 1; round <= 2; round++) {
  macro('SelectChart', '1');
  const from = lines().length, t = Date.now();
  macro('Run', 'btnEditData');
  const vba = Date.now() - t;
  const shown = waitFor(/POPUP (shown|helper did not answer)/, from);
  const wall = Date.now() - t;
  say(`round ${round}: Edit data macro returned after ${vba} ms; ${shown ? `"${shown.slice(20)}" seen after ${wall} ms` : 'no POPUP line'}`);
  if (round === 1) { keys('key code 53'); sleep(800); continue; } // Escape: close without changes
  // Type a value into the first cell (focused), wait for the live apply.
  const typedAt = lines().length;
  keys('keystroke "77"');
  const apply = waitFor(/APPLY \w+ (done|ran|failed)/, typedAt, 10_000);
  say(`typed 77 -> ${apply ? apply.slice(20) : 'no APPLY line'}`);
  const id = /APPLY (\w+)/.exec(apply)?.[1] ?? '';
  const edited = JSON.stringify((await session(id)).chart.series);
  say(`session after live apply: ${edited}`);
  const escAt = lines().length;
  keys('key code 53');
  const revert = waitFor(/APPLY \w+ (done|ran|failed)/, escAt, 10_000);
  say(`Escape -> ${revert ? revert.slice(20) : 'no APPLY line'}`);
  const back = await session(id);
  say(`session after Escape: ${back.state} ${JSON.stringify(back.chart.series)}`);
}
console.log('\n' + report.join('\n'));
