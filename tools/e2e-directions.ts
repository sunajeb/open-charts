// Live check of the chart gallery's directions in real PowerPoint for Mac (macOS only, not in CI):
//   node --experimental-strip-types --no-warnings tools/e2e-directions.ts
// Installs the add-in, then through the add-in's own entry point (the Dispatch the ribbon uses):
// inserts charts, turns the selected one with the direction buttons, picks a direction before an
// insert, checks a direction the type cannot take is refused; lists the ribbon's Charts group
// through the accessibility API (no screen capture); saves the deck, exports slide PNGs and quits.
// Run it only while holding /private/tmp/open-charts-ppt.lock (PowerPoint is shared).
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { addinMain } from '../src/addin/cli.ts';
import { addinHome } from '../src/addin/server.ts';
import { buildAddin } from '../src/addin/package.ts';
import { readCharts } from '../src/ooxml/pptx-read.ts';

const ROOT = fileURLToPath(new URL('../', import.meta.url));
const OUT = join(ROOT, 'out', 'e2e-directions');
const DECK = join(OUT, 'directions.pptx');
const PNG_DIR = join(OUT, 'png');
const LOG = join(addinHome(), 'addin.log');
if (process.platform !== 'darwin') { console.log('macOS only; skipped.'); process.exit(0); }

const osa = (script: string, timeoutSec = 120): string =>
  execFileSync('osascript', ['-e', `with timeout of ${timeoutSec} seconds\n${script}\nend timeout`], { encoding: 'utf8' }).trim();
const ppt = (body: string, timeoutSec?: number) => osa(`tell application "Microsoft PowerPoint"\n${body}\nend tell`, timeoutSec);
const q = (s: string) => JSON.stringify(s);
const logLines = () => (existsSync(LOG) ? readFileSync(LOG, 'utf8').split('\n').filter(Boolean) : []);
const sleep = (ms: number) => Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);

rmSync(OUT, { recursive: true, force: true });
mkdirSync(PNG_DIR, { recursive: true });
await addinMain(['install'], { register: true });
const { prefix } = buildAddin();
if (ppt('get loaded of add in "open-charts"') !== 'true') throw new Error('PowerPoint did not load open-charts.ppam');

let ready = false;
function macro(proc: string, ...args: string[]): string {
  const before = logLines().length;
  for (let attempt = 1; ; attempt++) {
    try { ppt(`run VB macro macro name ${q(`${prefix}_${proc}`)} list of parameters {${args.map(q).join(', ')}}`, 300); break; } catch (err) {
      if (ready || attempt >= 8) throw err;
      sleep(5000);
    }
  }
  ready = true;
  return logLines().slice(before).join('\n');
}
const quit = () => { try { ppt('quit saving no', 60); } catch { /* gone */ } };
process.on('exit', (code) => { if (code !== 0) quit(); });

const results: { step: string; log: string }[] = [];
const run = (step: string, id: string) => { const log = macro('Run', id); results.push({ step, log }); return log; };
ppt('activate\nmake new presentation');
let n = 0;
const slide = () => { n++; ppt(`make new slide at end of active presentation with properties {layout:slide layout blank}\ngo to slide view of document window 1 number ${n}`); };

// 1. Insert up, then turn the selected chart with each button its type allows.
slide(); run('insert stacked', 'chart_stacked'); macro('SelectChart', '1'); run('stacked ◀', 'dirLeft');
slide(); run('insert clustered', 'chart_clustered'); macro('SelectChart', '1'); run('clustered ▼', 'dirDown');
slide(); run('insert build-down waterfall', 'chart_waterfallDown'); macro('SelectChart', '1'); run('waterfall ▶', 'dirRight');
slide(); run('insert line', 'chart_line'); macro('SelectChart', '1'); run('line ▶ (profile)', 'dirRight');
slide(); run('insert Mekko', 'chart_mekko'); macro('SelectChart', '1'); run('mekko ▶', 'dirRight');
slide(); run('insert combination', 'chart_combo'); macro('SelectChart', '1'); run('combo ▼', 'dirDown');
// 2. A direction the type cannot take: refused with an explanation, the chart unchanged.
slide(); run('insert area', 'chart_area'); macro('SelectChart', '1');
const refused = run('area ▼ (refused)', 'dirDown');
// 3. Nothing selected: the arrow picks the next insert's direction (that insert only).
slide(); macro('SelectChart', '0'); run('pick ◀ with nothing selected', 'dirLeft'); run('insert 100% after ◀', 'chart_percent');
// 4. Turned and turned back up: the direction option goes away again.
slide(); run('insert stacked', 'chart_stacked'); macro('SelectChart', '1'); run('stacked ▼', 'dirDown'); macro('SelectChart', '1'); run('stacked ▲', 'dirUp');

// 5. The ribbon: the open-charts tab's Charts group, read through the accessibility API.
let ribbon = '';
try {
  ribbon = osa(`tell application "System Events" to tell process "Microsoft PowerPoint"
  set frontmost to true
  set w to window 1
  set out to ""
  repeat with e in (entire contents of w)
    try
      set d to description of e
      set nm to name of e
      if nm is in {"Insert chart", "Direction", "Up", "Down", "Left", "Right", "Edit data", "Update chart", "open-charts"} then set out to out & (role of e) & " " & nm & linefeed
    end try
  end repeat
  return out
end tell`, 120);
} catch (e) { ribbon = `accessibility query failed: ${(e as Error).message.split('\n')[0]}`; }

macro('ExportCopy', DECK, 'pptx');
macro('ExportCopy', PNG_DIR, 'png');
ppt('close active presentation saving no\nquit saving no');

const charts = existsSync(DECK) ? readCharts(new Uint8Array(readFileSync(DECK))) : [];
const specs = charts.flatMap((s) => s.charts.map((c) => { const j = JSON.parse(c.spec ?? '{}'); return `${j.type} ${j.options?.direction ?? 'up'}`; }));
const report = [
  ...results.map((r) => `${r.step}: ${r.log.split('\n').filter((l) => /RESULT|MESSAGE|Tell|can grow/.test(l)).join(' | ') || r.log.split('\n').pop()}`),
  '', 'Saved deck, chart per slide (type direction):', ...specs.map((s, i) => `  slide ${i + 1}: ${s}`),
  '', 'Ribbon (accessibility):', ribbon,
];
writeFileSync(join(OUT, 'report.txt'), report.join('\n'));
console.log(report.join('\n'));
const want = ['stacked left', 'clustered down', 'waterfall right', 'line right', 'mekko right', 'combo down', 'area up', 'percent left', 'stacked up'];
const bad = want.filter((w, i) => specs[i] !== w);
if (!/can grow up, left, right, not down/.test(refused)) bad.push('area ▼ was not refused with an explanation');
if (bad.length) { console.error(`FAILED: ${bad.join('; ')}`); process.exit(1); }
console.log('e2e-directions: OK');
