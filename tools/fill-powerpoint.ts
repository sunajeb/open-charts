// Live check of `open-charts fill` in PowerPoint for Mac (macOS only, not in CI):
//   node --experimental-strip-types tools/fill-powerpoint.ts [filled.pptx ...]
// Fills the example template (unless decks are given; plain and with the Ember style), opens each
// deck in PowerPoint and fails if PowerPoint does not open it cleanly (a repair prompt blocks the
// open, which then times out). Exports every slide as PNG to out/fill-powerpoint/<deck>/ through the
// installed add-in's ExportCopy (the sandboxed PowerPoint cannot write outside its container; the
// add-in's bridge copies out), closes the deck, and quits PowerPoint unless other presentations are
// open. Needs `open-charts addin install` (any build).
// Only one agent may drive PowerPoint at a time: take /private/tmp/open-charts-ppt.lock first.
import { execFileSync } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync } from 'node:fs';
import { homedir } from 'node:os';
import { basename, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzip } from '../src/ooxml/unzip.ts';

const ROOT = fileURLToPath(new URL('../', import.meta.url));
const OUT = join(ROOT, 'out', 'fill-powerpoint');
// A folder PowerPoint's sandbox may read (the add-in's).
const READ_DIR = join(homedir(), 'Library', 'Application Support', 'Microsoft', 'open-charts');

if (process.platform !== 'darwin') { console.log('macOS only; skipped.'); process.exit(0); }

const osa = (script: string, timeoutSec = 90) =>
  execFileSync('osascript', ['-e', `with timeout of ${timeoutSec} seconds\n${script}\nend timeout`], { encoding: 'utf8' }).trim();
const ppt = (body: string, t?: number) => osa(`tell application "Microsoft PowerPoint"\n${body}\nend tell`, t);
const q = (s: string) => JSON.stringify(s);
const sleep = (ms: number) => Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);

// The installed add-in's macro prefix, from its ribbon callbacks.
const ppam = join(READ_DIR, 'open-charts.ppam');
if (!existsSync(ppam)) { console.error('Install the add-in first: open-charts addin install'); process.exit(1); }
const ribbon = [...unzip(new Uint8Array(readFileSync(ppam))).values()].find((e) => e.name.startsWith('customUI/'))?.text() ?? '';
const prefix = /onAction="(\w+?)_/.exec(ribbon)?.[1];
if (!prefix) { console.error('Cannot read the add-in macro prefix'); process.exit(1); }

function exportPngs(dest: string) {
  for (let attempt = 1; ; attempt++) {
    try {
      ppt(`run VB macro macro name ${q(`${prefix}_ExportCopy`)} list of parameters {${q(dest)}, "png"}`, 300);
      return;
    } catch (err) {
      if (attempt >= 8) throw err; // right after start-up PowerPoint is still compiling the add-in
      sleep(5000);
    }
  }
}

let decks = process.argv.slice(2);
if (!decks.length) {
  mkdirSync(OUT, { recursive: true });
  const cli = (...args: string[]) => execFileSync(process.execPath, ['--experimental-strip-types', '--no-warnings', 'src/cli.ts', 'fill', 'examples/template/template.pptx', 'examples/template/data.json', ...args], { cwd: ROOT, stdio: 'ignore' });
  cli('-o', join(OUT, 'filled.pptx'));
  cli('--style', 'examples/styles/ember.json', '-o', join(OUT, 'filled-ember.pptx'));
  decks = [join(OUT, 'filled.pptx'), join(OUT, 'filled-ember.pptx')];
}

const failures: string[] = [];
mkdirSync(READ_DIR, { recursive: true });
// Never close someone else's work: quit at the end only if no other presentation is open.
const open = () => { try { return Number(ppt('count of presentations', 20)); } catch { return 0; } };
try {
  for (const deck of decks) {
    const name = basename(deck, '.pptx');
    const copy = join(READ_DIR, `fill-check-${name}.pptx`);
    copyFileSync(deck, copy);
    try {
      ppt(`activate\nopen POSIX file ${q(copy)}`, 60);
    } catch (err) {
      failures.push(`${name}: PowerPoint did not open the deck cleanly (repair prompt?): ${(err as Error).message.split('\n')[0]}`);
      continue;
    }
    const opened = ppt('get name of active presentation');
    const slides = Number(ppt('count of slides of active presentation'));
    console.log(`${name}: opened as "${opened}", ${slides} slides`);
    const dest = join(OUT, name);
    rmSync(dest, { recursive: true, force: true });
    exportPngs(dest);
    const pngs = () => (existsSync(dest) ? readdirSync(dest).filter((f) => f.toLowerCase().endsWith('.png')) : []);
    for (const end = Date.now() + 30_000; Date.now() < end && pngs().length < slides; sleep(250));
    ppt('close active presentation saving no');
    rmSync(copy, { force: true });
    if (pngs().length !== slides) failures.push(`${name}: exported ${pngs().length} PNGs for ${slides} slides`);
    console.log(`  ${pngs().length} PNGs → ${dest}`);
  }
} finally {
  if (open() === 0) { try { ppt('quit', 30); } catch { /* already closed */ } } else console.log('Other presentations are open: PowerPoint left running.');
}
if (failures.length) { console.error(failures.join('\n')); process.exit(1); }
console.log('OK: every deck opened without a repair prompt and exported.');
