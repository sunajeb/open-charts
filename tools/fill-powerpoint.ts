// Live check of `open-charts fill` in PowerPoint for Mac (macOS only, not in CI):
//   node --experimental-strip-types tools/fill-powerpoint.ts [filled.pptx ...]
// Fills the example template (unless decks are given), opens each deck in PowerPoint, fails if
// PowerPoint does not open it cleanly (a repair prompt blocks the open, which then times out), exports
// every slide as PNG to out/fill-powerpoint/<deck>/, closes it and quits PowerPoint.
// Only one agent may drive PowerPoint at a time: take /private/tmp/open-charts-ppt.lock first.
import { execFileSync } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, readdirSync, rmSync } from 'node:fs';
import { homedir } from 'node:os';
import { basename, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = fileURLToPath(new URL('../', import.meta.url));
const OUT = join(ROOT, 'out', 'fill-powerpoint');
// Folders PowerPoint's sandbox may read (the add-in's) and write (its own container).
const READ_DIR = join(homedir(), 'Library', 'Application Support', 'Microsoft', 'open-charts');
const WRITE_DIR = join(homedir(), 'Library', 'Containers', 'com.microsoft.Powerpoint', 'Data', 'open-charts-fill');

if (process.platform !== 'darwin') { console.log('macOS only; skipped.'); process.exit(0); }

const osa = (script: string, timeoutSec = 90) =>
  execFileSync('osascript', ['-e', `with timeout of ${timeoutSec} seconds\n${script}\nend timeout`], { encoding: 'utf8' }).trim();
const ppt = (body: string, t?: number) => osa(`tell application "Microsoft PowerPoint"\n${body}\nend tell`, t);
const q = (s: string) => JSON.stringify(s);

let decks = process.argv.slice(2);
if (!decks.length) {
  mkdirSync(OUT, { recursive: true });
  const styled = join(OUT, 'filled-ember.pptx');
  execFileSync(process.execPath, ['--experimental-strip-types', '--no-warnings', 'src/cli.ts', 'fill', 'examples/template/template.pptx', 'examples/template/data.json', '-o', join(OUT, 'filled.pptx')], { cwd: ROOT, stdio: 'inherit' });
  execFileSync(process.execPath, ['--experimental-strip-types', '--no-warnings', 'src/cli.ts', 'fill', 'examples/template/template.pptx', 'examples/template/data.json', '--style', 'examples/styles/ember.json', '-o', styled], { cwd: ROOT, stdio: 'ignore' });
  decks = [join(OUT, 'filled.pptx'), styled];
}

const failures: string[] = [];
mkdirSync(READ_DIR, { recursive: true });
try {
  for (const deck of decks) {
    const name = basename(deck, '.pptx');
    const copy = join(READ_DIR, `fill-check-${name}.pptx`);
    copyFileSync(deck, copy);
    rmSync(WRITE_DIR, { recursive: true, force: true });
    mkdirSync(WRITE_DIR, { recursive: true });
    try {
      ppt(`activate\nopen POSIX file ${q(copy)}`, 60);
    } catch (err) {
      failures.push(`${name}: PowerPoint did not open the deck cleanly (repair prompt?): ${(err as Error).message.split('\n')[0]}`);
      continue;
    }
    const opened = ppt('get name of active presentation');
    const slides = Number(ppt('count of slides of active presentation'));
    console.log(`${name}: opened as "${opened}", ${slides} slides`);
    ppt(`save active presentation in POSIX file ${q(join(WRITE_DIR, name))} as save as PNG`, 120);
    ppt('close active presentation saving no');
    rmSync(copy, { force: true });
    const dest = join(OUT, name);
    rmSync(dest, { recursive: true, force: true });
    mkdirSync(dest, { recursive: true });
    const pngs: string[] = [];
    const walk = (d: string) => { for (const f of readdirSync(d, { withFileTypes: true })) f.isDirectory() ? walk(join(d, f.name)) : f.name.toLowerCase().endsWith('.png') && pngs.push(join(d, f.name)); };
    if (existsSync(WRITE_DIR)) walk(WRITE_DIR);
    for (const p of pngs) copyFileSync(p, join(dest, basename(p)));
    if (pngs.length !== slides) failures.push(`${name}: exported ${pngs.length} PNGs for ${slides} slides`);
    console.log(`  ${pngs.length} PNGs → ${dest}`);
  }
} finally {
  try { ppt('quit saving no', 30); } catch { /* already closed */ }
}
if (failures.length) { console.error(failures.join('\n')); process.exit(1); }
console.log('OK: every deck opened without a repair prompt and exported.');
