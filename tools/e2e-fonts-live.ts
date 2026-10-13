// Live font scenario (task B4) in real PowerPoint for Mac: does PowerPoint draw each label exactly as
// wide as the engine measured it, in fonts other than Arial, and do theme fonts reach the engine?
//   npm run e2e:fonts              (installs the add-in; opens PowerPoint, ~2 minutes; owner-approved window only)
//   npm run e2e:fonts -- --dry-run (no PowerPoint: builds the deck and prints what would be compared)
//
// 1. Writes a deck (the CLI writer) whose theme fonts are Georgia (headings) / Verdana (body): one
//    waterfall per font — Arial, Georgia, Verdana, Times New Roman, Helvetica Neue, Calibri, Aptos,
//    and "+mn-lt" (the theme's body font) — skipping fonts not installed here.
// 2. Opens it in PowerPoint; on each slide selects the chart and runs Update chart (the add-in sends
//    the deck's theme fonts, majorfont= / minorfont=, and draws the engine's layout natively).
// 3. Reads every label back (OCPREFIX_TextBounds: PowerPoint's own BoundWidth / BoundHeight) and
//    compares it with the engine's measurement: widths within 2 %, and no label wider than its box.
// 4. Exports slide PNGs to out/e2e-fonts/png, quits PowerPoint (not saving), writes
//    out/e2e-fonts/report.txt; exits 1 on any mismatch.
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { defaultStyle, layoutChart, resolveFont, type Style, type TextShape } from '../src/core/index.ts';
import { samples } from '../src/core/samples.ts';
import { chartContent } from '../src/deck.ts';
import { writePptx } from '../src/render/pptx.ts';
import { createMeasurer, findFont } from '../src/fonts.ts';

const ROOT = fileURLToPath(new URL('../', import.meta.url));
const OUT = join(ROOT, 'out', 'e2e-fonts');
const dryRun = process.argv.includes('--dry-run');
const themeFonts = { major: 'Georgia', minor: 'Verdana' };
const FONTS = ['Arial', 'Georgia', 'Verdana', 'Times New Roman', 'Helvetica Neue', 'Calibri', 'Aptos', '+mn-lt'];
const frame = { x: 48, y: 96, w: 864, h: 408 };
const spec = samples.waterfall!;

const m = createMeasurer();
const fonts = FONTS.filter((f) => f.startsWith('+') || findFont(f));
const skipped = FONTS.filter((f) => !fonts.includes(f));
rmSync(OUT, { recursive: true, force: true });
mkdirSync(OUT, { recursive: true });

/** What the engine expects PowerPoint to draw, per label role: text, font, measured width, box width. */
const expected = fonts.map((font) => {
  const style: Style = { ...defaultStyle, font, themeFonts };
  const scene = layoutChart(spec, frame, m, style);
  const fm = m.forFont!(resolveFont(font, style));
  const labels = new Map<string, { text: string; w: number; box: number; h: number }>();
  scene.shapes.forEach((s, i) => {
    if (s.kind !== 'text' || !(s as TextShape).text) return;
    const t = s as TextShape;
    const lines = t.text.split('\n');
    labels.set(scene.roles![i], { text: t.text, w: Math.max(...lines.map((l) => fm.width(l, t.size, t.bold))), box: t.w, h: lines.length * (t.lineHeight ?? fm.lineHeight(t.size)) });
  });
  return { font, resolved: resolveFont(font, style), scene, labels, content: chartContent({ ...spec, style: { font } }, scene, font, { font }) };
});
console.log(`Fonts: ${fonts.map((f) => (f.startsWith('+') ? `${f} (= ${themeFonts.minor})` : f)).join(', ')}${skipped.length ? `; not installed, skipped: ${skipped.join(', ')}` : ''}`);

if (dryRun) {
  for (const e of expected) console.log(`  ${e.font.padEnd(16)} ${e.labels.size} labels, e.g. ${[...e.labels].slice(0, 2).map(([r, l]) => `${r} "${l.text}" ${l.w.toFixed(1)} pt`).join(', ')}`);
  console.log('Dry run: PowerPoint not started.');
  process.exit(0);
}
if (process.platform !== 'darwin') { console.log('e2e:fonts runs on macOS with PowerPoint installed; skipped.'); process.exit(0); }

const { addinMain, PPAM_DIR } = await import('../src/addin/cli.ts');
const { buildAddin } = await import('../src/addin/package.ts');
const { addinHome } = await import('../src/addin/server.ts');
const LOG = join(addinHome(), 'addin.log');
const osa = (script: string, timeoutSec = 120): string =>
  execFileSync('osascript', ['-e', `with timeout of ${timeoutSec} seconds\n${script}\nend timeout`], { encoding: 'utf8' }).trim();
const ppt = (body: string, timeoutSec?: number) => osa(`tell application "Microsoft PowerPoint"\n${body}\nend tell`, timeoutSec);
const q = (s: string) => JSON.stringify(s);
const sleep = (ms: number) => Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
const logLines = () => (existsSync(LOG) ? readFileSync(LOG, 'utf8').split('\n').filter(Boolean) : []);

await addinMain(['install'], { register: true });
const { prefix } = buildAddin();
let ready = false;
function macro(proc: string, ...args: string[]): string {
  const before = logLines().length;
  for (let attempt = 1; ; attempt++) {
    try {
      ppt(`run VB macro macro name ${q(`${prefix}_${proc}`)} list of parameters {${args.map(q).join(', ')}}`, ready ? 300 : 60);
      break;
    } catch (err) {
      if (ready || attempt >= 5) throw err;
      sleep(5000); // the add-in is still compiling after PowerPoint opened
    }
  }
  ready = true;
  return logLines().slice(before).join('\n');
}

// 1–2. The deck, in a folder PowerPoint's sandbox may read; open it and update each chart.
const deck = join(PPAM_DIR, 'fonts-live.pptx');
writeFileSync(deck, writePptx(expected.map((e) => ({ title: e.font, scenes: [e.content] })), 'Arial', undefined, themeFonts));
const failures: string[] = [];
const report: string[] = [`open-charts live font scenario ${new Date().toISOString()}`, `theme fonts: ${themeFonts.major} / ${themeFonts.minor}`, ''];
try {
  ppt(`activate\nopen POSIX file ${q(deck)}`);
  if (ppt('get loaded of add in "open-charts"') !== 'true') throw new Error('PowerPoint did not load open-charts.ppam (macro prompt?).');
  expected.forEach((e, k) => {
    ppt(`go to slide view of document window 1 number ${k + 1}`);
    macro('SelectChart', '1');
    const upd = macro('Run', 'btnUpdate');
    const drawn = macro('TextBounds', '1').split('\n').map((l) => l.match(/TEXT (.*)$/)?.[1]).filter((x): x is string => !!x);
    let worst = 0, n = 0;
    for (const line of drawn) {
      // The text's own characters (the whole range's BoundWidth also counts the paragraph mark, about a space wider).
      const [role, fontName, bwAll, boxW, bh, , bwChars] = line.split('|');
      const bw = bwChars || bwAll;
      const want = e.labels.get(role);
      if (!want) continue;
      n++;
      const ratio = Number(bw) / want.w;
      worst = Math.max(worst, Math.abs(ratio - 1));
      if (Math.abs(ratio - 1) > 0.02) failures.push(`${e.font}: ${role} "${want.text}" PowerPoint ${bw} pt, engine ${want.w.toFixed(2)} pt (${((ratio - 1) * 100).toFixed(1)} %)`);
      if (Number(bw) > Number(boxW) + 0.5) failures.push(`${e.font}: ${role} "${want.text}" overflows its box (${bw} > ${boxW} pt)`);
      if (Number(bh) > want.h * 1.15 + 1) failures.push(`${e.font}: ${role} "${want.text}" is taller than measured (${bh} vs ${want.h.toFixed(1)} pt)`);
      if (!e.font.startsWith('+') && fontName !== e.font) failures.push(`${e.font}: ${role} drawn in "${fontName}"`);
    }
    if (!n) failures.push(`${e.font}: no labels read back (update: ${upd.split('\n').slice(-1)[0] ?? ''})`);
    report.push(`${e.font.padEnd(16)} ${n} labels, worst width difference ${(worst * 100).toFixed(2)} %`);
    console.log(report.at(-1));
  });
  // 4. PNGs to look at.
  macro('ExportCopy', join(OUT, 'png'), 'png');
} finally {
  try { ppt('close active presentation saving no\nquit saving no', 60); } catch { /* already closed */ }
  rmSync(deck, { force: true });
}
report.push('', failures.length ? `FAILED (${failures.length}):\n  ${failures.join('\n  ')}` : 'OK', `slide PNGs: ${join(OUT, 'png')}`);
writeFileSync(join(OUT, 'report.txt'), `${report.join('\n')}\n`);
console.log(report.slice(-2).join('\n'));
process.exit(failures.length ? 1 : 0);
