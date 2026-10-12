// End-to-end test of the PowerPoint add-in in real PowerPoint for Mac (macOS only, not in CI):
//   npm run e2e:powerpoint
// Installs the add-in, inserts every available chart type through the add-in's own VBA entry
// point (the same Dispatch the ribbon uses), updates charts (edited data, moved + resized),
// saves the deck, exports slide PNGs, quits PowerPoint, then checks the saved .pptx with the
// same reader as `open-charts inspect`.
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { addinMain, PPAM_DIR } from '../src/addin/cli.ts';
import { createMeasurer } from '../src/fonts.ts';
import { defaultStyle, layoutChart } from '../src/core/index.ts';
import { chartContent } from '../src/deck.ts';
import { writePptx } from '../src/render/pptx.ts';
import { decodePng, diffImages, encodePng } from '../test/goldens/png.ts';
import { addinHome } from '../src/addin/edit.ts';
import { specKey } from '../src/addin/engine.ts';
import { buildAddin } from '../src/addin/package.ts';
import { galleryItems, insertableItems, sampleFor } from '../src/addin/registry.ts';
import { readCharts } from '../src/ooxml/pptx-read.ts';
import { formatInspect } from '../src/inspect.ts';

const ROOT = fileURLToPath(new URL('../', import.meta.url));
const OUT = join(ROOT, 'out', 'e2e-powerpoint');
const DECK = join(OUT, 'addin.pptx');
const PNG_DIR = join(OUT, 'png');
const CLI_PNG_DIR = join(OUT, 'cli-png');
const LOG = join(addinHome(), 'addin.log');

if (process.platform !== 'darwin') {
  console.log('e2e:powerpoint runs on macOS with PowerPoint installed; skipped.');
  process.exit(0);
}

const osa = (script: string, timeoutSec = 120): string =>
  execFileSync('osascript', ['-e', `with timeout of ${timeoutSec} seconds\n${script}\nend timeout`], { encoding: 'utf8' }).trim();
const ppt = (body: string, timeoutSec?: number) => osa(`tell application "Microsoft PowerPoint"\n${body}\nend tell`, timeoutSec);
const q = (s: string) => JSON.stringify(s);
const logLines = () => (existsSync(LOG) ? readFileSync(LOG, 'utf8').split('\n').filter(Boolean) : []);

rmSync(OUT, { recursive: true, force: true });
mkdirSync(PNG_DIR, { recursive: true });

// 1. Install (engine copy, ppam, bridge script) and register + load in PowerPoint.
addinMain(['install'], { register: true });
const { prefix } = buildAddin();
const loaded = ppt('get loaded of add in "open-charts"');
if (loaded !== 'true') {
  throw new Error('PowerPoint did not load open-charts.ppam. If PowerPoint shows a macro prompt, click "Enable Macros" and run again.');
}

// 2. Macro calls (procedures in a loaded add-in are found by plain name). Right after loading,
// PowerPoint is still compiling the add-in's VBA and answers error -18, so the first call waits.
const sleep = (ms: number) => Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
let ready = false;
function macro(proc: string, ...args: string[]) {
  const before = logLines().length;
  for (let attempt = 1; ; attempt++) {
    try {
      ppt(`run VB macro macro name ${q(`${prefix}_${proc}`)} list of parameters {${args.map(q).join(', ')}}`, 300);
      break;
    } catch (err) {
      if (ready || attempt >= 8) throw err;
      sleep(5000);
    }
  }
  ready = true;
  return logLines().slice(before).join('\n');
}

const results: string[] = [];
const quitPowerPoint = () => { try { ppt('quit saving no', 60); } catch { /* already gone */ } };
process.on('exit', (code) => { if (code !== 0) quitPowerPoint(); });
const abortIfFailed = (what: string) => {
  const bad = results.join('\n').split('\n').filter((l) => /RESULT/.test(l) && !/ OK$/.test(l) && !/not available/.test(l));
  if (bad.length) { console.error(`${what} failed:\n  ${bad.join('\n  ')}`); process.exit(1); }
};

ppt('activate\nmake new presentation');
let slideNo = 0;
const newSlide = () => {
  slideNo++;
  ppt(`set s to make new slide at end of active presentation with properties {layout:slide layout blank}
go to slide view of document window 1 number ${slideNo}`);
  return slideNo;
};

// 3. Insert every available type, one per slide.
const items = insertableItems();
const slideOf: Record<string, number> = {};
for (const id of items) {
  slideOf[id] = newSlide();
  results.push(macro('Run', `chart_${id}`));
}

abortIfFailed('Inserting charts');

// 4. A type the engine does not support yet must answer with a message, not a crash.
const unsupported = galleryItems().find((g) => !items.includes(g.id));
if (unsupported) {
  const n = newSlide();
  results.push(macro('Run', `chart_${unsupported.id}`));
  ppt(`delete slide ${n} of active presentation`);
  slideNo--;
}

// 5. Update with edited data: what the editor's save would leave behind for the waterfall chart.
const wf = sampleFor('waterfall');
if ('spec' in wf && slideOf.waterfall) {
  const edited = structuredClone(wf.spec);
  edited.categories.splice(5, 0, 'One-offs');
  edited.series[0].values.splice(5, 0, 6);
  mkdirSync(join(addinHome(), 'edits'), { recursive: true });
  writeFileSync(join(addinHome(), 'edits', `${specKey(JSON.stringify(wf.spec))}.json`), JSON.stringify(edited));
  ppt(`go to slide view of document window 1 number ${slideOf.waterfall}`);
  macro('SelectChart', '1');
  results.push(macro('Run', 'btnUpdate'));
}

// 6. Update after the user moved and shrank the chart: it re-lays out in the new place and size.
if (slideOf.stacked) {
  ppt(`set g to shape 1 of slide ${slideOf.stacked} of active presentation
set left position of g to 300
set top of g to 150
set width of g to 400
set height of g to 300
go to slide view of document window 1 number ${slideOf.stacked}`);
  macro('SelectChart', '1');
  results.push(macro('Run', 'btnUpdate'));
}

abortIfFailed('Updating charts');

// 7. Save, export PNGs, quit.
// PowerPoint is sandboxed: the add-in saves into its own temp folder and the bridge copies out.
const exportCopy = (dest: string, kind: string) => macro('ExportCopy', dest, kind);
results.push(exportCopy(DECK, 'pptx'));
results.push(exportCopy(PNG_DIR, 'png'));
ppt('close active presentation saving no\nquit saving no');
if (!existsSync(DECK)) throw new Error(`The deck was not exported:\n${results.slice(-2).join('\n')}`);

if (!readdirSync(PNG_DIR, { recursive: true }).some((f) => String(f).toLowerCase().endsWith('.png'))) {
  execFileSync('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', OUT, DECK], { stdio: 'ignore' });
  execFileSync('pdftoppm', ['-png', '-r', '80', join(OUT, 'addin.pdf'), join(PNG_DIR, 'slide')]);
}

// 8. Check the saved deck.
console.log(results.filter(Boolean).join('\n'));
const slides = readCharts(new Uint8Array(readFileSync(DECK)));
console.log(formatInspect(DECK, slides));
const failures: string[] = [];
const failedRuns = results.join('\n').split('\n').filter((l) => /RESULT/.test(l) && !/ OK$/.test(l) && !(unsupported && l.includes(`chart_${unsupported.id}`)));
failures.push(...failedRuns);
if (unsupported && !results.join('\n').includes('not available')) failures.push(`no "not available" message for ${unsupported.id}`);
for (const id of items) {
  const s = slides.find((x) => x.index === slideOf[id]);
  const c = s?.charts[0];
  if (!c) { failures.push(`slide ${slideOf[id]} (${id}): no open-charts chart found`); continue; }
  if (!c.spec) failures.push(`${id}: chart has no OC_SPEC`);
  if (c.contract !== '1' || !c.frame || !c.font) failures.push(`${id}: group tags incomplete`);
  if (!c.children || c.roles.length !== c.children) failures.push(`${id}: ${c.roles.length} of ${c.children} shapes carry OC_ROLE`);
}
const wfChart = slides.find((x) => x.index === slideOf.waterfall)?.charts[0];
if (slideOf.waterfall && !wfChart?.spec?.includes('One-offs')) failures.push('waterfall: edited data was not applied by Update chart');
const st = slides.find((x) => x.index === slideOf.stacked)?.charts[0];
if (st?.bounds && (Math.abs(st.bounds.x - 300) > 40 || st.bounds.w > 420)) failures.push(`stacked: not re-laid out in the moved frame (bounds ${JSON.stringify(st.bounds)})`);
// 9. Same charts written by the CLI's own .pptx writer (same spec, same frame), rendered by the
// same PowerPoint: the add-in's native shapes must look the same.
const style = defaultStyle;
const cliSlides = slides.map((sl) => {
  const c = sl.charts[0];
  const { style: overrides, ...chart } = JSON.parse(c.spec!);
  const scene = layoutChart(chart, c.frame!, createMeasurer(), { ...style, ...overrides });
  return { scenes: [chartContent(chart, scene, style.font, overrides)] };
});
const compareDeck = join(PPAM_DIR, 'compare.pptx'); // a folder PowerPoint's sandbox may read
writeFileSync(compareDeck, writePptx(cliSlides, style.font));
ppt(`activate\nopen POSIX file ${q(compareDeck)}`);
ready = false; // PowerPoint restarted: the add-in compiles again
macro('ExportCopy', CLI_PNG_DIR, 'png');
ppt('close active presentation saving no\nquit saving no');
rmSync(compareDeck, { force: true });
for (const sl of slides) {
  const a = join(PNG_DIR, `slide${sl.index}.png`), b = join(CLI_PNG_DIR, `slide${sl.index}.png`);
  if (!existsSync(a) || !existsSync(b)) { failures.push(`slide ${sl.index}: PNG missing for comparison`); continue; }
  const d = diffImages(decodePng(readFileSync(a)), decodePng(readFileSync(b)));
  console.log(`slide ${sl.index} add-in vs CLI writer: ${d.ok ? 'same' : 'DIFFERENT'} (${(d.fraction * 100).toFixed(3)}% pixels differ)`);
  if (!d.ok) {
    failures.push(`slide ${sl.index}: add-in rendering differs from the CLI writer (${d.reason})`);
    if (d.diff) writeFileSync(join(OUT, `diff-${sl.index}.png`), encodePng(d.diff));
  }
}

console.log(`PNGs: ${PNG_DIR} (CLI writer: ${CLI_PNG_DIR})`);
if (failures.length) {
  console.error(`\nFAILED:\n  ${failures.join('\n  ')}`);
  process.exit(1);
}
console.log(`\nOK: ${items.length} chart types inserted and saved, updates applied.`);
