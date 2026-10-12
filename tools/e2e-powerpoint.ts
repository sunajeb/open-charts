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
import { decodePng, diffImages, encodePng, type Image } from '../test/goldens/png.ts';
import { addinHome } from '../src/addin/server.ts';
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
await addinMain(["install"], { register: true });
const { prefix } = buildAddin();
const loaded = ppt('get loaded of add in "open-charts"');
if (loaded !== 'true') {
  throw new Error('PowerPoint did not load open-charts.ppam. If PowerPoint shows a macro prompt, click "Enable Macros" and run again.');
}

// 2. Macro calls (procedures in a loaded add-in are found by plain name). Right after loading,
// PowerPoint is still compiling the add-in's VBA and answers error -18, so the first call waits.
const sleep = (ms: number) => Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
let ready = false;
function macro(proc: string, ...args: string[]): string {
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
  const out = logLines().slice(before).join('\n');
  if (/still updating a chart/.test(out)) { sleep(500); return macro(proc, ...args); }
  return out;
}

/** Waits for the add-in's log to show a line matching `re` after line `from`. */
function waitForLog(re: RegExp, from: number, ms = 20_000): string {
  for (const end = Date.now() + ms; Date.now() < end; sleep(100)) {
    const hit = logLines().slice(from).find((l) => re.test(l));
    if (hit) return hit;
  }
  return '';
}

const results: string[] = [];
const failures: string[] = [];
const timings = { insert: [] as number[], update: [] as number[], resize: [] as number[], edit: 0, apply: 0 };
const quitPowerPoint = () => { try { ppt('quit saving no', 60); } catch { /* already gone */ } };
process.on('exit', (code) => { if (code !== 0) quitPowerPoint(); });
const abortIfFailed = (what: string) => {
  const bad = results.join('\n').split('\n').filter((l) => /RESULT/.test(l) && !/ OK \(\d+ ms\)/.test(l) && !/not available/.test(l));
  if (bad.length) { console.error(`${what} failed:\n  ${bad.join('\n  ')}`); process.exit(1); }
};

ppt('activate\nmake new presentation');
let slideNo = 0;
const newSlide = (layout = 'blank') => {
  slideNo++;
  ppt(`set s to make new slide at end of active presentation with properties {layout:slide layout ${layout}}
go to slide view of document window 1 number ${slideNo}`);
  return slideNo;
};

// 3. Insert every available type, one per slide.
const items = insertableItems();
const slideOf: Record<string, number> = {};
for (const id of items) {
  slideOf[id] = newSlide();
  const t = Date.now();
  results.push(macro('Run', `chart_${id}`));
  timings.insert.push(Date.now() - t);
}

// 3b. On a title-and-content slide the chart takes the empty content placeholder's place.
const placeholderSlide = newSlide('text slide');
const placeholderBox = ppt(`get {left position, top, width, height} of shape 2 of slide ${placeholderSlide} of active presentation`)
  .split(', ').map(Number);
results.push(macro('Run', 'chart_waterfall'));

abortIfFailed('Inserting charts');

// 4. A type the engine does not support yet must answer with a message, not a crash.
const unsupported = galleryItems().find((g) => !items.includes(g.id));
if (unsupported) {
  const n = newSlide();
  results.push(macro('Run', `chart_${unsupported.id}`));
  ppt(`delete slide ${n} of active presentation`);
  slideNo--;
}

// 5. Edit data -> Apply in the editor -> the chart on the slide updates by itself. Apply is what
// the editor POSTs to the session; the add-in's server then runs ApplyEdit in PowerPoint.
const wf = sampleFor('waterfall');
if ('spec' in wf && slideOf.waterfall) {
  ppt(`go to slide view of document window 1 number ${slideOf.waterfall}`);
  macro('SelectChart', '1');
  let t = Date.now();
  results.push(macro('Run', 'btnEditData'));
  timings.edit = Date.now() - t;
  const id = specKey(JSON.stringify(wf.spec));
  const session = await fetch(`http://127.0.0.1:4174/api/session/${id}`).then((r) => (r.ok ? r.json() : undefined), () => undefined);
  if (!session) failures.push(`Edit data did not open session ${id} on the add-in server`);
  const edited = structuredClone(wf.spec);
  edited.categories.splice(5, 0, 'One-offs');
  edited.series[0].values.splice(5, 0, 6);
  const before = logLines().length;
  t = Date.now();
  await fetch(`http://127.0.0.1:4174/api/session/${id}`, { method: 'POST', body: JSON.stringify({ chart: edited, state: 'applied' }) });
  const applied = waitForLog(/APPLIED|APPLY .* failed/, before);
  timings.apply = Date.now() - t;
  results.push(applied || 'APPLIED: nothing happened within 20 s');
  if (!/APPLIED \w+ OK/.test(applied)) {
    failures.push(`Apply in the editor did not update the chart (${applied || 'timeout'}); falling back to Update chart`);
    macro('SelectChart', '1');
    results.push(macro('Run', 'btnUpdate'));
  }
}

// 6. Resize without breaking: shrink a chart to half, then double it back; clicking elsewhere
// (a selection change) re-lays it out at the new size.
for (const [id, factors] of [['line', ['0.5', '2']], ['clustered', ['0.5']]] as const) {
  if (!slideOf[id]) continue;
  ppt(`go to slide view of document window 1 number ${slideOf[id]}`);
  for (const f of factors) {
    macro('SelectChart', '1');
    macro('ResizeChart', '1', f);
    const before = logLines().length, t = Date.now();
    macro('SelectChart', '0'); // the user clicks elsewhere: a selection change
    const done = waitForLog(/RELAYOUT/, before);
    timings.resize.push(Date.now() - t);
    results.push(done || `RELAYOUT ${id} x${f}: no automatic re-layout within 20 s`);
    if (!/RELAYOUT OK/.test(done)) failures.push(`${id} resized x${f}: ${done || 'no automatic re-layout (selection-change event not delivered?)'}`);
  }
}

// 6. Update after the user moved and shrank the chart: it re-lays out in the new place and size.
if (slideOf.stacked) {
  ppt(`set g to shape 1 of slide ${slideOf.stacked} of active presentation
set left position of g to 300
set top of g to 150
set width of g to 400
set height of g to 300
go to slide view of document window 1 number ${slideOf.stacked}`);
  const before = logLines().length;
  macro('SelectChart', '1');
  waitForLog(/RELAYOUT/, before, 5000); // the selection change re-lays the moved chart out by itself
  macro('SelectChart', '1');
  const t = Date.now();
  results.push(macro('Run', 'btnUpdate'));
  timings.update.push(Date.now() - t);
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
const failedRuns = results.join('\n').split('\n').filter((l) => /RESULT/.test(l) && !/ OK \(\d+ ms\)/.test(l) && !(unsupported && l.includes(`chart_${unsupported.id}`)));
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
for (const [id, w] of [['line', 864], ['clustered', 432]] as const) {
  const c = slides.find((x) => x.index === slideOf[id])?.charts[0];
  if (!c?.frame || !c.bounds) continue;
  if (Math.abs(c.frame.w - w) > 2) failures.push(`${id}: after resizing, OC_FRAME width is ${c.frame.w}, expected ${w} (not re-laid out)`);
  if (Math.abs(c.bounds.w - c.frame.w) > 1 || Math.abs(c.bounds.h - c.frame.h) > 1) failures.push(`${id}: group bounds ${JSON.stringify(c.bounds)} differ from OC_FRAME ${JSON.stringify(c.frame)}`);
}
const ph = slides.find((x) => x.index === placeholderSlide)?.charts[0];
if (!ph?.frame || [ph.frame.x, ph.frame.y, ph.frame.w, ph.frame.h].some((v, i) => Math.abs(v - placeholderBox[i]) > 1)) {
  failures.push(`placeholder slide: chart frame ${JSON.stringify(ph?.frame)} is not the content placeholder ${placeholderBox.join(',')}`);
}
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
  const ia = decodePng(readFileSync(a)), ib = decodePng(readFileSync(b));
  const d = diffImages(ia, ib);
  const off = beyondOnePixel(ia, ib);
  console.log(`slide ${sl.index} add-in vs CLI writer: ${d.ok ? 'identical' : off ? 'DIFFERENT' : 'same within 3 px'} (${(d.fraction * 100).toFixed(3)}% pixels differ, ${off} beyond a 3-px shift)`);
  if (d.diff && !d.ok) writeFileSync(join(OUT, `diff-${sl.index}.png`), encodePng(d.diff));
  if (off > ia.width * ia.height * 2.5e-5) failures.push(`slide ${sl.index}: add-in rendering differs from the CLI writer (${off} pixels beyond a 3-px shift)`);
}

/** Pixels of either image with no match (all channels within 48) in the other within ±R px: glyph placement that differs by a pixel or two is not a difference. */
function beyondOnePixel(a: Image, b: Image, R = 3): number {
  if (a.width !== b.width || a.height !== b.height) return a.width * a.height;
  const near = (p: Image, q: Image, x: number, y: number) => {
    const i = (y * p.width + x) * 4;
    for (let dy = -R; dy <= R; dy++) for (let dx = -R; dx <= R; dx++) {
      const xx = x + dx, yy = y + dy;
      if (xx < 0 || yy < 0 || xx >= q.width || yy >= q.height) continue;
      const j = (yy * q.width + xx) * 4;
      if (Math.abs(p.data[i] - q.data[j]) <= 48 && Math.abs(p.data[i + 1] - q.data[j + 1]) <= 48 && Math.abs(p.data[i + 2] - q.data[j + 2]) <= 48) return true;
    }
    return false;
  };
  let n = 0;
  for (let y = 0; y < a.height; y++) for (let x = 0; x < a.width; x++) if (!near(a, b, x, y) || !near(b, a, x, y)) n++;
  return n;
}

const ms = (v: number[]) => { const a = [...v].sort((x, y) => x - y); return a.length ? `median ${a[a.length >> 1]} ms, max ${a[a.length - 1]} ms (n=${a.length})` : 'n/a'; };
const part = (re: RegExp, what: string) => results.join('\n').split('\n').filter((l) => re.test(l)).map((l) => Number(new RegExp(what + ' (\\d+) ms').exec(l)?.[1])).filter(Number.isFinite);
const vbaMs = (re: RegExp) => results.join('\n').split('\n').filter((l) => re.test(l)).map((l) => Number(/\((\d+) ms\)/.exec(l)?.[1])).filter(Number.isFinite);
console.log(`\nTimings (wall clock, AppleScript run VB macro included):
  insert      ${ms(timings.insert.slice(1))} (first, cold: ${timings.insert[0]} ms)
  insert VBA  ${ms(vbaMs(/RESULT chart_/).slice(1))}: engine ${ms(part(/RESULT chart_/, 'engine').slice(1))}, draw ${ms(part(/RESULT chart_/, 'draw').slice(1))}
  update      ${ms(timings.update)}   VBA ${ms(vbaMs(/RESULT btnUpdate/))}
  resize      ${ms(timings.resize)} (selection change incl. re-layout)
  edit data   ${timings.edit} ms to session + editor launch
  apply       ${timings.apply} ms from editor Apply to chart updated   VBA ${ms(vbaMs(/APPLIED/))}`);
console.log(`PNGs: ${PNG_DIR} (CLI writer: ${CLI_PNG_DIR})`);
if (failures.length) {
  console.error(`\nFAILED:\n  ${failures.join('\n  ')}`);
  process.exit(1);
}
console.log(`\nOK: ${items.length} chart types inserted and saved, updates applied.`);
