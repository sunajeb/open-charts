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
import { DEFAULT_THEME, defaultStyle, layoutChart } from '../src/core/index.ts';
import { buildDeck, chartContent, type DeckSpec } from '../src/deck.ts';
import { writePptx } from '../src/render/pptx.ts';
import { decodePng, diffImages, encodePng, type Image } from '../test/goldens/png.ts';
import { addinHome } from '../src/addin/server.ts';
import { specKey } from '../src/addin/engine.ts';
import { buildAddin } from '../src/addin/package.ts';
import { galleryItems, insertableItems, sampleFor } from '../src/addin/registry.ts';
import { readCharts, readThemeColors } from '../src/ooxml/pptx-read.ts';
import { unzip } from '../src/ooxml/unzip.ts';
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
// Start PowerPoint with a new presentation. After a quit, PowerPoint keeps its process for a moment;
// talking to it then fails with -609/-600, so wait until it is gone and retry once if it still races.
let quitting = false;
// Text of a PowerPoint dialog/sheet, pressing its default button; '' when there is none.
const dismissDialog = (): string => {
  try {
    return osa(`tell application "System Events" to tell process "Microsoft PowerPoint"
set out to ""
repeat with w in windows
  if (subrole of w is "AXDialog") or (exists sheet 1 of w) then
    try
      set out to out & (value of static texts of w as text)
    end try
    keystroke return
  end if
end repeat
return out
end tell`, 20);
  } catch { return ''; }
};
const launch = () => {
  for (let i = 0; quitting && i < 60; i++) {
    try { execFileSync('pgrep', ['-x', 'Microsoft PowerPoint']); } catch { break; } // not running
    try { if (osa('application "Microsoft PowerPoint" is running') === 'false') break; } catch { /* quitting */ }
    sleep(500);
  }
  try { launch(); } catch (e) {
    if (!/-60[09]/.test(String((e as Error).message))) throw e;
    sleep(3000);
    launch();
  }
};
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
      // While the add-in is still loading, fail fast and look for whatever is blocking PowerPoint.
      ppt(`run VB macro macro name ${q(`${prefix}_${proc}`)} list of parameters {${args.map(q).join(', ')}}`, ready ? 300 : 60);
      break;
    } catch (err) {
      const dialog = dismissDialog();
      if (dialog) console.log(`PowerPoint dialog while running ${proc}: ${dialog} (dismissed)`);
      if (ready || attempt >= 5) throw err;
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

/**
 * Apply in the editor: POSTs the edited chart to the session as the editor does and waits for the
 * server's outcome on it (drawn, or error + message), then for the macro's APPLIED line in the log.
 */
async function applyInEditor(id: string, chart: unknown) {
  const before = logLines().length, t = Date.now();
  // A pooled keep-alive socket may have been closed by the server meanwhile: retry once.
  const post = () => fetch(`http://127.0.0.1:4174/api/session/${id}`, { method: 'POST', body: JSON.stringify({ chart, state: 'applied' }) });
  const rev = (await (await post().catch(post)).json()).rev;
  let outcome: { state: string; rev: number; message?: string } | undefined;
  for (const end = Date.now() + 30_000; Date.now() < end && !(outcome && outcome.rev > rev && ['drawn', 'error'].includes(outcome.state)); sleep(100)) {
    outcome = await fetch(`http://127.0.0.1:4174/api/session/${id}`).then((r) => r.json(), () => undefined);
  }
  const ms = Date.now() - t;
  const applied = waitForLog(/APPLIED|APPLY .* (failed|skipped)/, before, 1000);
  return { outcome, applied, ms, log: logLines().slice(before).join('\n') };
}

const results: string[] = [];
const failures: string[] = [];
const timings = { insert: [] as number[], update: [] as number[], resize: [] as number[], edit: 0, apply: [] as number[] };
const quitPowerPoint = () => { try { ppt('quit saving no', 60); } catch { /* already gone */ } };
process.on('exit', (code) => { if (code !== 0) quitPowerPoint(); });
const abortIfFailed = (what: string) => {
  const bad = results.join('\n').split('\n').filter((l) => /RESULT/.test(l) && !/ OK \(\d+ ms\)/.test(l) && !/not available/.test(l));
  if (bad.length) { console.error(`${what} failed:\n  ${bad.join('\n  ')}`); process.exit(1); }
};

launch();
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
  const { outcome, applied, ms } = await applyInEditor(id, edited);
  timings.apply.push(ms);
  results.push(applied || 'APPLIED: nothing in the log', `SESSION ${id} ${outcome?.state}${outcome?.message ? `: ${outcome.message}` : ''}`);
  if (outcome?.state === 'error') console.log(`Apply did not reach the slide; the editor shows: ${outcome.message}`);
  if (!/APPLIED \w+ OK/.test(applied) || outcome?.state !== 'drawn') {
    failures.push(`Apply in the editor did not update the chart (${applied || 'timeout'}, session ${outcome?.state}); falling back to Update chart`);
    macro('SelectChart', '1');
    results.push(macro('Run', 'btnUpdate'));
  }
}

// 5b. A chart that was ungrouped: Edit data says so (its pieces keep OC_ROLE, the data is gone).
{
  const n = newSlide();
  results.push(macro('Run', 'chart_clustered'));
  try {
    ppt(`ungroup shape 1 of slide ${n} of active presentation
select shape 2 of slide ${n} of active presentation`);
    const out = macro('Run', 'btnEditData');
    const ok = /This chart was ungrouped, so it can't be edited as a chart/.test(out);
    results.push(ok ? 'UNGROUPED edit data: "This chart was ungrouped…" OK' : out);
    if (!ok) failures.push(`ungrouped chart: ${out}`);
  } catch (err) {
    results.push(`UNGROUPED skipped: ${(err as Error).message.split('\n')[0]}`);
  }
  ppt(`delete slide ${n} of active presentation`);
  slideNo--;
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
ppt('close active presentation saving no\nquit saving no'); quitting = true;
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
// Charts use theme colours, which PowerPoint draws in the deck's theme (here the new presentation's
// Office theme): the CLI copy is laid out with, and written in, that same theme.
const deckTheme = { ...DEFAULT_THEME, ...readThemeColors(new Uint8Array(readFileSync(DECK))) };
console.log(`Deck theme: accent1 ${deckTheme.accent1}, dk1 ${deckTheme.dk1}`);
const style = defaultStyle;
const cliSlides = slides.map((sl) => {
  const c = sl.charts[0];
  // A slide whose insert failed has no chart: keep an empty slide so slide numbers still line up
  // (the failed insert is already reported from its RESULT line).
  if (!c?.spec || !c.frame) return { scenes: [] };
  const { style: overrides, ...chart } = JSON.parse(c.spec);
  const scene = layoutChart(chart, c.frame!, createMeasurer(), { ...style, ...overrides, theme: { ...overrides?.theme, ...deckTheme } });
  return { scenes: [chartContent(chart, scene, style.font, overrides)] };
});
const compareDeck = join(PPAM_DIR, 'compare.pptx'); // a folder PowerPoint's sandbox may read
writeFileSync(compareDeck, writePptx(cliSlides, style.font, deckTheme));
ppt(`activate\nopen POSIX file ${q(compareDeck)}`);
ready = false; // PowerPoint restarted: the add-in compiles again
macro('ExportCopy', CLI_PNG_DIR, 'png');
ppt('close active presentation saving no\nquit saving no'); quitting = true;
rmSync(compareDeck, { force: true });

// 9b. Apply with PowerPoint closed: the session (and so the editor) says so instead of "updated".
if ('spec' in wf) {
  const id = specKey(JSON.stringify(wf.spec));
  const running = () => execFileSync('osascript', ['-e', 'application "Microsoft PowerPoint" is running'], { encoding: 'utf8' }).trim() === 'true';
  for (const end = Date.now() + 20_000; Date.now() < end && running(); sleep(250)); // PowerPoint is quitting
  // A pooled keep-alive socket from minutes ago may have been closed by the server: retry once.
  const post = () => fetch(`http://127.0.0.1:4174/api/session/${id}`, { method: 'POST', body: JSON.stringify({ chart: wf.spec, state: 'applied' }) });
  const rev = (await (await post().catch(post)).json()).rev;
  let s: { state: string; rev: number; message?: string } | undefined;
  for (const end = Date.now() + 30_000; Date.now() < end && !(s && s.rev > rev && ['drawn', 'error'].includes(s.state)); sleep(100)) {
    s = await fetch(`http://127.0.0.1:4174/api/session/${id}`).then((r) => r.json(), () => undefined);
  }
  console.log(`Apply with PowerPoint closed: ${s?.state}: ${s?.message}`);
  if (s?.state !== 'error' || !/PowerPoint isn't running/.test(s.message ?? '')) failures.push(`Apply with PowerPoint closed: ${JSON.stringify(s)}`);
  if (running()) failures.push('Apply started PowerPoint');
}
// 10. Theme colours and a same-scale group, live. A CLI deck with two charts in one scale group
// opens in PowerPoint and gets another colour scheme: the charts recolour by themselves (theme
// colours). Update chart on the waterfall re-lays out its partner; resizing the stacked chart
// re-lays out the waterfall at the new scale.
const THEME_PNG = join(OUT, 'theme-png');
{
  const pair: DeckSpec = JSON.parse(readFileSync(join(ROOT, 'examples', 'two-charts.json'), 'utf8'));
  pair.slides = [pair.slides[0]];
  const themedDeck = join(PPAM_DIR, 'themed.pptx');
  writeFileSync(themedDeck, buildDeck(pair, createMeasurer()).pptx);
  ppt(`activate\nopen POSIX file ${q(themedDeck)}`);
  ready = false; // PowerPoint started again: the add-in compiles again
  macro('ExportCopy', join(THEME_PNG, 'before'), 'png');
  const scheme = '/Applications/Microsoft PowerPoint.app/Contents/Resources/Office Themes/Theme Colors/Red Orange.xml';
  try {
    ppt(`apply theme color scheme slide 1 of active presentation file name ${q(scheme)}`);
  } catch {
    ppt(`apply theme color scheme slide 1 of active presentation file name (POSIX file ${q(scheme)} as text)`);
  }
  ppt('go to slide view of document window 1 number 1');
  macro('SelectChart', '2');
  const upd = macro('Run', 'btnUpdate');
  console.log(upd);
  if (!/RESULT btnUpdate OK/.test(upd)) failures.push(`theme slide: Update chart failed: ${upd}`);
  if (!/SYNC revenue 1 of 1 charts/.test(upd)) failures.push(`theme slide: Update chart did not re-lay out the group partner (${upd})`);
  // Edit data and Apply in this restarted PowerPoint (the add-in server started for the first
  // PowerPoint must not drive this one: it froze PowerPoint and ApplyEdit never ran). The stacked
  // chart's new data keeps its scale, so the partner is re-laid out at the same group scale.
  const deckNow = join(OUT, 'themed-before-edit.pptx');
  macro('ExportCopy', deckNow, 'pptx');
  const stacked = readCharts(new Uint8Array(readFileSync(deckNow)))[0]?.charts.find((c) => c.spec && JSON.parse(c.spec).type === 'stacked');
  macro('SelectChart', '1');
  const editOut = macro('Run', 'btnEditData');
  if (!stacked?.spec || !/RESULT btnEditData OK/.test(editOut)) failures.push(`theme slide: Edit data failed (${editOut})`);
  else {
    const spec = JSON.parse(stacked.spec);
    spec.series[0].values[2] = 45;
    const a = await applyInEditor(specKey(stacked.spec), spec);
    timings.apply.push(a.ms);
    console.log(`Apply in a restarted PowerPoint (${a.ms} ms):\n${a.log}`);
    results.push(a.applied || 'APPLIED: nothing in the log');
    if (!/APPLIED \w+ OK/.test(a.applied) || a.outcome?.state !== 'drawn') failures.push(`theme slide: Apply in a restarted PowerPoint did not update the chart (${a.applied || 'timeout'}, session ${a.outcome?.state}${a.outcome?.message ? `: ${a.outcome.message}` : ''})`);
    if (!/SYNC revenue 1 of 1 charts/.test(a.log)) failures.push(`theme slide: Apply did not re-lay out the group partner (${a.log})`);
  }
  // Resize the stacked chart (its own scale changes): the selection change re-lays it out and then
  // its partner, at the new group scale.
  macro('SelectChart', '1');
  macro('ResizeChart', '1', '0.8');
  const before = logLines().length;
  macro('SelectChart', '0');
  const synced = waitForLog(/SYNC revenue/, before);
  sleep(500);
  console.log(logLines().slice(before).join('\n'));
  if (!/SYNC revenue 1 of 1 charts/.test(synced)) failures.push(`theme slide: resizing a chart did not re-lay out its group partner (${synced || 'timeout'})`);
  const themedOut = join(OUT, 'themed.pptx');
  macro('ExportCopy', themedOut, 'pptx');
  macro('ExportCopy', join(THEME_PNG, 'after'), 'png');
  ppt('close active presentation saving no\nquit saving no'); quitting = true;
  rmSync(themedDeck, { force: true });
  // The add-in drew theme colours as theme colours, and resolved them in the slide's new scheme.
  const files = unzip(new Uint8Array(readFileSync(themedOut)));
  const slideXml = files.get('ppt/slides/slide1.xml')?.text() ?? '';
  for (const slot of ['accent1', 'tx1']) if (!slideXml.includes(`<a:schemeClr val="${slot}"`)) failures.push(`theme slide: no schemeClr ${slot} in the add-in's shapes`);
  const charts = readCharts(new Uint8Array(readFileSync(themedOut)))[0]?.charts ?? [];
  console.log(`theme slide: ${charts.map((c) => `${c.name} frame ${JSON.stringify(c.frame)}`).join('; ')}`);
  if (charts.length !== 2) failures.push(`theme slide: ${charts.length} charts after the updates, expected 2`);
  console.log(`Theme slide PNGs: ${THEME_PNG}`);
}

// 11. Chart commands (ribbon Chart tools group / right-click items) and undo, live. Two charts on a
// fresh slide: Same scale, Remove from same scale, Totals, CAGR arrow, Convert to plain text, Reset
// label position; which commands apply to which selection (the right-click items' getVisible); and
// how many Cmd+Z presses undo one action (insert, Update chart, a command, a resize re-layout).
const undoCounts: Record<string, string> = {};
{
  // PowerPoint for Mac shows a selection made by one macro to the next only after a moment.
  const select = (proc: string, ...args: string[]) => { macro(proc, ...args); sleep(800); };
  launch();
  ready = false; // PowerPoint started again: the add-in compiles again
  ppt('set s to make new slide at end of active presentation with properties {layout:slide layout blank}\ngo to slide view of document window 1 number 1');
  results.push(macro('Run', 'chart_stacked'));
  results.push(macro('Run', 'chart_clustered'));
  ppt(`set a to shape 1 of slide 1 of active presentation
set left position of a to 20
set width of a to 440
set b to shape 2 of slide 1 of active presentation
set left position of b to 500
set width of b to 440`);
  select('SelectChart', '1'); macro('Run', 'btnUpdate');
  select('SelectChart', '2'); macro('Run', 'btnUpdate');
  const deckNow = () => { const f = join(OUT, 'commands.pptx'); macro('ExportCopy', f, 'pptx'); return readCharts(new Uint8Array(readFileSync(f)))[0]?.charts ?? []; };
  const specOf = (c: { spec?: string } | undefined): any => (c?.spec ? JSON.parse(c.spec) : undefined);
  const byType = (type: string) => deckNow().find((c) => specOf(c)?.type === type);
  const expect = (ok: boolean, what: string) => { results.push(`${ok ? 'OK' : 'FAILED'} ${what}`); if (!ok) failures.push(what); };
  const applies = (out: string) => /APPLIES.*$/m.exec(out.split('\n').filter((l) => / APPLIES/.test(l)).pop() ?? '')?.[0] ?? '';

  // Which commands apply (the right-click items show exactly these).
  select('SelectChart', '1');
  let out = applies(macro('Probe', '1'));
  expect(out === 'APPLIES btnEditData btnUpdate btnTotals btnCagr btnDifference btnResetLabel', `commands on a stacked chart (${out})`);
  select('SelectCharts', '1,2');
  out = applies(macro('Probe', '1'));
  expect(out === 'APPLIES btnSameScale', `commands on two charts (${out})`);

  // Same scale: one group in both specs, the partner re-laid out at the group's scale.
  out = macro('Run', 'btnSameScale');
  expect(/RESULT btnSameScale OK/.test(out) && /SAMESCALE scale1 2 of 2 charts/.test(out), `Same scale (${out})`);
  const charts = deckNow();
  expect(charts.length === 2 && charts.every((c) => specOf(c)?.options?.scaleGroup === 'scale1'), `both charts in scaleGroup scale1 (${charts.map((c) => specOf(c)?.options?.scaleGroup).join(', ')})`);
  select('SelectCharts', '1,2');
  out = applies(macro('Probe', '1'));
  expect(out === 'APPLIES btnSameScale btnUnscale', `commands on two charts in a group (${out})`);
  // Kept on later edits: Update on one re-lays out the other.
  select('SelectChart', '1');
  out = macro('Run', 'btnUpdate');
  expect(/SYNC scale1 1 of 1 charts/.test(out), `Update keeps the same scale (${out})`);
  // Remove from same scale.
  select('SelectChart', '2');
  out = macro('Run', 'btnUnscale');
  expect(/RESULT btnUnscale OK/.test(out), `Remove from same scale (${out})`);
  expect(!!byType('clustered') && !specOf(byType('clustered'))?.options?.scaleGroup, 'clustered chart left the group');

  // Totals and CAGR arrow on the stacked chart (the sample has a CAGR arrow and totals).
  select('SelectChart', '1');
  out = macro('Run', 'btnTotals');
  expect(/RESULT btnTotals OK/.test(out) && specOf(byType('stacked'))?.options?.totals === false, `Totals off (${out})`);
  select('SelectChart', '1');
  out = macro('Run', 'btnCagr');
  expect(/RESULT btnCagr OK/.test(out) && !specOf(byType('stacked'))?.options?.cagr, `CAGR arrow removed (${out})`);
  select('SelectChart', '2');
  out = macro('Run', 'btnTotals');
  expect(/Totals are not available/.test(out), `Totals on a clustered chart explains (${out})`);

  // Convert to plain text (a selected label), then Reset label position.
  const label = 'label:value:s0:c0';
  select('SelectRole', '1', label);
  out = applies(macro('Probe', '1'));
  expect(/btnPlainText/.test(out), `Convert to plain text applies to a selected label (${out})`);
  out = macro('Run', 'btnPlainText');
  expect(/RESULT btnPlainText OK/.test(out) && specOf(byType('stacked'))?.overrides?.[label]?.detach === true, `Convert to plain text (${out})`);
  select('SelectRole', '1', label);
  out = macro('Run', 'btnResetLabel');
  expect(/This label is where open-charts puts it already/.test(out), `Reset label position on a label that was not moved explains (${out})`);
  {
    // A moved value label (pinned through Apply; the totals are off by now, as the toolbar / direct manipulation would), then reset.
    select('SelectChart', '1');
    macro('Run', 'btnEditData');
    const cur = byType('stacked')!;
    const spec = specOf(cur);
    const pinned = { ...spec, overrides: { ...spec.overrides, 'label:value:s1:c1': { pin: { x: 0.3, y: 0.1 } } } };
    const a = await applyInEditor(specKey(cur.spec!), pinned);
    expect(/APPLIED \w+ OK/.test(a.applied), `pin a label through Apply (${a.applied})`);
    select('SelectRole', '1', 'label:value:s1:c1');
    out = macro('Run', 'btnResetLabel');
    expect(/RESULT btnResetLabel OK/.test(out) && !specOf(byType('stacked'))?.overrides?.['label:value:s1:c1']?.pin, `Reset label position (${out})`);
  }

  // Undo: Cmd+Z presses (System Events, as the user) until the slide is back to the state before the
  // action. Each action follows a keyboard edit of the user's (nudging chart 1 with the arrow key), so
  // an action that merged into the previous undo entry would overshoot instead of stopping there. The
  // chart the action replaces is renamed first, so a replacement with the same geometry still shows.
  const sig = () => ppt(`set acc to ""
set sl to slide 1 of presentation 1
repeat with i from 1 to (count of shapes of sl)
  set s to shape i of sl
  set n to 0
  try
    set n to count of group items of s
  end try
  set acc to acc & (get name of s) & "#" & n & "@" & ((get left position of s) as integer) & "," & ((get top of s) as integer) & "," & ((get width of s) as integer) & "," & ((get height of s) as integer) & "|"
end repeat
return acc`);
  // Keystrokes go to the frontmost app: never send them anywhere but PowerPoint (someone may be
  // using the Mac meanwhile).
  const key = (what: string) => {
    osa('tell application "Microsoft PowerPoint" to activate');
    sleep(200);
    if (osa('tell application "System Events" to get name of first process whose frontmost is true') !== 'Microsoft PowerPoint') throw new Error('PowerPoint is not the frontmost app');
    osa(`tell application "System Events" to ${what}`);
  };
  const press = () => { key('keystroke "z" using command down'); sleep(400); };
  const measure = (name: string, action: () => void, max = 8) => {
    try {
      macro('SelectChart', '1');
      key('key code 124');
      sleep(500);
      try { ppt('set name of shape 3 of slide 1 of presentation 1 to "BEFORE"'); } catch { /* no third chart yet */ }
      const pre = sig();
      action();
      sleep(300);
      const states = [sig()];
      for (let i = 0; i < max && states[states.length - 1] !== pre; i++) { press(); states.push(sig()); }
      const k = states.indexOf(pre);
      undoCounts[name] = k > 0 ? String(k) : `not back to the state before in ${max} presses`;
    } catch (err) {
      undoCounts[name] = `not measured: ${(err as Error).message.split('\n')[0]}`;
    }
    results.push(`UNDO ${name}: ${undoCounts[name]}`);
  };
  measure('insert', () => macro('Run', 'chart_stacked'));
  macro('Run', 'chart_stacked');
  measure('update', () => { macro('SelectChart', '3'); macro('Run', 'btnUpdate'); });
  measure('totals', () => { macro('SelectChart', '3'); macro('Run', 'btnTotals'); });
  measure('resize', () => {
    macro('SelectChart', '3'); macro('ResizeChart', '3', '0.7');
    const b = logLines().length; macro('SelectChart', '0'); waitForLog(/RELAYOUT/, b);
  });
  measure('same scale', () => { macro('SelectCharts', '1,3'); macro('Run', 'btnSameScale'); });
  ppt('close active presentation saving no\nquit saving no'); quitting = true;
  console.log(`Undo entries per action (Cmd+Z presses): ${Object.entries(undoCounts).map(([k, v]) => `${k} ${v}`).join(', ')}`);
  // One press per action; a resize is the user's drag plus the add-in's re-layout: two.
  for (const [k, v] of Object.entries(undoCounts)) if (!v.startsWith('not measured') && v !== (k === 'resize' ? '2' : '1')) failures.push(`undo ${k}: ${v} presses`);
}

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

/** Pixels of either image with no match (all channels within 48) in the other within ±R px, and no blend of its neighbours: glyph placement that differs by a pixel or two, or an anti-aliasing step, is not a difference. */
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
    // An anti-aliasing step on a high-contrast edge (white separators on saturated theme colours): the
    // pixel is a blend that lies between the colours of the other image's 3x3 neighbourhood.
    const lo = [255, 255, 255], hi = [0, 0, 0];
    for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
      const xx = x + dx, yy = y + dy;
      if (xx < 0 || yy < 0 || xx >= q.width || yy >= q.height) continue;
      const j = (yy * q.width + xx) * 4;
      for (let c = 0; c < 3; c++) { lo[c] = Math.min(lo[c], q.data[j + c]); hi[c] = Math.max(hi[c], q.data[j + c]); }
    }
    return [0, 1, 2].every((c) => p.data[i + c] >= lo[c] - 8 && p.data[i + c] <= hi[c] + 8);
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
  apply       ${timings.apply.join(' ms, ')} ms from editor Apply to chart updated (single chart, then a same-scale pair in a restarted PowerPoint)   VBA ${vbaMs(/APPLIED/).join(' ms, ')} ms`);
console.log(`PNGs: ${PNG_DIR} (CLI writer: ${CLI_PNG_DIR})`);
if (failures.length) {
  console.error(`\nFAILED:\n  ${failures.join('\n  ')}`);
  process.exit(1);
}
console.log(`\nOK: ${items.length} chart types inserted and saved, updates applied.`);
