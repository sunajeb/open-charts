// Live test of the agenda in real PowerPoint for Mac (macOS only, not in CI):
//   npm run e2e:agenda
// Installs the add-in, builds a deck of 30 content slides, then drives the add-in's own entry points
// (the same Dispatch the ribbon buttons use, and the selection-change handler the automatic update
// runs in): insert and rename chapters, nest, outdent, rename on another slide, move a chapter,
// split off a backup, duplicate / delete / move slides in the sorter, remove and insert hidden
// agenda slides, a table of contents. After every step every agenda slide is read back and checked.
// Then it exports the deck and PNGs and quits PowerPoint.
// Only one PowerPoint test at a time: hold /private/tmp/open-charts-ppt.lock while it runs.
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { addinMain } from '../src/addin/cli.ts';
import { addinHome } from '../src/addin/server.ts';
import { buildAddin } from '../src/addin/package.ts';
import { unzip } from '../src/ooxml/unzip.ts';

const ROOT = fileURLToPath(new URL('../', import.meta.url));
const OUT = join(ROOT, 'out', 'e2e-agenda');
const DECK = join(OUT, 'agenda.pptx');
const PNG_DIR = join(OUT, 'png');
const LOG = join(addinHome(), 'addin.log');

if (process.platform !== 'darwin') { console.log('e2e:agenda runs on macOS with PowerPoint installed; skipped.'); process.exit(0); }

const osa = (script: string, timeoutSec = 120) =>
  execFileSync('osascript', ['-e', `with timeout of ${timeoutSec} seconds\n${script}\nend timeout`], { encoding: 'utf8' }).trim();
const ppt = (body: string, timeoutSec?: number) => osa(`tell application "Microsoft PowerPoint"\n${body}\nend tell`, timeoutSec);
const q = (s: string) => JSON.stringify(s);
const logLines = () => (existsSync(LOG) ? readFileSync(LOG, 'utf8').split('\n').filter(Boolean) : []);
const sleep = (ms: number) => Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);

rmSync(OUT, { recursive: true, force: true });
mkdirSync(PNG_DIR, { recursive: true });
await addinMain(['install'], { register: true });
const { prefix } = buildAddin();
if (ppt('get loaded of add in "open-charts"') !== 'true') throw new Error('PowerPoint did not load open-charts.ppam.');

/** The text of a dialog PowerPoint shows (a VBA error, a message box), after pressing its default button. */
function dismissDialog(): string {
  try {
    return osa(`tell application "System Events" to tell process "Microsoft PowerPoint"
set t to ""
repeat with w in windows
  try
    set t to t & (value of every static text of w as text) & " | "
  end try
end repeat
key code 36
return t
end tell`, 20);
  } catch (err) { return `(could not read the dialog: ${(err as Error).message.split('\n')[0]})`; }
}

let ready = false;
function macro(proc: string, ...args: string[]): string {
  const before = logLines().length;
  for (let attempt = 1; ; attempt++) {
    try { ppt(`run VB macro macro name ${q(`${prefix}_${proc}`)} list of parameters {${args.map(q).join(', ')}}`, ready ? 90 : 300); break; } catch (err) {
      const msg = (err as Error).message;
      if (/-1712|timed out/.test(msg)) {
        console.error(`${proc}(${args.join(', ')}) timed out. Dialog: ${dismissDialog()}\nLog:\n${logLines().slice(before).join('\n')}`);
        throw err;
      }
      if (ready || attempt >= 8) throw err;
      sleep(5000);
    }
  }
  ready = true;
  return logLines().slice(before).join('\n');
}

const quit = () => { try { ppt('quit saving no', 60); } catch { /* gone */ } };
process.on('exit', (code) => { if (code !== 0) quit(); });
const failures: string[] = [];
const timings: string[] = [];
const check = (ok: boolean, what: string) => { console.log(`${ok ? 'ok  ' : 'FAIL'} ${what}`); if (!ok) failures.push(what); };

interface Row { n: number; kind: string; hidden: boolean; split: boolean; headings: string; current: string }
function dump(): Row[] {
  const out = macro('AgendaDump');
  const line = out.split('\n').reverse().find((l) => l.includes('DUMP '))!;
  return line.slice(line.indexOf('DUMP ') + 5).split(' ;; ').filter((s) => s.startsWith('SLIDE')).map((s) => {
    const [head, headings = ''] = s.split(':');
    const [, n, kind, ...flags] = head.trim().split(' ');
    const cur = headings.trim().split(' ').find((h) => h.endsWith('*')) ?? '';
    return { n: Number(n), kind, hidden: flags.includes('hidden'), split: flags.includes('split'), headings: headings.trim(), current: cur.replace(/^>+/, '').replace(/\*$/, '') };
  });
}
const agendaRows = (rows = dump()) => rows.filter((r) => r.kind !== '-');
const slideOf = (title: string, rows = dump()) => rows.find((r) => r.kind === 'chapter' && r.current === title)?.n ?? -1;
const show = (rows: Row[]) => agendaRows(rows).map((r) => `  ${r.n}${r.hidden ? ' (hidden)' : ''}: ${r.headings}`).join('\n');

/** Runs a ribbon command (silently, as the automation entry point does) and times it. */
function run(id: string) {
  const t = Date.now();
  const out = macro('Run', id);
  timings.push(`${id} ${Date.now() - t} ms`);
  const res = out.split('\n').find((l) => l.includes(`RESULT ${id}`)) ?? '';
  if (!/ OK \(/.test(res)) console.log(`  ${res || out}`);
  return res;
}
/** Words separated by spaces in headings: multi-word titles use "_" here and in the deck. */
function chapter(afterSlide: number, title: string) {
  macro('AgendaSelect', String(afterSlide), '');
  run('btnAgendaChapter');
  const rows = dump();
  const n = rows.find((r) => r.kind === 'chapter' && r.current === 'Chapter')?.n ?? -1;
  check(n === afterSlide + 1, `Chapter after slide ${afterSlide} inserts agenda slide ${afterSlide + 1} reading "Chapter" (got ${n})`);
  macro('AgendaSelect', String(n), '');
  macro('AgendaEdit', String(n), 'Chapter', title, '0');
  macro('AgendaSorter', 'none', '', ''); // the click elsewhere: automatic update
  return n;
}
const allShow = (pred: (r: Row) => boolean, what: string, rows = dump()) => {
  const bad = agendaRows(rows).filter((r) => !pred(r));
  check(!bad.length, `${what}${bad.length ? `\n${show(bad)}` : ''}`);
};

// 1. A deck of 30 content slides.
ppt('activate\nmake new presentation');
for (let i = 1; i <= 30; i++) {
  ppt(`set s to make new slide at end of active presentation with properties {layout:slide layout title only}
set content of text range of text frame of shape 1 of s to "Content ${i}"`);
}
ppt('go to slide view of document window 1 number 1');

// 2. Insert and rename chapters.
chapter(1, 'Introduction');
chapter(6, 'Market');
const size = chapter(7, 'Market_size');
macro('AgendaSelect', String(size), '');
macro('AgendaEdit', String(size), 'Market_size', '', '2');
macro('AgendaSorter', 'none', '', '');
let rows = dump();
check(rows.find((r) => r.current === 'Market')!.hidden, 'nesting Market_size directly after Market hides Market\'s agenda slide');
check(rows.find((r) => r.current === 'Market_size')!.headings === 'Introduction Market >Market_size*', 'Market_size\'s slide: path and the subchapter, nothing else' + `\n${show(rows)}`);
chapter(12, 'Growth');
const comp = chapter(18, 'Competition');
rows = dump();
check(rows.find((r) => r.current === 'Competition')!.headings.includes('>Competition*'), 'a new chapter takes the level of the one before it');
macro('AgendaSelect', String(comp), '');
run('btnAgendaOutdent');
chapter(25, 'Plan');
rows = dump();
console.log(show(rows));
allShow((r) => /^Introduction\*? Market\*? (>Market_size\*? >Growth\*? )?Competition\*? Plan\*?$/.test(r.headings.trim()), 'every agenda slide lists the five chapters; subchapters only on Market\'s family', rows);
check(rows.length === 36, `36 slides (30 + 6 agenda slides), got ${rows.length}`);

// 3. Rename on another chapter's slide.
const planAt = slideOf('Plan');
macro('AgendaSelect', String(planAt), '');
macro('AgendaEdit', String(planAt), 'Competition', 'Competitors', '0');
macro('AgendaSorter', 'none', '', '');
allShow((r) => r.headings.includes('Competitors') && !r.headings.includes('Competition'), 'a rename on the Plan slide reaches every agenda slide');

// 4. Move Plan up: its slides move along, before Competitors.
rows = dump();
const planSlides = slideOf('Plan', rows), compSlides = slideOf('Competitors', rows);
macro('AgendaSelect', String(planSlides), 'Plan');
run('btnAgendaUp');
rows = dump();
check(slideOf('Plan', rows) === compSlides && slideOf('Competitors', rows) === compSlides + (36 - planSlides + 1), `Plan's agenda slide and its ${36 - planSlides} slides moved before Competitors`);
allShow((r) => /Plan\*? Competitors\*?\s*$/.test(r.headings), 'new order on every agenda slide', rows);

// 5. Split: a backup section with its own agenda.
macro('AgendaSelect', '34', '');
run('btnAgendaSplit');
chapter(34, 'Backup_data');
rows = dump();
check(rows.find((r) => r.current === 'Backup_data')!.headings.trim() === 'Backup_data*', 'the backup has its own agenda');
allShow((r) => r.current === 'Backup_data' || !r.headings.includes('Backup_data'), 'the main agenda does not list the backup chapter', rows);

// 6. Slide sorter: duplicate, delete, move.
macro('AgendaSorter', 'duplicate', '2', '');
rows = dump();
check(agendaRows(rows).filter((r) => r.kind === 'chapter' && r.current === 'Introduction').length === 2 && agendaRows(rows).every((r) => r.current === 'Backup_data' || (r.headings.match(/Introduction/g) ?? []).length === 2), 'a duplicated agenda slide is a new chapter on every agenda slide');
macro('AgendaSorter', 'delete', '3', '');
allShow((r) => r.current === 'Backup_data' || (r.headings.match(/Introduction/g) ?? []).length === 1, 'deleting it removes the chapter everywhere');
const market = slideOf('Market');
macro('AgendaSorter', 'move', String(market + 3), String(market + 1)); // a content slide between Market and Market_size
check(!dump().find((r) => r.current === 'Market')!.hidden, 'content between Market and its first subchapter: Market is shown again');
macro('AgendaSorter', 'move', String(market + 1), String(market + 3));
check(dump().find((r) => r.current === 'Market')!.hidden, 'and hidden again once that content is gone');

// 7. Remove and insert hidden agenda slides.
const before = dump().length;
run('btnAgendaRemoveHidden');
rows = dump();
check(rows.length === before - 1 && slideOf('Market', rows) < 0, 'Remove Hidden Agenda Slides deletes Market\'s slide');
allShow((r) => r.headings.includes('Market ') || r.current === 'Backup_data', 'Market stays in every agenda', rows);
run('btnAgendaInsertHidden');
rows = dump();
check(rows.length === before && slideOf('Market', rows) === slideOf('Market_size', rows) - 1 && rows.find((r) => r.current === 'Market')!.hidden, 'Insert Hidden Agenda Slides brings it back, hidden, before Market_size');

// 8. Table of contents before the first chapter slide.
macro('AgendaSelect', '20', '');
run('btnAgendaToc');
rows = dump();
const toc = rows.find((r) => r.kind === 'toc');
check(!!toc && toc.n === slideOf('Introduction', rows) - 1 && toc.headings.trim() === 'Introduction Market Plan Competitors', `table of contents before the first chapter: ${toc?.n}: ${toc?.headings}`);

// 9. Nothing changed: Update agenda redraws nothing and is quick.
const t0 = Date.now();
run('btnAgendaUpdate');
console.log(`Update agenda on ${rows.length} slides, nothing changed: ${Date.now() - t0} ms`);
console.log(show(dump()));

// 10. Export, quit, read back.
macro('ExportCopy', DECK, 'pptx');
macro('ExportCopy', PNG_DIR, 'png');
ppt('close active presentation saving no');
quit();
if (!existsSync(DECK)) failures.push('the deck was not exported');
else {
  const z = unzip(new Uint8Array(readFileSync(DECK)));
  const hiddenParts = [...z.keys()].filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k) && /<p:sld [^>]*show="0"/.test(z.get(k)!.text()));
  check(hiddenParts.length === 1, `the saved deck has exactly one hidden slide (Market's): ${hiddenParts.join(', ')}`);
  const tags = [...z.keys()].filter((k) => k.startsWith('ppt/tags/')).map((k) => z.get(k)!.text()).join('');
  check((tags.match(/name="OC_KIND" val="chapter"/g) ?? []).length === 7 && /OC_SPLIT/.test(tags), 'chapter, toc and split tags saved');
}
console.log(`\nTimings:\n  ${timings.join('\n  ')}\nPNGs: ${PNG_DIR}`);
if (failures.length) { console.error(`\n${failures.length} failure(s):\n  ${failures.join('\n  ')}`); process.exit(1); }
console.log('\nOK: agenda live test passed.');
