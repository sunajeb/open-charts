// The agenda on a deck of 30 content slides, through the add-in's own entry points (the Dispatch the
// ribbon buttons use, and the selection-change handler the automatic update runs in): insert and
// rename chapters, nest, outdent, rename on another slide, move a chapter, split off a backup,
// duplicate / delete / move slides in the sorter, remove and insert hidden agenda slides, a table of
// contents. After every step every agenda slide is read back and checked; then the exported deck.
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { createMeasurer } from '../../../src/fonts.ts';
import { unzip } from '../../../src/ooxml/unzip.ts';
import type { Ctx } from '../scenarios.ts';

interface Row { n: number; kind: string; hidden: boolean; split: boolean; headings: string; current: string }

export async function run(ctx: Ctx) {
  const { pp, check } = ctx;
  const DECK = join(ctx.out, 'agenda.pptx'), PNG_DIR = join(ctx.out, 'png');
  const macro = (proc: string, ...args: string[]) => pp.macro(proc, ...args);

  const dump = (): Row[] => {
    const out = macro('AgendaDump', '0');
    const line = out.split('\n').reverse().find((l) => l.includes('DUMP '));
    if (!line) throw new Error(`AgendaDump wrote no DUMP line: ${out.split('\n').pop()}`);
    return line.slice(line.indexOf('DUMP ') + 5).split(' ;; ').filter((s) => s.startsWith('SLIDE')).map((s) => {
      const [head, headings = ''] = s.split(':');
      const [, n, kind, ...flags] = head.trim().split(' ');
      const cur = headings.trim().split(' ').find((h) => h.endsWith('*')) ?? '';
      return { n: Number(n), kind, hidden: flags.includes('hidden'), split: flags.includes('split'), headings: headings.trim(), current: cur.replace(/^>+/, '').replace(/\*$/, '') };
    });
  };
  const agendaRows = (rows = dump()) => rows.filter((r) => r.kind !== '-');
  const slideOf = (title: string, rows = dump()) => rows.find((r) => r.kind === 'chapter' && r.current === title)?.n ?? -1;
  const show = (rows: Row[]) => agendaRows(rows).map((r) => `  ${r.n}${r.hidden ? ' (hidden)' : ''}: ${r.headings}`).join('\n');
  const run = (id: string) => {
    const r = pp.run(id);
    ctx.time(id, r.ms);
    if (!/ OK \(/.test(r.result)) ctx.note(`${id}: ${r.result || r.log}`);
    return r.result;
  };
  /** Words separated by spaces in headings: multi-word titles use "_" here and in the deck. */
  const chapter = (afterSlide: number, title: string) => {
    macro('AgendaSelect', String(afterSlide), '');
    run('btnAgendaChapter');
    const n = dump().find((r) => r.kind === 'chapter' && r.current === 'Chapter')?.n ?? -1;
    check(n === afterSlide + 1, `Chapter after slide ${afterSlide} inserts agenda slide ${afterSlide + 1} reading "Chapter" (got ${n})`);
    macro('AgendaSelect', String(n), '');
    macro('AgendaEdit', String(n), 'Chapter', title, '0');
    macro('AgendaSorter', 'none', '', ''); // the click elsewhere: automatic update
    return n;
  };
  const allShow = (pred: (r: Row) => boolean, what: string, rows = dump()) => {
    const bad = agendaRows(rows).filter((r) => !pred(r));
    check(!bad.length, `${what}${bad.length ? `\n${show(bad)}` : ''}`);
  };

  // 1. A deck of 30 content slides (one AppleScript call).
  pp.newPresentation();
  pp.tell(`repeat with i from 1 to 30
  set s to make new slide at end of active presentation with properties {layout:slide layout title only}
  set content of text range of text frame of shape 1 of s to "Content " & i
end repeat
go to slide view of document window 1 number 1`, 60);

  // 2. Insert and rename chapters.
  chapter(1, 'Introduction');
  chapter(6, 'Market');
  const size = chapter(7, 'Market_size');
  macro('AgendaSelect', String(size), '');
  macro('AgendaEdit', String(size), 'Market_size', '', '2');
  macro('AgendaSorter', 'none', '', '');
  let rows = dump();
  check(!!rows.find((r) => r.current === 'Market')?.hidden, 'nesting Market_size directly after Market hides Market\'s agenda slide');
  check(rows.find((r) => r.current === 'Market_size')?.headings === 'Introduction Market >Market_size*', `Market_size's slide: path and the subchapter, nothing else\n${show(rows)}`);
  chapter(12, 'Growth');
  const comp = chapter(18, 'Competition');
  rows = dump();
  check(!!rows.find((r) => r.current === 'Competition')?.headings.includes('>Competition*'), 'a new chapter takes the level of the one before it');
  macro('AgendaSelect', String(comp), '');
  run('btnAgendaOutdent');
  chapter(25, 'Plan');
  rows = dump();
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
  check(rows.find((r) => r.current === 'Backup_data')?.headings.trim() === 'Backup_data*', 'the backup has its own agenda');
  allShow((r) => r.current === 'Backup_data' || !r.headings.includes('Backup_data'), 'the main agenda does not list the backup chapter', rows);

  // 6. Slide sorter: duplicate, delete, move.
  macro('AgendaSorter', 'duplicate', '2', '');
  rows = dump();
  check(agendaRows(rows).filter((r) => r.kind === 'chapter' && r.current === 'Introduction').length === 2 && agendaRows(rows).every((r) => r.current === 'Backup_data' || (r.headings.match(/Introduction/g) ?? []).length === 2), 'a duplicated agenda slide is a new chapter on every agenda slide');
  macro('AgendaSorter', 'delete', '3', '');
  allShow((r) => r.current === 'Backup_data' || (r.headings.match(/Introduction/g) ?? []).length === 1, 'deleting it removes the chapter everywhere');
  const market = slideOf('Market');
  macro('AgendaSorter', 'move', String(market + 3), String(market + 1)); // a content slide between Market and Market_size
  check(!dump().find((r) => r.current === 'Market')?.hidden, 'content between Market and its first subchapter: Market is shown again');
  macro('AgendaSorter', 'move', String(market + 1), String(market + 3));
  check(!!dump().find((r) => r.current === 'Market')?.hidden, 'and hidden again once that content is gone');

  // 7. Remove and insert hidden agenda slides.
  const before = dump().length;
  run('btnAgendaRemoveHidden');
  rows = dump();
  check(rows.length === before - 1 && slideOf('Market', rows) < 0, 'Remove Hidden Agenda Slides deletes Market\'s slide');
  allShow((r) => r.headings.includes('Market ') || r.current === 'Backup_data', 'Market stays in every agenda', rows);
  run('btnAgendaInsertHidden');
  rows = dump();
  check(rows.length === before && slideOf('Market', rows) === slideOf('Market_size', rows) - 1 && !!rows.find((r) => r.current === 'Market')?.hidden, 'Insert Hidden Agenda Slides brings it back, hidden, before Market_size');

  // 8. Table of contents before the first chapter slide.
  macro('AgendaSelect', '20', '');
  run('btnAgendaToc');
  rows = dump();
  const toc = rows.find((r) => r.kind === 'toc');
  check(!!toc && toc.n === slideOf('Introduction', rows) - 1 && toc.headings.trim() === 'Introduction Market Plan Competitors', `table of contents before the first chapter: ${toc?.n}: ${toc?.headings}`);

  // 9. Nothing changed: Update agenda redraws nothing.
  run('btnAgendaUpdate');
  ctx.note(`agenda at the end:\n${show(dump())}`);

  // 10. Export and read back.
  pp.exportCopy(DECK, 'pptx');
  pp.exportCopy(PNG_DIR, 'png');
  if (!check(existsSync(DECK), 'the deck was exported')) return;
  const z = unzip(new Uint8Array(readFileSync(DECK)));
  const hiddenParts = [...z.keys()].filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k) && /<p:sld [^>]*show="0"/.test(z.get(k)!.text()));
  check(hiddenParts.length === 1, `the saved deck has exactly one hidden slide (Market's): ${hiddenParts.join(', ')}`);
  const tags = [...z.keys()].filter((k) => k.startsWith('ppt/tags/')).map((k) => z.get(k)!.text()).join('');
  check((tags.match(/name="OC_KIND" val="chapter"/g) ?? []).length === 7 && /OC_SPLIT/.test(tags), 'chapter, toc and split tags saved');
  // As PowerPoint saved them: every heading box holds its text (indent + bold text width), and the
  // agenda elements of the main section all sit at the same place.
  const m = createMeasurer();
  const EMU = 12700;
  const clipped: string[] = [];
  const origins = new Set<string>();
  for (const k of [...z.keys()].filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k))) {
    const xml = z.get(k)!.text();
    if (!/name="Agenda"/.test(xml)) continue;
    const grp = xml.match(/<p:grpSp><p:nvGrpSpPr><p:cNvPr id="\d+" name="Agenda"[\s\S]*?<a:off x="(\d+)" y="(\d+)"/);
    if (grp && !/Backup_data/.test(xml)) origins.add(`${Math.round(+grp[1] / EMU)},${Math.round(+grp[2] / EMU)}`);
    for (const sp of xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) ?? []) {
      const marL = sp.match(/marL="(\d+)"/), cx = sp.match(/<a:ext cx="(\d+)"/), sz = sp.match(/sz="(\d+)"/), t = sp.match(/<a:t>([^<]*)<\/a:t>/);
      if (!marL || !cx || !sz || !t) continue;
      const need = +marL[1] / EMU + m.width(t[1], +sz[1] / 100, true);
      if (need > +cx[1] / EMU + 0.5) clipped.push(`${k}: "${t[1]}" needs ${need.toFixed(1)} pt, box ${(+cx[1] / EMU).toFixed(1)} pt`);
    }
  }
  check(!clipped.length, `no heading is clipped${clipped.length ? `:\n  ${clipped.join('\n  ')}` : ''}`);
  check(origins.size === 1, `the main section's agendas all sit at one place: ${[...origins].join(' ')}`);
}
