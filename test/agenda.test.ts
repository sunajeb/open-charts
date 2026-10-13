// Agenda: the model (hierarchy, 12 style states, visibility, hidden parents, sections, moves), the
// layout, the CLI deck, and the add-in's update logic against a simulated PowerPoint deck that
// applies the engine's answers the way the VBA does (addin/vba/OpenCharts.bas, Agenda section).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createMeasurer } from '../src/fonts.ts';
import {
  AGENDA_STATES, agendaMetrics, applyMove, defaultAgendaFrame, headingStyle, layoutAgenda, normalizeLevels, numbersOf,
  parentsOf, planDeck, planMove, stateOf, visibleChapters, type AgendaSlide, type Chapter, type StoredAgenda,
} from '../src/core/agenda.ts';
import { validateDeck } from '../src/core/index.ts';
import { buildDeck, deckAgenda, type DeckSpec } from '../src/deck.ts';
import { handleRequest, parseRequest } from '../src/addin/engine.ts';
import { parseSnapshot, reconcile } from '../src/addin/agenda.ts';
import { escapeField, parseResponse, SHAPE_COLUMNS, unescapeField } from '../src/addin/format.ts';
import { unzip } from '../src/ooxml/unzip.ts';

const m = createMeasurer();
const ch = (spec: string): Chapter[] => spec.split(' ').map((s, i) => {
  const [title, level] = s.split(':');
  return { id: `c${i + 1}`, title, level: Number(level ?? 1) };
});

// A · A1 · A1a · A1b · A2 · B · B1 · B1a · C
const TREE = ch('A A1:2 A1a:3 A1b:3 A2:2 B B1:2 B1a:3 C');
const idx = (t: string) => TREE.findIndex((c) => c.title === t);

// ---- Model -------------------------------------------------------------------------------------

test('hierarchy: parents by text level, levels normalised', () => {
  assert.deepEqual(parentsOf(TREE).map((p) => (p < 0 ? '-' : TREE[p].title)), ['-', 'A', 'A1', 'A1', 'A', '-', 'B', 'B1', '-']);
  assert.deepEqual(normalizeLevels(ch('X:2 Y:4 Z:1')).map((c) => c.level), [1, 2, 1]);
  assert.deepEqual(numbersOf(TREE), [1, 1, 1, 2, 2, 2, 1, 1, 3]);
});

test('the twelve style states, exactly as the manual defines them', () => {
  const s = (i: string, cur: string | 'toc') => stateOf(TREE, idx(i), cur === 'toc' ? 'toc' : idx(cur));
  assert.equal(s('A', 'A'), 'current-top'); // 1
  assert.equal(s('A', 'A1a'), 'top-contains-current'); // 2
  assert.equal(s('B', 'A1'), 'top'); // 3
  assert.equal(s('A1', 'A1'), 'current-sub'); // 4
  assert.equal(s('A1', 'A'), 'child-of-current-top'); // 5
  assert.equal(s('A2', 'A1a'), 'in-subtree'); // 6: sibling of A1, which is on the path
  assert.equal(s('A1b', 'A1a'), 'in-subtree'); // 6: sibling of the current chapter
  assert.equal(s('A1a', 'A1'), 'in-subtree'); // 6: immediate child of the current chapter
  assert.equal(s('A1a', 'A'), 'under-current-top'); // 7
  assert.equal(s('A1a', 'A2'), 'same-top'); // 8: below A, like the current A2, not in its subtree
  assert.equal(s('B1', 'A1'), 'other'); // 9
  assert.equal(s('A', 'toc'), 'toc-top'); // 10
  assert.equal(s('A1', 'toc'), 'toc-sub'); // 11
  assert.equal(s('A1', 'A1a'), 'path'); // 12
  assert.equal(AGENDA_STATES.length, 12);
});

test('subchapters show on chapter slides of the same level and on the parent chapter\'s slide', () => {
  const shown = (cur: string | 'toc') => visibleChapters(TREE, cur === 'toc' ? 'toc' : idx(cur)).map((i) => TREE[i].title).join(' ');
  assert.equal(shown('A'), 'A A1 A2 B C');
  assert.equal(shown('A1'), 'A A1 A1a A1b A2 B C');
  assert.equal(shown('A1a'), 'A A1 A1a A1b A2 B C');
  assert.equal(shown('B1'), 'A B B1 B1a C');
  assert.equal(shown('C'), 'A B C');
  assert.equal(shown('toc'), 'A B C');
  assert.equal(visibleChapters(TREE, 'toc', { tocLevels: 2 }).length, 6);
});

test('presets and per-state overrides', () => {
  assert.equal(headingStyle('current-top').fill, 'accent1');
  assert.equal(headingStyle('current-top', { preset: 'plain' }).fill, '');
  assert.ok(headingStyle('top', { preset: 'boxed' }).line);
  assert.equal(headingStyle('current-sub', { fontSize: 20 }).size, 17);
  const o = headingStyle('current-top', { styles: { 'current-top': { fill: '00FF00', size: 30 } } });
  assert.deepEqual([o.fill, o.size, o.bold], ['00FF00', 30, true]);
});

const slidesOf = (spec: string): AgendaSlide[] => spec.split(' ').map((t) => {
  if (t === '.') return {};
  if (t === 'T') return { kind: 'toc' };
  if (t === '|') return { split: true };
  return { kind: 'chapter', chapter: t };
});

test('hidden parent: hidden when its first subchapter\'s agenda slide follows directly, shown again once content is between', () => {
  const cs: Chapter[] = [{ id: 'a', title: 'A', level: 1 }, { id: 'a1', title: 'A1', level: 2 }, { id: 'b', title: 'B', level: 1 }, { id: 'b1', title: 'B1', level: 2 }];
  const p = planDeck(slidesOf('a a1 . b . b1 .'), cs);
  assert.deepEqual(p.slides.map((s) => s.hidden), [true, false, false, false, false, false, false]);
  assert.deepEqual(planDeck(slidesOf('a . a1 . b b1'), cs).slides.map((s) => s.hidden), [false, false, false, false, true, false]);
  // A chapter followed by a sibling (not a child) is never hidden.
  assert.deepEqual(planDeck(slidesOf('a b'), [cs[0], cs[2]]).slides.map((s) => s.hidden), [false, false]);
});

test('sections: chapters after a splitter form an independent agenda; slide-less chapters stay in place', () => {
  const cs: Chapter[] = [{ id: 'a', title: 'A', level: 1 }, { id: 'b', title: 'B', level: 1 }, { id: 'x', title: 'X', level: 1 }, { id: 'y', title: 'Y', level: 1 }];
  const p = planDeck([...slidesOf('T a . b .'), { split: true }, ...slidesOf('x . y')], cs);
  assert.deepEqual(p.sections.map((s) => s.map((c) => c.title).join('')), ['AB', 'XY']);
  assert.deepEqual(p.slides.map((s) => s.section), [0, 0, 0, 0, 0, 1, 1, 1, 1]);
  assert.equal(p.slides[0].current, 'toc');
  // A removed parent (no slide) is still listed, before its first child.
  const q = planDeck(slidesOf('a1 .'), [{ id: 'a', title: 'A', level: 1 }, { id: 'a1', title: 'A1', level: 2 }]);
  assert.deepEqual(q.sections[0].map((c) => c.title), ['A', 'A1']);
});

test('moving a chapter moves its slides along, never across a splitter', () => {
  const deck = slidesOf('a . . b . c . | x .');
  const names = ['a', '1', '2', 'b', '3', 'c', '4', '|', 'x', '5'];
  const up = planMove(deck, 'b', -1)!;
  assert.deepEqual(applyMove(names, up), ['b', '3', 'a', '1', '2', 'c', '4', '|', 'x', '5']);
  const down = planMove(deck, 'a', 1)!;
  assert.deepEqual(applyMove(names, down), ['b', '3', 'a', '1', '2', 'c', '4', '|', 'x', '5']);
  // The last chapter before the backup does not take the backup along, and cannot move down.
  assert.equal(planMove(deck, 'c', 1), undefined);
  assert.equal(planMove(deck, 'a', -1), undefined);
  assert.equal(planMove(deck, 'x', -1), undefined);
});

// ---- Layout ------------------------------------------------------------------------------------

test('layout: one native text box per shown chapter, bullets and text levels, same width on every slide', () => {
  const met = agendaMetrics(TREE, {}, m);
  const frame = defaultAgendaFrame(met, { w: 960, h: 540 });
  assert.ok(Math.abs(frame.x + frame.w / 2 - 480) < 1 && Math.abs(frame.y + frame.h / 2 - 270) < 1, 'centred on the slide');
  const a = layoutAgenda(TREE, idx('A1'), frame, m);
  assert.deepEqual(a.roles, ['A', 'A1', 'A1a', 'A1b', 'A2', 'B', 'C'].map((t) => `chapter:c${idx(t) + 1}`));
  const sh = a.shapes as import('../src/core/types.ts').TextShape[];
  assert.deepEqual(sh.map((s) => s.level), [1, 2, 3, 3, 2, 1, 1]);
  assert.ok(sh.every((s) => s.bullet?.char === '•'));
  assert.ok(sh[1].x > sh[0].x && sh[2].x > sh[1].x, 'subchapters are indented');
  assert.ok(sh.every((s, i) => i === 0 || s.y >= sh[i - 1].y + sh[i - 1].h), 'stacked top down');
  assert.ok(sh.every((s) => Math.abs(s.x + s.w - (frame.x + frame.w)) < 0.01), 'right edges line up');
  assert.ok(sh[sh.length - 1].y + sh[sh.length - 1].h <= frame.y + frame.h + 0.5, 'fits the frame (tallest slide)');
  for (const s of sh) assert.ok(m.width(s.text, s.size, true) + (s.bullet!.at + s.bullet!.hang) <= s.w + 0.5, `"${s.text}" fits its box`);
  const n = layoutAgenda(TREE, 'toc', frame, m, { numbering: true });
  assert.deepEqual(n.shapes.map((s) => (s as { bullet?: { number?: number } }).bullet?.number), [1, 2, 3]);
});

// ---- CLI deck ----------------------------------------------------------------------------------

test('CLI: deck spec agenda slides become native agenda slides with tags; hidden parents hidden', () => {
  const deck = JSON.parse(readFileSync('examples/agenda.json', 'utf8')) as DeckSpec;
  assert.deepEqual(validateDeck(deck), []);
  const { pptx, slides } = buildDeck(deck, m);
  assert.equal(slides.length, deck.slides.length);
  const z = unzip(pptx);
  const xml = (n: number) => z.get(`ppt/slides/slide${n}.xml`)!.text();
  const market = deck.slides.findIndex((s) => s.chapter === 'Market') + 1;
  assert.match(xml(market), /<p:sld [^>]*show="0"/, 'Market has no content before Market size: hidden');
  assert.doesNotMatch(xml(market + 1), /show="0"/);
  const tags = [...z.keys()].filter((k) => k.startsWith('ppt/tags/')).map((k) => z.get(k)!.text()).join('\n');
  assert.match(tags, /name="OC_KIND" val="chapter"/);
  assert.match(tags, /name="OC_KIND" val="toc"/);
  assert.match(tags, /name="OC_SPLIT" val="1"/);
  assert.match(tags, /name="OC_AGENDA"/);
  assert.match(tags, /name="OC_ROLE" val="chapter:c\d+"/);
  assert.match(xml(2), /<a:buChar char="•"\/>/);
  assert.match(xml(market), /lvl="1"/, 'subchapters carry their text level');
  // The backup's agenda lists only its own chapters.
  const backup = deck.slides.findIndex((s) => s.chapter === 'Methodology') + 1;
  assert.doesNotMatch(xml(backup), /Executive summary/);
  assert.match(xml(backup), /Data sources/);
  assert.equal(deckAgenda(deck).chapters.length, 9);
});

test('CLI: agenda validation', () => {
  const errs = validateDeck({ agenda: { preset: 'fancy', styles: { nope: {}, top: { fill: 'zz' } } }, slides: [{ chapter: '' }, { chapter: 'A', level: 0 }, { toc: true, chart: { type: 'stacked', categories: ['a'], series: [{ name: 's', values: [1] }] } }, { level: 2 }] });
  assert.deepEqual(errs.map((e) => e.path).sort(), ['/agenda/preset', '/agenda/styles/nope', '/agenda/styles/top/fill', '/slides/0/chapter', '/slides/1/level', '/slides/2/chart', '/slides/3', '/slides/3/level'].sort());
  assert.deepEqual(validateDeck({ slides: [{ title: 'Backup', split: true }] }), []);
});

// ---- Add-in: a simulated PowerPoint deck -------------------------------------------------------

interface FakeHeading { id: string; text: string; level: number; fill: string; line: string; color: string; bold: boolean; size: number }
interface FakeSlide {
  name: string; kind: '' | 'chapter' | 'toc'; chapter: string; split: boolean; hidden: boolean;
  group?: { frame: string; tags: Record<string, string>; headings: FakeHeading[] };
}

/** What the VBA does with the engine's answers, on an array of slides. */
class FakeDeck {
  slides: FakeSlide[];
  current = 0;
  redraws = 0;
  n = 0;
  constructor(count: number) { this.slides = Array.from({ length: count }, (_, i) => this.blank(`s${i + 1}`)); }
  blank(name: string): FakeSlide { return { name, kind: '', chapter: '', split: false, hidden: false }; }

  snapshot(target = '') {
    const lines = [['D', String(this.current + 1), target].map(escapeField).join('\t')];
    this.slides.forEach((s, i) => {
      const g = s.group;
      lines.push(['L', String(i + 1), s.kind, s.chapter, s.split ? '1' : '0', s.hidden ? '1' : '0', g?.frame ?? '', g?.tags.OC_AGENDA ?? '', g?.tags.OC_HASH ?? ''].map(escapeField).join('\t'));
      for (const h of g?.headings ?? []) {
        const style = i === this.current ? [h.fill, h.line, h.color, h.bold ? '1' : '0', String(h.size)] : [];
        lines.push(['H', String(i + 1), h.id, h.text, String(h.level), ...style].map(escapeField).join('\t'));
      }
    });
    return lines.join('\n');
  }

  async run(command: string, target = '') {
    const res = await handleRequest(parseRequest(['agenda', command, '0,0,960,540', '', 'theme=222222,FFFFFF,1F3A5F,EEEEEE,1F3A5F,3E6A9E,7FA3CC,B9CDE5,A6A6A6,D9D9D9', this.snapshot(target)].join('\n')), m);
    return this.apply(res, command);
  }

  /** Applies P records, then structural X records, then (after new/move) runs update as the VBA does. */
  async apply(res: string, command: string): Promise<string[]> {
    const messages: string[] = [];
    const lines = res.split('\n').filter(Boolean).map((l) => l.split('\t').map(unescapeField));
    assert.equal(lines[0][0], 'OC');
    let p: FakeSlide | undefined;
    let pending: { redraw: boolean; frame: string; tags: Record<string, string>; headings: FakeHeading[] } | undefined;
    const flush = () => {
      if (p && pending) {
        if (pending.redraw) { p.group = { frame: pending.frame, tags: pending.tags, headings: pending.headings }; this.redraws++; }
        else Object.assign(p.group!.tags, pending.tags);
      }
      pending = undefined;
    };
    const structural: string[][] = [];
    for (const f of lines.slice(1)) {
      if (f[0] === 'P') {
        flush();
        p = this.slides[Number(f[1]) - 1];
        p.chapter = f[2]; p.hidden = f[3] === '1'; p.kind = (f[5] as FakeSlide['kind']) || p.kind;
        pending = { redraw: f[4] === '1', frame: '', tags: {}, headings: [] };
      } else if (f[0] === 'F' && pending) pending.frame = f[1];
      else if (f[0] === 'T' && pending) pending.tags[f[1]] = f[2];
      else if (f[0] === 'S' && pending) {
        const r = Object.fromEntries(SHAPE_COLUMNS.map((c, i) => [c, f[i + 1] ?? ''])) as Record<string, string>;
        if (r.role.startsWith('chapter:')) {
          pending.headings.push({ id: r.role.slice(8), text: r.text, level: Number(r.level) || 1, fill: r.fill.split('@')[0], line: r.lineColor, color: r.color, bold: r.bold === '1', size: Number(r.size) });
        }
      } else if (f[0] === 'M') messages.push(f[1]);
      else if (f[0] === 'X') structural.push(f);
      else if (f[0] === 'E') throw new Error(f[1]);
    }
    flush();
    // Colours: the engine sends theme names; PowerPoint reports the RGB it shows.
    const theme: Record<string, string> = { accent1: '1F3A5F', bg1: 'FFFFFF', tx1: '222222', lt1: 'FFFFFF', dk1: '222222' };
    for (const s of this.slides) for (const h of s.group?.headings ?? []) for (const k of ['fill', 'line', 'color'] as const) h[k] = theme[h[k]] ?? h[k];
    let after = false;
    for (const x of structural.filter((x) => x[1] !== 'delete')) {
      if (x[1] === 'new') {
        const pos = Number(x[2]) - 1, rm = Number(x[5]) - 1;
        const copy: FakeSlide = rm >= 0 ? structuredClone(this.slides[rm]) : this.blank('');
        Object.assign(copy, { name: `agenda${++this.n}`, kind: x[3], chapter: x[4], split: false, hidden: false });
        this.slides.splice(pos, 0, copy);
        this.current = pos;
        after = true;
      } else if (x[1] === 'move') {
        this.slides = applyMove(this.slides, { from: Number(x[2]) - 1, count: Number(x[3]), to: Number(x[4]) - 1 });
        after = true;
      }
    }
    for (const x of structural.filter((x) => x[1] === 'delete')) this.slides.splice(Number(x[2]) - 1, 1);
    if (after && command !== 'update') messages.push(...(await this.run('update')));
    return messages;
  }

  /** Agenda slides as "title*" (current) lists, for assertions. */
  view() {
    return this.slides.filter((s) => s.kind).map((s) => {
      const cur = s.group?.headings.find((h) => h.id === s.chapter);
      return `${s.hidden ? '(hidden) ' : ''}${s.kind === 'toc' ? 'TOC' : cur?.text ?? '?'}: ${s.group?.headings.map((h) => `${'>'.repeat(h.level - 1)}${h.text}${h.id === s.chapter ? '*' : ''}`).join(' ')}`;
    });
  }

  find(title: string) { return this.slides.findIndex((s) => s.kind === 'chapter' && s.group?.headings.find((h) => h.id === s.chapter)?.text === title); }

  /** The user types over a heading on slide `i`. */
  rename(i: number, id: string, text: string) { this.slides[i].group!.headings.find((h) => h.id === id)!.text = text; }
  chapterIdOf(title: string) { return this.slides[this.find(title)].chapter; }
}

/** Inserts a chapter after slide `after` (0-based) and names it, as a user would. */
async function addChapter(d: FakeDeck, after: number, title: string) {
  d.current = after;
  await d.run('chapter');
  const at = d.current;
  const s = d.slides[at];
  assert.equal(s.kind, 'chapter');
  assert.equal(s.group!.headings.find((h) => h.id === s.chapter)!.text, 'Chapter', 'the new heading reads "Chapter"');
  d.rename(at, s.chapter, title);
  await d.run('update');
  return at;
}

test('add-in: insert, rename, nest, reorder, split, duplicate, delete on a 40-slide deck; every agenda slide follows', async () => {
  const d = new FakeDeck(40);
  // Insert / rename: chapters before slides 1, 9, 17, 25 (each insert shifts the rest by one).
  await addChapter(d, -1 + 0, 'Summary');
  assert.equal(d.slides[0].kind, '', 'Chapter goes after the current slide');
  await addChapter(d, 9, 'Market');
  await addChapter(d, 12, 'Size');
  await addChapter(d, 19, 'Competition');
  await addChapter(d, 28, 'Plan');
  assert.deepEqual(d.view().map((v) => v.split(':')[0]), ['Summary', 'Market', 'Size', 'Competition', 'Plan']);
  for (const v of d.view()) assert.match(v, /Summary\*? Market\*? Size\*? Competition\*? Plan\*?$/, 'every agenda slide lists every chapter');

  // Nest: Size one level down (as PowerPoint's indent button would change its text level).
  d.current = d.find('Size');
  await d.run('indent', d.chapterIdOf('Size'));
  assert.ok(d.view().find((v) => v.startsWith('Size'))!.includes('>Size*'));
  assert.ok(!d.view().find((v) => v.startsWith('Summary'))!.includes('Size'), 'a subchapter is not shown on other chapters\' slides');
  assert.ok(d.view().find((v) => v.includes('Market*'))!.includes('>Size'), 'shown on the parent chapter\'s slide');

  // Rename on any agenda slide: the Plan slide renames Competition; every slide follows.
  const plan = d.find('Plan');
  d.rename(plan, d.chapterIdOf('Competition'), 'Competitors');
  d.current = 0;
  await d.run('update');
  assert.ok(d.view().every((v) => !v.includes('Competition') && v.includes('Competitors')));

  // Nesting via the text level (Shift+Alt+→ in PowerPoint): demote Plan under Competitors.
  const planAt = d.find('Plan');
  d.slides[planAt].group!.headings.find((h) => h.id === d.slides[planAt].chapter)!.level = 2;
  await d.run('update');
  assert.ok(d.view().find((v) => v.startsWith('Plan'))!.includes('>Plan*'));

  // Reorder: Competitors (with its slides) up past Size, then past Market — like moving a line of
  // the agenda up (Shift+Alt+Up). Plan's slide is not part of it: only the chapter's own slides move.
  const before = d.slides.map((s) => s.name);
  const compAt = d.find('Competitors');
  const compSlides = before.slice(compAt, d.find('Plan'));
  d.current = compAt;
  await d.run('up', d.chapterIdOf('Competitors'));
  assert.ok(d.find('Competitors') < d.find('Size') && d.find('Competitors') > d.find('Market'));
  await d.run('up', d.chapterIdOf('Competitors'));
  const marketAt = d.find('Market');
  assert.deepEqual(d.slides.slice(d.find('Competitors'), d.find('Competitors') + compSlides.length).map((s) => s.name), compSlides, 'its slides moved along');
  assert.ok(d.find('Competitors') < marketAt);
  for (const v of d.view()) assert.match(v, /^[^:]*: Summary\*? Competitors\*? Market\*?/, `new order everywhere: ${v}`);

  // Split: a backup section from slide 36 on, with its own chapter.
  d.slides[35].split = true;
  await addChapter(d, 36, 'Appendix');
  const appendix = d.view().find((v) => v.startsWith('Appendix'))!;
  assert.equal(appendix, 'Appendix: Appendix*');
  assert.ok(d.view().filter((v) => !v.startsWith('Appendix')).every((v) => !v.includes('Appendix')), 'independent agendas');

  // Duplicate an agenda slide (slide sorter): a new chapter with the same name.
  const sum = d.find('Summary');
  d.slides.splice(sum + 1, 0, structuredClone(d.slides[sum]));
  await d.run('update');
  assert.equal(d.view().filter((v) => v.startsWith('Summary')).length, 2);
  assert.notEqual(d.slides[sum].chapter, d.slides[sum + 1].chapter, 'the copy got its own chapter id');

  // Delete it again: the chapter disappears from every agenda.
  d.slides.splice(sum + 1, 1);
  await d.run('update');
  assert.ok(d.view().filter((v) => !v.startsWith('Appendix')).every((v) => (v.split(': ')[1].match(/Summary/g) ?? []).length === 1));

  // Restyle: one current top-level heading turned green restyles every current top-level heading.
  const mk = d.find('Market');
  d.current = mk;
  d.slides[mk].group!.headings.find((h) => h.id === d.slides[mk].chapter)!.fill = '00B050';
  await d.run('update');
  for (const s of d.slides.filter((x) => x.kind === 'chapter' && x.group!.headings.find((h) => h.id === x.chapter)!.level === 1)) {
    assert.equal(s.group!.headings.find((h) => h.id === s.chapter)!.fill, '00B050');
  }
  // Nothing changed: an update redraws nothing.
  const r = d.redraws;
  await d.run('update');
  assert.equal(d.redraws, r, 'unchanged agenda slides are not redrawn');
});

test('add-in: hidden parent slides; remove and insert them again', async () => {
  const d = new FakeDeck(10);
  await addChapter(d, 1, 'A');
  await addChapter(d, 2, 'A1'); // directly after A: no content of A's own
  d.current = d.find('A1');
  await d.run('indent', d.chapterIdOf('A1'));
  assert.deepEqual(d.view().map((v) => v.split(':')[0]), ['(hidden) A', 'A1']);
  // Content between A and A1: A is shown again.
  d.slides.splice(d.find('A') + 1, 0, d.blank('content'));
  await d.run('update');
  assert.deepEqual(d.view().map((v) => v.split(':')[0]), ['A', 'A1']);
  d.slides.splice(d.find('A') + 1, 1);
  await d.run('update');
  assert.equal(d.view()[0].split(':')[0], '(hidden) A');

  // Remove hidden agenda slides: A's slide goes, A stays in the agenda.
  const n = d.slides.length;
  await d.run('removeHidden');
  assert.equal(d.slides.length, n - 1);
  assert.deepEqual(d.view(), ['A1: A >A1*']);
  await d.run('update');
  assert.deepEqual(d.view(), ['A1: A >A1*'], 'A survives later updates without a slide');
  // Insert hidden agenda slides: A gets its slide back, before A1.
  await d.run('insertHidden');
  assert.deepEqual(d.view(), ['(hidden) A: A* >A1', 'A1: A >A1*']);
});

test('add-in: table of contents goes before the first chapter slide', async () => {
  const d = new FakeDeck(6);
  await addChapter(d, 1, 'One');
  await addChapter(d, 3, 'Two');
  d.current = 5;
  await d.run('toc');
  const toc = d.slides.findIndex((s) => s.kind === 'toc');
  assert.equal(toc, d.find('One') - 1);
  assert.deepEqual(d.view()[0], 'TOC: One Two');
});

test('add-in: snapshot parsing and reconcile keep the newest stored list', () => {
  const stored = (rev: number, title: string): StoredAgenda => ({ v: 1, rev, chapters: [{ id: 'a', title, level: 1 }], shown: [] });
  const line = (n: number, st: StoredAgenda) => ['L', String(n), 'chapter', 'a', '0', '0', '0,0,10,10', JSON.stringify(st), 'h'].map(escapeField).join('\t');
  const snap = parseSnapshot(['D\t1\t', line(1, stored(3, 'New')), line(2, stored(2, 'Old'))].join('\n'));
  assert.equal(snap.slides.length, 2);
  const model = reconcile(snap, () => 'fresh');
  assert.deepEqual(model.chapters.map((c) => c.title), ['New', 'New']);
  assert.equal(model.slides[1].chapter, 'fresh', 'the duplicate gets a new id');
  assert.equal(model.rev, 4);
});

test('add-in response: a new chapter slide, then S records with the heading\'s text level and bullet', async () => {
  const d = new FakeDeck(2);
  d.current = 0;
  const res = await handleRequest(parseRequest(['agenda', 'chapter', '0,0,960,540', '', d.snapshot()].join('\n')), m);
  assert.match(res, /^X\tnew\t2\tchapter\tc[0-9a-f]+\t0$/m, 'no agenda slide to copy yet: roleModel 0');
  d.slides.splice(1, 0, { name: 'new', kind: 'chapter', chapter: 'cnew', split: false, hidden: false });
  const upd = parseResponse(await handleRequest(parseRequest(['agenda', 'update', '0,0,960,540', '', d.snapshot()].join('\n')), m));
  assert.equal(upd.errors.length, 0);
  assert.deepEqual(upd.shapes.map((x) => [x.role, x.text, x.level, x.bullet]), [['chapter:cnew', 'Chapter', '1', 'u8226']]);
  assert.ok(upd.tags.some(([k]) => k === 'OC_AGENDA') && !upd.tags.some(([k]) => k === 'OC_V'), 'an agenda is not a chart');
  assert.equal(parseResponse(await handleRequest(parseRequest(['agenda', 'bogus', '0,0,960,540', '', ''].join('\n')), m)).messages.length, 1);
});

test('ribbon: an Agenda group in the open-charts tab; every agenda button reaches AgendaCommand', async () => {
  const { buildAddin, AGENDA_BUTTONS, AGENDA_MENU } = await import('../src/addin/package.ts');
  const { ribbon, source } = buildAddin();
  assert.match(ribbon, /<tab id="tabOpenCharts"[\s\S]*<group id="grpCharts"[\s\S]*<group id="grpAgenda" label="Agenda"[\s\S]*<\/tab>/);
  const ids = [...AGENDA_BUTTONS.map((b) => b[0]), ...AGENDA_MENU.map((b) => b[0])];
  assert.deepEqual(ids.sort(), ['Chapter', 'Down', 'Indent', 'InsertHidden', 'Join', 'Outdent', 'Placeholder', 'RemoveHidden', 'Split', 'Toc', 'Up', 'Update'].sort());
  for (const id of ids) {
    assert.match(ribbon, new RegExp(`id="btnAgenda${id}"[^>]*onAction=`));
    if (id !== 'Update') assert.match(source, new RegExp(`Case "${id}"|, "${id}"`), `VBA handles ${id}`);
  }
  assert.match(source, /ElseIf Left\$\(id, 9\) = "btnAgenda" Then\s+AgendaCommand Mid\$\(id, 10\)/);
  assert.match(source, /CheckResized\s+AgendaAutoCheck/, 'selection changes run the automatic agenda update');
});
