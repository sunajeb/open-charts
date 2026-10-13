// Agenda: a deck-level model of chapters and the agenda element drawn on every agenda slide
// (docs/USER-GUIDE.md "Agenda"; reference behaviour: the reference product's manual, "Agendas").
//
// - A deck is a sequence of slides. An agenda slide starts a chapter; the chapter's slides are the
//   ones up to the next agenda slide. A table-of-contents slide lists the top-level chapters
//   without highlighting. A split slide starts a new section with its own, independent agenda.
// - Chapters nest by text level (1 = top). The parent of a chapter is the closest earlier chapter
//   of a lower level in the same section.
// - Each agenda slide shows the top-level chapters plus every chapter whose parent is the slide's
//   chapter or one of its ancestors ("subchapters are only shown on chapter slides of the same
//   level and on the slide of the parent chapter").
// - Every heading has one of twelve style states (AGENDA_STATES, in the manual's order); a style
//   (preset + per-state overrides) gives each state its fill, outline, text colour, weight, size.
// - The agenda slide of a chapter whose first subchapter's agenda slide follows it directly (no
//   content of its own) is hidden in the slide show.
import type { Frame, Scene, Style, TextShape } from './types.ts';
import type { Measurer } from './text.ts';
import { resolveColor, themeOf } from './theme.ts';

export interface Chapter { id: string; title: string; level: number }

export const AGENDA_STATES = [
  'current-top', // 1 chapter is current and at top level
  'top-contains-current', // 2 top-level, contains the current chapter
  'top', // 3 top-level, does not contain the current chapter
  'current-sub', // 4 current, not at top level
  'child-of-current-top', // 5 immediate child of the current chapter, which is top-level
  'in-subtree', // 6 sibling of a chapter on the path to the current one, or a child of the current
  'under-current-top', // 7 below the current chapter, which is top-level (not an immediate child)
  'same-top', // 8 below the same top-level chapter as the current one
  'other', // 9 any other chapter
  'toc-top', // 10 table of contents, top level
  'toc-sub', // 11 table of contents, not top level
  'path', // 12 on the path from the top-level chapter to the current one
] as const;
export type AgendaState = typeof AGENDA_STATES[number];

/** How one kind of heading looks. Colours are hex or theme slots; fill/line "" = none. */
export interface HeadingStyle { fill?: string; line?: string; color?: string; bold?: boolean; size?: number }

export const AGENDA_PRESETS = ['flat', 'plain', 'boxed'] as const;
export type AgendaPreset = typeof AGENDA_PRESETS[number];

export interface AgendaOptions {
  /** Predefined style (default "flat"). */
  preset?: AgendaPreset;
  /** Number chapters (1., 2., … per level) instead of bullets. */
  numbering?: boolean;
  /** Font size of top-level headings in pt (default 20); subchapters are 85 % of it. */
  fontSize?: number;
  /** Levels shown on a table-of-contents slide (default 1: top-level chapters only). */
  tocLevels?: number;
  /** Per-state overrides on top of the preset. */
  styles?: Partial<Record<AgendaState, HeadingStyle>>;
}

export const AGENDA_OPTION_KEYS = ['preset', 'numbering', 'fontSize', 'tocLevels', 'styles'] as const;
export const HEADING_STYLE_KEYS = ['fill', 'line', 'color', 'bold', 'size'] as const;

const SUB = 0.85;
const NONE = '';
const MUTED = '7F7F7F';

const PRESETS: Record<AgendaPreset, Record<AgendaState, HeadingStyle>> = {
  flat: {
    'current-top': { fill: 'accent1', color: 'bg1', bold: true },
    'top-contains-current': { fill: 'D9D9D9', color: 'tx1', bold: true },
    top: { fill: 'F2F2F2', color: 'tx1' },
    'current-sub': { fill: 'accent1', color: 'bg1', bold: true },
    'child-of-current-top': { fill: NONE, color: 'tx1' },
    'in-subtree': { fill: NONE, color: 'tx1' },
    'under-current-top': { fill: NONE, color: MUTED },
    'same-top': { fill: NONE, color: MUTED },
    other: { fill: NONE, color: MUTED },
    'toc-top': { fill: NONE, color: 'tx1' },
    'toc-sub': { fill: NONE, color: '595959' },
    path: { fill: 'D9D9D9', color: 'tx1', bold: true },
  },
  plain: {
    'current-top': { fill: NONE, color: 'accent1', bold: true },
    'top-contains-current': { fill: NONE, color: 'tx1', bold: true },
    top: { fill: NONE, color: 'tx1' },
    'current-sub': { fill: NONE, color: 'accent1', bold: true },
    'child-of-current-top': { fill: NONE, color: 'tx1' },
    'in-subtree': { fill: NONE, color: 'tx1' },
    'under-current-top': { fill: NONE, color: MUTED },
    'same-top': { fill: NONE, color: MUTED },
    other: { fill: NONE, color: MUTED },
    'toc-top': { fill: NONE, color: 'tx1' },
    'toc-sub': { fill: NONE, color: '595959' },
    path: { fill: NONE, color: 'tx1', bold: true },
  },
  boxed: {
    'current-top': { fill: 'accent1', line: 'accent1', color: 'bg1', bold: true },
    'top-contains-current': { fill: NONE, line: 'accent1', color: 'tx1', bold: true },
    top: { fill: NONE, line: 'BFBFBF', color: 'tx1' },
    'current-sub': { fill: 'accent1', line: 'accent1', color: 'bg1', bold: true },
    'child-of-current-top': { fill: NONE, line: 'BFBFBF', color: 'tx1' },
    'in-subtree': { fill: NONE, line: 'BFBFBF', color: 'tx1' },
    'under-current-top': { fill: NONE, line: 'D9D9D9', color: MUTED },
    'same-top': { fill: NONE, line: 'D9D9D9', color: MUTED },
    other: { fill: NONE, line: 'D9D9D9', color: MUTED },
    'toc-top': { fill: NONE, line: 'BFBFBF', color: 'tx1' },
    'toc-sub': { fill: NONE, line: 'D9D9D9', color: '595959' },
    path: { fill: NONE, line: 'accent1', color: 'tx1', bold: true },
  },
};

const TOP_STATES = new Set<AgendaState>(['current-top', 'top-contains-current', 'top', 'toc-top']);

/** The effective style of a state: preset, then overrides; size filled in (pt). */
export function headingStyle(state: AgendaState, opts: AgendaOptions = {}): Required<HeadingStyle> {
  const base = PRESETS[opts.preset ?? 'flat'][state];
  const o = opts.styles?.[state] ?? {};
  const size = (opts.fontSize ?? 20) * (TOP_STATES.has(state) ? 1 : SUB);
  return { fill: NONE, line: NONE, color: 'tx1', bold: false, size: +size.toFixed(1), ...base, ...stripUndefined(o) };
}

const stripUndefined = <T extends object>(o: T): Partial<T> => Object.fromEntries(Object.entries(o).filter(([, v]) => v !== undefined)) as Partial<T>;

// ---- Hierarchy -------------------------------------------------------------------------------

/** Levels made consistent: the first chapter is top-level and a level is at most one below the one before. */
export function normalizeLevels(chapters: Chapter[]): Chapter[] {
  let prev = 0;
  return chapters.map((c) => {
    const level = Math.max(1, Math.min(Math.round(c.level) || 1, prev + 1));
    prev = level;
    return level === c.level ? c : { ...c, level };
  });
}

/** Index of each chapter's parent (-1 for top level). */
export function parentsOf(chapters: Chapter[]): number[] {
  const out: number[] = [];
  const stack: number[] = [];
  chapters.forEach((c, i) => {
    while (stack.length && chapters[stack[stack.length - 1]].level >= c.level) stack.pop();
    out.push(stack.length ? stack[stack.length - 1] : -1);
    stack.push(i);
  });
  return out;
}

const ancestors = (parent: number[], i: number) => {
  const out: number[] = [];
  for (let p = parent[i]; p >= 0; p = parent[p]) out.push(p);
  return out;
};
const topOf = (parent: number[], i: number) => { let t = i; while (parent[t] >= 0) t = parent[t]; return t; };
const depthOf = (parent: number[], i: number) => ancestors(parent, i).length + 1;

/** The style state of chapter `i` on the agenda slide of chapter `current` (or a table of contents). */
export function stateOf(chapters: Chapter[], i: number, current: number | 'toc', parent = parentsOf(chapters)): AgendaState {
  const top = parent[i] < 0;
  if (current === 'toc') return top ? 'toc-top' : 'toc-sub';
  const curTop = parent[current] < 0;
  const up = ancestors(parent, current);
  if (i === current) return curTop ? 'current-top' : 'current-sub';
  if (top) return up.includes(i) ? 'top-contains-current' : 'top';
  if (curTop) {
    if (parent[i] === current) return 'child-of-current-top';
    return topOf(parent, i) === current ? 'under-current-top' : 'other';
  }
  if (up.includes(i)) return 'path';
  if (parent[i] === current || up.includes(parent[i])) return 'in-subtree';
  return topOf(parent, i) === topOf(parent, current) ? 'same-top' : 'other';
}

/** Indices of the chapters shown on the agenda slide of `current` (or a table of contents). */
export function visibleChapters(chapters: Chapter[], current: number | 'toc', opts: AgendaOptions = {}, parent = parentsOf(chapters)): number[] {
  if (current === 'toc') return chapters.map((_, i) => i).filter((i) => depthOf(parent, i) <= (opts.tocLevels ?? 1));
  const path = new Set([current, ...ancestors(parent, current)]);
  return chapters.map((_, i) => i).filter((i) => parent[i] < 0 || path.has(parent[i]));
}

/** 1-based position of each chapter among its siblings (for numbering). */
export function numbersOf(chapters: Chapter[], parent = parentsOf(chapters)): number[] {
  const count = new Map<number, number>();
  return chapters.map((_, i) => {
    const n = (count.get(parent[i]) ?? 0) + 1;
    count.set(parent[i], n);
    return n;
  });
}

// ---- Deck structure --------------------------------------------------------------------------

/** What a slide is to the agenda. `chapter` is the id of the chapter an agenda slide starts. */
export interface AgendaSlide { kind?: 'chapter' | 'toc'; chapter?: string; split?: boolean }

export interface SlidePlan {
  section: number;
  /** For agenda slides: the chapter it starts (index into its section's chapters), or "toc". */
  current?: number | 'toc';
  /** Hide in the slide show: a chapter with no slides of its own before its first subchapter. */
  hidden: boolean;
}

export interface DeckPlan {
  /** Chapters of each section, in order (levels normalised). */
  sections: Chapter[][];
  slides: SlidePlan[];
}

/**
 * Sections, chapter order and per-slide state of a deck. `chapters` holds every chapter in deck
 * order, including chapters without an agenda slide (their slide was removed); a slide-less chapter
 * belongs to the section of the chapter after it.
 */
export function planDeck(slides: AgendaSlide[], chapters: Chapter[]): DeckPlan {
  const sectionOfSlide: number[] = [];
  let sec = 0;
  slides.forEach((s, i) => { if (s.split && i > 0) sec++; sectionOfSlide.push(sec); });
  const nSections = sec + 1;
  const slideOf = new Map<string, number>();
  slides.forEach((s, i) => { if (s.kind === 'chapter' && s.chapter && !slideOf.has(s.chapter)) slideOf.set(s.chapter, i); });
  const raw: Chapter[][] = Array.from({ length: nSections }, () => []);
  let pending: Chapter[] = [];
  for (const c of chapters) {
    const at = slideOf.get(c.id);
    if (at === undefined) { pending.push(c); continue; }
    raw[sectionOfSlide[at]].push(...pending, c);
    pending = [];
  }
  if (pending.length) raw[nSections - 1].push(...pending);
  const sections = raw.map(normalizeLevels);
  const index = sections.map((cs) => new Map(cs.map((c, i) => [c.id, i])));
  const parents = sections.map((cs) => parentsOf(cs));
  const plans: SlidePlan[] = slides.map((s, i) => {
    const section = sectionOfSlide[i];
    if (s.kind === 'toc') return { section, current: 'toc', hidden: false };
    if (s.kind !== 'chapter' || !s.chapter || slideOf.get(s.chapter) !== i) return { section, hidden: false };
    const cur = index[section].get(s.chapter)!;
    const next = slides[i + 1];
    const child = cur + 1;
    const hidden = next?.kind === 'chapter' && child < sections[section].length && parents[section][child] === cur &&
      next.chapter === sections[section][child].id && sectionOfSlide[i + 1] === section;
    return { section, current: cur, hidden };
  });
  return { sections, slides: plans };
}

// ---- Layout ----------------------------------------------------------------------------------

export interface AgendaMetrics { w: number; h: number }

const INDENT = 1.2; // per level, × font size
const PAD_X = 0.4; // left/right padding inside a heading box, × size
const PAD_Y = 0.25;
const GAP = 2; // pt between heading boxes

function rowHeight(size: number, m: Measurer) { return m.lineHeight(size) + 2 * PAD_Y * size; }

function bulletText(opts: AgendaOptions, n: number) { return opts.numbering ? `${n}.` : '•'; }

/** Width and the tallest height of a section's agenda element over all its agenda slides. */
export function agendaMetrics(chapters: Chapter[], opts: AgendaOptions, m: Measurer, tocToo = false): AgendaMetrics {
  const parent = parentsOf(chapters);
  const nums = numbersOf(chapters, parent);
  const base = opts.fontSize ?? 20;
  let w = 0;
  chapters.forEach((c, i) => {
    const d = depthOf(parent, i);
    const size = d === 1 ? base : base * SUB;
    const hang = m.width(bulletText(opts, nums[i]), size, true) + size * 0.5;
    w = Math.max(w, (d - 1) * INDENT * base + 2 * PAD_X * size + hang + m.width(c.title, size, true));
  });
  const height = (cur: number | 'toc') => visibleChapters(chapters, cur, opts, parent)
    .reduce((h, i) => h + rowHeight(depthOf(parent, i) === 1 ? base : base * SUB, m) + GAP, -GAP);
  let h = 0;
  chapters.forEach((_, i) => { h = Math.max(h, height(i)); });
  if (tocToo) h = Math.max(h, height('toc'));
  return { w: Math.ceil(Math.max(w, base * 4)), h: Math.max(0, Math.ceil(h)) };
}

/** Default frame: centred on the slide. */
export function defaultAgendaFrame(metrics: AgendaMetrics, slide: { w: number; h: number }): Frame {
  const w = Math.min(metrics.w, slide.w - 48);
  const h = Math.max(metrics.h, 20);
  return { x: +((slide.w - w) / 2).toFixed(2), y: +((slide.h - h) / 2).toFixed(2), w, h };
}

export interface AgendaHeading { id: string; title: string; level: number; state: AgendaState; style: Required<HeadingStyle> }

/**
 * The agenda element of one slide: one text shape per shown chapter (role `chapter:<id>`), top
 * down from the frame's top-left, all as wide as the frame (less their indent).
 */
export function layoutAgenda(chapters: Chapter[], current: number | 'toc', frame: Frame, m: Measurer, opts: AgendaOptions = {}, style?: Partial<Style>): Scene & { headings: AgendaHeading[] } {
  const parent = parentsOf(chapters);
  const nums = numbersOf(chapters, parent);
  const theme = themeOf(style);
  const col = (c: string) => (c ? resolveColor(c, theme) : undefined);
  const base = opts.fontSize ?? 20;
  const shapes: TextShape[] = [];
  const roles: string[] = [];
  const headings: AgendaHeading[] = [];
  let y = frame.y;
  for (const i of visibleChapters(chapters, current, opts, parent)) {
    const c = chapters[i];
    const d = depthOf(parent, i);
    const state = stateOf(chapters, i, current, parent);
    const hs = headingStyle(state, opts);
    const size = hs.size;
    const h = rowHeight(d === 1 ? base : base * SUB, m);
    const x = frame.x + (d - 1) * INDENT * base;
    const at = PAD_X * size;
    const btext = bulletText(opts, nums[i]);
    const hang = m.width(btext, size, true) + size * 0.5;
    shapes.push({
      kind: 'text', x: r2(x), y: r2(y), w: r2(Math.max(size, frame.x + frame.w - x)), h: r2(h), text: c.title, size,
      color: col(hs.color) ?? resolveColor('tx1', theme), bold: hs.bold || undefined, align: 'left', valign: 'middle',
      fill: col(hs.fill), stroke: col(hs.line), level: d,
      bullet: opts.numbering ? { number: nums[i], at: r2(at), hang: r2(hang) } : { char: btext, at: r2(at), hang: r2(hang) },
    });
    roles.push(`chapter:${c.id}`);
    headings.push({ id: c.id, title: c.title, level: d, state, style: hs });
    y += h + GAP;
  }
  return { frame, shapes, roles, hidden: [], headings };
}

const r2 = (v: number) => +v.toFixed(2);

// ---- Editing ---------------------------------------------------------------------------------

/** Slide ranges: the slides that belong to the chapter whose agenda slide is at `start`. */
export function chapterRange(slides: AgendaSlide[], start: number): { start: number; count: number } {
  let end = start + 1;
  while (end < slides.length && slides[end].kind !== 'chapter' && slides[end].kind !== 'toc' && !slides[end].split) end++;
  return { start, count: end - start };
}

/**
 * Moving a chapter up or down moves its slides along: the agenda slide and every slide up to the
 * next agenda slide trade places with the previous (next) chapter's slides in the same section.
 * Returns the move as "take `count` slides at `from` and put them so the first lands at `to`"
 * (0-based, `to` counted after removal), or undefined when it cannot move.
 */
export function planMove(slides: AgendaSlide[], chapterId: string, dir: -1 | 1): { from: number; count: number; to: number } | undefined {
  const at = slides.findIndex((s) => s.kind === 'chapter' && s.chapter === chapterId);
  if (at < 0) return undefined;
  const me = chapterRange(slides, at);
  const sectionStart = (() => { let i = at; while (i > 0 && !slides[i].split) i--; return i; })();
  if (dir < 0) {
    let p = at - 1;
    while (p >= sectionStart && slides[p].kind !== 'chapter') p--;
    if (p < sectionStart || (slides[at].split)) return undefined;
    return { from: at, count: me.count, to: p };
  }
  const n = me.start + me.count;
  if (n >= slides.length || slides[n].kind !== 'chapter' || slides[n].split) return undefined;
  const next = chapterRange(slides, n);
  return { from: at, count: me.count, to: at + next.count };
}

/** Applies a planned move to an array (tests, and the add-in's model before PowerPoint moves slides). */
export function applyMove<T>(items: T[], mv: { from: number; count: number; to: number }): T[] {
  const out = items.slice();
  const taken = out.splice(mv.from, mv.count);
  out.splice(mv.to, 0, ...taken);
  return out;
}

// ---- Stored form -----------------------------------------------------------------------------

/**
 * What every agenda element stores in its OC_AGENDA tag: the whole deck's chapter list (so any
 * agenda slide can rebuild every other), the agenda options, and the headings as drawn on this
 * slide (to tell a rename, a level change or a restyle by the user from what was drawn).
 */
export interface StoredAgenda {
  v: 1;
  /** Bumped by every update; the highest wins when lists disagree (a slide pasted from another deck). */
  rev: number;
  /** Every chapter of the deck in order; `slide: false` for a chapter whose agenda slide was removed. */
  chapters: (Chapter & { slide?: false })[];
  options?: AgendaOptions;
  /** Headings drawn on this slide: id, title, level, state and the colours they were drawn in (hex). */
  shown: { id: string; title: string; level: number; state: AgendaState; fill: string; line: string; color: string; bold: boolean; size: number }[];
}

/** One line per heading, as the add-in's VBA reads the drawn headings back: id, level, text. */
export function headingSignature(headings: { id: string; level: number; title: string }[]): string {
  return headings.map((h) => `${h.id}\t${h.level}\t${h.title}`).join('\n');
}
