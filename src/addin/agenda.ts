// Agenda in the PowerPoint add-in: VBA sends a snapshot of the deck (what each slide is to the agenda,
// the agenda elements and the headings as they are now), this works out the chapters — including
// renames, level changes and restyles the user made on any agenda slide, slides added, duplicated,
// reordered or deleted in the slide sorter — and answers what VBA must change: per agenda slide the
// new element (or only its tags when it looks the same), its chapter id and hidden state; and for
// the structural commands, the slides to insert, move or delete.
//
// Request (engine.ts parseRequest): op "agenda", arg = command, frame = "0,0,slideW,slideH",
// bounds = the agenda placeholder of the "open-charts agenda" layout ("" if none), theme=…, then:
//   D <tab> currentSlide <tab> targetChapter         slide numbers are 1-based; target: the chapter
//                                                    of the selected heading, else ""
//   L <tab> n <tab> kind <tab> chapter <tab> split <tab> hidden <tab> bounds <tab> OC_AGENDA <tab> OC_HASH <tab> OC_FRAME
//                                                    one per slide; kind "chapter" | "toc" | ""
//   H <tab> n <tab> chapter <tab> text <tab> level [<tab> fill <tab> line <tab> color <tab> bold <tab> size]
//                                                    the headings of agenda slide n as they are now;
//                                                    the style columns for the current slide only
// Commands: update | chapter | toc | up | down | indent | outdent | removeHidden | insertHidden.
//
// Answer (format.ts records):
//   P <tab> n <tab> chapter <tab> hidden <tab> redraw <tab> kind   then, if redraw, G/F/T/S records of
//                                                    the new element, else only its T records
//   X <tab> new <tab> position <tab> kind <tab> chapter <tab> roleModel   insert an agenda slide so it
//                                                    becomes slide `position` (copy of slide roleModel,
//                                                    or a new slide when 0); in order, positions final
//   X <tab> move <tab> from <tab> count <tab> to     take `count` slides from `from`, so the first
//                                                    becomes slide `to`
//   X <tab> delete <tab> n                           delete slide n (after the P records are applied)
//   M <tab> message
import { randomBytes } from 'node:crypto';
import type { Frame, Measurer, ThemeColors } from '../core/index.ts';
import { fitAgenda, planDeck, planMove, chapterRange, normalizeLevels,
  type AgendaOptions, type AgendaSlide, type AgendaState, type Chapter, type HeadingStyle, type StoredAgenda } from '../core/agenda.ts';
import { agendaContent } from '../agenda.ts';
import { shapeName } from '../render/pptx.ts';
import { parseFrame } from '../render/tags.ts';
import { FORMAT_VERSION, num, record, shapeRecord, unescapeField } from './format.ts';

export interface SnapshotHeading { id: string; text: string; level: number; style?: { fill: string; line: string; color: string; bold: boolean; size: number } }
export interface SnapshotSlide {
  kind: '' | 'chapter' | 'toc'; chapter: string; split: boolean; hidden: boolean;
  bounds?: Frame; stored?: StoredAgenda; hash: string; headings: SnapshotHeading[];
  /** The element's OC_FRAME: where it was drawn (bounds elsewhere = the user moved it). */
  frame?: Frame;
}
export interface Snapshot { current: number; target: string; slides: SnapshotSlide[] }

export const NEW_CHAPTER_TITLE = 'Chapter';

/** Parses the D/L/H lines (1-based slide numbers become 0-based indices). */
export function parseSnapshot(text: string): Snapshot {
  const snap: Snapshot = { current: 0, target: '', slides: [] };
  for (const line of text.replace(/\r\n?/g, '\n').split('\n')) {
    if (!line) continue;
    const f = line.split('\t').map(unescapeField);
    if (f[0] === 'D') { snap.current = Math.max(0, Number(f[1]) - 1); snap.target = f[2] ?? ''; }
    else if (f[0] === 'L') {
      const i = Number(f[1]) - 1;
      let stored: StoredAgenda | undefined;
      try { stored = f[7] ? JSON.parse(f[7]) : undefined; } catch { stored = undefined; }
      snap.slides[i] = {
        kind: f[2] === 'chapter' || f[2] === 'toc' ? f[2] : '', chapter: f[3] ?? '', split: f[4] === '1', hidden: f[5] === '1',
        bounds: parseFrame(f[6]), stored: stored?.v === 1 && Array.isArray(stored.chapters) ? stored : undefined, hash: f[8] ?? '', headings: [],
        frame: parseFrame(f[9]),
      };
    } else if (f[0] === 'H') {
      const s = snap.slides[Number(f[1]) - 1];
      if (!s) continue;
      const h: SnapshotHeading = { id: f[2], text: (f[3] ?? '').replace(/[\r\n\v]+/g, ' ').trim(), level: Number(f[4]) || 1 };
      if (f.length >= 10) h.style = { fill: f[5].toUpperCase(), line: f[6].toUpperCase(), color: f[7].toUpperCase(), bold: f[8] === '1', size: Number(f[9]) || 0 };
      s.headings.push(h);
    }
  }
  for (let i = 0; i < snap.slides.length; i++) snap.slides[i] ??= { kind: '', chapter: '', split: false, hidden: false, hash: '', headings: [] };
  return snap;
}

export interface Model {
  /** Every chapter, deck order; `slide: false` when its agenda slide was removed. */
  chapters: StoredAgenda['chapters'];
  /** Slides as the agenda sees them (chapter ids made unique). */
  slides: AgendaSlide[];
  options: AgendaOptions;
  rev: number;
}

const newId = () => `c${randomBytes(4).toString('hex')}`;

/**
 * The deck's chapters now. The chapter list stored on the agenda elements (the newest) is the base;
 * the agenda slides decide order and existence (moved, duplicated or deleted slides); headings that
 * differ from what was drawn are the user's edits (text = rename, text level = nesting, format =
 * restyle of every heading of that state).
 */
export function reconcile(snap: Snapshot, makeId = newId): Model {
  let base: StoredAgenda | undefined;
  snap.slides.forEach((s, i) => {
    if (!s.stored) return;
    if (!base || s.stored.rev > base.rev || (s.stored.rev === base.rev && i === snap.current)) base = s.stored;
  });
  const info = new Map<string, { title: string; level: number }>();
  for (const s of snap.slides) for (const c of s.stored?.chapters ?? []) if (!info.has(c.id)) info.set(c.id, { title: c.title, level: c.level });
  for (const c of base?.chapters ?? []) info.set(c.id, { title: c.title, level: c.level });
  const options: AgendaOptions = structuredClone(base?.options ?? {});

  // Edits on the agenda slides; the current slide's win.
  const order = snap.slides.map((_, i) => i).filter((i) => i !== snap.current).concat(snap.slides[snap.current] ? [snap.current] : []);
  for (const i of order) {
    const s = snap.slides[i];
    if (!s.stored) continue;
    const drawn = new Map(s.stored.shown.map((h) => [h.id, h]));
    for (const h of s.headings) {
      const d = drawn.get(h.id);
      const c = info.get(h.id);
      if (!d || !c) continue;
      if (h.text && h.text !== d.title) c.title = h.text;
      if (h.level !== d.level) c.level = Math.max(1, c.level + (h.level - d.level));
      if (h.style) restyle(options, d.state, d, h.style);
    }
  }

  // Order and existence from the slides.
  const seen = new Set<string>();
  const slides: AgendaSlide[] = snap.slides.map((s) => {
    const split = s.split ? { split: true } : {};
    if (s.kind === 'toc') return { kind: 'toc', ...split };
    if (s.kind !== 'chapter') return split;
    let id = s.chapter;
    if (!id || seen.has(id)) {
      // A duplicated agenda slide (same id twice) or one without an id: a new chapter like it.
      const like = id ? info.get(id) : undefined;
      id = makeId();
      info.set(id, { title: like?.title ?? NEW_CHAPTER_TITLE, level: like?.level ?? 0 });
    }
    seen.add(id);
    return { kind: 'chapter', chapter: id, ...split };
  });
  const onSlides = slides.filter((s) => s.kind === 'chapter').map((s) => s.chapter!);
  const baseList = base?.chapters ?? [];
  // Chapters without a slide stay just before the chapter that followed them.
  const slideless = new Map<string, Chapter[]>();
  let run: Chapter[] = [];
  for (const c of baseList) {
    if (c.slide === false && !seen.has(c.id)) { run.push({ id: c.id, title: info.get(c.id)!.title, level: info.get(c.id)!.level }); continue; }
    if (run.length) { slideless.set(c.id, run); run = []; }
  }
  const chapters: StoredAgenda['chapters'] = [];
  for (const id of onSlides) {
    for (const c of slideless.get(id) ?? []) chapters.push({ ...c, slide: false });
    const c = info.get(id) ?? { title: NEW_CHAPTER_TITLE, level: 0 };
    // A new chapter: at the level of the chapter before it.
    const level = c.level || (chapters.length ? chapters[chapters.length - 1].level : 1);
    chapters.push({ id, title: c.title, level });
  }
  return { chapters: normalizeKeepSlide(chapters), slides, options, rev: (base?.rev ?? 0) + 1 };
}

function normalizeKeepSlide(cs: StoredAgenda['chapters']): StoredAgenda['chapters'] {
  const n = normalizeLevels(cs);
  return n.map((c, i) => (cs[i].slide === false ? { id: c.id, title: c.title, level: c.level, slide: false } : { id: c.id, title: c.title, level: c.level }));
}

/** A heading formatted differently from how it was drawn restyles every heading of its state. */
function restyle(options: AgendaOptions, state: AgendaState, drawn: StoredAgenda['shown'][number], now: NonNullable<SnapshotHeading['style']>) {
  const change: HeadingStyle = {};
  if (now.fill !== drawn.fill) change.fill = now.fill;
  if (now.line !== drawn.line) change.line = now.line;
  if (now.color && now.color !== drawn.color) change.color = now.color;
  if (now.bold !== drawn.bold) change.bold = now.bold;
  if (now.size > 0 && Math.abs(now.size - drawn.size) > 0.25) change.size = Math.round(now.size * 2) / 2;
  if (!Object.keys(change).length) return;
  options.styles = { ...options.styles, [state]: { ...options.styles?.[state], ...change } };
}

export interface AgendaRequest { command: string; slide: { w: number; h: number }; placeholder?: Frame; theme?: Partial<ThemeColors>; snapshot: string }

const header = () => record('OC', FORMAT_VERSION);

/** The answer for one agenda command. */
export function handleAgenda(req: AgendaRequest, m: Measurer, makeId = newId): string {
  const snap = parseSnapshot(req.snapshot);
  const model = reconcile(snap, makeId);
  const out = [header()];
  const target = snap.target || chapterAt(model.slides, snap.current);
  switch (req.command) {
    case 'chapter':
    case 'toc': {
      // A new chapter slide goes after the current slide, or after the selected chapter's slides;
      // a table of contents before the first chapter slide (after the current slide if none).
      let pos: number;
      if (req.command === 'toc') {
        const first = model.slides.findIndex((s) => s.kind === 'chapter');
        pos = first >= 0 ? first : snap.current + 1;
      } else if (snap.target) {
        const at = model.slides.findIndex((s) => s.chapter === snap.target);
        pos = at >= 0 ? at + chapterRange(model.slides, at).count : snap.current + 1;
      } else pos = snap.current + 1;
      pos = Math.min(pos, model.slides.length);
      const model2 = roleModel(model.slides, pos);
      out.push(record('X', 'new', String(pos + 1), req.command, req.command === 'chapter' ? makeId() : '', String(model2 + 1)));
      return out.join('\n') + '\n';
    }
    case 'up':
    case 'down': {
      if (!target) return message(out, 'Select a chapter (an agenda slide or a heading on it) first.');
      const mv = planMove(model.slides, target, req.command === 'up' ? -1 : 1);
      if (!mv) return message(out, `This chapter is already the ${req.command === 'up' ? 'first' : 'last'} of its agenda.`);
      out.push(record('X', 'move', String(mv.from + 1), String(mv.count), String(mv.to + 1)));
      return out.join('\n') + '\n';
    }
    case 'indent':
    case 'outdent': {
      if (!target) return message(out, 'Select a chapter (an agenda slide or a heading on it) first.');
      const c = model.chapters.find((x) => x.id === target);
      if (c) c.level = Math.max(1, c.level + (req.command === 'indent' ? 1 : -1));
      model.chapters = normalizeKeepSlide(model.chapters);
      break;
    }
    case 'insertHidden': {
      const at = new Map(model.slides.map((s, i) => [s.chapter, i] as const));
      let shift = 0;
      let pending: string[] = [];
      for (const c of model.chapters) {
        if (c.slide === false) { pending.push(c.id); continue; }
        const i = at.get(c.id);
        if (i === undefined) continue;
        for (const id of pending) { out.push(record('X', 'new', String(i + shift + 1), 'chapter', id, String(i + shift + 1))); shift++; }
        pending = [];
      }
      if (out.length === 1) return message(out, 'There are no removed agenda slides to insert.');
      return out.join('\n') + '\n';
    }
    case 'update':
    case 'removeHidden':
      break;
    default:
      return message(out, `This add-in and its engine do not match (agenda "${req.command}").`);
  }

  let deletes: number[] = [];
  if (req.command === 'removeHidden') {
    const plan = planDeck(model.slides, model.chapters);
    deletes = plan.slides.map((p, i) => (p.hidden ? i : -1)).filter((i) => i >= 0);
    const gone = new Set(deletes.map((i) => model.slides[i].chapter));
    model.chapters = model.chapters.map((c) => (gone.has(c.id) ? { ...c, slide: false as const } : c));
    if (!deletes.length) out.push(record('M', 'There are no hidden agenda slides.'));
  }
  // Removed slides no longer count: plan the rest as they will be.
  const kept = model.slides.map((s, i) => (deletes.includes(i) ? undefined : s));
  const keptSlides = kept.filter((s): s is AgendaSlide => !!s);
  const keptIndex = kept.map((s, i) => (s ? i : -1)).filter((i) => i >= 0);
  const plan = planDeck(keptSlides, model.chapters);
  const tocIn = (k: number) => plan.slides.some((p) => p.section === k && p.current === 'toc');
  const font = 'Arial';
  const style = req.theme ? { theme: req.theme } : undefined;
  // Where the user put an agenda element by hand (it is no longer where it was drawn) it stays; every
  // other agenda slide of a section goes where its first moved one is, else to the layout's agenda
  // placeholder, else the body area, so a section's agendas line up.
  const moved = (s: SnapshotSlide) => !!(s.bounds && s.frame && (Math.abs(s.bounds.x - s.frame.x) > 1 || Math.abs(s.bounds.y - s.frame.y) > 1));
  const anchor = new Map<number, Frame>();
  plan.slides.forEach((p, j) => {
    const s = snap.slides[keptIndex[j]];
    if (p.current !== undefined && moved(s) && !anchor.has(p.section)) anchor.set(p.section, { ...s.bounds!, w: s.frame!.w, h: s.frame!.h });
  });
  const fits = new Map<string, ReturnType<typeof fitAgenda>>();
  const fitFor = (k: number, at?: Frame) => {
    const key = `${k}:${at ? [at.x, at.y, at.w].join(',') : ''}`;
    if (!fits.has(key)) fits.set(key, fitAgenda(plan.sections[k], model.options, m, req.slide, tocIn(k), at));
    return fits.get(key)!;
  };
  plan.slides.forEach((p, j) => {
    const i = keptIndex[j];
    const s = snap.slides[i];
    const a = keptSlides[j];
    if (p.current === undefined) return;
    const own = moved(s) ? { ...s.bounds!, w: s.frame!.w, h: s.frame!.h } : undefined;
    const fit = fitFor(p.section, own ?? anchor.get(p.section) ?? req.placeholder);
    const frame = fit.frame;
    const content = agendaContent({ section: plan.sections[p.section], current: p.current, frame, all: model.chapters, options: model.options, drawOptions: fit.options, rev: model.rev }, m, style);
    const redraw = !s.bounds || content.hash !== s.hash;
    out.push(record('P', String(i + 1), a.chapter ?? '', p.hidden ? '1' : '0', redraw ? '1' : '0', a.kind ?? ''));
    if (redraw) {
      out.push(record('G', content.name, content.altText), record('F', [frame.x, frame.y, frame.w, frame.h].map(num).join(',')));
      for (const [k, v] of content.tags ?? []) out.push(record('T', k, v));
      content.scene.shapes.forEach((sh, k) => {
        const role = content.scene.roles?.[k] ?? '';
        out.push(shapeRecord(sh, font, role, role.startsWith('chapter:') ? `Chapter ${role.slice(8)}` : shapeName(sh, k + 1, role || undefined)));
      });
    } else {
      for (const [k, v] of content.tags ?? []) out.push(record('T', k, v));
    }
  });
  for (const i of deletes.sort((x, y) => y - x)) out.push(record('X', 'delete', String(i + 1)));
  return out.join('\n') + '\n';
}

function message(out: string[], text: string) { out.push(record('M', text)); return out.join('\n') + '\n'; }

/** The chapter slide `i` belongs to (its own, or the closest agenda slide before it). */
export function chapterAt(slides: AgendaSlide[], i: number): string {
  for (let k = Math.min(i, slides.length - 1); k >= 0; k--) {
    if (slides[k].kind === 'chapter') return slides[k].chapter!;
    if (slides[k].split) break;
  }
  return '';
}

/** The slide a new agenda slide at `pos` copies: the closest agenda slide before it, else after it; -1 for none. */
export function roleModel(slides: AgendaSlide[], pos: number): number {
  for (let k = pos - 1; k >= 0; k--) if (slides[k].kind) return k;
  for (let k = pos; k < slides.length; k++) if (slides[k].kind) return k;
  return -1;
}
