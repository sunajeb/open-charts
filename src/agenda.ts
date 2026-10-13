// The agenda element as a host draws it: the laid-out headings (src/core/agenda.ts) plus the tags
// that keep it editable — shared by the CLI deck writer and the PowerPoint add-in.
//
// Group tags: OC_AGENDA (StoredAgenda JSON), OC_FRAME ("x,y,w,h"), OC_SIG (the headings as VBA
// reads them back: id, level, text per line) and OC_HASH (what was drawn, so an update skips slides
// that did not change). Each heading carries OC_ROLE "chapter:<id>". Slide tags: OC_KIND
// ("chapter" | "toc"), OC_CHAPTER (the chapter an agenda slide starts) and OC_SPLIT ("1": the
// slide starts a new agenda section). No OC_V: an agenda is not a chart.
import type { Frame, Measurer, Style } from './core/index.ts';
import { colorHex, themeOf, resolveColor } from './core/theme.ts';
import { headingSignature, layoutAgenda, type AgendaOptions, type Chapter, type StoredAgenda } from './core/agenda.ts';
import type { SceneContent } from './render/pptx.ts';
import type { Tags } from './render/tags.ts';

export const AGENDA_TAG = 'OC_AGENDA';
export const AGENDA_GROUP_NAME = 'Agenda';

const num = (v: number) => String(+v.toFixed(2));
const hexOf = (c: string, style?: Partial<Style>) => (c ? colorHex(resolveColor(c, themeOf(style))).toUpperCase() : '');

export interface AgendaDraw {
  /** Chapters of the slide's section. */
  section: Chapter[];
  current: number | 'toc';
  frame: Frame;
  /** Every chapter of the deck (stored on each agenda element). */
  all: StoredAgenda['chapters'];
  options?: AgendaOptions;
  /** The options laid out with (`options` plus the fitted font size); default `options`. */
  drawOptions?: AgendaOptions;
  rev: number;
}

export function agendaContent(d: AgendaDraw, m: Measurer, style?: Partial<Style>): SceneContent & { hash: string } {
  const scene = layoutAgenda(d.section, d.current, d.frame, m, d.drawOptions ?? d.options, style);
  const stored: StoredAgenda = {
    v: 1, rev: d.rev, chapters: d.all, ...(d.options && Object.keys(d.options).length ? { options: d.options } : {}),
    shown: scene.headings.map((h) => ({
      id: h.id, title: h.title, level: h.level, state: h.state,
      fill: hexOf(h.style.fill, style), line: hexOf(h.style.line, style), color: hexOf(h.style.color, style), bold: h.style.bold, size: h.style.size,
    })),
  };
  const f = d.frame;
  const hash = fnv(JSON.stringify([scene.shapes, scene.roles, f]));
  const tags: Tags = [
    [AGENDA_TAG, JSON.stringify(stored)],
    ['OC_FRAME', [f.x, f.y, f.w, f.h].map(num).join(',')],
    ['OC_SIG', headingSignature(scene.headings)],
    ['OC_HASH', hash],
  ];
  const alt = d.current === 'toc' ? 'Table of contents' : `Agenda, current chapter: ${d.section[d.current]?.title ?? ''}`;
  return { scene, name: AGENDA_GROUP_NAME, altText: `${alt} (${scene.headings.map((h) => h.title).join(', ')})`, tags, hash };
}

/** 64-bit FNV-1a of a string as 16 hex digits (pure: this module also runs in the browser bundle). */
function fnv(text: string): string {
  let h = 0xcbf29ce484222325n;
  for (let i = 0; i < text.length; i++) h = ((h ^ BigInt(text.charCodeAt(i))) * 0x100000001b3n) & 0xffffffffffffffffn;
  return h.toString(16).padStart(16, '0');
}
