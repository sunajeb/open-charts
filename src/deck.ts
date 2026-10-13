import type { ChartSpec, Frame, Scene, Style } from './core/index.ts';
import { checkScene, defaultStyle, layoutGroup, fileStyle, measurerFor, mergeStyles, migrate, resolveFont, scaleGroupOf, styleFor, themeOf, type StyleFile } from './core/index.ts';
import type { Measurer } from './core/text.ts';
import { SLIDE, writePptx, type SceneContent } from './render/pptx.ts';
import { chartTags, chunkedTags } from './render/tags.ts';
import { baseOf } from './core/overrides.ts';
import { sceneToSvg } from './render/svg.ts';
import { fitAgenda, planDeck, type AgendaOptions, type AgendaSlide, type Chapter } from './core/agenda.ts';
import { agendaContent } from './agenda.ts';
import type { Tags } from './render/tags.ts';

/** One chart of a slide that holds several (`charts`); without a frame, the charts share the chart area side by side. */
export interface SlideChart { frame?: Frame; chart: ChartSpec }

/**
 * A slide: one chart (`chart`) in the chart area, or several (`charts`), each in its own frame. Or an
 * agenda slide: `chapter` starts a chapter (at text `level`, 1 = top), `toc` is a table of contents.
 * `split` starts a new, independent agenda section at this slide (e.g. the backup).
 */
export interface DeckSlide {
  title?: string; chart?: ChartSpec; charts?: SlideChart[];
  chapter?: string; level?: number; toc?: boolean; split?: boolean;
}

/** Agenda options of a deck (src/core/agenda.ts) plus the title of agenda slides (default "Agenda"). */
export type DeckAgenda = AgendaOptions & { title?: string };

export interface DeckSpec {
  /** Spec format version (see migrateDeck / validateDeck in src/core). */
  specVersion?: 1;
  $schema?: string;
  style?: Partial<Style>;
  agenda?: DeckAgenda;
  slides: DeckSlide[];
}

export interface BuiltSlide {
  title?: string;
  /** The slide's first chart (every slide has at least one). */
  scene: Scene;
  /** Every chart on the slide, in spec order. */
  scenes: Scene[];
  svg: string;
  issues: string[];
}

/** Chart area on a 16:9 slide, below the title. */
export const CHART_FRAME = { x: 48, y: 96, w: SLIDE.w - 96, h: SLIDE.h - 132 };
/** Space between charts that share the chart area. */
export const CHART_GAP = 24;

const TYPE_NAMES: Record<ChartSpec['type'], string> = {
  stacked: 'Stacked column chart', clustered: 'Clustered column chart', percent: '100% column chart',
  waterfall: 'Waterfall chart', mekko: 'Mekko chart', line: 'Line chart',
  area: 'Stacked area chart', area100: '100% area chart', pie: 'Pie chart', doughnut: 'Doughnut chart',
  harvey: 'Harvey ball chart',
  combo: 'Combination chart', mekkoxy: 'Mekko XY chart', scatter: 'Scatter chart', bubble: 'Bubble chart',
  table: 'Table', gantt: 'Gantt chart',
};

export function altText(chart: ChartSpec): string {
  return `${TYPE_NAMES[chart.type]} of ${chart.series.filter((s) => !s.hidden).map((s) => s.name).join(', ')} across ${chart.categories.join(', ')}`;
}

/** The charts of a slide with their frames: `chart` fills the chart area; `charts` without frames split it into equal columns. */
export function slideCharts(slide: DeckSlide): { frame: Frame; chart: ChartSpec }[] {
  if (!slide.charts) return [{ frame: CHART_FRAME, chart: slide.chart! }];
  const n = slide.charts.length;
  const w = (CHART_FRAME.w - CHART_GAP * (n - 1)) / n;
  return slide.charts.map((c, i) => ({ chart: c.chart, frame: c.frame ?? { ...CHART_FRAME, x: CHART_FRAME.x + i * (w + CHART_GAP), w } }));
}

/** True for a slide the agenda draws (a chapter's agenda slide or a table of contents). */
const isAgendaSlide = (s: DeckSlide) => s.chapter !== undefined || !!s.toc;

/** Chapter ids, deck order: c1, c2, … */
export const chapterIdOf = (k: number) => `c${k + 1}`;

/** The agenda's view of the deck: chapters in order and what each slide is. */
export function deckAgenda(deck: DeckSpec): { slides: AgendaSlide[]; chapters: Chapter[] } {
  const chapters: Chapter[] = [];
  const slides = deck.slides.map((s): AgendaSlide => {
    const split = s.split ? { split: true } : {};
    if (s.chapter !== undefined) {
      const id = chapterIdOf(chapters.length);
      chapters.push({ id, title: s.chapter, level: s.level ?? 1 });
      return { kind: 'chapter', chapter: id, ...split };
    }
    return s.toc ? { kind: 'toc', ...split } : split;
  });
  return { slides, chapters };
}

/**
 * A deck under a style file: each chart's style becomes the file's style for its type, then the
 * deck's, then the chart's own (later wins), so the stored spec re-lays out the same anywhere.
 */
export function applyStyleFile(deck: DeckSpec, file: StyleFile): DeckSpec {
  const chart = (c: ChartSpec): ChartSpec => ({ ...c, style: mergeStyles(styleFor(file, c.type), deck.style, c.style) });
  return {
    ...deck,
    style: mergeStyles(fileStyle(file), deck.style),
    slides: deck.slides.map((s) => ({
      ...s,
      ...(s.chart ? { chart: chart(s.chart) } : {}),
      ...(s.charts ? { charts: s.charts.map((c) => ({ ...c, chart: chart(c.chart) })) } : {}),
    })),
  };
}

export function buildDeck(deck: DeckSpec, m: Measurer, styleFile?: StyleFile): { pptx: Uint8Array; slides: BuiltSlide[] } {
  // Under a style file every chart carries its full style; the deck style only sets the slide font and theme.
  const deckStyle = styleFile ? undefined : deck.style;
  if (styleFile) deck = applyStyleFile(deck, styleFile);
  const style = { ...defaultStyle, ...deck.style };
  const agenda = deck.slides.some((s) => isAgendaSlide(s) || s.split) ? agendaSlides(deck, m) : [];
  const built = deck.slides.map((s, i) => {
    if (agenda[i]?.content) {
      const a = agenda[i];
      const title = s.title ?? deck.agenda?.title ?? 'Agenda';
      return {
        charts: [],
        slide: { title, scene: a.content!.scene, scenes: [a.content!.scene], svg: sceneToSvg([a.content!.scene], SLIDE, title, style.font), issues: [] },
        content: [a.content!],
        extra: { title, tags: a.tags, hidden: a.hidden },
      };
    }
    if (!s.chart && !s.charts) {
      const scene: Scene = { frame: CHART_FRAME, shapes: [], hidden: [] };
      return { charts: [], slide: { title: s.title, scene, scenes: [], svg: sceneToSvg([], SLIDE, s.title, style.font), issues: [] }, content: [], extra: { title: s.title, tags: agenda[i]?.tags } };
    }
    // Older (v0) specs lose options their chart type does not use; say so instead of failing.
    const migrated: string[] = [];
    const charts = slideCharts(s);
    const scenes = layoutGroup(charts.map((c) => ({ spec: migrate(c.chart, migrated), frame: c.frame })), m, style);
    const at = (k: number) => (charts.length > 1 ? `chart ${k + 1}: ` : '');
    const issues = [
      ...migrated.map((w) => `warning: ${w}`),
      ...scenes.flatMap((scene, k) => [
        ...checkScene(scene, { measurer: measurerFor(m, resolveFont(charts[k].chart.style?.font ?? style.font, { ...style, ...charts[k].chart.style })) }).map((i) => `${i.kind}: ${at(k)}${i.detail}`),
        ...scene.hidden.map((h) => `hidden: ${at(k)}${h}`),
        ...(scene.warnings ?? []).filter((w) => w.startsWith('font "')).map((w) => `warning: ${at(k)}${w}`),
      ]),
    ];
    const font = (k: number) => charts[k].chart.style?.font ?? style.font;
    return {
      charts,
      slide: { title: s.title, scene: scenes[0], scenes, svg: sceneToSvg(scenes, SLIDE, s.title, font(0)), issues },
      content: scenes.map((scene, k) => chartContent(charts[k].chart, scene, font(k), deckStyle)),
      extra: { title: s.title, tags: agenda[i]?.tags } as { title?: string; tags?: Tags; hidden?: boolean },
    };
  });
  const pptx = writePptx(built.map((b) => ({ ...b.extra, scenes: b.content })), style.font, themeOf(style), style.themeFonts);
  return { pptx, slides: built.map((b) => b.slide) };
}

/** A laid-out chart ready for the writer: meaningful group name, alt text and tag contract v1. */
export function chartContent(chart: ChartSpec, scene: Scene, font: string, style?: Partial<Style>, name = TYPE_NAMES[chart.type] ?? 'Chart'): SceneContent {
  // The stored spec carries the style it was drawn with (deck style under the chart's own), so the
  // chart re-lays out the same anywhere — theme slots stay slots and resolve in the deck it is in.
  const merged = style && Object.keys(style).length ? { ...style, ...chart.style } : chart.style;
  const spec = merged && Object.keys(merged).length ? { ...chart, style: merged } : chart;
  const tags = chartTags(spec, scene.frame, font);
  // Same-scale group: the group name and the chart's own scale, so a host can find the chart's
  // partners on the slide and the smallest scale among them without parsing the spec.
  const group = scaleGroupOf(migrate(chart));
  if (group && scene.valueScale) tags.push(['OC_SGROUP', group], ['OC_SCALE', scaleText(scene.valueScale.own)]);
  // OC_BASE: the formatting this layout draws, per role, so the next update can tell what the user
  // changed with PowerPoint's own tools and keep it as an override (core/overrides.ts mergeNativeEdits).
  if (scene.roles?.length && scene.roles[0] !== 'placeholder') tags.push(...chunkedTags('OC_BASE', JSON.stringify(baseOf(scene.shapes, scene.roles, font))));
  return { scene, name, altText: altText(chart), tags };
}

/** A scale (pt per unit) as plain decimal digits, 8 significant: VBA's Val() reads no exponent. */
export function scaleText(v: number): string {
  const digits = Math.min(20, Math.max(0, 7 - Math.floor(Math.log10(v))));
  return v.toFixed(digits).replace(/(\.\d*?)0+$/, '$1').replace(/\.$/, '');
}

interface AgendaSlideOut { content?: SceneContent; tags?: Tags; hidden?: boolean }

/** The agenda element, slide tags and hidden flag of every agenda slide of a deck. */
function agendaSlides(deck: DeckSpec, m: Measurer): AgendaSlideOut[] {
  const { title: _title, ...options } = deck.agenda ?? {};
  const { slides, chapters } = deckAgenda(deck);
  const plan = planDeck(slides, chapters);
  const fits = plan.sections.map((cs, k) => fitAgenda(cs, options, m, SLIDE, plan.slides.some((p) => p.section === k && p.current === 'toc')));
  return slides.map((a, i): AgendaSlideOut => {
    const tags: Tags = [];
    if (a.kind) tags.push(['OC_KIND', a.kind]);
    if (a.chapter) tags.push(['OC_CHAPTER', a.chapter]);
    if (a.split) tags.push(['OC_SPLIT', '1']);
    const p = plan.slides[i];
    if (p.current === undefined) return { tags: tags.length ? tags : undefined };
    const fit = fits[p.section];
    const content = agendaContent({ section: plan.sections[p.section], current: p.current, frame: fit.frame, all: chapters, options, drawOptions: fit.options, rev: 1 }, m, deck.style);
    return { content, tags, hidden: p.hidden };
  });
}
