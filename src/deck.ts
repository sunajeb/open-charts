import type { ChartSpec, Frame, Scene, Style } from './core/index.ts';
import { checkScene, defaultStyle, layoutGroup, migrate, scaleGroupOf, themeOf } from './core/index.ts';
import type { Measurer } from './core/text.ts';
import { SLIDE, writePptx, type SceneContent } from './render/pptx.ts';
import { chartTags } from './render/tags.ts';
import { sceneToSvg } from './render/svg.ts';

/** One chart of a slide that holds several (`charts`); without a frame, the charts share the chart area side by side. */
export interface SlideChart { frame?: Frame; chart: ChartSpec }

/** A slide: one chart (`chart`) in the chart area, or several (`charts`), each in its own frame. */
export interface DeckSlide { title?: string; chart?: ChartSpec; charts?: SlideChart[] }

export interface DeckSpec {
  /** Spec format version (see migrateDeck / validateDeck in src/core). */
  specVersion?: 1;
  $schema?: string;
  style?: Partial<Style>;
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

export function buildDeck(deck: DeckSpec, m: Measurer): { pptx: Uint8Array; slides: BuiltSlide[] } {
  const style = { ...defaultStyle, ...deck.style };
  const built = deck.slides.map((s) => {
    // Older (v0) specs lose options their chart type does not use; say so instead of failing.
    const migrated: string[] = [];
    const charts = slideCharts(s);
    const scenes = layoutGroup(charts.map((c) => ({ spec: migrate(c.chart, migrated), frame: c.frame })), m, style);
    const at = (k: number) => (charts.length > 1 ? `chart ${k + 1}: ` : '');
    const issues = [
      ...migrated.map((w) => `warning: ${w}`),
      ...scenes.flatMap((scene, k) => [
        ...checkScene(scene, { measurer: m }).map((i) => `${i.kind}: ${at(k)}${i.detail}`),
        ...scene.hidden.map((h) => `hidden: ${at(k)}${h}`),
      ]),
    ];
    const font = (k: number) => charts[k].chart.style?.font ?? style.font;
    return {
      charts,
      slide: { title: s.title, scene: scenes[0], scenes, svg: sceneToSvg(scenes, SLIDE, s.title, font(0)), issues },
      content: scenes.map((scene, k) => chartContent(charts[k].chart, scene, font(k), deck.style)),
    };
  });
  const pptx = writePptx(built.map((b, i) => ({ title: deck.slides[i].title, scenes: b.content })), style.font, themeOf(style));
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
  return { scene, name, altText: altText(chart), tags };
}

/** A scale (pt per unit) as plain decimal digits, 8 significant: VBA's Val() reads no exponent. */
export function scaleText(v: number): string {
  const digits = Math.min(20, Math.max(0, 7 - Math.floor(Math.log10(v))));
  return v.toFixed(digits).replace(/(\.\d*?)0+$/, '$1').replace(/\.$/, '');
}
