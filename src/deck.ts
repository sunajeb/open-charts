import type { ChartSpec, Scene, Style } from './core/index.ts';
import { checkScene, defaultStyle, layoutChart, migrate } from './core/index.ts';
import type { Measurer } from './core/text.ts';
import { SLIDE, writePptx, type SceneContent } from './render/pptx.ts';
import { chartTags } from './render/tags.ts';
import { sceneToSvg } from './render/svg.ts';

export interface DeckSpec {
  /** Spec format version (see migrateDeck / validateDeck in src/core). */
  specVersion?: 1;
  $schema?: string;
  style?: Partial<Style>;
  slides: { title?: string; chart: ChartSpec }[];
}

export interface BuiltSlide { title?: string; scene: Scene; svg: string; issues: string[] }

/** Chart area on a 16:9 slide, below the title. */
export const CHART_FRAME = { x: 48, y: 96, w: SLIDE.w - 96, h: SLIDE.h - 132 };

const TYPE_NAMES: Record<ChartSpec['type'], string> = {
  stacked: 'Stacked column chart', clustered: 'Clustered column chart', percent: '100% column chart',
  waterfall: 'Waterfall chart', mekko: 'Mekko chart', line: 'Line chart',
  area: 'Stacked area chart', area100: '100% area chart', pie: 'Pie chart', doughnut: 'Doughnut chart',
  harvey: 'Harvey ball chart',
  combo: 'Combination chart', mekkoxy: 'Mekko XY chart', scatter: 'Scatter chart', bubble: 'Bubble chart',
  table: 'Table', gantt: 'Gantt chart',
};

export function altText(chart: ChartSpec): string {
  return `${TYPE_NAMES[chart.type]} of ${chart.series.map((s) => s.name).join(', ')} across ${chart.categories.join(', ')}`;
}

export function buildDeck(deck: DeckSpec, m: Measurer): { pptx: Uint8Array; slides: BuiltSlide[] } {
  const style = { ...defaultStyle, ...deck.style };
  const slides = deck.slides.map((s) => {
    // Older (v0) specs lose options their chart type does not use; say so instead of failing.
    const migrated: string[] = [];
    const scene = layoutChart(migrate(s.chart, migrated), CHART_FRAME, m, style);
    const issues = [
      ...migrated.map((w) => `warning: ${w}`),
      ...checkScene(scene, { measurer: m }).map((i) => `${i.kind}: ${i.detail}`), ...scene.hidden.map((h) => `hidden: ${h}`),
    ];
    return { title: s.title, scene, svg: sceneToSvg([scene], SLIDE, s.title, style.font), issues };
  });
  const pptx = writePptx(
    deck.slides.map((s, i) => ({ title: s.title, scenes: [chartContent(s.chart, slides[i].scene, style.font, deck.style)] })),
    style.font,
  );
  return { pptx, slides };
}

/** A laid-out chart ready for the writer: meaningful group name, alt text and tag contract v1. */
export function chartContent(chart: ChartSpec, scene: Scene, font: string, style?: Partial<Style>, name = TYPE_NAMES[chart.type] ?? 'Chart'): SceneContent {
  const spec = style && Object.keys(style).length ? { ...chart, style } : chart;
  return { scene, name, altText: altText(chart), tags: chartTags(spec, scene.frame, font) };
}
