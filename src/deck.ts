import type { ChartSpec, Scene, Style } from './core/index.ts';
import { checkScene, defaultStyle, layoutChart } from './core/index.ts';
import type { Measurer } from './core/text.ts';
import { SLIDE, writePptx } from './render/pptx.ts';
import { sceneToSvg } from './render/svg.ts';

export interface DeckSpec {
  style?: Partial<Style>;
  slides: { title?: string; chart: ChartSpec }[];
}

export interface BuiltSlide { title?: string; scene: Scene; svg: string; issues: string[] }

/** Chart area on a 16:9 slide, below the title. */
export const CHART_FRAME = { x: 48, y: 96, w: SLIDE.w - 96, h: SLIDE.h - 132 };

const TYPE_NAMES: Record<ChartSpec['type'], string> = {
  stacked: 'Stacked column chart', clustered: 'Clustered column chart', percent: '100% column chart',
  waterfall: 'Waterfall chart', mekko: 'Mekko chart', line: 'Line chart',
};

export function altText(chart: ChartSpec): string {
  return `${TYPE_NAMES[chart.type]} of ${chart.series.map((s) => s.name).join(', ')} across ${chart.categories.join(', ')}`;
}

export function buildDeck(deck: DeckSpec, m: Measurer): { pptx: Uint8Array; slides: BuiltSlide[] } {
  const style = { ...defaultStyle, ...deck.style };
  const slides = deck.slides.map((s) => {
    const scene = layoutChart(s.chart, CHART_FRAME, m, style);
    const issues = [...checkScene(scene, { measurer: m }).map((i) => `${i.kind}: ${i.detail}`), ...scene.hidden.map((h) => `hidden: ${h}`)];
    return { title: s.title, scene, svg: sceneToSvg([scene], SLIDE, s.title, style.font), issues };
  });
  const pptx = writePptx(
    deck.slides.map((s, i) => ({ title: s.title, scenes: [{ scene: slides[i].scene, name: 'Chart', altText: altText(s.chart) }] })),
    style.font,
  );
  return { pptx, slides };
}
