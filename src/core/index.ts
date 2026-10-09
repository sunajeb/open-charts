import type { ChartSpec, Frame, Scene, Style } from './types.ts';
import type { Measurer } from './text.ts';
import { makeCtx } from './common.ts';
import { layoutColumns } from './charts/column.ts';
import { layoutWaterfall } from './charts/waterfall.ts';
import { layoutMekko } from './charts/mekko.ts';
import { layoutLine } from './charts/line.ts';

export * from './types.ts';
export { createMeasurer } from './text.ts';
export type { Measurer } from './text.ts';
export { checkScene } from './quality.ts';

export const defaultStyle: Style = {
  font: 'Arial',
  fontSize: 12,
  text: '222222',
  muted: '595959',
  axis: '7F7F7F',
  background: 'FFFFFF',
  palette: ['1F3A5F', '3E6A9E', '7FA3CC', 'B9CDE5', 'A6A6A6', 'D9D9D9'],
  positive: '3E8E5E',
  negative: 'C0504D',
};

type Layout = (ctx: ReturnType<typeof makeCtx>) => void;
const layouts: Record<ChartSpec['type'], Layout> = {
  stacked: layoutColumns,
  clustered: layoutColumns,
  percent: layoutColumns,
  waterfall: layoutWaterfall,
  mekko: layoutMekko,
  line: layoutLine,
};

export function layoutChart(spec: ChartSpec, frame: Frame, m: Measurer, style: Style = defaultStyle): Scene {
  validate(spec);
  const layout = layouts[spec.type];
  if (!layout) throw new Error(`Unknown chart type "${spec.type}"`);
  const ctx = makeCtx(spec, style, frame, m);
  layout(ctx);
  return { frame, shapes: ctx.shapes, hidden: ctx.hidden };
}

function validate(spec: ChartSpec) {
  if (!spec.categories?.length) throw new Error('Chart needs at least one category');
  if (!spec.series?.length) throw new Error('Chart needs at least one series');
  for (const s of spec.series) {
    if (s.values.length !== spec.categories.length) {
      throw new Error(`Series "${s.name}" has ${s.values.length} values but there are ${spec.categories.length} categories`);
    }
  }
}
