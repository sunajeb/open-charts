import type { ChartSpec, Frame, Scene, Style } from './types.ts';
import type { Measurer } from './text.ts';
import { makeCtx } from './common.ts';
import { prepareSpec } from './spec.ts';
import { layoutColumns } from './charts/column.ts';
import { layoutWaterfall } from './charts/waterfall.ts';
import { layoutMekko } from './charts/mekko.ts';
import { layoutLine } from './charts/line.ts';
import { layoutCombo } from './charts/combo.ts';
import { layoutMekkoXY } from './charts/mekkoxy.ts';
import { layoutScatter } from './charts/scatter.ts';
import { layoutTable } from './charts/table.ts';
import { layoutGantt } from './charts/gantt.ts';

export * from './types.ts';
export { measurerFromFonts, lineHeightFactor, wrap } from './text.ts';
export type { Measurer } from './text.ts';
export { checkScene } from './quality.ts';
export type { Issue, CheckOptions } from './quality.ts';
export {
  SPEC_VERSION, migrate, migrateDeck, validateSpec, validateDeck, InvalidSpecError,
  CHART_KEYS, OPTION_KEYS, SERIES_KEYS, DECK_KEYS, STYLE_KEYS, CHART_TYPES,
} from './spec.ts';
export type { SpecError } from './spec.ts';
export { formatValue, applyFormat, formatNumber, LOCALES, pickScale } from './numbers.ts';
export type { NumberStyle, Locale } from './numbers.ts';
export { parseCategoryDate, years360, cagrYears } from './dates.ts';

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
  combo: layoutCombo,
  mekkoxy: layoutMekkoXY,
  scatter: layoutScatter,
  bubble: layoutScatter,
  table: layoutTable,
  gantt: layoutGantt,
};

/**
 * Lay out one chart. The spec is migrated to the current version and validated first; an invalid
 * spec throws InvalidSpecError whose message lists every "/path: problem" (also on `.errors`).
 */
export function layoutChart(spec: ChartSpec, frame: Frame, m: Measurer, style: Style = defaultStyle): Scene {
  const s = prepareSpec(spec);
  const ctx = makeCtx(s, style, frame, m);
  layouts[s.type](ctx);
  if (ctx.roles.length !== ctx.shapes.length) throw new Error('internal: every shape needs a role');
  return { frame, shapes: ctx.shapes, roles: ctx.roles, hidden: ctx.hidden };
}
