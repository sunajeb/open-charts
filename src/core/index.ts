import type { ChartSpec, Frame, Scene, Style } from './types.ts';
import type { Measurer } from './text.ts';
import { makeCtx } from './common.ts';
import { prepareSpec } from './spec.ts';
import { resolveCollisions } from './resolve.ts';
import { minimumSize, placeholder, TooSmall } from './minsize.ts';
import { layoutColumns } from './charts/column.ts';
import { layoutWaterfall } from './charts/waterfall.ts';
import { layoutMekko } from './charts/mekko.ts';
import { layoutLine } from './charts/line.ts';
import { layoutArea } from './charts/area.ts';
import { layoutPie } from './charts/pie.ts';
import { layoutHarvey } from './charts/harvey.ts';
import { layoutCombo } from './charts/combo.ts';
import { layoutMekkoXY } from './charts/mekkoxy.ts';
import { layoutScatter } from './charts/scatter.ts';
import { layoutTable } from './charts/table.ts';
import { layoutGantt } from './charts/gantt.ts';

export * from './types.ts';
export { measurerFromFonts, lineHeightFactor, wrap } from './text.ts';
export type { Measurer } from './text.ts';
export { checkScene } from './quality.ts';
export { minimumSize, TooSmall } from './minsize.ts';
export type { Size } from './minsize.ts';
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
  area: layoutArea,
  area100: layoutArea,
  pie: layoutPie,
  doughnut: layoutPie,
  harvey: layoutHarvey,
  combo: layoutCombo,
  mekkoxy: layoutMekkoXY,
  scatter: layoutScatter,
  bubble: layoutScatter,
  table: layoutTable,
  gantt: layoutGantt,
};

/**
 * Lay out one chart. Below the chart type's minimum size (see `minimumSize`) the scene is a "too small"
 * placeholder (frame outline + message, role `placeholder`) and `hidden` says how large to make it.
 * Labels that still collide after the chart type has fitted them are hidden and reported (resolve.ts).
 * The spec is migrated to the current version and validated first; an invalid
 * spec throws InvalidSpecError whose message lists every "/path: problem" (also on `.errors`).
 */
export function layoutChart(spec: ChartSpec, frame: Frame, m: Measurer, style: Style = defaultStyle): Scene {
  const s = prepareSpec(spec);
  const min = minimumSize(s);
  // Lay out at the frame, or — when it is too small — find the size to ask for: at least the type's
  // minimum, grown by what the content says it needs (a table's columns, a category axis).
  let size = { w: Math.max(frame.w, min.w), h: Math.max(frame.h, min.h) };
  let why: string | undefined;
  for (let k = 0; k < 6; k++) {
    const ctx = makeCtx(s, style, { ...frame, ...size }, m);
    try {
      layouts[s.type](ctx);
    } catch (e) {
      if (!(e instanceof TooSmall)) throw e;
      why = e.message;
      size = { w: Math.max(size.w, Math.ceil(e.need.w ?? 0)), h: Math.max(size.h, Math.ceil(e.need.h ?? 0)) };
      continue;
    }
    if (size.w !== frame.w || size.h !== frame.h) break;
    if (ctx.roles.length !== ctx.shapes.length) throw new Error('internal: every shape needs a role');
    ctx.hidden.push(...resolveCollisions(ctx.shapes, ctx.roles, frame, m));
    return { frame, shapes: ctx.shapes, roles: ctx.roles, hidden: ctx.hidden };
  }
  return placeholder(s, frame, size, m, style, why);
}
