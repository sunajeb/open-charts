import type { ChartSpec, Frame, Scene, Style } from './types.ts';
import type { Measurer } from './text.ts';
import { makeCtx, type LayoutHints } from './common.ts';
export type { LayoutHints } from './common.ts';
import { resolveColor, resolveStyle, themeOf } from './theme.ts';
import { appliesTo, prepareSpec } from './spec.ts';
import { mirrorShapes } from './direction.ts';
import { resolveCollisions } from './resolve.ts';
import { restoreRoles, visibleSeries } from './visible.ts';
import { minimumSize, placeholder, TooSmall } from './minsize.ts';
import { compileOverrides, unmatched, type OverrideCtx } from './overrides.ts';
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
export { measurerFromFonts, measurerFromMetrics, fontMetrics, lineHeightFactor, wrap } from './text.ts';
export type { Measurer, FontMetrics } from './text.ts';
export { checkScene } from './quality.ts';
export {
  LABEL_FIELDS, ELEMENT_PROPERTIES, ELEMENT_STYLE_KEYS, ELEMENT_KINDS, ROLE_KINDS, kindOf, propertiesFor,
  compileOverrides, matches as roleMatches, specificity, detachElement, unpinElement, remapOverrides, PATTERNS,
  baseOf, mergeNativeEdits, rendered as renderedOf, sceneEntries, colorName,
} from './overrides.ts';
export type { ElementKind, ElementProperty, PropertyInfo, Dash, Arrowheads, MarkerStyle, Pattern, RenderedProps, RenderedMap, RenderedEntry } from './overrides.ts';
export { minimumSize, TooSmall } from './minsize.ts';
export type { Size } from './minsize.ts';
export type { Issue, CheckOptions } from './quality.ts';
export {
  SPEC_VERSION, migrate, migrateDeck, validateSpec, validateDeck, validateStyle, InvalidSpecError, OPTIONS, LIMITS, appliesTo, optionDefaults, editorSpec,
  CHART_KEYS, OPTION_KEYS, SERIES_KEYS, DECK_KEYS, STYLE_KEYS, STYLE_NUMBER_KEYS, SLIDE_KEYS, SLIDE_CHART_KEYS, CHART_TYPES, DIRECTIONS, directionsOf,
} from './spec.ts';
export type { SpecError, OptionInfo, OptionKey, Direction } from './spec.ts';
export { formatValue, applyFormat, formatNumber, LOCALES, pickScale } from './numbers.ts';
export type { NumberStyle, Locale } from './numbers.ts';
export { parseCategoryDate, years360, cagrYears } from './dates.ts';
export { LINK_SLOTS, isElementColor, resolveElementColor, parseThemeColor, tintHex, colorTheme } from './theme.ts';
export { THEME_SLOTS, MAPPED_SLOTS, COLOR_NAMES, DEFAULT_THEME, isColor, isColorName, themeOf, resolveColor, resolveStyle, colorHex, colorSlot } from './theme.ts';
export type { ThemeSlot, ThemeColors, ColorName } from './theme.ts';


/**
 * The open-charts style. Series, text and background use theme slots, so a chart takes the colours
 * of the presentation it is in; in previews and CLI decks the slots are the open-charts theme
 * (DEFAULT_THEME: the same navy-to-grey ramp and near-black text as before themes).
 */
export const defaultStyle: Style = {
  font: 'Arial',
  fontSize: 12,
  text: 'tx1',
  muted: '595959',
  axis: '7F7F7F',
  background: 'bg1',
  palette: ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'],
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
 * Series marked `hidden` are left out of the layout (roles keep their index in the spec).
 * The spec is migrated to the current version and validated first; an invalid
 * spec throws InvalidSpecError whose message lists every "/path: problem" (also on `.errors`).
 */
export function layoutChart(spec: ChartSpec, frame: Frame, m: Measurer, style: Style = defaultStyle, hints: LayoutHints = {}): Scene {
  const prepared = prepareSpec(spec);
  // The chart's own style wins over the one passed in; theme slots resolve once, here (theme.ts),
  // before hidden series are left out, so the per-row palette mapping moves resolved colours.
  const merged: Style = prepared.style ? { ...style, ...prepared.style } : style;
  const theme = themeOf(merged);
  const resolved: ChartSpec = prepared.series.some((x) => x.color)
    ? { ...prepared, series: prepared.series.map((x) => (x.color ? { ...x, color: resolveColor(x.color, theme) } : x)) }
    : prepared;
  // Hidden series (optional rows) are left out here; roles name the spec's series again at the end.
  const { spec: s, style: st, index, notes } = visibleSeries(resolved, resolveStyle(merged));
  style = st;
  // Overrides name series by their index in the spec; the chart types see visible indices.
  // The style's default element styles sit under the chart's own overrides, property by property.
  const own = compileOverrides(prepared.overrides);
  const defaults = compileOverrides(merged.elements);
  const lookup = (role: string) => { const d = defaults(role), o = own(role); return d && o ? { ...d, ...o } : d ?? o; };
  const seen = new Set<string>();
  const toSpec = (role: string) => (index.every((i, k) => i === k) ? role : role.replace(/(^|:)s(\d+)(?=:|$)/g, (_, pre, k) => `${pre}s${index[+k] ?? k}`));
  const any = (o?: object) => !!o && Object.keys(o).length > 0;
  const ov: OverrideCtx | undefined = any(prepared.overrides) || any(merged.elements)
    ? { theme, get: (role) => { const r = toSpec(role); seen.add(r); return lookup(r); } }
    : undefined;
  const min = minimumSize(s);
  // Lay out at the frame, or — when it is too small — find the size to ask for: at least the type's
  // minimum, grown by what the content says it needs (a table's columns, a category axis).
  let size = { w: Math.max(frame.w, min.w), h: Math.max(frame.h, min.h) };
  let why: string | undefined;
  for (let k = 0; k < 6; k++) {
    const ctx = makeCtx(s, style, { ...frame, ...size }, m, hints, ov);
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
    // "down" and "left" are the native "up" / "right" layouts mirrored (direction.ts).
    const dir = ctx.o.direction;
    if (dir === 'down' || dir === 'left') {
      mirrorShapes(ctx.shapes, ctx.roles, ctx.frame, dir === 'left' ? 'x' : 'y');
      // Where the value scale runs (drag handles) mirrors with the shapes.
      const at = ctx.valueScale?.at, f = ctx.frame;
      if (at) ctx.valueScale!.at = dir === 'left' ? [2 * f.x + f.w - at[0], at[1], 2 * f.x + f.w - at[2], at[3]] : [at[0], 2 * f.y + f.h - at[1], at[2], 2 * f.y + f.h - at[3]];
    }
    ctx.hidden.push(...resolveCollisions(ctx.shapes, ctx.roles, frame, m), ...notes);
    restoreRoles(ctx.roles, index);
    const warnings = any(prepared.overrides) ? unmatched(prepared.overrides, seen, prepared) : [];
    return {
      frame, shapes: ctx.shapes, roles: ctx.roles, hidden: ctx.hidden, ...(ctx.valueScale ? { valueScale: ctx.valueScale } : {}),
      ...(warnings.length ? { warnings } : {}),
    };
  }
  return placeholder(s, frame, size, m, style, why);
}

/** One chart of a slide for `layoutGroup`. */
export interface GroupItem { spec: ChartSpec; frame: Frame }

/** The same-scale group a chart belongs to (`options.scaleGroup` on a chart type that reads it). */
export const scaleGroupOf = (spec: ChartSpec): string | undefined => {
  const g = spec.options?.scaleGroup;
  return typeof g === 'string' && g && appliesTo('scaleGroup', spec.type) ? g : undefined;
};

/**
 * Lay out the charts of one slide. Charts whose `scaleGroup` names the same group draw one value unit
 * equally long: the smallest of their own scales, so the chart with the largest values (for its
 * height) keeps its scale and the others leave headroom (`Scene.valueScale`). Charts without a group,
 * or alone in theirs, are laid out exactly as by `layoutChart`.
 */
export function layoutGroup(items: GroupItem[], m: Measurer, style: Style = defaultStyle): Scene[] {
  const scenes = items.map((it) => layoutChart(it.spec, it.frame, m, style));
  const groups = new Map<string, number[]>();
  items.forEach((it, i) => {
    const g = scaleGroupOf(it.spec);
    if (g && scenes[i].valueScale) groups.set(g, [...(groups.get(g) ?? []), i]);
  });
  for (const members of groups.values()) {
    if (members.length < 2) continue;
    // A chart that follows the group reserves a wider axis, which can shorten its plot a little and
    // with it its own scale: settle in a few rounds (each one only ever lowers the target).
    for (let round = 0; round < 3; round++) {
      const target = Math.min(...members.map((i) => scenes[i].valueScale!.own));
      const off = members.filter((i) => Math.abs(scenes[i].valueScale!.ptPerUnit - target) > target * 1e-6);
      if (!off.length) break;
      for (const i of off) scenes[i] = layoutChart(items[i].spec, items[i].frame, m, style, { ptPerUnit: target });
    }
  }
  return scenes;
}
export { parseStyleFile, validateStyleFile, mergeStyles, styleFor, fileStyle, InvalidStyleFileError, STYLE_FILE_KEYS, STYLE_FILE_LIMITS } from './style-file.ts';
export type { StyleFile } from './style-file.ts';
export { dragHandles, dragSpec, resetSpec, describeDrag, gapFor, valueAt, niceStep } from './handles.ts';
export type { DragHandle, HandleKind, Relayout } from './handles.ts';
