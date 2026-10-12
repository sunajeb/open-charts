// All geometry is in points (1/72 inch). Origin is the slide's top-left corner.
import type { Scale, Unit } from './numbers.ts';
export type { Scale, Unit } from './numbers.ts';

export type ChartType = 'stacked' | 'clustered' | 'percent' | 'waterfall' | 'mekko' | 'line'
  | 'area' | 'area100' | 'pie' | 'doughnut' | 'harvey'
  | 'combo' | 'mekkoxy' | 'scatter' | 'bubble' | 'table' | 'gantt';

/**
 * A cell: a number, empty (null), or "e" in a waterfall = this category is a computed total.
 * Scatter / bubble: an [x, y] / [x, y, size] tuple (one point; the category is its label).
 * Gantt: a date string ("2026-03-15", "Mar 2026", "Q2 2026") read with parseCategoryDate.
 */
export type Value = number | null | 'e' | number[] | string;

export interface Series {
  name: string;
  values: Value[];
  /** Hex colour, "RRGGBB" or "#RRGGBB". */
  color?: string;
  /** Combination chart: draw this series as a line over the columns (same value axis). */
  as?: 'column' | 'line';
}

export interface ChartOptions {
  /** Total label above each stacked column. Default: on for stacked, off otherwise. */
  totals?: boolean;
  /**
   * Compound annual growth arrow between two category indices (uses column totals). The number of
   * years comes from `years`, else from dates read from the category labels (30/360), else from
   * the index distance.
   */
  cagr?: { from: number; to: number; years?: number } | null;
  /** Difference arrow between the totals of two categories. */
  difference?: { from: number; to: number } | null;
  /** Show a value axis with ticks. Default: off (values are labelled on the chart). */
  valueAxis?: boolean;
  /** Force the value axis to include zero. Default: on, except narrow-range line charts. */
  zeroBased?: boolean;
  /** Fixed number of decimals. Default: inferred from the data. */
  decimals?: number;
  /** Unit on every value label: a suffix string ("m", "%"), or {text, position, space} for "$42" / "42 €". */
  unit?: string | Unit;
  /** Excel-style number format for value labels, e.g. "#,##0.0" or "#,##0;(#,##0)". Overrides `decimals`. */
  format?: string;
  /** Locale for decimal mark, grouping and percent spacing: en-US (default), en-GB, en-IN, de-DE, fr-FR. */
  locale?: string;
  /** Magnitude scaling shared by the whole chart: "auto" picks k / m / bn from the data. Default "none". */
  scale?: Scale;
  /** Stacked, clustered and 100% charts: "horizontal" draws bars from the left. Default "vertical". */
  orientation?: 'vertical' | 'horizontal';
  /** Mekko: stack every column to 100% (default) or keep absolute segment heights. Widths always encode totals. */
  mekko?: 'percent' | 'unit';
  /** Gap between columns as a fraction of the column width. Default 0.5. */
  gap?: number;
  /** Series identification: labels to the right of the last column, a legend, or none. */
  seriesLabels?: 'right' | 'legend' | 'none';
  /** Waterfall: category indices drawn as totals anchored at the baseline. Default: first and last category. */
  totalsAt?: number[];
  /** Line: draw a step (ortho) line, horizontal to the next category and then vertical. Default off. */
  step?: boolean;
  /** 100% chart: values are already percentages (the datasheet's % toggle); columns keep their own totals. */
  percentData?: boolean;
  /** Combination chart: how the column series are drawn. Default "stacked". */
  combo?: 'stacked' | 'clustered';
  /** Category order (Mekko XY): sheet order (default), reverse, or by height. */
  sort?: 'sheet' | 'reverse' | 'ascending' | 'descending';
  /** Scatter and bubble: axis titles. */
  axisTitles?: { x?: string; y?: string };
  /** Gantt: shaded periods (holidays, phases) across every row, with an optional label. */
  shades?: { from: string; to: string; label?: string }[];
  /** Bubble: size legend (one bubble of a round size, top right); a string adds a caption, e.g. "Revenue ($m)". */
  sizeLegend?: boolean | string;
  /** Gantt: brackets labelling periods of time (phases), drawn under the rows. */
  brackets?: { from: string; to: string; label?: string }[];
  /** Gantt: a "today" line through every row at this date (it is never read from the clock). */
  today?: string;
  /** Pie and doughnut: what the slice labels show — shares (default), absolute values, or both ("118 (52%)"). */
  labelContent?: 'percent' | 'value' | 'both';
}

export interface ChartSpec {
  /** Spec format version. Unversioned input is v0 and is upgraded by `migrate()`. */
  specVersion?: 1;
  $schema?: string;
  type: ChartType;
  categories: string[];
  series: Series[];
  options?: ChartOptions;
}

export interface Style {
  font: string;
  fontSize: number;
  text: string;
  muted: string;
  axis: string;
  background: string;
  palette: string[];
  positive: string;
  negative: string;
}

export interface Frame { x: number; y: number; w: number; h: number }

export interface RectShape {
  kind: 'rect';
  x: number; y: number; w: number; h: number;
  fill: string;
  stroke?: string;
  strokeWidth?: number;
  /** Draw as an ellipse inscribed in the box (line chart markers). Same as `geom: 'ellipse'`. */
  ellipse?: boolean;
  /** Preset outline drawn in the box. Default 'rect'. */
  geom?: 'rect' | 'ellipse' | 'diamond' | 'roundRect' | 'triangle';
}

/**
 * One path command. Coordinates are slide points. `A` draws a circular arc centred on (cx, cy):
 * angles in degrees, 0 = 3 o'clock, positive = clockwise on screen; it starts with a straight
 * segment from the current point to the arc's start (or moves there if the path is empty).
 */
export type PathCmd =
  | { op: 'M' | 'L'; x: number; y: number }
  | { op: 'A'; cx: number; cy: number; r: number; start: number; sweep: number }
  | { op: 'Z' };

/** Freeform: area fills, pie and doughnut wedges, Harvey balls. Renderers draw the flattened outline. */
export interface PathShape {
  kind: 'path';
  commands: PathCmd[];
  fill?: string;
  stroke?: string;
  strokeWidth?: number;
  dash?: boolean;
}

export interface LineShape {
  kind: 'line';
  x1: number; y1: number; x2: number; y2: number;
  color: string;
  width: number;
  dash?: boolean;
  arrowEnd?: boolean;
}

export interface TextShape {
  kind: 'text';
  x: number; y: number; w: number; h: number;
  text: string;
  size: number;
  color: string;
  bold?: boolean;
  align: 'left' | 'center' | 'right';
  valign: 'top' | 'middle' | 'bottom';
  /** Distance between baselines of consecutive lines, from the measured font (B7). */
  lineHeight?: number;
  /** Optional box behind the text (e.g. CAGR bubble). */
  fill?: string;
  stroke?: string;
  rounded?: boolean;
  /** Coarse label kind (kept for compatibility; `Scene.roles` is the precise role). */
  role?: 'value' | 'total' | 'category' | 'series' | 'axis' | 'decoration' | 'title';
}

export type Shape = RectShape | LineShape | TextShape | PathShape;

export interface Scene {
  frame: Frame;
  shapes: Shape[];
  /**
   * Role of each shape, parallel to `shapes`, assigned by the engine only: `bar:s<i>:c<j>`,
   * `label:value:s<i>:c<j>`, `label:total:c<j>`, `label:category:c<j>`, `label:series:s<i>`, `axis`,
   * `axis:tick:<k>`, `gridline:<k>`, `arrow:cagr`, `arrow:diff`, `connector:c<j>`, `legend:s<i>`,
   * `marker:s<i>:c<j>`, `line:s<i>`, `area:s<i>` (area band), `wedge:s0:c<j>` (pie/doughnut slice),
   * `leader:c<j>` (pie label leader), `leader:s<i>:c<j>` (leader of a segment label placed outside its segment), `label:total` (doughnut grand total), `legend:c<j>` (legend of categories), `ball:s<i>:c<j>` (Harvey ball),
   * `shade:<k>` (Gantt shaded period), `bracket:<k>` (Gantt bracket), `today` (Gantt today line),
   * `legend:size` (bubble size legend); `leader:c<j>` also links a scatter label to its point. Always set by `layoutChart`.
   */
  roles?: string[];
  /** Labels the engine chose to hide (or shrink) because they did not fit, with the reason. */
  hidden: string[];
}
