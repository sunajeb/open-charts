// All geometry is in points (1/72 inch). Origin is the slide's top-left corner.
import type { Scale, Unit } from './numbers.ts';
import type { Overrides } from './overrides.ts';
export type { Scale, Unit } from './numbers.ts';
export type { Overrides, ElementStyle, ContentField, LabelField, NumberField } from './overrides.ts';

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
  /**
   * Leave this row out of the chart; its data stays in the spec (the datasheet's eye toggle). Hidden
   * series keep their index in roles and their palette colour, so showing one again changes nothing else.
   */
  hidden?: boolean;
}

/**
 * Chart options. Which chart types read each option is `OPTIONS[key].types` in spec.ts (an option
 * on any other type is a validation error); the per-type defaults are `optionDefaults(type)`. The
 * comments below summarise both.
 */
export interface ChartOptions {
  /** Total above each column (stacked, 100%, Mekko, Mekko XY, combo, area) or in the hole (doughnut). Default: on for stacked, mekko and mekkoxy; off for percent, area, doughnut and combo. */
  totals?: boolean;
  /**
   * Compound annual growth arrow between two category indices (uses column totals). The number of
   * years comes from `years`, else from dates read from the category labels (30/360), else from
   * the index distance.
   */
  cagr?: { from: number; to: number; years?: number } | null;
  /**
   * Total difference arrow between the totals of two categories. `mode` (manual: Arrow modes):
   * "absolute" (default, ↔, the signed difference), "relative" (→, change in percent from `from` to
   * `to`) or "reverse" (←, change in percent from `to` to `from`).
   */
  difference?: { from: number; to: number; mode?: DiffMode } | null;
  /** Level difference arrows between two segment tops / data points (or a value line); mode default "relative". */
  levelArrows?: LevelArrowSpec[];
  /** Horizontal reference lines at a value or at the mean of the column totals, labelled left or right. */
  valueLines?: ValueLineSpec[];
  /** 100% chart: the "100%" indicator beside the chart. Default: right, or left when series names are on the right (none when the left holds a value axis). */
  percentIndicator?: 'left' | 'right' | 'none';
  /** 100% chart: the datasheet's "100%=" row, one base per category (null = the column's sum). */
  percentBase?: (number | null)[];
  /** Category indices after which one extra category slot of space is inserted (repeat an index for more). */
  categoryGaps?: number[];
  /** Category indices after which the category axis shows a break symbol. */
  categoryBreaks?: number[];
  /** Value axis break: the range from–to is compressed and every bar that spans it shows the break. */
  axisBreak?: AxisBreakSpec | null;
  /** Combination chart: indices of line series drawn against a second value axis on the right. */
  secondaryAxis?: number[];
  /** Show a value axis with ticks. Default: on for line; off for the other types that read it (their values are labelled on the chart). */
  valueAxis?: boolean;
  /** Line, scatter and bubble: force the value axis to include zero (true) or fit the data (false). Default: zero, unless the data sits in a narrow band far from it. */
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
  /**
   * Which way the chart grows (`DIRECTIONS` in spec.ts lists them per type): "up" columns (default),
   * "down" columns hanging from the top, "right" bars from the left (a line: a profile chart), "left"
   * bars from the right with the categories on the right. Replaces `orientation` ("horizontal" = "right").
   */
  direction?: 'up' | 'down' | 'left' | 'right';
  /** Mekko: stack every column to 100% (default) or keep absolute segment heights. Widths always encode totals. */
  mekko?: 'percent' | 'unit';
  /** Stacked, clustered, 100%, waterfall and combo: gap between columns as a fraction of the column width, 0–10. Default 0.5. */
  gap?: number;
  /** Series identification: labels to the right of the last column, a legend, or none. */
  seriesLabels?: 'right' | 'legend' | 'none';
  /** Waterfall: category indices drawn as totals anchored at the baseline. Default: the first category and any "e" cells, else the first and last. */
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
  /**
   * Same-scale group: charts on one slide with the same group name draw values at one scale (the
   * smallest of their own), so equal values are equally tall across the charts.
   */
  scaleGroup?: string;
  /**
   * Top of the value axis (the drag handle on the axis sets it): the axis reaches at least this
   * value, leaving headroom above the data; it never cuts data off (a value below the data's top is
   * ignored). Left out or null: the axis fits the data.
   */
  axisMax?: number | null;
}

/** Difference arrow mode (manual: Arrow modes): relative →, absolute ↔, reverse ← (relative, other direction). */
export type DiffMode = 'relative' | 'absolute' | 'reverse';
/** End of a level difference arrow: the top of a segment / a data point (series omitted = the column total), or a value line. */
export type LevelEnd = { category: number; series?: number } | { valueLine: number };
export interface LevelArrowSpec {
  from: LevelEnd;
  to: LevelEnd;
  /** Default "relative". */
  mode?: DiffMode;
  /** "beside" (default): right of the later column (below the later bar); "between": in the gap before it. */
  position?: 'beside' | 'between';
}
export interface ValueLineSpec {
  /** A value, or "mean": the arithmetic mean of the column totals (clustered: of the bars). */
  value: number | 'mean';
  /** Label text; "{value}" is replaced by the formatted value. Default: the value ("Ø value" for a mean). */
  label?: string;
  /** Label beside the chart: "right" (default) or "left". */
  position?: 'left' | 'right';
}
export interface AxisBreakSpec {
  from: number;
  to: number;
  /** "straight" (default) or "wiggle". */
  style?: 'straight' | 'wiggle';
}

export interface ChartSpec {
  /** Spec format version. Unversioned input is v0 and is upgraded by `migrate()`. */
  specVersion?: 1;
  $schema?: string;
  type: ChartType;
  categories: string[];
  series: Series[];
  options?: ChartOptions;
  /** Style overrides for this chart (over the deck's style). */
  style?: Partial<Style>;
  /**
   * Per-element formatting: role (`label:value:s1:c2`) or selector (`label:value:s1:*`, `bar:*:c2`,
   * `label:category:*`) → ElementStyle. Exact roles win over series / category selectors, which win
   * over kind selectors (core/overrides.ts, DATA-FORMAT.md "Overrides").
   */
  overrides?: Overrides;
}

/**
 * Fonts, colours and number defaults (DATA-FORMAT.md "Style"). Every colour is a hex colour
 * ("1F3A5F") or a theme slot (`accent1`…`accent6`, `dk1` `lt1` `dk2` `lt2`, or the slide-mapped
 * `tx1` `bg1` `tx2` `bg2`); a slot follows the theme of the presentation the chart is in.
 * Inside a scene, a colour that came from a slot reads "1F3A5F@accent1" (see core/theme.ts).
 */
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
  /** Hex colour of each theme slot, for previews, the CLI deck's theme and contrast (default: the open-charts theme). */
  theme?: Partial<Record<'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6' | 'dk1' | 'lt1' | 'dk2' | 'lt2', string>>;
  /** Number defaults for every chart; a chart's own options win. */
  numbers?: Pick<ChartOptions, 'locale' | 'decimals' | 'scale'>;
  /**
   * Default element styles: role or selector → ElementStyle, the same vocabulary as a chart's
   * `overrides` (DATA-FORMAT.md "Overrides"), e.g. {"gridline:*": {"line": "D9D9D9", "dash": "dot"}}.
   * A chart's own overrides win, property by property. `pin` and `detach` are per element only.
   */
  elements?: Overrides;
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
  /** Dashed ('dot': dotted) outline (an override). */
  dash?: boolean | 'dot';
  /** OOXML preset pattern in `fill` on the background (an override). */
  pattern?: string;
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
  /** Dashed outline; 'dot': dotted. */
  dash?: boolean | 'dot';
  /** OOXML preset pattern in `fill` on the background (an override). */
  pattern?: string;
}

export interface LineShape {
  kind: 'line';
  x1: number; y1: number; x2: number; y2: number;
  color: string;
  width: number;
  /** Dashed; 'dot': dotted. */
  dash?: boolean | 'dot';
  arrowEnd?: boolean;
  /** Arrowhead at the start (an override). */
  arrowStart?: boolean;
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
  /** Width of the box's outline (an override; default 0.75). */
  strokeWidth?: number;
  rounded?: boolean;
  /** Font face, when it differs from the chart's (an override). */
  font?: string;
  /**
   * Text direction (an override): -90 = turned 90° left (reads bottom to top), 90 = turned 90° right.
   * x y w h are then the turned text's footprint on the slide, not the unrotated box.
   */
  rotate?: -90 | 90;
  /** Placed by hand (an override's `pin`): collisions never hide it. */
  pinned?: boolean;
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
   * `legend:size` (bubble size legend), `placeholder` (frame outline and message of a chart below its
   * minimum size); `leader:c<j>` also links a scatter label to its point. Always set by `layoutChart`.
   */
  roles?: string[];
  /** Labels the engine chose to hide (or shrink) because they did not fit, with the reason. */
  hidden: string[];
  /** Value scale of charts with a value axis (same-scale groups, `layoutGroup`). */
  valueScale?: ValueScale;
  /** Non-fatal notes about the spec, e.g. an override that matches nothing after a data edit. */
  warnings?: string[];
}

/**
 * The value scale a chart drew with: its domain, how long one value unit is (points), and how long
 * it would be on its own (`own`; smaller `ptPerUnit` = the chart follows a same-scale group).
 */
export interface ValueScale {
  min: number; max: number; ptPerUnit: number; own: number;
  /** Where `min` and `max` are on the slide: [x1, y1, x2, y2] (the plot's value direction; drag handles). */
  at?: [number, number, number, number];
}
