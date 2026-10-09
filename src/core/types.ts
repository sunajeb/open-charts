// All geometry is in points (1/72 inch). Origin is the slide's top-left corner.

export type ChartType = 'stacked' | 'clustered' | 'percent' | 'waterfall' | 'mekko' | 'line';

export interface Series {
  name: string;
  values: (number | null)[];
  color?: string;
}

export interface ChartOptions {
  /** Total label above each stacked column. Default: on for stacked, off otherwise. */
  totals?: boolean;
  /** Compound annual growth arrow between two category indices (uses column totals). */
  cagr?: { from: number; to: number } | null;
  /** Difference arrow between the totals of two categories. */
  difference?: { from: number; to: number } | null;
  /** Show a value axis with ticks. Default: off (values are labelled on the chart). */
  valueAxis?: boolean;
  /** Force the value axis to include zero. Default: on, except narrow-range line charts. */
  zeroBased?: boolean;
  /** Fixed number of decimals. Default: inferred from the data. */
  decimals?: number;
  /** Suffix appended to every value label, e.g. "m" or "k". */
  unit?: string;
  /** Gap between columns as a fraction of the column width. Default 0.5. */
  gap?: number;
  /** Series identification: labels to the right of the last column, a legend, or none. */
  seriesLabels?: 'right' | 'legend' | 'none';
}

export interface ChartSpec {
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
  /** Draw as an ellipse inscribed in the box (line chart markers). */
  ellipse?: boolean;
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
  /** Optional box behind the text (e.g. CAGR bubble). */
  fill?: string;
  stroke?: string;
  rounded?: boolean;
  /** Role used by quality checks (labels must not overlap each other). */
  role?: 'value' | 'total' | 'category' | 'series' | 'axis' | 'decoration' | 'title';
}

export type Shape = RectShape | LineShape | TextShape;

export interface Scene {
  frame: Frame;
  shapes: Shape[];
  /** Labels the engine chose to hide because they did not fit. */
  hidden: string[];
}
