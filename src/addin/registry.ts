// The chart gallery: the tiles the ribbon gallery and the editor's chart-type gallery show, in the
// reference product's order (5 per row), each with its icon (icons.ts), label and the directions it
// can be inserted in (DIRECTIONS in core/spec.ts). Engine types the catalogue does not list are added
// after a rebuild. A tile is insertable when the engine supports its type and there is default data
// (src/core/samples.ts, or the tile's own `sample`); otherwise clicking it explains that it is not
// available yet.
import { CHART_TYPES, directionsOf, validateSpec, type Direction } from '../core/spec.ts';
import { samples } from '../core/samples.ts';
import type { ChartSpec } from '../core/types.ts';
import { ICONS, iconSvg } from './icons.ts';

export interface ChartTypeEntry {
  /** Tile id, unique; also what the ribbon sends to the engine ("<id>" or "<id>@<direction>"). */
  id: string;
  /** Engine chart type. */
  type: string;
  /** Options merged over the sample (variants such as the unit-axis Mekko). */
  options?: Record<string, unknown>;
  /** Default data when the type's sample does not fit the tile (build-up / build-down waterfall). */
  sample?: ChartSpec;
  /** Label under the icon; "\n" marks where a two-line label breaks. */
  label: string;
  /** Icon name in icons.ts (default: the id). */
  icon?: string;
  /** Gallery section (screen readers, the ribbon's tooltips). */
  family?: string;
  /** One line shown in the tooltip. */
  description?: string;
}

const years = ['2021', '2022', '2023', '2024', '2025'];
const waterfallSample = (categories: string[], values: (number | 'e')[], totalsAt?: number[]): ChartSpec => ({
  specVersion: 1, type: 'waterfall', categories,
  series: [{ name: 'Value', values, color: 'accent2' }],
  options: { unit: 'm', ...(totalsAt ? { totalsAt } : {}) },
});

/** Catalogue order = gallery order: the reference's 16 tiles, then the other open-charts types. */
export const CATALOGUE: ChartTypeEntry[] = [
  { id: 'stacked', type: 'stacked', label: 'Stacked\nColumn/Bar', family: 'Column and bar', description: 'Segments stacked per category, totals on top; CAGR arrow.' },
  { id: 'percent', type: 'percent', label: '100%\nColumn/Bar', family: 'Column and bar', description: 'Each column is 100%; segments show the share.' },
  { id: 'clustered', type: 'clustered', label: 'Clustered\nColumn/Bar', family: 'Column and bar', description: 'Series side by side per category.' },
  {
    id: 'waterfall', type: 'waterfall', label: 'Build-Up\nWaterfall', family: 'Waterfall', description: 'Components that build up to a total.',
    sample: waterfallSample(['Materials', 'Labour', 'Energy', 'Logistics', 'Overhead', 'Total cost'], [38, 24, 16, 12, 10, 'e'], [5]),
  },
  {
    id: 'waterfallDown', type: 'waterfall', label: 'Build-Down\nWaterfall', family: 'Waterfall', description: 'A total broken down into its components, step by step.',
    sample: waterfallSample(['Revenue', 'Cost of sales', 'Operating expenses', 'Depreciation', 'Interest', 'Tax', 'Net income'], [420, -180, -95, -30, -15, -25, 'e']),
  },
  { id: 'mekko', type: 'mekko', label: 'Mekko\n(% Axis)', family: 'Mekko', description: 'Column width = category total, segments = share of the column.' },
  { id: 'combo', type: 'combo', label: 'Combination', family: 'Column and bar', description: 'Stacked columns with a line series on top.' },
  { id: 'line', type: 'line', label: 'Line/Profile', family: 'Line and area', description: 'Values over time, one line per series; sideways it is a profile chart.' },
  { id: 'area', type: 'area', label: 'Stacked\nArea', family: 'Line and area', description: 'Stacked areas over time.' },
  { id: 'area100', type: 'area100', label: '100%\nArea', family: 'Line and area', description: 'Stacked areas, each category is 100%.' },
  { id: 'mekkoUnit', type: 'mekko', options: { mekko: 'unit', valueAxis: true }, label: 'Mekko\n(Unit Axis)', family: 'Mekko', description: 'Column width = category total, segments in units on a value axis.' },
  { id: 'pie', type: 'pie', label: 'Pie', family: 'Pie', description: 'Shares of a whole as wedges.' },
  { id: 'doughnut', type: 'doughnut', label: 'Doughnut', family: 'Pie', description: 'Ring of shares with the total in the middle.' },
  { id: 'scatter', type: 'scatter', label: 'Scatter', family: 'XY', description: 'Points by two values; labels next to points.' },
  { id: 'bubble', type: 'bubble', label: 'Bubble', family: 'XY', description: 'Scatter with bubble size as a third value.' },
  { id: 'gantt', type: 'gantt', label: 'Gantt/\nTimeline', family: 'Timeline and table', description: 'Tasks and milestones on a date scale.' },
  { id: 'mekkoxy', type: 'mekkoxy', label: 'Mekko XY', family: 'Mekko', description: 'Bar width and height both carry a value (cost curve).' },
  { id: 'harvey', type: 'harvey', label: 'Harvey\nBalls', family: 'Timeline and table', description: 'Harvey balls rating options against criteria.' },
  { id: 'table', type: 'table', label: 'Table', family: 'Timeline and table', description: 'Formatted number table.' },
];

const label = (type: string) => type.charAt(0).toUpperCase() + type.slice(1).replace(/([a-z])([A-Z])/g, '$1 $2');

/** Gallery entries: the catalogue, then engine types the catalogue does not list yet. */
export function galleryItems(): ChartTypeEntry[] {
  const known = new Set(CATALOGUE.map((c) => c.type));
  const extra = [...CHART_TYPES, ...Object.keys(samples)].filter((t, i, a) => !known.has(t) && a.indexOf(t) === i);
  return [...CATALOGUE, ...extra.map((type) => ({ id: type, type, label: label(type), family: 'More' }))];
}

/** The icon a tile shows (a type without its own icon borrows a generic one). */
export const iconOf = (e: ChartTypeEntry) => (ICONS[e.icon ?? e.id] ? e.icon ?? e.id : 'stacked');

/** Directions a tile can be inserted in (its type's, `['up']` when it has no arrows). */
export const tileDirections = (e: ChartTypeEntry): readonly Direction[] => directionsOf(e.type);

/** One-line label (the "\n" of a two-line label becomes a space, "Gantt/\nTimeline" stays joined). */
export const flatLabel = (e: ChartTypeEntry) => e.label.replace(/\/\n/g, '/').replace(/\n/g, ' ');

const entry = (id: string): ChartTypeEntry =>
  galleryItems().find((g) => g.id === id) ?? { id, type: id, label: label(id) };

/**
 * Default data for a gallery item — "<id>" or "<id>@<direction>" (the ribbon's direction buttons) —
 * or why it cannot be inserted. A direction the tile's type cannot take is ignored (inserted up).
 */
export function sampleFor(item: string): { spec: ChartSpec } | { error: string } {
  const [id, dir] = item.split('@');
  const e = entry(id);
  const supported = (CHART_TYPES as readonly string[]).includes(e.type);
  const sample = e.sample ?? (samples as Record<string, ChartSpec | undefined>)[e.type];
  const name = flatLabel(e);
  if (!supported || !sample) return { error: `${name} charts are not available in this version of open-charts yet.` };
  const spec = structuredClone(sample);
  if (e.options) spec.options = { ...spec.options, ...structuredClone(e.options) } as ChartSpec['options'];
  if (dir && dir !== 'up' && tileDirections(e).includes(dir as Direction)) spec.options = { ...spec.options, direction: dir as Direction };
  // A variant whose options this engine does not know yet.
  if (validateSpec(spec).length) return { error: `${name} charts are not available in this version of open-charts yet.` };
  return { spec };
}

/** Gallery items that can be inserted right now (used by the end-to-end test). */
export function insertableItems(): string[] {
  return galleryItems().map((g) => g.id).filter((id) => 'spec' in sampleFor(id));
}

const TYPE_NAMES: Record<string, string> = { percent: '100%', area100: '100% area', combo: 'Combination', mekkoxy: 'Mekko XY', harvey: 'Harvey ball' };

/** Display name of an engine chart type ("Waterfall", "Mekko XY"). */
export function typeLabel(type: string): string {
  return TYPE_NAMES[type] ?? label(type);
}

/** The gallery as the editor shows it: tiles with their inline SVG icon and directions (server.ts sends it). */
export function editorGallery() {
  return galleryItems().map((e) => ({
    id: e.id, type: e.type, label: e.label, options: e.options ?? {}, directions: tileDirections(e),
    description: e.description ?? '', icon: iconSvg(iconOf(e)),
  }));
}
