// Chart types the add-in offers. The ribbon gallery is generated from this list plus any type the
// engine knows (CHART_TYPES) that is not listed here, so new chart types appear after a rebuild.
// A type is insertable when the engine supports it and src/core/samples.ts has default data for it;
// otherwise clicking it explains that it is not available yet.
import { CHART_TYPES, validateSpec } from '../core/spec.ts';
import { samples } from '../core/samples.ts';
import type { ChartSpec } from '../core/types.ts';

export interface ChartTypeEntry {
  /** Gallery item id, unique; also what the ribbon sends to the engine. */
  id: string;
  /** Engine chart type. */
  type: string;
  /** Options merged over the type's sample (variants such as step line). */
  options?: Record<string, unknown>;
  label: string;
  /** Built-in Office icon. */
  imageMso: string;
}

/** Catalogue order = gallery order (tasks/todo.md chart catalogue). */
export const CATALOGUE: ChartTypeEntry[] = [
  { id: 'stacked', type: 'stacked', label: 'Stacked', imageMso: 'ChartTypeColumnInsertGallery' },
  { id: 'clustered', type: 'clustered', label: 'Clustered', imageMso: 'ChartTypeColumnInsertGallery' },
  { id: 'percent', type: 'percent', label: '100%', imageMso: 'ChartTypeColumnInsertGallery' },
  { id: 'percentAxis', type: 'percent', options: { valueAxis: true }, label: '100% with % axis', imageMso: 'ChartTypeColumnInsertGallery' },
  { id: 'waterfall', type: 'waterfall', label: 'Waterfall', imageMso: 'ChartTypeColumnInsertGallery' },
  { id: 'combo', type: 'combo', label: 'Combination', imageMso: 'ChartTypeOtherInsertGallery' },
  { id: 'line', type: 'line', label: 'Line', imageMso: 'ChartTypeLineInsertGallery' },
  { id: 'step', type: 'line', options: { step: true }, label: 'Step line', imageMso: 'ChartTypeLineInsertGallery' },
  { id: 'area', type: 'area', label: 'Area', imageMso: 'ChartTypeAreaInsertGallery' },
  { id: 'area100', type: 'area100', label: '100% area', imageMso: 'ChartTypeAreaInsertGallery' },
  { id: 'mekko', type: 'mekko', label: 'Mekko', imageMso: 'ChartTypeBarInsertGallery' },
  { id: 'mekkoxy', type: 'mekkoxy', label: 'Mekko XY', imageMso: 'ChartTypeBarInsertGallery' },
  { id: 'scatter', type: 'scatter', label: 'Scatter', imageMso: 'ChartTypeXYScatterInsertGallery' },
  { id: 'bubble', type: 'bubble', label: 'Bubble', imageMso: 'ChartTypeXYScatterInsertGallery' },
  { id: 'pie', type: 'pie', label: 'Pie', imageMso: 'ChartTypePieInsertGallery' },
  { id: 'doughnut', type: 'doughnut', label: 'Doughnut', imageMso: 'ChartTypePieInsertGallery' },
  { id: 'gantt', type: 'gantt', label: 'Gantt', imageMso: 'ChartTypeBarInsertGallery' },
  { id: 'harvey', type: 'harvey', label: 'Harvey ball', imageMso: 'ChartTypePieInsertGallery' },
  { id: 'table', type: 'table', label: 'Table', imageMso: 'TableInsert' },
];

const label = (type: string) => type.charAt(0).toUpperCase() + type.slice(1).replace(/([a-z])([A-Z])/g, '$1 $2');

/** Gallery entries: the catalogue, then engine types the catalogue does not list yet. */
export function galleryItems(): ChartTypeEntry[] {
  const known = new Set(CATALOGUE.map((c) => c.type));
  const extra = [...CHART_TYPES, ...Object.keys(samples)].filter((t, i, a) => !known.has(t) && a.indexOf(t) === i);
  return [...CATALOGUE, ...extra.map((type) => ({ id: type, type, label: label(type), imageMso: 'ChartInsert' }))];
}

const entry = (id: string): ChartTypeEntry =>
  galleryItems().find((g) => g.id === id) ?? { id, type: id, label: label(id), imageMso: 'ChartInsert' };

/** Default data for a gallery item (its type's sample plus the item's options), or why it cannot be inserted. */
export function sampleFor(id: string): { spec: ChartSpec } | { error: string } {
  const e = entry(id);
  const supported = (CHART_TYPES as readonly string[]).includes(e.type);
  const sample = (samples as Record<string, ChartSpec | undefined>)[e.type];
  if (!supported || !sample) return { error: `${e.label} charts are not available in this version of open-charts yet.` };
  const spec = structuredClone(sample);
  if (e.options) spec.options = { ...spec.options, ...structuredClone(e.options) } as ChartSpec['options'];
  // A variant whose options this engine does not know yet (e.g. step line before it lands).
  if (validateSpec(spec).length) return { error: `${e.label} charts are not available in this version of open-charts yet.` };
  return { spec };
}

/** Gallery items that can be inserted right now (used by the end-to-end test). */
export function insertableItems(): string[] {
  return galleryItems().map((g) => g.id).filter((id) => 'spec' in sampleFor(id));
}
