// Optional rows: a series marked `hidden` stays in the spec (and the datasheet) but is left out of the
// chart. Chart types never see hidden series: layoutChart lays out the visible ones, then maps every
// role back to the series' index in the spec, so roles, tags and overrides keep naming the same row.
import type { ChartSpec, LevelEnd, Style } from './types.ts';

/** Palette offset per chart type: series `s` takes `palette[s + offset]` (waterfall keeps palette[0] for totals, Mekko XY's width series has no colour). */
const PALETTE_OFFSET: Partial<Record<ChartSpec['type'], number>> = {
  stacked: 0, clustered: 0, percent: 0, mekko: 0, line: 0, area: 0, area100: 0, combo: 0, scatter: 0, bubble: 0,
  waterfall: 1, mekkoxy: -1,
};

export interface Visible {
  spec: ChartSpec;
  style: Style;
  /** Index in the original spec of each visible series. */
  index: number[];
  /** Decorations left out because they point at a hidden series (for `Scene.hidden`). */
  notes: string[];
}

/**
 * The spec as the chart types see it: hidden series removed. Colours stay put — each visible series
 * keeps the palette colour of its original row (the palette is re-ordered for the visible rows) — and
 * waterfall categories marked "e" only in a hidden row stay totals; `secondaryAxis` and level arrows
 * follow the visible rows (an arrow at a hidden row is left out). Without hidden series the inputs
 * are returned unchanged.
 */
export function visibleSeries(spec: ChartSpec, style: Style): Visible {
  const index = spec.series.flatMap((s, i) => (s.hidden ? [] : [i]));
  if (index.length === spec.series.length) return { spec, style, index, notes: [] };
  const series = index.map((i) => { const { hidden: _, ...s } = spec.series[i]; return { ...s, values: s.values.slice() }; });
  if (spec.type === 'waterfall') {
    spec.categories.forEach((_, c) => {
      if (spec.series.some((s) => s.values[c] === 'e') && !series.some((s) => s.values[c] === 'e')) series[0].values[c] = 'e';
    });
  }
  const off = PALETTE_OFFSET[spec.type];
  let palette = style.palette;
  if (off !== undefined && palette.length) {
    const n = palette.length;
    palette = palette.slice();
    index.forEach((orig, k) => {
      if (k + off < 0 || orig + off < 0) return;
      palette[k + off] = style.palette[(orig + off) % n];
    });
    // A palette longer than the original would wrap differently: never grow it past the slots used.
    for (let j = 0; j < palette.length; j++) palette[j] ??= style.palette[j % n];
  }
  // Options that name series by index follow the visible rows; what points at a hidden row is left out.
  const notes: string[] = [];
  const at = (i: number) => index.indexOf(i);
  let options = spec.options;
  if (options?.secondaryAxis) options = { ...options, secondaryAxis: options.secondaryAxis.filter((i) => at(i) >= 0).map(at) };
  if (options?.levelArrows) {
    const end = (e: LevelEnd): LevelEnd | null => ('series' in e && e.series != null ? (at(e.series) >= 0 ? { ...e, series: at(e.series) } : null) : e);
    options = { ...options, levelArrows: options.levelArrows.flatMap((a, k) => {
      const from = end(a.from), to = end(a.to);
      if (from && to) return [{ ...a, from, to }];
      notes.push(`level arrow ${k + 1}: points at a hidden series`);
      return [];
    }) };
  }
  return { spec: { ...spec, series, ...(options ? { options } : {}) }, style: palette === style.palette ? style : { ...style, palette }, index, notes };
}

/** Map `s<k>` in roles (visible index) back to `s<i>` (index in the spec), in place. */
export function restoreRoles(roles: string[], index: number[]) {
  if (index.every((i, k) => i === k)) return;
  for (let r = 0; r < roles.length; r++) roles[r] = roles[r].replace(/(^|:)s(\d+)(?=:|$)/g, (_, pre, k) => `${pre}s${index[+k] ?? k}`);
}
