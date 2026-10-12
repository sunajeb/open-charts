import type { Ctx } from '../common.ts';
import { axisFormatter, categoryHeight, categoryLabels, legend, num, text, valueAxis } from '../common.ts';
import { niceTicks, textOn } from '../numbers.ts';
import { overlaps, type Box } from '../place.ts';

/**
 * Waterfall (bridge) chart.
 *
 * Spec convention
 * - Every series value is a delta: positive segments float upwards from the running sum, negative
 *   ones downwards. With several series the positives of a category stack upwards and the negatives
 *   downwards from the entry level; the next category continues from entry + net.
 * - `options.totalsAt` lists category indices drawn as totals: a column from the baseline to the
 *   running sum. Default: the first and the last category. A total at index 0 is the start value and
 *   its values seed the running sum; values of any later total are ignored (use null).
 * - Colours: totals palette[0]; single-series deltas style.positive / style.negative (or the series
 *   colour when given); multi-series deltas palette[1..] per series (or the series colour).
 * - Labels: signed delta labels inside the segment when they fit, else above (increase) / below
 *   (decrease); bold unsigned totals above (or below, if negative) the total column.
 */
export function layoutWaterfall(ctx: Ctx) {
  const { spec, style, frame, o, fs, lh, pad, m } = ctx;
  const nCat = spec.categories.length;
  const nSer = spec.series.length;
  const cats = [...Array(nCat).keys()];
  const sers = [...Array(nSer).keys()];
  const raw = (s: number, c: number) => num(spec.series[s].values[c]) ?? 0;
  // A cell "e" (as in the datasheet convention) marks its category as a computed total.
  const eCats = cats.filter((c) => sers.some((s) => spec.series[s].values[c] === 'e'));
  const totals = new Set((o.totalsAt ?? (eCats.length ? [0, ...eCats] : nCat > 1 ? [0, nCat - 1] : [0])).filter((i) => i >= 0 && i < nCat));
  if (nSer > 1 && o.seriesLabels === 'right') ctx.o.seriesLabels = 'legend';

  // Running sum: entry/exit level and the stacked segments of every column.
  interface Seg { s: number; v: number; lo: number; hi: number }
  const cols = cats.map((c) => ({ c, total: totals.has(c), entry: 0, exit: 0, segs: [] as Seg[] }));
  let level = 0;
  for (const col of cols) {
    const net = sers.reduce((a, s) => a + raw(s, col.c), 0);
    if (col.total) {
      if (col.c === 0) level = net; // the start value seeds the running sum
      col.exit = level;
      col.segs.push({ s: 0, v: level, lo: Math.min(0, level), hi: Math.max(0, level) });
      continue;
    }
    col.entry = level;
    let up = level, down = level;
    for (const s of sers) {
      const v = raw(s, col.c);
      if (v > 0) { col.segs.push({ s, v, lo: up, hi: up + v }); up += v; }
      else if (v < 0) { col.segs.push({ s, v, lo: down + v, hi: down }); down += v; }
    }
    if (!col.segs.length) col.segs.push({ s: 0, v: 0, lo: level, hi: level }); // zero delta still gets a "0" label
    level += net;
    col.exit = level;
  }

  // Value domain covers the baseline and every segment boundary.
  const bounds = cols.flatMap((col) => col.segs.flatMap((g) => [g.lo, g.hi]).concat(col.entry, col.exit));
  let dmin = Math.min(0, ...bounds), dmax = Math.max(0, ...bounds);
  if (dmax === dmin) dmax = dmin + 1;
  const ticks = o.valueAxis ? niceTicks(dmin, dmax) : null;
  if (ticks) [dmin, dmax] = [ticks[0], ticks[ticks.length - 1]];
  const axisFmt = axisFormatter(ctx, ticks);
  const axisW = ticks ? Math.max(...ticks.map((t) => m.width(axisFmt(t), fs))) + pad * 1.5 : 0;

  // Reserve space: legend on top, a label row above the highest bar, a label row below the lowest
  // bar when something dips below the baseline, categories at the bottom.
  const serColor = (s: number) =>
    (spec.series[s].color ?? (nSer > 1 ? style.palette[(s + 1) % style.palette.length] : '')).replace('#', '');
  // Legend swatches use the same colours as the delta bars (palette[0] is reserved for totals).
  const legendH = legend(ctx, 'box', (s) => serColor(s) || ctx.color(s));
  const labelH = lh + pad * 0.3;
  const belowH = dmin < 0 ? labelH : 0;
  const slot = (frame.w - axisW) / nCat;
  const colW = slot / (1 + o.gap);
  const catH = categoryHeight(ctx, slot);
  const px = frame.x + axisW;
  const pw = frame.w - axisW;
  const py = frame.y + legendH + labelH;
  const ph = frame.y + frame.h - catH - belowH - py;
  if (ph < lh * 2) ctx.hidden.push('chart frame too small');
  const yOf = (v: number) => py + ph - ((v - dmin) / (dmax - dmin)) * ph;
  const cx = (c: number) => px + slot * (c + 0.5);
  const xl = (c: number) => cx(c) - colW / 2;
  const plotBottom = yOf(dmin) + belowH;

  const segColor = (col: typeof cols[number], g: Seg) =>
    col.total ? style.palette[0].replace('#', '') : serColor(g.s) || (g.v >= 0 ? style.positive : style.negative).replace('#', '');

  if (ticks) valueAxis(ctx, ticks, yOf, px, axisW, axisFmt);

  // Bars.
  const bars: Box[] = [];
  for (const col of cols) {
    for (const g of col.segs) {
      const y = yOf(g.hi), h = Math.max(0.01, yOf(g.lo) - y);
      const box = { x: xl(col.c), y, w: colW, h };
      bars.push(box);
      ctx.add(`bar:s${g.s}:c${col.c}`, { kind: 'rect', ...box, fill: segColor(col, g), stroke: style.background, strokeWidth: 0.75 });
    }
  }

  // Connectors: thin horizontal line from the exit level of one column to the start of the next.
  for (let c = 0; c + 1 < nCat; c++) {
    const y = yOf(cols[c].exit);
    ctx.add(`connector:c${c}`, { kind: 'line', x1: xl(c) + colW, y1: y, x2: xl(c + 1), y2: y, color: style.axis, width: 0.5 });
  }

  // Baseline over the bars so columns crossing zero still show it.
  ctx.add('axis', { kind: 'line', x1: px, y1: yOf(0), x2: px + pw, y2: yOf(0), color: style.axis, width: 0.75 });
  categoryLabels(ctx, cats.map(cx), slot, plotBottom);

  // Labels: inside the segment when it fits, otherwise outside the column, never over a bar or another label.
  const placed: Box[] = [];
  const inFrame = (b: Box) => b.y >= frame.y - 0.5 && b.y + b.h <= plotBottom + 0.5 && b.x >= frame.x && b.x + b.w <= frame.x + frame.w;
  const free = (b: Box) => inFrame(b) && !placed.some((p) => overlaps(p, b)) && !bars.some((p) => overlaps(p, b, 0.5));
  for (const col of cols) {
    const { c } = col;
    let above = Math.min(...col.segs.map((g) => yOf(g.hi))) - pad * 0.3;
    let below = Math.max(...col.segs.map((g) => yOf(g.lo))) + pad * 0.3;
    // Segments farthest from the entry level go first, so outside labels stack away from the column.
    const reach = (g: Seg) => (g.v >= 0 ? g.hi - col.entry : col.entry - g.lo);
    const order = [...col.segs].sort((a, b) => reach(b) - reach(a));
    for (const g of order) {
      const label = col.total ? ctx.fmt(g.v) : ctx.fmt(g.v, { signed: true });
      const bold = col.total;
      const role = bold ? `label:total:c${c}` : `label:value:s${g.s}:c${c}`;
      const tw = m.width(label, fs, bold) + 2;
      const y1 = yOf(g.hi), h = yOf(g.lo) - y1;
      const fill = segColor(col, g);
      const inside = { x: cx(c) - tw / 2, y: y1 + h / 2 - lh / 2, w: tw, h: lh };
      if (!bold && tw <= colW && lh * 0.9 <= h && !placed.some((p) => overlaps(p, inside))) {
        placed.push(inside);
        text(ctx, role, { ...inside, text: label, color: textOn(fill), role: 'value' });
        continue;
      }
      const up = { x: cx(c) - tw / 2, y: above - lh, w: tw, h: lh };
      const dn = { x: cx(c) - tw / 2, y: below, w: tw, h: lh };
      const spot = (g.v >= 0 ? [up, dn] : [dn, up]).find(free);
      if (!spot) { ctx.hidden.push(`${spec.categories[c]}: no room for label "${label}"`); continue; }
      placed.push(spot);
      if (spot === up) above = spot.y - pad * 0.15; else below = spot.y + lh + pad * 0.15;
      text(ctx, role, { ...spot, text: label, bold, role: bold ? 'total' : 'value' });
    }
  }
}
