import type { Ctx } from '../common.ts';
import { categoryBreakSymbol, slotLayout, slotUnits } from '../decorations.ts';
import {
  axisFormatter, barCategories, barCategoryLabels, bottomAxisHeight, bottomValueAxis, categoryHeight, categoryLabels, legend, num, text, valueAxis, valueTicks, axisWidth, sharedScale,
} from '../common.ts';
import { niceTicks, textOn } from '../numbers.ts';
import { overlaps, type Box } from '../place.ts';
import { placePending, type PendingLabel } from '../labels.ts';
import type { RectShape } from '../types.ts';
import { sideways } from '../direction.ts';

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
  if (sideways(ctx.o)) { layoutWaterfallBars(ctx, waterfallModel(ctx)); return; }
  const { spec, style, frame, o, fs, lh, pad, m } = ctx;
  const md = waterfallModel(ctx);
  const { nCat, cats, cols, segColor } = md;
  let { dmin, dmax, ticks, axisFmt } = md;
  const axisW = axisWidth(ctx, dmin, ticks, axisFmt);

  // Reserve space: legend on top, a label row above the highest bar, a label row below the lowest
  // bar when something dips below the baseline, categories at the bottom.
  // Legend swatches use the same colours as the delta bars (palette[0] is reserved for totals).
  const legendH = legend(ctx, 'box', (s) => md.serColor(s) || ctx.color(s));
  const labelH = lh + pad * 0.3;
  const belowH = dmin < 0 ? labelH : 0;
  const slot = (frame.w - axisW) / slotUnits(nCat, o.categoryGaps);
  const colW = slot / (1 + o.gap);
  const catH = categoryHeight(ctx, slot);
  const px = frame.x + axisW;
  const pw = frame.w - axisW;
  const py = frame.y + legendH + labelH;
  const ph = frame.y + frame.h - catH - belowH - py;
  ({ dmin, dmax, ticks } = sharedScale(ctx, dmin, dmax, ph, ticks));
  if (ticks) axisFmt = axisFormatter(ctx, ticks);
  if (ph < lh * 2) ctx.hidden.push('chart frame too small');
  const yOf = (v: number) => py + ph - ((v - dmin) / (dmax - dmin)) * ph;
  const cx = slotLayout(nCat, o.categoryGaps, px, pw).centre;
  const xl = (c: number) => cx(c) - colW / 2;
  const plotBottom = yOf(dmin) + belowH;

  if (ticks) valueAxis(ctx, ticks, yOf, px, axisW, axisFmt);

  // Bars.
  const bars: Box[] = [];
  const segRect = new Map<Seg, RectShape>();
  for (const col of cols) {
    for (const g of col.segs) {
      const y = yOf(g.hi), h = Math.max(0.01, yOf(g.lo) - y);
      const box = { x: xl(col.c), y, w: colW, h };
      bars.push(box);
      const rect = { kind: 'rect', ...box, fill: segColor(col, g), stroke: style.background, strokeWidth: 0.75 } as const;
      segRect.set(g, rect);
      ctx.add(`bar:s${g.s}:c${col.c}`, rect);
    }
  }

  // Connectors: thin horizontal line from the exit level of one column to the start of the next.
  for (let c = 0; c + 1 < nCat; c++) {
    const y = yOf(cols[c].exit);
    ctx.add(`connector:c${c}`, { kind: 'line', x1: xl(c) + colW, y1: y, x2: xl(c + 1), y2: y, color: style.axis, width: 0.5 });
  }

  // Baseline over the bars so columns crossing zero still show it.
  ctx.add('axis', { kind: 'line', x1: px, y1: yOf(0), x2: px + pw, y2: yOf(0), color: style.axis, width: 0.75 });
  for (const i of new Set(o.categoryBreaks ?? [])) {
    if (i >= 0 && i < nCat - 1) categoryBreakSymbol(ctx, `break:cat:${i}`, false, (cx(i) + cx(i + 1)) / 2, yOf(0));
  }
  categoryLabels(ctx, cats.map(cx), slot, plotBottom);

  // Labels: inside the segment when it fits, otherwise outside the column, never over a bar or another label.
  const placed: Box[] = [];
  const pending: PendingLabel[] = [];
  const inFrame = (b: Box) => b.y >= frame.y - 0.5 && b.y + b.h <= plotBottom + 0.5 && b.x >= frame.x && b.x + b.w <= frame.x + frame.w;
  const free = (b: Box) => inFrame(b) && !placed.some((p) => overlaps(p, b)) && !bars.some((p) => overlaps(p, b, 0.5));
  for (const col of cols) {
    const { c } = col;
    const top = Math.min(...col.segs.map((g) => yOf(g.hi))), bottom = Math.max(...col.segs.map((g) => yOf(g.lo)));
    let above = top - pad * 0.3, below = bottom + pad * 0.3, upFree = true, dnFree = true;
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
      // Above / below the column only for the outermost segment on that side, before any other label
      // takes that spot: anything further away would need a leader across other segments.
      const up = upFree && yOf(g.hi) <= top + 0.01 ? { x: cx(c) - tw / 2, y: above - lh, w: tw, h: lh } : null;
      const dn = dnFree && yOf(g.lo) >= bottom - 0.01 ? { x: cx(c) - tw / 2, y: below, w: tw, h: lh } : null;
      const spot = (g.v >= 0 ? [up, dn] : [dn, up]).find((b): b is Box => !!b && free(b));
      if (!spot) {
        // Neither above nor below: on the segment (boxed) or beside the column with a leader.
        pending.push({ s: g.s, c, text: label, seg: segRect.get(g)!, sides: ['right', 'left'], boxCover: bold ? undefined : col.segs.map((x) => segRect.get(x)!), reason: `${spec.categories[c]}: no room for label "${label}"`, bold, role });
        continue;
      }
      placed.push(spot);
      if (spot === up) { above = spot.y - pad * 0.15; upFree = false; } else { below = spot.y + lh + pad * 0.15; dnFree = false; }
      text(ctx, role, { ...spot, text: label, bold, role: bold ? 'total' : 'value' });
    }
  }
  placePending(ctx, pending);
}

/** Running sum, colours and value domain, shared by columns and bars. */
function waterfallModel(ctx: Ctx) {
  const { spec, style, o } = ctx;
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
  const ticks = o.valueAxis ? niceTicks(dmin, dmax, valueTicks(ctx)) : null;
  if (ticks) [dmin, dmax] = [ticks[0], ticks[ticks.length - 1]];
  const axisFmt = axisFormatter(ctx, ticks);
  const serColor = (s: number) =>
    (spec.series[s].color ?? (nSer > 1 ? style.palette[(s + 1) % style.palette.length] : '')).replace('#', '');
  const segColor = (col: typeof cols[number], g: Seg) =>
    col.total ? style.palette[0].replace('#', '') : serColor(g.s) || (g.v >= 0 ? style.positive : style.negative).replace('#', '');
  return { nCat, nSer, cats, sers, cols, dmin, dmax, ticks, axisFmt, serColor, segColor };
}
type Model = ReturnType<typeof waterfallModel>;
/** One segment of a waterfall column: series, signed value and its extent on the value axis. */
interface Seg { s: number; v: number; lo: number; hi: number }

/**
 * Waterfall bars (direction right; left is this mirrored): categories down the left, the running sum
 * growing to the right, connectors running down from one bar's end to the next bar. A label sits in
 * its segment when it fits, else beyond the bar's end in the direction of the step (right of an
 * increase, left of a decrease), else above or below the bar.
 */
function layoutWaterfallBars(ctx: Ctx, md: Model) {
  const { spec, style, frame, o, fs, lh, pad, m } = ctx;
  const { nCat, cols, segColor } = md;
  let { dmin, dmax, ticks, axisFmt } = md;
  type Col = Model['cols'][number];
  const legendH = legend(ctx, 'box', (s) => md.serColor(s) || ctx.color(s));
  const axisH = bottomAxisHeight(ctx, ticks);
  const { catW, fit } = barCategories(ctx);
  const labelOf = (col: Col, g: Seg) => (col.total ? ctx.fmt(g.v) : ctx.fmt(g.v, { signed: true }));
  const widthOf = (col: Col, g: Seg) => m.width(labelOf(col, g), fs, col.total) + 2;

  const py = frame.y + legendH;
  const ph = frame.y + frame.h - axisH - py;
  const slot = ph / slotUnits(nCat, o.categoryGaps);
  const barT = slot / (1 + o.gap);
  const cy = slotLayout(nCat, o.categoryGaps, py, ph).centre;
  const fitsIn = (col: Col, g: Seg, len: number) => !col.total && widthOf(col, g) <= len && lh * 0.9 <= barT;

  // Plot between the category column and the room the outside labels need at either end. That room
  // depends on which labels fit inside, which depends on the scale: settle it in a few rounds.
  const x0 = frame.x + catW + pad;
  let left = 0, right = 0, px = x0, pw = frame.x + frame.w - x0;
  for (let round = 0; round < 4; round++) {
    px = x0 + left;
    pw = frame.x + frame.w - right - px;
    const at = (v: number) => px + ((v - dmin) / (dmax - dmin)) * pw;
    let needL = 0, needR = 0;
    for (const col of cols) {
      const hi = Math.max(...col.segs.map((g) => g.hi)), lo = Math.min(...col.segs.map((g) => g.lo));
      for (const g of col.segs) {
        if (fitsIn(col, g, at(g.hi) - at(g.lo))) continue;
        const w = widthOf(col, g) + pad * 0.4;
        if (g.v >= 0) needR = Math.max(needR, at(hi) + w - (px + pw));
        else needL = Math.max(needL, px - (at(lo) - w));
      }
    }
    const nl = Math.max(0, left + needL), nr = Math.max(0, right + needR);
    if (Math.abs(nl - left) < 0.5 && Math.abs(nr - right) < 0.5) break;
    [left, right] = [nl, nr];
  }
  ({ dmin, dmax, ticks } = sharedScale(ctx, dmin, dmax, pw, ticks));
  if (ticks) axisFmt = axisFormatter(ctx, ticks);
  if (pw < lh * 2 || ph < lh * nCat * 0.5) ctx.hidden.push('chart frame too small');
  const xOf = (v: number) => px + ((v - dmin) / (dmax - dmin)) * pw;
  const yA = py + ph;

  if (ticks) bottomValueAxis(ctx, ticks, xOf, yA, axisFmt);

  const segRect = new Map<Seg, RectShape>();
  for (const col of cols) {
    for (const g of col.segs) {
      const x = xOf(g.lo);
      const rect = { kind: 'rect', x, y: cy(col.c) - barT / 2, w: Math.max(0.01, xOf(g.hi) - x), h: barT, fill: segColor(col, g), stroke: style.background, strokeWidth: 0.75 } as const;
      segRect.set(g, rect);
      ctx.add(`bar:s${g.s}:c${col.c}`, rect);
    }
  }
  // Connectors: from the exit level of one bar down to the next.
  for (let c = 0; c + 1 < nCat; c++) {
    const x = xOf(cols[c].exit);
    ctx.add(`connector:c${c}`, { kind: 'line', x1: x, y1: cy(c) + barT / 2, x2: x, y2: cy(c + 1) - barT / 2, color: style.axis, width: 0.5 });
  }
  const xBase = xOf(0);
  ctx.add('axis', { kind: 'line', x1: xBase, y1: py, x2: xBase, y2: yA, color: style.axis, width: 0.75 });
  for (const i of new Set(o.categoryBreaks ?? [])) {
    if (i >= 0 && i < nCat - 1) categoryBreakSymbol(ctx, `break:cat:${i}`, true, (cy(i) + cy(i + 1)) / 2, xBase);
  }
  barCategoryLabels(ctx, fit, catW, cy);

  // Labels: inside, else beyond the bar's end (the outermost segment on that side only), else above or below.
  const placed: Box[] = [];
  const pending: PendingLabel[] = [];
  const bars = [...segRect.values()];
  const inFrame = (b: Box) => b.x >= frame.x + catW - 0.5 && b.x + b.w <= frame.x + frame.w + 0.5 && b.y >= py - 0.5 && b.y + b.h <= yA + 0.5;
  const free = (b: Box) => inFrame(b) && !placed.some((p) => overlaps(p, b)) && !bars.some((p) => overlaps(p, b, 0.5));
  for (const col of cols) {
    const { c } = col;
    const hiX = Math.max(...col.segs.map((g) => xOf(g.hi))), loX = Math.min(...col.segs.map((g) => xOf(g.lo)));
    let rightEnd = hiX + pad * 0.3, leftEnd = loX - pad * 0.3, rFree = true, lFree = true;
    const reach = (g: Seg) => (g.v >= 0 ? g.hi - col.entry : col.entry - g.lo);
    for (const g of [...col.segs].sort((a, b) => reach(b) - reach(a))) {
      const label = labelOf(col, g);
      const bold = col.total;
      const role = bold ? `label:total:c${c}` : `label:value:s${g.s}:c${c}`;
      const tw = widthOf(col, g);
      const x1 = xOf(g.lo), len = xOf(g.hi) - x1;
      const inside = { x: x1 + len / 2 - tw / 2, y: cy(c) - lh / 2, w: tw, h: lh };
      if (fitsIn(col, g, len) && !placed.some((p) => overlaps(p, inside))) {
        placed.push(inside);
        text(ctx, role, { ...inside, text: label, color: textOn(segColor(col, g)), role: 'value' });
        continue;
      }
      const r = rFree && xOf(g.hi) >= hiX - 0.01 ? { x: rightEnd, y: cy(c) - lh / 2, w: tw, h: lh } : null;
      const l = lFree && xOf(g.lo) <= loX + 0.01 ? { x: leftEnd - tw, y: cy(c) - lh / 2, w: tw, h: lh } : null;
      const spot = (g.v >= 0 ? [r, l] : [l, r]).find((b): b is Box => !!b && free(b));
      if (!spot) {
        pending.push({ s: g.s, c, text: label, seg: segRect.get(g)!, sides: ['above', 'below'], boxCover: bold ? undefined : col.segs.map((x) => segRect.get(x)!), reason: `${spec.categories[c]}: no room for label "${label}"`, bold, role });
        continue;
      }
      placed.push(spot);
      if (spot === r) { rightEnd = spot.x + tw + pad * 0.15; rFree = false; } else { leftEnd = spot.x - pad * 0.15; lFree = false; }
      text(ctx, role, { ...spot, text: label, bold, align: spot === r ? 'left' : 'right', role: bold ? 'total' : 'value' });
    }
  }
  placePending(ctx, pending);
}
