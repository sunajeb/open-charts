import type { Ctx } from '../common.ts';
import { axisFormatter, fitCategories, legend, num, seriesLabelWidth, seriesLabelsRight, text, valueAxis, valueTicks } from '../common.ts';
import { niceTicks, roundToTotal, textOn } from '../numbers.ts';
import { overlaps } from '../place.ts';

/**
 * Mekko (Marimekko) chart: a stacked column chart whose column widths are proportional to the
 * category totals, so segment area encodes absolute value. Columns touch, separated by thin
 * background-coloured lines.
 *
 * - `options.mekko: 'percent'` (default) stacks every column to 100%; segment labels are shares,
 *   rounded together so each column reads exactly 100%. `'unit'` keeps absolute segment heights
 *   and labels values.
 * - Bold absolute totals sit on top of each column; each column's share of the grand total sits
 *   under its category label, since that is what the width encodes.
 * - Labels above and below narrow columns may borrow room from their neighbours; a label that
 *   would have to leave its own column is hidden and recorded instead.
 */
export function layoutMekko(ctx: Ctx) {
  const { spec, style, frame, o, fs, lh, pad, m } = ctx;
  const percent = (o.mekko ?? 'percent') === 'percent';
  const nCat = spec.categories.length;
  const cats = [...Array(nCat).keys()];
  const sers = spec.series.map((_, i) => i);
  const raw = (s: number, c: number) => Math.max(0, num(spec.series[s].values[c]) ?? 0);
  if (spec.series.some((s) => s.values.some((v) => (num(v) ?? 0) < 0))) ctx.hidden.push('negative values ignored in Mekko chart');

  const totals = cats.map((c) => sers.reduce((a, s) => a + raw(s, c), 0));
  const grand = totals.reduce((a, v) => a + v, 0);
  if (grand <= 0) { ctx.hidden.push('Mekko chart needs positive totals'); return; }
  const pctDecimals = o.decimals ?? 0;
  const shares = cats.map((c) => (totals[c] ? roundToTotal(sers.map((s) => (raw(s, c) / totals[c]) * 100), pctDecimals) : sers.map(() => 0)));
  const colShares = roundToTotal(totals.map((t) => (t / grand) * 100), 0);
  // Unit Mekko: segment labels are rounded together so they add up to the total on top.
  const rounded = cats.map((c) => (o.format ? null : roundToTotal(sers.map((s) => raw(s, c) / ctx.scale.div), ctx.decimals)));

  // Value domain: 0–100% or 0–max total.
  let dmax = percent ? 100 : Math.max(...totals);
  const ticks = o.valueAxis ? niceTicks(0, dmax, valueTicks(ctx)) : null;
  if (ticks) dmax = ticks[ticks.length - 1];
  const axisFmt = axisFormatter(ctx, ticks, percent);
  const axisW = ticks ? Math.max(...ticks.map((t) => m.width(axisFmt(t), fs))) + pad * 1.5 : 0;

  // A very narrow last column cannot anchor series labels on the right (they would seem to describe
  // a sliver): name the series in a legend instead, unless the spec asked for labels on the right.
  const px = frame.x + axisW;
  const narrowLast = (frame.w - axisW - seriesLabelWidth(ctx, 0)) * (totals[nCat - 1] / grand) < fs * 2;
  if (narrowLast && o.seriesLabels === 'right' && spec.options?.seriesLabels === undefined) ctx.o.seriesLabels = 'legend';

  // Reserve space: legend and totals on top, axis left, series labels right, categories + shares below.
  const legendH = legend(ctx);
  const showTotals = !!o.totals;
  // Without labels on the right, a narrow last column still needs room to centre its labels over it.
  const lastLabelW = Math.max(m.width(ctx.fmt(totals[nCat - 1]), fs, true), Math.min(fs * 6, m.width(spec.categories[nCat - 1], fs))) + 2;
  const rightW = ctx.o.seriesLabels === 'right' ? seriesLabelWidth(ctx, 0) : narrowLast ? lastLabelW / 2 : 0;
  const pw = frame.w - axisW - rightW;
  const colW = totals.map((t) => (pw * t) / grand);
  const x0 = cats.map((c) => px + colW.slice(0, c).reduce((a, w) => a + w, 0));
  const cx = (c: number) => x0[c] + colW[c] / 2;

  // Totals: bold, centred on their column, nudged sideways over narrow columns, kept a clear gap
  // apart. Totals that still collide are staggered into a second row above; beyond that, hidden.
  const totalItems = showTotals ? cats.filter((c) => totals[c]).map((c) => ({ c, cx: cx(c), w: m.width(ctx.fmt(totals[c]), fs, true) + 2 })) : [];
  const row1 = placeRow(totalItems, colW, px, frame.x + frame.w, pad * 1.5);
  const row2 = placeRow(totalItems.filter((it) => !row1.has(it.c)), colW, px, frame.x + frame.w, pad * 1.5);
  const totalsH = showTotals ? (lh + pad * 0.3) * (row2.size ? 2 : 1) : 0;

  // Category labels wrap inside their column; narrow columns wrap to a minimum width and borrow room.
  const minLabelW = Math.min(fs * 6, pw / nCat);
  const fit = fitCategories(ctx, cats.map((c) => Math.max(colW[c], minLabelW)));
  if (fit.size < fs) ctx.hidden.push(`category labels shrunk to ${fit.size} pt to fit`);
  const catLines = fit.lines;
  const shareText = cats.map((c) => ctx.pct(colShares[c], 0));
  const catH = Math.max(...catLines.map((l) => l.length)) * fit.lh + lh + pad * 0.6;

  const py = frame.y + legendH + totalsH;
  const ph = frame.y + frame.h - catH - py;
  if (ph < lh * 2) ctx.hidden.push('chart frame too small');
  const yOf = (v: number) => py + ph - (v / dmax) * ph;

  if (ticks) valueAxis(ctx, ticks, yOf, px, axisW, axisFmt);

  const lastSegs: { s: number; y: number }[] = [];
  const tops: number[] = [];

  for (const c of cats) {
    if (!totals[c]) { ctx.hidden.push(`${spec.categories[c]}: zero total, no column`); continue; }
    let up = 0;
    for (const s of sers) {
      const v = raw(s, c);
      if (!v) continue;
      const hv = percent ? (v / totals[c]) * 100 : v;
      const y1 = yOf(up + hv), y2 = yOf(up), h = y2 - y1;
      up += hv;
      const fill = ctx.color(s);
      ctx.add(`bar:s${s}:c${c}`, { kind: 'rect', x: x0[c], y: y1, w: colW[c], h, fill, stroke: style.background, strokeWidth: 0.75 });
      const label = percent ? ctx.pct(shares[c][s], pctDecimals) : rounded[c] ? ctx.fmt(rounded[c]![s], { scaled: true }) : ctx.fmt(v);
      const tw = m.width(label, fs);
      if (tw + 2 <= colW[c] && lh * 0.9 <= h) {
        text(ctx, `label:value:s${s}:c${c}`, { x: cx(c) - tw / 2 - 1, y: y1 + h / 2 - lh / 2, w: tw + 2, h: lh, text: label, color: textOn(fill), role: 'value' });
      } else ctx.hidden.push(`${spec.series[s].name} / ${spec.categories[c]}: segment too small`);
      if (c === nCat - 1) lastSegs.push({ s, y: y1 + h / 2 });
    }
    tops[c] = yOf(up);
  }

  const rowY = new Map<number, number>();
  for (const it of totalItems) {
    const p = row1.get(it.c);
    if (!p) continue;
    const y = tops[it.c] - lh - pad * 0.2;
    rowY.set(it.c, y);
    text(ctx, `label:total:c${it.c}`, { x: p.x, y, w: it.w, h: lh, text: ctx.fmt(totals[it.c]), bold: true, role: 'total' });
  }
  for (const it of totalItems) {
    const p = row2.get(it.c);
    if (!p) { if (!row1.has(it.c)) ctx.hidden.push(`${spec.categories[it.c]}: total label too wide for column`); continue; }
    // Above its own column and above any first-row total it reaches over.
    const below = [...row1].filter(([, q]) => q.x < p.x + it.w && p.x < q.x + q.w).map(([c]) => rowY.get(c)!);
    const y = Math.min(tops[it.c] - lh - pad * 0.2, ...below) - lh - pad * 0.3;
    text(ctx, `label:total:c${it.c}`, { x: p.x, y, w: it.w, h: lh, text: ctx.fmt(totals[it.c]), bold: true, role: 'total' });
  }

  // Baseline.
  const bottom = yOf(0);
  ctx.add('axis', { kind: 'line', x1: px, y1: bottom, x2: px + pw, y2: bottom, color: style.axis, width: 0.75 });

  // Category name with the column's share of the grand total beneath; falls back to the share alone.
  const catItems = cats.map((c) => ({
    c, cx: cx(c),
    w: fit.hidden.has(c) ? m.width(shareText[c], fs) + 2 : Math.max(m.width(catLines[c].join('\n'), fit.size), m.width(shareText[c], fs)) + 2,
    alt: m.width(shareText[c], fs) + 2,
  }));
  const catPos = placeRow(catItems, colW, px, frame.x + frame.w, pad);
  for (const c of cats) {
    const p = catPos.get(c);
    const y = bottom + pad * 0.4;
    if (!p) { ctx.hidden.push(`${spec.categories[c]}: category label too wide for column`); continue; }
    let sy = y;
    if (p.narrowed || fit.hidden.has(c)) ctx.hidden.push(`${spec.categories[c]}: category label too wide for column`);
    else {
      const h = catLines[c].length * fit.lh;
      text(ctx, `label:category:c${c}`, { x: p.x, y, w: p.w, h, text: catLines[c].join('\n'), valign: 'top', role: 'category', size: fit.size });
      sy += h;
    }
    text(ctx, `label:category:c${c}`, { x: p.x, y: sy, w: p.w, h: lh, text: shareText[c], color: style.muted, valign: 'top', role: 'decoration' });
  }

  if (o.seriesLabels === 'right') seriesLabelsRight(ctx, lastSegs.reverse(), px + pw, py, bottom);
}

export interface RowItem { c: number; cx: number; w: number; alt?: number }

/**
 * Lay a row of labels out along x, each centred on its column. Labels that would collide move
 * apart symmetrically (one left, the other right, as the manual places close segment labels); a label that would no longer cover its column
 * centre, or still collides, is narrowed to its `alt` width or dropped (narrowest column first).
 */
export function placeRow(items: RowItem[], colW: number[], min: number, max: number, gap: number) {
  const live = items.map((i) => ({ ...i, narrowed: false }));
  for (;;) {
    const xs = cluster(live.map((i) => i.cx), live.map((i) => i.w), min, max, gap);
    const boxes = live.map((it, k) => ({ x: xs[k] - it.w / 2, y: 0, w: it.w, h: 1 }));
    const bad = live.filter((it, k) =>
      Math.abs(xs[k] - it.cx) > it.w / 2 - 2 || boxes.some((b, j) => j !== k && overlaps(b, boxes[k], -gap / 2)));
    if (!bad.length) return new Map(live.map((it, k) => [it.c, { x: boxes[k].x, w: it.w, narrowed: it.narrowed }]));
    const worst = bad.reduce((a, b) => (colW[b.c] < colW[a.c] ? b : a));
    if (worst.alt != null && worst.alt < worst.w) { worst.w = worst.alt; worst.narrowed = true; }
    else live.splice(live.indexOf(worst), 1);
  }
}

/**
 * Spread labels of the given widths along one axis, keeping order and a gap between neighbours.
 * Labels that would collide form a cluster placed where it moves its members least (each as far
 * from its centre as the others), kept within [min, max]. Returns the new centres.
 */
export function cluster(centres: number[], widths: number[], min: number, max: number, gap: number): number[] {
  const order = centres.map((_, i) => i).sort((a, b) => centres[a] - centres[b]);
  type Group = { ids: number[]; start: number; size: number };
  const groups: Group[] = [];
  const place = (g: Group) => {
    let off = 0, sum = 0;
    for (const i of g.ids) { sum += centres[i] - (off + widths[i] / 2); off += widths[i] + gap; }
    g.size = off - gap;
    g.start = Math.max(min, Math.min(sum / g.ids.length, max - g.size));
  };
  for (const i of order) {
    const g = { ids: [i], start: 0, size: 0 };
    place(g);
    groups.push(g);
    while (groups.length > 1) {
      const a = groups[groups.length - 2], b = groups[groups.length - 1];
      if (a.start + a.size + gap <= b.start + 1e-9) break;
      a.ids.push(...b.ids);
      groups.pop();
      place(a);
    }
  }
  const out = centres.slice();
  for (const g of groups) {
    let x = g.start;
    for (const i of g.ids) { out[i] = x + widths[i] / 2; x += widths[i] + gap; }
  }
  return out;
}
