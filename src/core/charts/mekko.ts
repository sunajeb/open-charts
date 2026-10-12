import type { Ctx } from '../common.ts';
import { axisFormatter, fitCategories, legend, num, seriesLabelWidth, seriesLabelsRight, text, valueAxis } from '../common.ts';
import { niceTicks, roundToTotal, textOn } from '../numbers.ts';
import { overlaps, spread } from '../place.ts';

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
  const ticks = o.valueAxis ? niceTicks(0, dmax) : null;
  if (ticks) dmax = ticks[ticks.length - 1];
  const axisFmt = axisFormatter(ctx, ticks, percent);
  const axisW = ticks ? Math.max(...ticks.map((t) => m.width(axisFmt(t), fs))) + pad * 1.5 : 0;

  // Reserve space: legend and totals on top, axis left, series labels right, categories + shares below.
  const legendH = legend(ctx);
  const showTotals = spec.options?.totals ?? true;
  const totalsH = showTotals ? lh + pad * 0.3 : 0;
  const rightW = seriesLabelWidth(ctx, 0);
  const px = frame.x + axisW;
  const pw = frame.w - axisW - rightW;
  const colW = totals.map((t) => (pw * t) / grand);
  const x0 = cats.map((c) => px + colW.slice(0, c).reduce((a, w) => a + w, 0));
  const cx = (c: number) => x0[c] + colW[c] / 2;

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
  const totalItems: (RowItem & { y: number; label: string })[] = [];

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
    if (showTotals) {
      const label = ctx.fmt(totals[c]);
      totalItems.push({ c, cx: cx(c), w: m.width(label, fs, true) + 2, y: yOf(up) - lh - pad * 0.2, label });
    }
  }

  // Totals: bold, centred on the column, nudged sideways over narrow columns (the row sits above the series labels).
  const totalPos = placeRow(totalItems, colW, px, frame.x + frame.w, pad * 0.5);
  for (const it of totalItems) {
    const p = totalPos.get(it.c);
    if (p) text(ctx, `label:total:c${it.c}`, { x: p.x, y: it.y, w: it.w, h: lh, text: it.label, bold: true, role: 'total' });
    else ctx.hidden.push(`${spec.categories[it.c]}: total label too wide for column`);
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
 * Lay a row of labels out along x, each centred on its column. Labels wider than their column
 * are nudged sideways so they do not collide; a label that would no longer cover its column
 * centre, or still collides, is narrowed to its `alt` width or dropped (narrowest column first).
 */
export function placeRow(items: RowItem[], colW: number[], min: number, max: number, gap: number) {
  const live = items.map((i) => ({ ...i, narrowed: false }));
  for (;;) {
    const xs = spread(live.map((i) => i.cx), live.map((i) => i.w), min, max, gap);
    const boxes = live.map((it, k) => ({ x: xs[k] - it.w / 2, y: 0, w: it.w, h: 1 }));
    const bad = live.filter((it, k) =>
      Math.abs(xs[k] - it.cx) > it.w / 2 - 2 || boxes.some((b, j) => j !== k && overlaps(b, boxes[k], -gap / 2)));
    if (!bad.length) return new Map(live.map((it, k) => [it.c, { x: boxes[k].x, w: it.w, narrowed: it.narrowed }]));
    const worst = bad.reduce((a, b) => (colW[b.c] < colW[a.c] ? b : a));
    if (worst.alt != null && worst.alt < worst.w) { worst.w = worst.alt; worst.narrowed = true; }
    else live.splice(live.indexOf(worst), 1);
  }
}
