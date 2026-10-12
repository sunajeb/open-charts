import type { Ctx } from '../common.ts';
import { fitCategories, legend, num, seriesLabelWidth, seriesLabelsRight, text, valueAxis } from '../common.ts';
import { formatValue, inferDecimals, niceTicks, roundToTotal, textOn } from '../numbers.ts';
import { placeRow, type RowItem } from './mekko.ts';
import { spread } from '../place.ts';

/**
 * Mekko XY (Mekko with units, bar-Mekko, cost curve): width and height are set independently. The
 * first series is the X extent (column widths, like the reference datasheet's "X extent" row);
 * every further series is a stacked height. Columns touch, as in every Mekko chart.
 *
 * - `options.sort`: sheet order (default), reverse, ascending or descending by column height.
 * - One height series: its value sits above each column. Several: segment labels inside and a bold
 *   total on top. Labels over narrow columns borrow room from neighbours or are hidden and reported.
 *   A segment of the last column too small for its label gets it outside, with a leader.
 * - Under each column: its name and, unless the baseline is a value axis, its width.
 * - `valueAxis`: a value axis for heights on the left and the baseline becomes a value axis for the
 *   cumulative widths (both with nice ticks).
 */
export function layoutMekkoXY(ctx: Ctx) {
  const { spec, style, frame, o, fs, lh, pad, m } = ctx;
  const nCat = spec.categories.length;
  const cats = [...Array(nCat).keys()];
  const hs = spec.series.map((_, i) => i).slice(1);
  const width = (c: number) => Math.max(0, num(spec.series[0].values[c]) ?? 0);
  const h = (s: number, c: number) => num(spec.series[s].values[c]) ?? 0;
  const color = (s: number) => (spec.series[s].color ?? style.palette[(s - 1) % style.palette.length]).replace('#', '');
  const up = cats.map((c) => hs.reduce((a, s) => a + Math.max(0, h(s, c)), 0));
  const down = cats.map((c) => hs.reduce((a, s) => a + Math.min(0, h(s, c)), 0));
  const total = cats.map((c) => up[c] + down[c]);
  const grandW = cats.reduce((a, c) => a + width(c), 0);
  if (grandW <= 0) { ctx.hidden.push('Mekko XY needs positive widths'); return; }
  if (hs.length === 1 && spec.options?.seriesLabels === undefined) ctx.o.seriesLabels = 'none';

  // Category order (the manual's "Y extent" orders sort by the column's total height).
  const order = [...cats];
  if (o.sort === 'reverse') order.reverse();
  else if (o.sort === 'ascending') order.sort((a, b) => total[a] - total[b]);
  else if (o.sort === 'descending') order.sort((a, b) => total[b] - total[a]);

  // Heights get their own decimals; widths are a different quantity and are shown without the unit.
  const hDecimals = o.decimals ?? inferDecimals(hs.flatMap((s) => cats.map((c) => h(s, c) / ctx.scale.div)));
  const fmtH = (v: number) => ctx.fmt(v, { decimals: hDecimals });
  const wDecimals = inferDecimals(cats.map(width));
  const fmtW = (v: number, d = wDecimals) => formatValue(v, { locale: o.locale, decimals: d });

  let dmin = Math.min(0, ...down), dmax = Math.max(0, ...up);
  if (dmax === dmin) dmax = dmin + 1;
  const ticks = o.valueAxis ? niceTicks(dmin, dmax) : null;
  if (ticks) [dmin, dmax] = [ticks[0], ticks[ticks.length - 1]];
  const yFmt = (v: number) => ctx.fmt(v, { decimals: ticks ? inferDecimals(ticks.map((t) => t / ctx.scale.div)) : 0 });
  const axisW = ticks ? Math.max(...ticks.map((t) => m.width(yFmt(t), fs))) + pad * 1.5 : 0;
  // The columns always span the plot; the width axis shows the nice ticks that fall inside it.
  const xTicks = o.valueAxis ? niceTicks(0, grandW).filter((t) => t <= grandW * (1 + 1e-9)) : null;
  const xMax = grandW;
  const xFmt = (v: number) => fmtW(v, xTicks ? inferDecimals(xTicks) : 0);

  const legendH = legend(ctx, 'box', color, undefined, (s) => s > 0);
  // Axis titles from the series names: the height series (when there is one) above the plot, the width series under it.
  const yTitleH = hs.length === 1 ? lh + pad * 0.3 : 0;
  const topH = lh + pad * 0.3;
  const axisH = xTicks ? lh + pad * 0.8 : 0;
  const widthLine = !xTicks;
  const belowH = dmin < 0 ? topH : 0;
  const py = frame.y + legendH + yTitleH + topH;
  const geometry = (outW: number) => {
    const rightW = seriesLabelWidth(ctx, 0) + outW;
    const px = frame.x + axisW;
    const pw = frame.w - axisW - rightW;
    const colW = cats.map((c) => (pw * width(c)) / xMax);
    const x0 = new Array<number>(nCat);
    let acc = px;
    for (const c of order) { x0[c] = acc; acc += colW[c]; }
    const minLabelW = Math.min(fs * 6, pw / nCat);
    const fit = fitCategories(ctx, cats.map((c) => Math.max(colW[c], minLabelW)));
    const catH = axisH + Math.max(...fit.lines.map((l) => l.length)) * fit.lh + (widthLine ? lh : 0) + pad * 0.6 + lh;
    const ph = frame.y + frame.h - catH - belowH - py;
    return { px, pw, colW, x0, fit, ph, yOf: (v: number) => py + ph - ((v - dmin) / (dmax - dmin)) * ph };
  };
  // Segment labels that do not fit their segment go outside it with a leader, as the manual's rule for
  // small segments says. Beside the last column that space is free (the series names move right to
  // make room); elsewhere it is the neighbouring column, so those labels are hidden and reported.
  const last = order[nCat - 1];
  const segLabel = (s: number, c: number, rounded: number[] | null, k: number) => (rounded ? ctx.fmt(rounded[k], { scaled: true, decimals: hDecimals }) : fmtH(h(s, c)));
  const roundedOf = (c: number) => (o.format ? null : roundToTotal(hs.map((s) => h(s, c) / ctx.scale.div), hDecimals));
  const fitsIn = (label: string, w: number, hh: number) => m.width(label, fs) + 2 <= w && lh * 0.9 <= hh;
  let g = geometry(0);
  const outside = (gg: typeof g) => (hs.length > 1 ? hs.filter((s, k) => h(s, last) > 0 &&
    !fitsIn(segLabel(s, last, roundedOf(last), k), gg.colW[last], gg.yOf(0) - gg.yOf(h(s, last)))) : []);
  const outW0 = Math.max(0, ...outside(g).map((s) => m.width(segLabel(s, last, roundedOf(last), hs.indexOf(s)), fs) + 2 + pad * 1.2));
  if (outW0) g = geometry(outW0);
  const outW = outW0 ? Math.max(0, ...outside(g).map((s) => m.width(segLabel(s, last, roundedOf(last), hs.indexOf(s)), fs) + 2 + pad * 1.2)) : 0;
  const { px, pw, colW, x0, fit, ph, yOf } = g;
  const cx = (c: number) => x0[c] + colW[c] / 2;
  if (fit.size < fs) ctx.hidden.push(`category labels shrunk to ${fit.size} pt to fit`);
  if (ph < lh * 2) ctx.hidden.push('chart frame too small');

  const lastSegs: { s: number; y: number }[] = [];
  const outs: { s: number; y: number; label: string; w: number }[] = [];
  const tops: (RowItem & { y: number; label: string })[] = [];
  for (const c of order) {
    if (!colW[c]) { ctx.hidden.push(`${spec.categories[c]}: zero width, no column`); continue; }
    let u = 0, d = 0;
    const rounded = roundedOf(c);
    hs.forEach((s, k) => {
      const v = h(s, c);
      if (!v) return;
      const [a, b] = v > 0 ? [u, u + v] : [d + v, d];
      if (v > 0) u += v; else d += v;
      const y1 = yOf(b), hh = yOf(a) - y1;
      const fill = color(s);
      ctx.add(`bar:s${s}:c${c}`, { kind: 'rect', x: x0[c], y: y1, w: colW[c], h: hh, fill, stroke: style.background, strokeWidth: 0.75 });
      if (c === last) lastSegs.push({ s, y: y1 + hh / 2 });
      if (hs.length === 1) return;
      const label = segLabel(s, c, rounded, k);
      const tw = m.width(label, fs);
      if (fitsIn(label, colW[c], hh)) {
        text(ctx, `label:value:s${s}:c${c}`, { x: cx(c) - tw / 2 - 1, y: y1 + hh / 2 - lh / 2, w: tw + 2, h: lh, text: label, color: textOn(fill), role: 'value' });
      } else if (c === last && v > 0 && outW) outs.push({ s, y: y1 + hh / 2, label, w: tw + 2 });
      else ctx.hidden.push(`${spec.series[s].name} / ${spec.categories[c]}: segment too small`);
    });
    const label = fmtH(total[c]);
    const neg = total[c] < 0;
    tops.push({ c, cx: cx(c), w: m.width(label, fs, hs.length > 1) + 2, y: neg ? yOf(d) + pad * 0.2 : yOf(u) - lh - pad * 0.2, label });
  }

  // Heights (one series) or totals (several) above each column, nudged sideways over narrow columns.
  const showTop = hs.length === 1 || (spec.options?.totals ?? true);
  if (showTop) {
    for (const group of [tops.filter((t) => total[t.c] >= 0), tops.filter((t) => total[t.c] < 0)]) {
      const at = placeRow(group, colW, px, px + pw, pad * 0.5);
      for (const it of group) {
        const p = at.get(it.c);
        const role = hs.length === 1 ? `label:value:s${hs[0]}:c${it.c}` : `label:total:c${it.c}`;
        if (p) text(ctx, role, { x: p.x, y: it.y, w: it.w, h: lh, text: it.label, bold: hs.length > 1, role: hs.length === 1 ? 'value' : 'total' });
        else ctx.hidden.push(`${spec.categories[it.c]}: ${hs.length === 1 ? 'value' : 'total'} label too wide for column`);
      }
    }
  }

  // The value axis over the first column's edge, not hidden behind it.
  if (ticks) valueAxis(ctx, ticks, yOf, px, axisW, yFmt);
  const base = yOf(0);
  ctx.add('axis', { kind: 'line', x1: px, y1: base, x2: px + pw, y2: base, color: style.axis, width: 0.75 });
  const bottom = Math.max(base, yOf(dmin)) + belowH;

  // Baseline as a value axis for the cumulative widths.
  if (xTicks) {
    const xOf = (v: number) => px + (v / xMax) * pw;
    xTicks.forEach((t, k) => {
      const label = xFmt(t);
      const w = m.width(label, fs) + 2;
      const x = Math.min(Math.max(xOf(t) - w / 2, frame.x), frame.x + frame.w - w);
      ctx.add(`axis:tick:${ticks ? ticks.length + k : k}`, { kind: 'line', x1: xOf(t), y1: bottom, x2: xOf(t), y2: bottom + 3, color: style.axis, width: 0.75 });
      text(ctx, `axis:tick:${ticks ? ticks.length + k : k}`, { x, y: bottom + pad * 0.5, w, h: lh, text: label, color: style.muted, role: 'axis' });
    });
  }

  // Category names (and widths) under their columns.
  const wText = cats.map((c) => fmtW(width(c)));
  const items = order.map((c) => ({
    c, cx: cx(c),
    w: fit.hidden.has(c) ? (widthLine ? m.width(wText[c], fs) + 2 : 0) : Math.max(m.width(fit.lines[c].join('\n'), fit.size), widthLine ? m.width(wText[c], fs) : 0) + 2,
    alt: widthLine ? m.width(wText[c], fs) + 2 : undefined,
  })).filter((it) => it.w > 0);
  const at = placeRow(items, colW, frame.x, frame.x + frame.w, pad);
  for (const c of order) {
    const p = at.get(c);
    const y = bottom + axisH + pad * 0.4;
    if (!p || p.narrowed || fit.hidden.has(c)) ctx.hidden.push(`${spec.categories[c]}: category label too wide for column`);
    if (!p) continue;
    let sy = y;
    if (!p.narrowed && !fit.hidden.has(c)) {
      const hh = fit.lines[c].length * fit.lh;
      text(ctx, `label:category:c${c}`, { x: p.x, y, w: p.w, h: hh, text: fit.lines[c].join('\n'), valign: 'top', role: 'category', size: fit.size });
      sy += hh;
    }
    if (widthLine) text(ctx, `label:category:c${c}`, { x: p.x, y: sy, w: p.w, h: lh, text: wText[c], color: style.muted, valign: 'top', role: 'decoration' });
  }

  // Outside labels of the last column: beside it, spread apart, each with a leader to its segment.
  const edge = x0[last] + colW[last];
  const ys = spread(outs.map((it) => it.y), outs.map(() => lh), py, base, 0);
  outs.forEach((it, k) => {
    const lx = edge + pad * 0.6;
    ctx.add(`leader:c${last}`, { kind: 'line', x1: edge, y1: it.y, x2: lx - 1, y2: ys[k], color: style.muted, width: 0.5 });
    text(ctx, `label:value:s${it.s}:c${last}`, { x: lx, y: ys[k] - lh / 2, w: it.w, h: lh, text: it.label, align: 'left', role: 'value' });
  });
  if (o.seriesLabels === 'right') seriesLabelsRight(ctx, lastSegs.reverse(), px + pw + outW, py, base);

  const title = (name: string, x: number, y: number, align: 'left' | 'right') => {
    const w = Math.min(m.width(name, fs) + 2, frame.w);
    text(ctx, 'axis', { x: align === 'left' ? x : x - w, y, w, h: lh, text: name, align, color: style.muted, role: 'axis' });
  };
  title(spec.series[0].name, px + pw, frame.y + frame.h - lh, 'right');
  if (hs.length === 1) title(spec.series[hs[0]].name, frame.x, frame.y + legendH, 'left');
}
