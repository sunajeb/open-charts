import type { Ctx } from '../common.ts';
import {
  arrowRowHeight, bracketArrow, cagrLabel, categoryHeight, categoryLabels, legend,
  seriesLabelWidth, seriesLabelsRight, text, valueAxis,
} from '../common.ts';
import { formatNumber, inferDecimals, niceTicks, roundToTotal, textOn } from '../numbers.ts';

/** Stacked, 100% and clustered column charts. */
export function layoutColumns(ctx: Ctx) {
  const { spec, style, frame, o, fs, lh, pad, m } = ctx;
  const nCat = spec.categories.length;
  const nSer = spec.series.length;
  const percent = spec.type === 'percent';
  const clustered = spec.type === 'clustered';
  const raw = (s: number, c: number) => spec.series[s].values[c] ?? 0;

  const cats = [...Array(nCat).keys()];
  const sers = [...Array(nSer).keys()];
  const pos = cats.map((c) => sers.reduce((a, s) => a + Math.max(0, raw(s, c)), 0));
  const neg = cats.map((c) => sers.reduce((a, s) => a + Math.min(0, raw(s, c)), 0));
  const total = cats.map((c) => pos[c] + neg[c]);
  const shown = (s: number, c: number) => (percent ? (pos[c] ? (Math.max(0, raw(s, c)) / pos[c]) * 100 : 0) : raw(s, c));
  if (percent && neg.some((v) => v < 0)) ctx.hidden.push('negative values ignored in 100% chart');
  // 100% labels are rounded together so every column reads exactly 100%.
  const pctDecimals = o.decimals ?? 0;
  const shares = cats.map((c) => roundToTotal(sers.map((s) => shown(s, c)), pctDecimals));

  // Value domain.
  let [dmin, dmax] = percent ? [0, 100]
    : clustered ? [Math.min(0, ...cats.flatMap((c) => sers.map((s) => raw(s, c)))), Math.max(0, ...cats.flatMap((c) => sers.map((s) => raw(s, c))))]
    : [Math.min(0, ...neg), Math.max(0, ...pos)];
  if (dmax === dmin) dmax = dmin + 1;
  const ticks = o.valueAxis ? niceTicks(dmin, dmax) : null;
  if (ticks) [dmin, dmax] = [ticks[0], ticks[ticks.length - 1]];
  const axisDecimals = ticks ? inferDecimals(ticks) : 0;
  const axisFmt = (v: number) => formatNumber(v, axisDecimals, percent ? '%' : o.unit ?? '');

  // Reserve space: legend and arrows on top, totals above columns, axis on the left, series labels right, categories below.
  const legendH = legend(ctx);
  const arrows = [o.cagr, o.difference].filter((a) => a && !clustered);
  const arrowsH = arrows.length * arrowRowHeight(ctx);
  const showTotals = !clustered && !!o.totals;
  const totalsH = showTotals ? lh + pad * 0.3 : 0;
  const barLabelH = clustered ? lh + pad * 0.3 : 0;
  const axisW = ticks ? Math.max(...ticks.map((t) => m.width(axisFmt(t), fs))) + pad * 1.5 : 0;

  let rightW = 0, slot = 0, colW = 0, catH = 0;
  for (let i = 0; i < 3; i++) {
    slot = (frame.w - axisW - rightW) / nCat;
    colW = slot / (1 + o.gap);
    catH = categoryHeight(ctx, slot);
    rightW = clustered ? 0 : seriesLabelWidth(ctx, (slot - colW) / 2);
  }
  const px = frame.x + axisW;
  const pw = frame.w - axisW - rightW;
  const py = frame.y + legendH + arrowsH + totalsH + barLabelH;
  const ph = frame.y + frame.h - catH - (clustered && dmin < 0 ? barLabelH : 0) - py;
  if (ph < lh * 2) ctx.hidden.push('chart frame too small');
  const yOf = (v: number) => py + ph - ((v - dmin) / (dmax - dmin)) * ph;
  const cx = (c: number) => px + slot * (c + 0.5);

  if (ticks) valueAxis(ctx, ticks, yOf, px, axisW, axisFmt);

  const tops: number[] = [];
  const lastSegs: { name: string; y: number }[] = [];

  for (const c of cats) {
    if (clustered) {
      const groupW = colW;
      const barW = groupW / nSer;
      let top = yOf(0);
      for (const s of sers) {
        const v = raw(s, c);
        const x = cx(c) - groupW / 2 + s * barW;
        const y1 = yOf(Math.max(0, v)), y2 = yOf(Math.min(0, v));
        ctx.shapes.push({ kind: 'rect', x, y: y1, w: barW, h: Math.max(0.01, y2 - y1), fill: ctx.color(s), stroke: style.background, strokeWidth: 0.5 });
        const label = ctx.fmt(v);
        const tw = m.width(label, fs);
        if (tw <= barW + pad * 0.4) {
          const ly = v >= 0 ? y1 - lh - pad * 0.15 : y2 + pad * 0.15;
          text(ctx, { x: x + barW / 2 - tw / 2 - 1, y: ly, w: tw + 2, h: lh, text: label, role: 'value' });
        } else ctx.hidden.push(`${spec.series[s].name} / ${spec.categories[c]}: label too wide`);
        top = Math.min(top, y1);
      }
      tops.push(top);
      continue;
    }

    let up = 0, down = 0;
    const x = cx(c) - colW / 2;
    for (const s of sers) {
      const v = percent ? Math.max(0, shown(s, c)) : raw(s, c);
      if (!v) continue;
      const [a, b] = v > 0 ? [up, up + v] : [down + v, down];
      if (v > 0) up += v; else down += v;
      const y1 = yOf(b), y2 = yOf(a), h = y2 - y1;
      const fill = ctx.color(s);
      ctx.shapes.push({ kind: 'rect', x, y: y1, w: colW, h, fill, stroke: style.background, strokeWidth: 0.75 });
      const label = percent ? formatNumber(shares[c][s], pctDecimals, '%') : ctx.fmt(raw(s, c));
      const tw = m.width(label, fs);
      if (tw + 2 <= colW && lh * 0.9 <= h) {
        text(ctx, { x: x + colW / 2 - tw / 2 - 1, y: y1 + h / 2 - lh / 2, w: tw + 2, h: lh, text: label, color: textOn(fill), role: 'value' });
      } else ctx.hidden.push(`${spec.series[s].name} / ${spec.categories[c]}: segment too small`);
      if (c === nCat - 1) lastSegs.push({ name: spec.series[s].name, y: y1 + h / 2 });
    }
    let top = yOf(percent ? 100 : up);
    if (showTotals) {
      const label = ctx.fmt(total[c]);
      const tw = m.width(label, fs, true);
      top -= lh + pad * 0.2;
      text(ctx, { x: cx(c) - tw / 2 - 1, y: top, w: tw + 2, h: lh, text: label, bold: true, role: 'total' });
    }
    tops.push(top);
  }

  // Baseline.
  ctx.shapes.push({ kind: 'line', x1: px, y1: yOf(0), x2: px + pw, y2: yOf(0), color: style.axis, width: 0.75 });
  categoryLabels(ctx, cats.map(cx), slot, Math.max(yOf(0), yOf(dmin)) + (clustered && dmin < 0 ? barLabelH : 0));

  if (o.seriesLabels === 'right' && !clustered) {
    seriesLabelsRight(ctx, lastSegs.reverse(), cx(nCat - 1) + colW / 2, py - totalsH, py + ph);
  }

  // Growth and difference arrows, one row each, above the totals.
  let row = 0;
  const totalsFor = (c: number) => (percent ? pos[c] : total[c]);
  for (const [kind, ref] of [['cagr', o.cagr], ['difference', o.difference]] as const) {
    if (!ref || clustered) continue;
    const { from: a, to: b } = ref;
    if (a < 0 || b >= nCat || a >= b) { ctx.hidden.push(`${kind}: invalid category range`); continue; }
    const label = kind === 'cagr'
      ? cagrLabel(ctx, totalsFor(a), totalsFor(b), b - a)
      : ctx.fmt(totalsFor(b) - totalsFor(a), o.unit, true);
    if (!label) { ctx.hidden.push(`${kind}: needs positive totals`); continue; }
    const yLine = frame.y + legendH + row * arrowRowHeight(ctx) + arrowRowHeight(ctx) / 2;
    bracketArrow(ctx, cx(a), tops[a] - pad * 0.3, cx(b), tops[b] - pad * 0.3, yLine, label);
    row++;
  }
}
