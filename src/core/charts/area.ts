import type { Ctx } from '../common.ts';
import { axisFormatter, fitCategories, reportFit, legend, num, seriesLabelWidth, seriesLabelsRight, text, valueAxis } from '../common.ts';
import { niceTicks, roundToTotal, textOn } from '../numbers.ts';
import { overlaps, type Box } from '../place.ts';
import type { PathCmd } from '../types.ts';

/**
 * Stacked area (`area`) and 100% stacked area (`area100`) charts. Each series is a filled band between
 * the cumulative sums below and including it; categories sit on points that span the plot edge to edge.
 *
 * - Missing values are interpolated linearly between their neighbours (the whole chart, since the
 *   bands stack); missing values at either end count as zero. Interpolated points carry no label.
 * - Point labels sit inside their band at each category when the band is tall enough across the
 *   label's width; 100% labels are shares rounded together so every category reads exactly 100%.
 * - `options.totals` adds bold totals above the top band (stacked area only).
 * - Series are named right of the last point (default) or in a legend.
 */
export function layoutArea(ctx: Ctx) {
  const { spec, style, frame, o, fs, lh, pad, m } = ctx;
  const percent = spec.type === 'area100';
  const nCat = spec.categories.length;
  const cats = [...Array(nCat).keys()];
  const sers = spec.series.map((_, i) => i);
  if (spec.series.some((s) => s.values.some((v) => (num(v) ?? 0) < 0))) ctx.hidden.push('negative values ignored in area chart');

  // Values with gaps interpolated; `given` marks the points that came from the datasheet.
  const given = sers.map((s) => cats.map((c) => num(spec.series[s].values[c]) != null));
  const vals = sers.map((s) => interpolate(cats.map((c) => Math.max(0, num(spec.series[s].values[c]) ?? NaN))));
  const totals = cats.map((c) => sers.reduce((a, s) => a + vals[s][c], 0));
  const shown = (s: number, c: number) => (percent ? (totals[c] ? (vals[s][c] / totals[c]) * 100 : 0) : vals[s][c]);
  const pctDecimals = o.decimals ?? 0;
  const shares = cats.map((c) => roundToTotal(sers.map((s) => shown(s, c)), pctDecimals));
  const rounded = cats.map((c) => (o.format ? null : roundToTotal(sers.map((s) => vals[s][c] / ctx.scale.div), ctx.decimals)));
  const label = (s: number, c: number) =>
    percent ? ctx.pct(shares[c][s], pctDecimals) : rounded[c] ? ctx.fmt(rounded[c]![s], { scaled: true }) : ctx.fmt(vals[s][c]);

  let dmax = percent ? 100 : Math.max(...totals);
  if (dmax <= 0) { ctx.hidden.push('area chart has no positive values'); dmax = 1; }
  const ticks = o.valueAxis ? niceTicks(0, dmax) : null;
  if (ticks) dmax = ticks[ticks.length - 1];
  const axisFmt = axisFormatter(ctx, ticks, percent);
  const axisW = ticks ? Math.max(...ticks.map((t) => m.width(axisFmt(t), fs))) + pad * 1.5 : 0;
  const showTotals = !percent && !!o.totals;

  // Horizontal layout: points run edge to edge; the first and last category labels may overhang the
  // plot by half their width, so the plot is inset just enough to keep them in the frame.
  const legendH = legend(ctx);
  const rightW = seriesLabelWidth(ctx, 0);
  let pxL = frame.x + axisW, pxR = frame.x + frame.w - rightW, step = 0;
  for (let i = 0; i < 3; i++) {
    step = nCat > 1 ? (pxR - pxL) / (nCat - 1) : pxR - pxL;
    const fit = fitCategories(ctx, step);
    const half = (c: number) => (fit.hidden.has(c) ? 0 : Math.min(step / 2, Math.max(...fit.lines[c].map((l) => m.width(l, fit.size))) / 2 + 1));
    // With a value axis the first label starts at the axis instead of centring on the point.
    pxL = frame.x + (axisW ? axisW : half(0));
    pxR = frame.x + frame.w - Math.max(rightW, half(nCat - 1));
  }
  step = nCat > 1 ? (pxR - pxL) / (nCat - 1) : pxR - pxL;
  const fit = fitCategories(ctx, step);
  const catH = fit.height;
  const xOf = (c: number) => (nCat > 1 ? pxL + c * step : (pxL + pxR) / 2);

  const totalsH = showTotals ? lh + pad * 0.3 : 0;
  const py = frame.y + legendH + totalsH + lh * 0.5;
  const ph = frame.y + frame.h - catH - py;
  if (ph < lh * 2) ctx.hidden.push('chart frame too small');
  const yOf = (v: number) => py + ph - (v / dmax) * ph;
  const bottom = yOf(0);

  if (ticks) valueAxis(ctx, ticks, yOf, pxL, axisW, axisFmt);

  // Cumulative edges: lower[s][c] / upper[s][c] in slide y.
  const lower: number[][] = [], upper: number[][] = [];
  const acc = cats.map(() => 0);
  for (const s of sers) {
    lower.push(cats.map((c) => yOf(acc[c])));
    cats.forEach((c) => { acc[c] += shown(s, c); });
    upper.push(cats.map((c) => yOf(acc[c])));
  }
  // A single category is drawn as a flat band across the plot.
  const xs = nCat > 1 ? cats.map(xOf) : [pxL, pxR];
  const pick = (arr: number[]) => (nCat > 1 ? arr : [arr[0], arr[0]]);
  const edgeAt = (arr: number[], x: number) => {
    const ys = pick(arr);
    if (x <= xs[0]) return ys[0];
    for (let k = 1; k < xs.length; k++) {
      if (x <= xs[k]) return ys[k - 1] + ((ys[k] - ys[k - 1]) * (x - xs[k - 1])) / (xs[k] - xs[k - 1]);
    }
    return ys[ys.length - 1];
  };
  /** Tightest band edges over [x1, x2] (vertices included): lowest upper edge and highest lower edge. */
  const span = (s: number, x1: number, x2: number) => {
    const probe = [x1, x2, ...xs.filter((x) => x > x1 && x < x2)];
    return {
      top: Math.max(...probe.map((x) => edgeAt(upper[s], x))),
      bottom: Math.min(...probe.map((x) => edgeAt(lower[s], x))),
    };
  };

  for (const s of sers) {
    if (cats.every((c) => !shown(s, c))) { ctx.hidden.push(`${spec.series[s].name}: no values`); continue; }
    const up = pick(upper[s]), lo = pick(lower[s]);
    const commands: PathCmd[] = [
      ...xs.map((x, k) => ({ op: k ? 'L' : 'M', x, y: up[k] }) as PathCmd),
      ...xs.map((_, k) => ({ op: 'L', x: xs[xs.length - 1 - k], y: lo[xs.length - 1 - k] }) as PathCmd),
      { op: 'Z' },
    ];
    ctx.add(`area:s${s}`, { kind: 'path', commands, fill: ctx.color(s), stroke: style.background, strokeWidth: 0.75 });
  }

  // Point labels inside their band, clamped into the plot at the ends.
  const placed: Box[] = [];
  const inset = pad * 0.5;
  const clampX = (x: number, w: number) => Math.min(Math.max(x, pxL + inset), pxR - inset - w);
  for (const s of sers) {
    for (const c of cats) {
      if (!given[s][c] || !shown(s, c)) continue;
      const t = label(s, c);
      const w = m.width(t, fs) + 2;
      const x = clampX(xOf(c) - w / 2, w);
      const room = span(s, x, x + w);
      const box = { x, y: (room.top + room.bottom) / 2 - lh / 2, w, h: lh };
      if (room.bottom - room.top < lh * 0.9 || placed.some((p) => overlaps(p, box, 0))) {
        ctx.hidden.push(`${spec.series[s].name} / ${spec.categories[c]}: band too thin for label`);
        continue;
      }
      placed.push(box);
      text(ctx, `label:value:s${s}:c${c}`, { ...box, text: t, color: textOn(ctx.color(s)), role: 'value' });
    }
  }

  // Totals above the top band, lifted clear of the edge across the label's width.
  if (showTotals) {
    const top = sers.length - 1;
    for (const c of cats) {
      if (!totals[c]) continue;
      const t = ctx.fmt(totals[c]);
      const w = m.width(t, fs, true) + 2;
      const x = clampX(xOf(c) - w / 2, w);
      const edge = Math.min(...[x, x + w, ...xs.filter((v) => v > x && v < x + w)].map((v) => edgeAt(upper[top], v)));
      const box = { x, y: edge - lh - pad * 0.2, w, h: lh };
      if (box.y < frame.y || placed.some((p) => overlaps(p, box, 0))) { ctx.hidden.push(`${spec.categories[c]}: no room for total`); continue; }
      placed.push(box);
      text(ctx, `label:total:c${c}`, { ...box, text: t, bold: true, role: 'total' });
    }
  }

  ctx.add('axis', { kind: 'line', x1: pxL, y1: bottom, x2: pxR, y2: bottom, color: style.axis, width: 0.75 });
  // Category labels as wide as their text, centred on their point; kept inside the plot's ends.
  reportFit(ctx, fit);
  for (const c of cats) {
    if (fit.hidden.has(c)) continue;
    const lines = fit.lines[c];
    const w = Math.min(step, Math.max(...lines.map((l) => m.width(l, fit.size))) + 2);
    let x = xOf(c) - w / 2;
    if (c === 0 && axisW) x = Math.max(x, pxL);
    text(ctx, `label:category:c${c}`, {
      x, y: bottom + pad * 0.4, w, h: lines.length * fit.lh, text: lines.join('\n'), valign: 'top', role: 'category', size: fit.size,
    });
  }

  if (o.seriesLabels === 'right') {
    const last = nCat - 1;
    const items = sers
      .filter((s) => shown(s, last) > 0)
      .map((s) => ({ s, y: (upper[s][last] + lower[s][last]) / 2 }))
      .reverse();
    seriesLabelsRight(ctx, items, pxR, py - totalsH, bottom);
  }
}

/** Fill NaN gaps by linear interpolation between known neighbours; leading/trailing gaps become 0. */
function interpolate(v: number[]): number[] {
  const out = v.slice();
  for (let i = 0; i < out.length; i++) {
    if (!Number.isNaN(out[i])) continue;
    const a = i - 1;
    let b = i + 1;
    while (b < out.length && Number.isNaN(v[b])) b++;
    out[i] = a >= 0 && b < out.length ? out[a] + ((v[b] - out[a]) * (i - a)) / (b - a) : 0;
  }
  return out;
}
