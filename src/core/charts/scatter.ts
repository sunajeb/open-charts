import type { Ctx } from '../common.ts';
import type { TextShape } from '../types.ts';
import { legend, text, valueAxis } from '../common.ts';
import { formatValue, inferDecimals, niceTicks, textOn } from '../numbers.ts';
import { overlaps, segmentHitsBox, type Box } from '../place.ts';

interface Point { s: number; c: number; x: number; y: number; size: number }

/**
 * Scatter and bubble charts. Each series is a group (colour, legend entry); each category is one
 * point, labelled with the category name, whose cell is [x, y] (scatter) or [x, y, size] (bubble).
 *
 * - Two value axes with nice ticks; each starts at zero unless the data sits in a narrow band far
 *   from it (`zeroBased` forces either way). The domain grows so markers and bubbles stay inside.
 * - Bubble area is proportional to size; smaller bubbles are drawn in front of larger ones.
 * - Point labels try the spots around their marker (and inside a bubble that is big enough) and
 *   never cover another label or marker; a label with no free spot is hidden and reported.
 * - `axisTitles: {x, y}` names the axes.
 */
export function layoutScatter(ctx: Ctx) {
  const { spec, style, frame, o, fs, lh, pad, m } = ctx;
  const bubble = spec.type === 'bubble';
  const points: Point[] = [];
  spec.series.forEach((ser, s) => ser.values.forEach((v, c) => {
    if (!Array.isArray(v)) return;
    const [x, y, size = 0] = v;
    points.push({ s, c, x, y, size: Math.max(0, size) });
  }));
  if (!points.length) { ctx.hidden.push('chart has no points'); return; }
  if (spec.options?.seriesLabels === undefined || o.seriesLabels === 'right') ctx.o.seriesLabels = spec.series.length > 1 ? 'legend' : 'none';

  // Groups are told apart by their markers, so the legend shows circles.
  const legendH = legend(ctx, 'circle');
  const titles = o.axisTitles ?? {};
  const yTitleH = titles.y ? lh + pad * 0.3 : 0;
  const xTitleH = titles.x ? lh : 0;
  const unit = o.unit;
  const fmtAxis = (ticks: number[]) => {
    const d = inferDecimals(ticks.map((t) => t / ctx.scale.div));
    return (v: number) => formatValue(v / ctx.scale.div, { locale: o.locale, decimals: d, unit, magnitude: ctx.scale.suffix });
  };

  // Plot box (refined once the tick labels are known). The bubble size legend (the reference's
  // "Add Bubble Size Legend") sits top right: one bubble of a round size, drawn to the chart's scale.
  const maxSize = Math.max(0, ...points.map((p) => p.size));
  const r0 = Math.max(3, fs * 0.3);
  const xAxisH = lh + pad * 0.8;
  const sizeValue = bubble && o.sizeLegend && maxSize > 0 ? roundDown(maxSize / 2) : 0;
  const rMaxFor = (ph: number) => (bubble ? Math.max(r0, Math.min(frame.w, ph) * 0.09) : r0);
  const plotH = (topH: number) => frame.y + frame.h - xTitleH - xAxisH - (frame.y + topH + lh / 2);
  let topH = legendH + yTitleH;
  if (sizeValue) topH = Math.max(topH, 2 * rMaxFor(plotH(topH)) * Math.sqrt(sizeValue / maxSize) + pad);
  const py = frame.y + topH + lh / 2;
  const ph = plotH(topH);
  const rMax = rMaxFor(ph);
  const radius = (p: Point) => (bubble ? (maxSize > 0 ? Math.max(1.5, rMax * Math.sqrt(p.size / maxSize)) : r0) : r0);

  // Domain: data (with zero unless narrow), widened so every marker fits inside the axes.
  const domain = (vals: number[], lenPx: number, rad: number[]) => {
    const lo = Math.min(...vals), hi = Math.max(...vals);
    const narrow = (lo > 0 && (hi - lo) / hi < 0.6) || (hi < 0 && (hi - lo) / -lo < 0.6);
    const zero = o.zeroBased ?? !narrow;
    let a = zero ? Math.min(0, lo) : lo, b = zero ? Math.max(0, hi) : hi;
    if (a === b) { a -= 1; b += 1; }
    let ticks = niceTicks(a, b);
    for (let k = 0; k < 3; k++) {
      const perPx = (ticks[ticks.length - 1] - ticks[0]) / lenPx;
      const a2 = Math.min(a, ...vals.map((v, i) => v - rad[i] * perPx)), b2 = Math.max(b, ...vals.map((v, i) => v + rad[i] * perPx));
      // A marker that would cross an axis widens the domain past it (even past zero).
      const next = niceTicks(a2, b2);
      if (next[0] === ticks[0] && next[next.length - 1] === ticks[ticks.length - 1]) break;
      ticks = next;
    }
    return ticks;
  };
  const rad = points.map(radius);
  const yTicks = domain(points.map((p) => p.y), ph, rad);
  const yFmt = fmtAxis(yTicks);
  const axisW = Math.max(...yTicks.map((t) => m.width(yFmt(t), fs))) + pad * 1.5;
  const px = frame.x + axisW;
  let xTicks = domain(points.map((p) => p.x), frame.w - axisW, rad);
  let xFmt = fmtAxis(xTicks);
  const overhang = () => (m.width(xFmt(xTicks[xTicks.length - 1]), fs) + 2) / 2;
  let pw = frame.x + frame.w - overhang() - px;
  xTicks = domain(points.map((p) => p.x), pw, rad);
  xFmt = fmtAxis(xTicks);
  pw = frame.x + frame.w - overhang() - px;
  if (ph < lh * 3 || pw < lh * 3) ctx.hidden.push('chart frame too small');

  const [x0, x1] = [xTicks[0], xTicks[xTicks.length - 1]];
  const [y0, y1] = [yTicks[0], yTicks[yTicks.length - 1]];
  const xOf = (v: number) => px + ((v - x0) / (x1 - x0)) * pw;
  const yOf = (v: number) => py + ph - ((v - y0) / (y1 - y0)) * ph;

  // Axes: values on the left, the x axis along the bottom (tick roles numbered across both).
  valueAxis(ctx, yTicks, yOf, px, axisW, yFmt);
  const base = py + ph;
  ctx.add('axis', { kind: 'line', x1: px, y1: base, x2: px + pw, y2: base, color: style.axis, width: 0.75 });
  const yLabels = ctx.shapes.filter((sh, i): sh is TextShape => sh.kind === 'text' && ctx.roles[i].startsWith('axis:tick:'));
  xTicks.forEach((t, k) => {
    const role = `axis:tick:${yTicks.length + k}`;
    const label = xFmt(t);
    const w = m.width(label, fs) + 2;
    ctx.add(role, { kind: 'line', x1: xOf(t), y1: base, x2: xOf(t), y2: base + 3, color: style.axis, width: 0.75 });
    const box = { x: xOf(t) - w / 2, y: base + pad * 0.5, w, h: lh };
    // In the corner the y axis's lowest label already marks the origin.
    if (!yLabels.some((yl) => overlaps(yl, box))) text(ctx, role, { ...box, text: label, color: style.muted, role: 'axis' });
  });
  if (titles.y) {
    const w = Math.min(m.width(titles.y, fs) + 2, frame.w);
    text(ctx, 'axis', { x: frame.x, y: frame.y + topH - yTitleH, w, h: lh, text: titles.y, align: 'left', color: style.muted, role: 'axis' });
  }
  if (titles.x) {
    const w = Math.min(m.width(titles.x, fs) + 2, frame.w);
    text(ctx, 'axis', { x: px + pw - w, y: frame.y + frame.h - lh, w, h: lh, text: titles.x, align: 'right', color: style.muted, role: 'axis' });
  }

  // Markers, largest first so small bubbles sit in front.
  const boxes = points.map((p, i) => ({ x: xOf(p.x) - rad[i], y: yOf(p.y) - rad[i], w: 2 * rad[i], h: 2 * rad[i] }));
  const order = points.map((_, i) => i).sort((a, b) => rad[b] - rad[a]);
  for (const i of order) {
    const p = points[i];
    ctx.add(`marker:s${p.s}:c${p.c}`, { kind: 'rect', ...boxes[i], fill: ctx.color(p.s), stroke: style.background, strokeWidth: 0.75, geom: 'ellipse' });
  }

  // Bubble size legend, top right, clear of the series legend and the y-axis title.
  if (sizeValue) {
    const rl = rMax * Math.sqrt(sizeValue / maxSize);
    const caption = typeof o.sizeLegend === 'string' ? `${o.sizeLegend}: ` : '';
    const label = caption + formatValue(sizeValue, { locale: o.locale, decimals: inferDecimals([sizeValue]) });
    const w = m.width(label, fs) + 2;
    const bx = frame.x + frame.w - w - pad * 0.5 - 2 * rl;
    const lb = { x: frame.x + frame.w - w, y: frame.y + rl - lh / 2, w, h: lh };
    const taken = ctx.shapes.filter((sh, k) => sh.kind === 'text' && /^(legend|axis)/.test(ctx.roles[k]) && sh.y < py) as TextShape[];
    if (taken.some((t) => overlaps(t, { x: bx, y: frame.y, w: frame.x + frame.w - bx, h: 2 * rl }))) ctx.hidden.push('bubble size legend: no room');
    else {
      ctx.add('legend:size', { kind: 'rect', x: bx, y: frame.y, w: 2 * rl, h: 2 * rl, fill: style.background, stroke: style.muted, strokeWidth: 0.75, geom: 'ellipse' });
      text(ctx, 'legend:size', { ...lb, text: label, align: 'left', color: style.muted, role: 'series' });
    }
  }

  // Point labels: the category name next to its marker, in the first free spot; points that have
  // none get their label further out with a leader line ("connecting lines if necessary").
  const plot: Box = { x: px, y: frame.y + topH, w: pw + overhang(), h: base - frame.y - topH };
  const inPlot = (b: Box) => b.x >= plot.x + 1 && b.y >= plot.y && b.x + b.w <= plot.x + plot.w && b.y + b.h <= base - 1;
  const placed: Box[] = [];
  const leaders: { x1: number; y1: number; x2: number; y2: number }[] = [];
  const freeBox = (b: Box, i: number, centre: Box | null) => inPlot(b) && !placed.some((q) => overlaps(q, b)) &&
    !leaders.some((l) => segmentHitsBox(l.x1, l.y1, l.x2, l.y2, b)) &&
    boxes.every((mb, j) => (j === i ? b === centre || !overlaps(mb, b, 0.5) : !overlaps(mb, b, 0)));
  // Crowded points first: they have the fewest options.
  const crowd = points.map((_, i) => boxes.filter((b, j) => j !== i && overlaps(grow(b, lh * 2), boxes[i], 0)).length);
  const labelOrder = points.map((_, i) => i).sort((a, b) => crowd[b] - crowd[a] || a - b);
  const label = (i: number) => spec.categories[points[i].c];
  const put = (i: number, spot: Box, centre = false) => {
    const p = points[i];
    placed.push(spot);
    text(ctx, `label:value:s${p.s}:c${p.c}`, { ...spot, text: label(i), role: 'value', ...(centre ? { color: textOn(ctx.color(p.s)) } : {}) });
  };
  for (const i of labelOrder) {
    const name = label(i);
    if (!name) continue;
    const w = m.width(name, fs) + 2;
    const cx = xOf(points[i].x), cy = yOf(points[i].y), r = rad[i], g = pad * 0.3;
    const d = r * Math.SQRT1_2;
    const centre: Box | null = bubble && w <= r * 1.7 && lh <= r * 1.2 ? { x: cx - w / 2, y: cy - lh / 2, w, h: lh } : null;
    const spots: Box[] = centre ? [centre] : [];
    spots.push(
      { x: cx + r + g, y: cy - lh / 2, w, h: lh },
      { x: cx - r - g - w, y: cy - lh / 2, w, h: lh },
      { x: cx - w / 2, y: cy - r - g - lh, w, h: lh },
      { x: cx - w / 2, y: cy + r + g, w, h: lh },
      { x: cx + d + g * 0.5, y: cy - d - lh, w, h: lh },
      { x: cx - d - g * 0.5 - w, y: cy - d - lh, w, h: lh },
      { x: cx + d + g * 0.5, y: cy + d, w, h: lh },
      { x: cx - d - g * 0.5 - w, y: cy + d, w, h: lh },
    );
    const spot = spots.find((b) => freeBox(b, i, centre));
    if (spot) put(i, spot, spot === centre); else leader(i, w, cx, cy, r);
  }

  /** Leader-line label: rings further out in sixteen directions; the leader runs from the marker's edge to the label and crosses no other marker or label. */
  function leader(i: number, w: number, cx: number, cy: number, r: number) {
    let found: { box: Box; x1: number; y1: number; x2: number; y2: number } | null = null;
    for (const dist of [lh * 1.2, lh * 2, lh * 3, lh * 4.5]) {
      for (let k = 0; k < 16 && !found; k++) {
        const a = (k * Math.PI) / 8, ux = Math.cos(a), uy = -Math.sin(a);
        const ax = cx + ux * (r + dist), ay = cy + uy * (r + dist);
        // The label hangs off the anchor on the side facing away from the point.
        const box = { x: ux > 0.3 ? ax : ux < -0.3 ? ax - w : ax - w / 2, y: uy > 0.3 ? ay : uy < -0.3 ? ay - lh : ay - lh / 2, w, h: lh };
        const ex = Math.min(Math.max(cx, box.x), box.x + box.w), ey = Math.min(Math.max(cy, box.y), box.y + box.h);
        const len = Math.hypot(ex - cx, ey - cy);
        if (len <= r + 1) continue;
        const l = { x1: cx + ((ex - cx) / len) * (r + 1), y1: cy + ((ey - cy) / len) * (r + 1), x2: ex, y2: ey };
        const clearLine = boxes.every((mb, j) => j === i || !segmentHitsBox(l.x1, l.y1, l.x2, l.y2, grow(mb, 1))) &&
          placed.every((q) => !segmentHitsBox(l.x1, l.y1, l.x2, l.y2, q));
        if (clearLine && freeBox(box, i, null)) found = { box, ...l };
      }
      if (found) break;
    }
    if (!found) { ctx.hidden.push(`${label(i)}: no room for point label`); return; }
    const p = points[i];
    leaders.push(found);
    ctx.add(`leader:c${p.c}`, { kind: 'line', x1: found.x1, y1: found.y1, x2: found.x2, y2: found.y2, color: style.muted, width: 0.5 });
    put(i, found.box);
  }
}

/** The largest 1, 2 or 5 × 10^k not above v. */
function roundDown(v: number): number {
  const p = 10 ** Math.floor(Math.log10(v));
  return [5, 2, 1].map((k) => k * p).find((x) => x <= v) ?? p;
}

const grow = (b: Box, d: number): Box => ({ x: b.x - d, y: b.y - d, w: b.w + 2 * d, h: b.h + 2 * d });
