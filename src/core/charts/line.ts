import type { Ctx } from '../common.ts';
import {
  axisFormatter, categoryHeight, categoryLabels, legend, num, seriesLabelsRight, text, valueAxis, valueTicks, axisWidth, sharedScale,
} from '../common.ts';
import { niceTicks } from '../numbers.ts';
import { overlaps, segmentHitsBox, type Box } from '../place.ts';
import type { LineShape } from '../types.ts';

/**
 * Line chart: one point per category, markers, value labels placed above or below without collisions.
 * `options.step` draws an ortho (step) line: horizontal to the next category, then vertical.
 */
export function layoutLine(ctx: Ctx) {
  const { spec, style, frame, o, fs, lh, pad, m } = ctx;
  const nCat = spec.categories.length;
  const sers = spec.series.map((_, i) => i);
  const val = (s: number, c: number) => num(spec.series[s].values[c]);
  const all = spec.series.flatMap((s) => s.values.map(num)).filter((v): v is number => v != null);
  if (!all.length) { ctx.hidden.push('line chart has no values'); return; }

  // Include zero unless the data sits in a narrow band far from it (e.g. retention 95–120%).
  const lo = Math.min(...all), hi = Math.max(...all);
  const narrow = lo > 0 && (hi - lo) / hi < 0.4;
  const zero = o.zeroBased ?? !narrow;
  const [vmin, vmax] = zero ? [Math.min(0, lo), Math.max(0, hi)] : [lo, hi];
  let ticks = !o.valueAxis ? null : niceTicks(vmin, vmax, valueTicks(ctx));
  let [dmin, dmax] = ticks ? [ticks[0], ticks[ticks.length - 1]] : [vmin, vmax];
  if (dmax === dmin) dmax = dmin + 1;
  let axisFmt = axisFormatter(ctx, ticks);
  const axisW = axisWidth(ctx, dmin, ticks, axisFmt);

  // Label every point when the chart is sparse enough to stay readable, otherwise only the last one.
  const labelAll = nCat * sers.length <= 24;
  const r = Math.max(3, fs * 0.3);
  const lastLabelHalf = Math.max(0, ...sers.map((s) => {
    const v = val(s, nCat - 1);
    return v == null ? 0 : m.width(ctx.fmt(v), fs) / 2;
  }));

  const legendH = legend(ctx, 'line');
  let slot = 0, catH = 0, rightW = 0;
  const nameW = o.seriesLabels === 'right' ? Math.max(...spec.series.map((s) => m.width(s.name, fs))) : 0;
  for (let i = 0; i < 3; i++) {
    slot = (frame.w - axisW - rightW) / nCat;
    catH = categoryHeight(ctx, slot);
    rightW = o.seriesLabels === 'right' ? Math.max(0, nameW + lastLabelHalf + r + pad * 2 - slot / 2) : 0;
  }
  const px = frame.x + axisW;
  const pw = frame.w - axisW - rightW;
  // Headroom for a label above a point at the top of the scale.
  const py = frame.y + legendH + lh + pad * 0.5 + r;
  const ph = frame.y + frame.h - catH - py - lh * 0.5;
  ({ dmin: dmin, dmax: dmax, ticks } = sharedScale(ctx, dmin, dmax, ph, ticks));
  if (ticks) axisFmt = axisFormatter(ctx, ticks);
  const yOf = (v: number) => py + ph - ((v - dmin) / (dmax - dmin)) * ph;
  const xOf = (c: number) => px + slot * (c + 0.5);

  if (ticks) valueAxis(ctx, ticks, yOf, px, axisW, axisFmt);
  const bottom = yOf(dmin);
  ctx.add('axis', { kind: 'line', x1: px, y1: bottom, x2: px + pw, y2: bottom, color: style.axis, width: 0.75 });
  if (dmin < 0 && dmax > 0) {
    ctx.add('axis', { kind: 'line', x1: px, y1: yOf(0), x2: px + pw, y2: yOf(0), color: style.axis, width: 0.5, dash: true });
  }
  categoryLabels(ctx, spec.categories.map((_, c) => xOf(c)), slot, bottom);

  // Lines, broken at missing values.
  for (const s of sers) {
    for (let c = 1; c < nCat; c++) {
      const a = val(s, c - 1), b = val(s, c);
      if (a == null || b == null) continue;
      const seg = (x1: number, y1: number, x2: number, y2: number) =>
        ctx.add(`line:s${s}`, { kind: 'line', x1, y1, x2, y2, color: ctx.color(s), width: 2 });
      // Step (ortho) line: hold the previous value across to the next category, then jump vertically.
      if (o.step) {
        seg(xOf(c - 1), yOf(a), xOf(c), yOf(a));
        if (a !== b) seg(xOf(c), yOf(a), xOf(c), yOf(b));
      } else seg(xOf(c - 1), yOf(a), xOf(c), yOf(b));
    }
  }
  const markers: Box[] = [];
  for (const s of sers) {
    for (let c = 0; c < nCat; c++) {
      const v = val(s, c);
      if (v == null) continue;
      const box = { x: xOf(c) - r, y: yOf(v) - r, w: r * 2, h: r * 2 };
      markers.push(box);
      ctx.add(`marker:s${s}:c${c}`, { kind: 'rect', ...box, fill: ctx.color(s), stroke: style.background, strokeWidth: 0.75, ellipse: true });
    }
  }

  // Value labels: try above the point, then below; never on top of another label or marker.
  // Data lines are found by role, not by stroke width (B10).
  const placed: Box[] = [];
  const segments = ctx.shapes.filter((sh, i): sh is LineShape => sh.kind === 'line' && ctx.roles[i].startsWith('line:'));
  const free = (b: Box) =>
    b.y >= frame.y && b.y + b.h <= bottom && b.x >= frame.x && b.x + b.w <= frame.x + frame.w &&
    !placed.some((p) => overlaps(p, b)) && !markers.some((p) => overlaps(p, b, 0)) &&
    !segments.some((l) => segmentHitsBox(l.x1, l.y1, l.x2, l.y2, b));
  // Right-hand series names sit beside the last point, so its labels must not go to the right (B11).
  const namesRight = o.seriesLabels === 'right';
  const lastBoxes: Box[] = [];
  for (let c = 0; c < nCat; c++) {
    if (!labelAll && c !== nCat - 1) continue;
    const pts = sers.filter((s) => val(s, c) != null).sort((a, b) => yOf(val(a, c)!) - yOf(val(b, c)!));
    pts.forEach((s, rank) => {
      const v = val(s, c)!;
      const label = ctx.fmt(v);
      const w = m.width(label, fs) + 2;
      const x = xOf(c) - w / 2;
      const yc = yOf(v);
      const above = { x, y: yc - r - pad * 0.3 - lh, w, h: lh };
      const below = { x, y: yc + r + pad * 0.3, w, h: lh };
      const side = (dir: 1 | -1, dy: number) => ({ x: xOf(c) + dir * (r + pad * 0.4) - (dir < 0 ? w : 0), y: yc - lh / 2 + dy, w, h: lh });
      // The top series prefers above, the bottom series prefers below, so labels move away from each other.
      // Step lines read along their plateaus, so every label sits above its point where it can.
      const vertical = rank === pts.length - 1 && pts.length > 1 && !o.step ? [below, above] : [above, below];
      const sides = namesRight && c === nCat - 1
        ? [side(-1, -lh / 2), side(-1, lh / 2)]
        : [side(1, -lh / 2), side(-1, -lh / 2), side(1, lh / 2), side(-1, lh / 2)];
      // A label may sit on its own point's line ends only when nothing else fits; prefer clean spots first.
      const spot = [...vertical, ...sides].find(free) ?? vertical.find((b) => !placed.some((p) => overlaps(p, b)) && !markers.some((p) => overlaps(p, b, 0)));
      if (!spot) { ctx.hidden.push(`${spec.series[s].name} / ${spec.categories[c]}: no room for label`); return; }
      placed.push(spot);
      if (c === nCat - 1) lastBoxes.push(spot);
      text(ctx, `label:value:s${s}:c${c}`, { ...spot, text: label, color: style.text, role: 'value' });
    });
  }

  if (namesRight) {
    const items = sers
      .map((s) => {
        let c = nCat - 1;
        while (c >= 0 && val(s, c) == null) c--;
        return c < 0 ? null : { s, y: yOf(val(s, c)!) };
      })
      .filter((i): i is { s: number; y: number } => i !== null);
    // Start right of whatever was actually placed at the last point (labels, markers).
    const edge = Math.max(xOf(nCat - 1) + Math.max(r, lastLabelHalf), ...lastBoxes.map((b) => b.x + b.w));
    seriesLabelsRight(ctx, items, edge + pad * 0.5, py - lh, bottom);
  }
}
