import type { Ctx } from '../common.ts';
import { categoryHeight, categoryLabels, legend, seriesLabelsRight, text, valueAxis } from '../common.ts';
import { formatNumber, inferDecimals, niceTicks } from '../numbers.ts';
import { overlaps, segmentHitsBox, type Box } from '../place.ts';

/** Line chart: one point per category, markers, value labels placed above or below without collisions. */
export function layoutLine(ctx: Ctx) {
  const { spec, style, frame, o, fs, lh, pad, m } = ctx;
  const nCat = spec.categories.length;
  const sers = spec.series.map((_, i) => i);
  const val = (s: number, c: number) => spec.series[s].values[c];
  const all = spec.series.flatMap((s) => s.values).filter((v): v is number => v != null);
  if (!all.length) { ctx.hidden.push('line chart has no values'); return; }

  // Include zero unless the data sits in a narrow band far from it (e.g. retention 95–120%).
  const lo = Math.min(...all), hi = Math.max(...all);
  const narrow = lo > 0 && (hi - lo) / hi < 0.4;
  const zero = o.zeroBased ?? !narrow;
  const [vmin, vmax] = zero ? [Math.min(0, lo), Math.max(0, hi)] : [lo, hi];
  const ticks = o.valueAxis === false ? null : niceTicks(vmin, vmax);
  let [dmin, dmax] = ticks ? [ticks[0], ticks[ticks.length - 1]] : [vmin, vmax];
  if (dmax === dmin) dmax = dmin + 1;
  const axisDecimals = ticks ? inferDecimals(ticks) : 0;
  const axisFmt = (v: number) => formatNumber(v, axisDecimals, o.unit ?? '');
  const axisW = ticks ? Math.max(...ticks.map((t) => m.width(axisFmt(t), fs))) + pad * 1.5 : 0;

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
  const py = frame.y + legendH + lh + pad * 0.5;
  const ph = frame.y + frame.h - catH - py - lh * 0.5;
  const yOf = (v: number) => py + ph - ((v - dmin) / (dmax - dmin)) * ph;
  const xOf = (c: number) => px + slot * (c + 0.5);

  if (ticks) valueAxis(ctx, ticks, yOf, px, axisW, axisFmt);
  const bottom = yOf(dmin);
  ctx.shapes.push({ kind: 'line', x1: px, y1: bottom, x2: px + pw, y2: bottom, color: style.axis, width: 0.75 });
  if (dmin < 0 && dmax > 0) {
    ctx.shapes.push({ kind: 'line', x1: px, y1: yOf(0), x2: px + pw, y2: yOf(0), color: style.axis, width: 0.5, dash: true });
  }
  categoryLabels(ctx, spec.categories.map((_, c) => xOf(c)), slot, bottom);

  // Lines, broken at missing values.
  for (const s of sers) {
    for (let c = 1; c < nCat; c++) {
      const a = val(s, c - 1), b = val(s, c);
      if (a == null || b == null) continue;
      ctx.shapes.push({ kind: 'line', x1: xOf(c - 1), y1: yOf(a), x2: xOf(c), y2: yOf(b), color: ctx.color(s), width: 2 });
    }
  }
  const markers: Box[] = [];
  for (const s of sers) {
    for (let c = 0; c < nCat; c++) {
      const v = val(s, c);
      if (v == null) continue;
      const box = { x: xOf(c) - r, y: yOf(v) - r, w: r * 2, h: r * 2 };
      markers.push(box);
      ctx.shapes.push({ kind: 'rect', ...box, fill: ctx.color(s), stroke: style.background, strokeWidth: 0.75, ellipse: true });
    }
  }

  // Value labels: try above the point, then below; never on top of another label or marker.
  const placed: Box[] = [];
  const segments = ctx.shapes.filter((sh) => sh.kind === 'line' && sh.width === 2) as { x1: number; y1: number; x2: number; y2: number }[];
  const free = (b: Box) =>
    b.y >= frame.y && b.y + b.h <= bottom && b.x >= frame.x && b.x + b.w <= frame.x + frame.w &&
    !placed.some((p) => overlaps(p, b)) && !markers.some((p) => overlaps(p, b, 0)) &&
    !segments.some((l) => segmentHitsBox(l.x1, l.y1, l.x2, l.y2, b));
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
      const vertical = rank === pts.length - 1 && pts.length > 1 ? [below, above] : [above, below];
      const order = [...vertical, side(1, -lh / 2), side(-1, -lh / 2), side(1, lh / 2), side(-1, lh / 2)];
      // A label may sit on its own point's line ends only when nothing else fits; prefer clean spots first.
      const spot = order.find(free) ?? vertical.find((b) => !placed.some((p) => overlaps(p, b)) && !markers.some((p) => overlaps(p, b, 0)));
      if (!spot) { ctx.hidden.push(`${spec.series[s].name} / ${spec.categories[c]}: no room for label`); return; }
      placed.push(spot);
      text(ctx, { ...spot, text: label, color: style.text, role: 'value' });
    });
  }

  if (o.seriesLabels === 'right') {
    const items = sers
      .map((s) => {
        let c = nCat - 1;
        while (c >= 0 && val(s, c) == null) c--;
        return c < 0 ? null : { name: spec.series[s].name, y: yOf(val(s, c)!) };
      })
      .filter((i): i is { name: string; y: number } => i !== null);
    seriesLabelsRight(ctx, items, xOf(nCat - 1) + Math.max(r, lastLabelHalf) + pad * 0.5, py - lh, bottom);
  }
}
