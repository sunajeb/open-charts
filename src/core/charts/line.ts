import type { Ctx } from '../common.ts';
import {
  axisFormatter, categoryHeight, categoryLabels, legend, num, seriesLabelsRight, text, valueAxis, valueTicks, axisWidth, sharedScale,
} from '../common.ts';
import { niceTicks } from '../numbers.ts';
import { overlaps, segmentHitsBox, type Box } from '../place.ts';
import type { LineShape } from '../types.ts';
import { diffLabel, drawLevelArrow, drawValueLine, valueLineLabelW, valueLineText } from '../decorations.ts';
import { arrowBubble } from '../common.ts';

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
  // Value lines at fixed values (manual: Value lines) widen the axis to show them.
  const vls = (o.valueLines ?? []).map((vl, k) => ({ k, v: vl.value as number, side: vl.position ?? 'right', label: valueLineText(ctx, vl, vl.value as number) }));
  const vlv = vls.map((v) => v.v);
  // A floating axis keeps air below the lowest point: on the bottom line its marker and label would sit
  // on the axis and in the category row.
  const flo = Math.min(lo - (hi - lo) * 0.1, ...vlv);
  const [vmin, vmax] = zero ? [Math.min(0, lo, ...vlv), Math.max(0, hi, ...vlv)] : [flo, Math.max(hi, ...vlv)];
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
  // Level difference arrows between points (manual: Level difference arrows), with their labels.
  const level = (e: { category: number; series?: number } | { valueLine: number }) =>
    ('valueLine' in e ? vls.find((v) => v.k === e.valueLine)?.v : val(e.series ?? 0, e.category)) ?? null;
  const levels = (o.levelArrows ?? []).flatMap((a, k) => {
    const [va, vb] = [level(a.from), level(a.to)];
    if (va == null || vb == null) { ctx.hidden.push(`level arrow ${k + 1}: a point has no value`); return []; }
    const mode = a.mode ?? 'relative';
    const label = diffLabel(ctx, va, vb, mode);
    if (!label) { ctx.hidden.push(`level arrow ${k + 1}: a relative difference needs a positive base`); return []; }
    const cs = [a.from, a.to].flatMap((e) => ('category' in e ? [e.category] : []));
    return [{ k, a, va, vb, mode, label, later: Math.max(...cs), earlier: Math.min(...cs) }];
  });
  const levelW = Math.max(0, ...levels.filter((l) => l.later === nCat - 1 && (l.a.position !== 'between' || l.earlier === l.later))
    .map((l) => arrowBubble(ctx, l.label).w + pad * 1.6 + r));
  const sideW = (side: 'left' | 'right') => Math.max(0, ...vls.filter((v) => v.side === side).map((v) => valueLineLabelW(ctx, v.label) + pad * 0.3));
  const leftW = sideW('left'), decoRightW = sideW('right');
  let slot = 0, catH = 0, rightW = 0;
  const nameW = o.seriesLabels === 'right' ? Math.max(...spec.series.map((s) => m.width(s.name, fs))) : 0;
  for (let i = 0; i < 3; i++) {
    slot = (frame.w - leftW - axisW - rightW) / nCat;
    catH = categoryHeight(ctx, slot);
    const names = o.seriesLabels === 'right' ? nameW + lastLabelHalf + r + pad * 2 : 0;
    rightW = names || levelW || decoRightW ? Math.max(0, names + levelW + decoRightW - slot / 2) : 0;
  }
  const px = frame.x + leftW + axisW;
  const pw = frame.w - leftW - axisW - rightW;
  // Headroom for a label above a point at the top of the scale.
  const py = frame.y + legendH + lh + pad * 0.5 + r;
  const ph = frame.y + frame.h - catH - py - lh * 0.5;
  ({ dmin, dmax, ticks } = sharedScale(ctx, dmin, dmax, ph, ticks, [px, py + ph, px, py]));
  if (ticks) axisFmt = axisFormatter(ctx, ticks);
  const yOf = (v: number) => py + ph - ((v - dmin) / (dmax - dmin)) * ph;
  const xOf = (c: number) => px + slot * (c + 0.5);

  if (ticks) valueAxis(ctx, ticks, yOf, px, axisW, axisFmt, frame.x + leftW);
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
    seriesLabelsRight(ctx, items, edge + pad * 0.5 + levelW, py - lh, bottom);
  }

  // Decorations last, around the labels: value lines, then level arrows.
  const area = { x: frame.x, y: frame.y + legendH, w: frame.w, h: frame.h - legendH };
  for (const vl of vls) {
    drawValueLine(ctx, {
      horizontal: false, w: yOf(vl.v), side: vl.side, label: vl.label, role: `valueline:${vl.k}`, bounds: area,
      u1: vl.side === 'left' ? frame.x + leftW : px, u2: vl.side === 'left' ? px + pw : frame.x + frame.w - decoRightW, reason: `value line ${vl.k + 1}`,
    });
  }
  for (const l of levels) {
    const between = l.a.position === 'between' && l.earlier !== l.later;
    const base = between ? (xOf(l.later - 1) + xOf(l.later)) / 2 : xOf(l.later) + r + pad * 0.8;
    const guides = [l.a.from, l.a.to].flatMap((e) => ('category' in e
      ? [{ u: xOf(e.category) + (between && e.category === l.later ? -r : r), w: yOf(level(e)!) }]
      : []));
    drawLevelArrow(ctx, {
      horizontal: false, us: between ? [base] : [base, base + pad * 0.5, base + pad], guides, mode: l.mode,
      wa: yOf(l.va), wb: yOf(l.vb), label: l.label, role: `arrow:level:${l.k}`, bounds: area, reason: `level arrow ${l.k + 1}`,
    });
  }
}
