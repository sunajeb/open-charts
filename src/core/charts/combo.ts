import type { Ctx } from '../common.ts';
import {
  axisFormatter, categoryHeight, categoryLabels, legend, num, seriesLabelWidth, seriesLabelsRight, text, valueAxis, valueTicks, axisWidth, sharedScale,
} from '../common.ts';
import { niceTicks, roundToTotal, textOn } from '../numbers.ts';
import { overlaps, segmentHitsBox, type Box } from '../place.ts';

/**
 * Combination chart: stacked (default) or clustered columns, plus every series marked `as: 'line'`
 * drawn as a line with markers over the same value axis.
 *
 * - Segment labels stay inside their segment: centred, else moved within it clear of the line, else
 *   hidden and reported.
 * - A total sits on its column and is raised in small steps only when the line runs through it.
 * - A line label sits directly above or below its point (nudged just clear of a column edge), else
 *   beside it; when the point lies inside a column and none of these is off the columns, straight
 *   above the column (under its total). It sits on a column — wholly inside one segment, in a contrasting
 *   colour, never across an edge — only when no spot off the columns is free. Totals and line labels
 *   of a category are placed together, keeping a small clearance from every other label; what finds
 *   no spot is hidden and reported.
 */
export function layoutCombo(ctx: Ctx) {
  const { spec, style, frame, o, fs, lh, pad, m } = ctx;
  const clustered = (o.combo ?? 'stacked') === 'clustered';
  if (spec.options?.seriesLabels === undefined) ctx.o.seriesLabels = clustered ? 'legend' : 'right';
  const nCat = spec.categories.length;
  const cats = [...Array(nCat).keys()];
  const isLine = (s: number) => spec.series[s].as === 'line';
  const colS = spec.series.map((_, i) => i).filter((s) => !isLine(s));
  const lineS = spec.series.map((_, i) => i).filter(isLine);
  const raw = (s: number, c: number) => num(spec.series[s].values[c]);
  const val = (s: number, c: number) => raw(s, c) ?? 0;

  const pos = cats.map((c) => colS.reduce((a, s) => a + Math.max(0, val(s, c)), 0));
  const neg = cats.map((c) => colS.reduce((a, s) => a + Math.min(0, val(s, c)), 0));
  const lineVals = lineS.flatMap((s) => cats.map((c) => raw(s, c))).filter((v): v is number => v != null);
  const colVals = colS.flatMap((s) => cats.map((c) => val(s, c)));
  let dmin = clustered ? Math.min(0, ...colVals, ...lineVals) : Math.min(0, ...neg, ...lineVals);
  let dmax = clustered ? Math.max(0, ...colVals, ...lineVals) : Math.max(0, ...pos, ...lineVals);
  if (dmax === dmin) dmax = dmin + 1;
  let ticks = o.valueAxis ? niceTicks(dmin, dmax, valueTicks(ctx)) : null;
  if (ticks) [dmin, dmax] = [ticks[0], ticks[ticks.length - 1]];
  let axisFmt = axisFormatter(ctx, ticks);
  const axisW = axisWidth(ctx, dmin, ticks, axisFmt);
  const showTotals = !clustered && !!o.totals && colS.length > 0;

  // Reserve space: legend, then room above the tallest mark for totals, bar labels and line labels.
  // Lines default to dark greys so they read over any column colour; a white halo separates them from the bars.
  const lineColors = [style.text, style.muted, style.axis];
  const color = (s: number) => (!isLine(s) || spec.series[s].color ? ctx.color(s) : lineColors[lineS.indexOf(s) % lineColors.length]);
  const legendH = legend(ctx, (s) => (isLine(s) ? 'line' : 'box'), color);
  // A line point inside a stack labels above the stack, under the total: room for both.
  const inStack = !clustered && lineS.some((s) => cats.some((c) => { const v = raw(s, c); return v != null && v > 0 && v < pos[c]; }));
  const aboveH = Math.max(showTotals ? lh + pad * 0.3 : 0, clustered ? lh + pad * 0.3 : 0, lineS.length ? lh + pad * 0.5 : 0,
    inStack ? (showTotals ? 2.5 : 1) * lh + pad * 0.5 : 0);
  const lastLineHalf = Math.max(0, ...lineS.map((s) => (raw(s, nCat - 1) == null ? 0 : m.width(ctx.fmt(raw(s, nCat - 1)!), fs) / 2)));
  let rightW = 0, slot = 0, colW = 0, catH = 0;
  for (let i = 0; i < 3; i++) {
    slot = (frame.w - axisW - rightW) / nCat;
    colW = slot / (1 + o.gap);
    catH = categoryHeight(ctx, slot);
    rightW = seriesLabelWidth(ctx, (slot - Math.max(colW, lastLineHalf * 2)) / 2);
  }
  const px = frame.x + axisW;
  const pw = frame.w - axisW - rightW;
  const py = frame.y + legendH + aboveH;
  const below = dmin < 0 ? lh + pad * 0.3 : 0;
  const ph = frame.y + frame.h - catH - below - py;
  ({ dmin: dmin, dmax: dmax, ticks } = sharedScale(ctx, dmin, dmax, ph, ticks));
  if (ticks) axisFmt = axisFormatter(ctx, ticks);
  if (ph < lh * 2) ctx.hidden.push('chart frame too small');
  const yOf = (v: number) => py + ph - ((v - dmin) / (dmax - dmin)) * ph;
  const cx = (c: number) => px + slot * (c + 0.5);
  const r = Math.max(3, fs * 0.3);
  const bottom = Math.max(yOf(0), yOf(dmin));

  if (ticks) valueAxis(ctx, ticks, yOf, px, axisW, axisFmt);

  // Line geometry first, so column labels can keep clear of it.
  const segs: { s: number; x1: number; y1: number; x2: number; y2: number }[] = [];
  const markers: (Box & { s: number; c: number })[] = [];
  for (const s of lineS) {
    for (let c = 0; c < nCat; c++) {
      const v = raw(s, c);
      if (v == null) continue;
      markers.push({ s, c, x: cx(c) - r, y: yOf(v) - r, w: 2 * r, h: 2 * r });
      const b = raw(s, c + 1);
      if (c + 1 < nCat && b != null) segs.push({ s, x1: cx(c), y1: yOf(v), x2: cx(c + 1), y2: yOf(b) });
    }
  }
  const crossed = (b: Box) => segs.some((l) => segmentHitsBox(l.x1, l.y1, l.x2, l.y2, b)) || markers.some((mk) => overlaps(mk, b, 0));
  // The line's white halo counts as part of the line; labels keep a small clearance from each other.
  const lineHits = (b: Box) => crossed(grow(b, 1));
  const gapL = pad * 0.25;
  const placed: Box[] = [];
  const clear = (b: Box, extra: Box[] = []) => ![...placed, ...extra].some((p) => overlaps(grow(p, gapL), b, 0));

  // Columns.
  const bars: (Box & { fill: string })[] = [];
  /** Top of each category's columns and of the value labels above them. */
  const colTop = cats.map(() => Infinity);
  const consistent = !clustered && !o.format;
  const rounded = cats.map((c) => (consistent ? roundToTotal(colS.map((s) => val(s, c) / ctx.scale.div), ctx.decimals) : null));
  const lastSegs: { s: number; y: number }[] = [];
  const tops = cats.map(() => 0);
  for (const c of cats) {
    if (clustered) {
      const barW = colW / Math.max(1, colS.length);
      colS.forEach((s, k) => {
        const v = raw(s, c);
        if (v == null) return;
        const x = cx(c) - colW / 2 + k * barW;
        const y1 = yOf(Math.max(0, v)), y2 = yOf(Math.min(0, v));
        const box = { x, y: y1, w: barW, h: Math.max(0.01, y2 - y1) };
        bars.push({ ...box, fill: ctx.color(s) });
        colTop[c] = Math.min(colTop[c], y1);
        ctx.add(`bar:s${s}:c${c}`, { kind: 'rect', ...box, fill: ctx.color(s), stroke: style.background, strokeWidth: 0.5 });
        const label = ctx.fmt(v);
        const tw = m.width(label, fs) + 2;
        const lb = { x: x + barW / 2 - tw / 2, y: v >= 0 ? y1 - lh - pad * 0.15 : y2 + pad * 0.15, w: tw, h: lh };
        if (tw > barW + pad * 0.4) ctx.hidden.push(`${spec.series[s].name} / ${spec.categories[c]}: label too wide`);
        else if (lineHits(lb)) ctx.hidden.push(`${spec.series[s].name} / ${spec.categories[c]}: label covered by a line`);
        else { placed.push(lb); colTop[c] = Math.min(colTop[c], lb.y); text(ctx, `label:value:s${s}:c${c}`, { ...lb, text: label, role: 'value' }); }
      });
      continue;
    }
    let up = 0, down = 0;
    const x = cx(c) - colW / 2;
    colS.forEach((s, k) => {
      const v = val(s, c);
      if (!v) { if (c === nCat - 1) lastSegs.push({ s, y: yOf(up) }); return; }
      const [a, b] = v > 0 ? [up, up + v] : [down + v, down];
      if (v > 0) up += v; else down += v;
      const y1 = yOf(b), h = yOf(a) - y1;
      const fill = ctx.color(s);
      const box = { x, y: y1, w: colW, h };
      bars.push({ ...box, fill });
      ctx.add(`bar:s${s}:c${c}`, { kind: 'rect', ...box, fill, stroke: style.background, strokeWidth: 0.75 });
      if (c === nCat - 1) lastSegs.push({ s, y: y1 + h / 2 });
      const label = rounded[c] ? ctx.fmt(rounded[c]![k], { scaled: true }) : ctx.fmt(v);
      const tw = m.width(label, fs) + 2;
      if (tw > colW || lh * 0.9 > h) { ctx.hidden.push(`${spec.series[s].name} / ${spec.categories[c]}: segment too small`); return; }
      // Centred, else moved up or down within the segment (never to its edge) until the line leaves it free.
      // When a line point sits in the segment, the label moves away from it to leave the point's label room.
      const slack = Math.max(0, (h - lh) / 2 - pad * 0.5);
      const pts = markers.filter((mk) => mk.c === c && mk.y + r >= y1 && mk.y + r <= y1 + h).map((mk) => mk.y + r);
      const dist = (b2: Box) => Math.min(Infinity, ...pts.map((yp) => Math.abs(b2.y + lh / 2 - yp)));
      const cands = [0, -0.5, 0.5, -1, 1].map((k2) => ({ x: cx(c) - tw / 2, y: y1 + h / 2 - lh / 2 + k2 * slack, w: tw, h: lh }));
      if (pts.length && dist(cands[0]) < lh * 1.6) cands.sort((p, q) => dist(q) - dist(p));
      const spot = cands.find((b2) => !lineHits(b2));
      if (!spot) { ctx.hidden.push(`${spec.series[s].name} / ${spec.categories[c]}: label covered by a line`); return; }
      placed.push(spot);
      text(ctx, `label:value:s${s}:c${c}`, { ...spot, text: label, color: textOn(fill), role: 'value' });
    });
    tops[c] = up;
    colTop[c] = yOf(up);
  }

  ctx.add('axis', { kind: 'line', x1: px, y1: yOf(0), x2: px + pw, y2: yOf(0), color: style.axis, width: 0.75 });
  categoryLabels(ctx, cats.map(cx), slot, bottom + below);

  // Lines and markers on top of the columns.
  for (const l of segs) {
    ctx.add(`line:s${l.s}`,
      { kind: 'line', x1: l.x1, y1: l.y1, x2: l.x2, y2: l.y2, color: style.background, width: 4 },
      { kind: 'line', x1: l.x1, y1: l.y1, x2: l.x2, y2: l.y2, color: color(l.s), width: 2 });
  }
  for (const mk of markers) {
    ctx.add(`marker:s${mk.s}:c${mk.c}`, { kind: 'rect', x: mk.x, y: mk.y, w: mk.w, h: mk.h, fill: color(mk.s), stroke: style.background, strokeWidth: 1, geom: 'ellipse' });
  }

  // Totals and line labels, one category at a time.
  const namesRight = o.seriesLabels === 'right';
  const inBounds = (b: Box) =>
    b.y >= frame.y + legendH && b.y + b.h <= bottom && b.x >= frame.x && b.x + b.w <= px + pw + (namesRight ? 0 : rightW);
  /** The fill under a label box: '' over free space, a segment's fill when wholly inside it, null across an edge. */
  const under = (b: Box): string | null => {
    const hit = bars.filter((p) => overlaps(p, b, 0));
    if (!hit.length) return '';
    const [h0] = hit;
    const inside = hit.length === 1 && b.x >= h0.x + 1 && b.x + b.w <= h0.x + h0.w - 1 && b.y >= h0.y + 1 && b.y + b.h <= h0.y + h0.h - 1;
    return inside ? h0.fill : null;
  };
  // On a segment, keep well apart from the segment's own label so the two read as separate numbers.
  const free = (b: Box, extra: Box[]) => {
    const u = inBounds(b) && !lineHits(b) ? under(b) : null;
    return u !== null && clear(u ? grow(b, lh * 0.4) : b, extra);
  };
  const lineSpots = (mk: Box & { s: number; c: number }, preferBelow: boolean): Box[] => {
    const w = m.width(ctx.fmt(raw(mk.s, mk.c)!), fs) + 2;
    const xc = mk.x + r, yc = mk.y + r, g = pad * 0.3;
    const vertical = (dir: 1 | -1) => [0, 0.15, 0.3, 0.45, 0.6, 0.8].map((k) => ({ x: xc - w / 2, y: dir < 0 ? yc - r - g - lh - k * lh : yc + r + g + k * lh, w, h: lh }));
    const side = (dir: 1 | -1, above: boolean) => ({ x: xc + dir * (r + g) - (dir < 0 ? w : 0), y: above ? yc - g - lh : yc + g, w, h: lh });
    const sides = namesRight && mk.c === nCat - 1 ? [side(-1, true), side(-1, false)] : [side(1, true), side(-1, true), side(1, false), side(-1, false)];
    // A point inside its columns: straight above them.
    const top = colTop[mk.c];
    const overStack = yc > top ? [0, 0.15, 0.3, 0.45].map((k) => ({ x: xc - w / 2, y: top - g - lh - k * lh, w, h: lh })) : [];
    return [...(preferBelow ? [...vertical(1), ...vertical(-1)] : [...vertical(-1), ...vertical(1)]), ...sides, ...overStack];
  };
  const labelAll = nCat * lineS.length <= 24;
  const lastBoxes: Box[] = [];
  for (const c of cats) {
    // Top line first; a lower line at the same category prefers the spot below its point.
    const mks = labelAll || c === nCat - 1 ? markers.filter((mk) => mk.c === c).sort((a, b) => a.y - b.y) : [];
    const totalLabel = showTotals ? ctx.fmt(pos[c] + neg[c]) : null;
    const tw = totalLabel ? m.width(totalLabel, fs, true) + 2 : 0;
    const totals: (Box | null)[] = totalLabel
      ? [0, 1, 2, 3, 4, 5, 6].map((k) => ({ x: cx(c) - tw / 2, y: yOf(tops[c]) - lh - pad * 0.2 - k * lh * 0.5, w: tw, h: lh }))
        .filter((b) => b.y >= frame.y + legendH && !lineHits(b) && clear(b))
      : [null];
    // Cheapest combination: every label placed, the total low, line labels in their first-choice spots
    // and off the columns (a label on a segment reads as the segment's value: a last resort).
    let best: { tb: Box | null; spots: (Box | null)[]; cost: number } | null = null;
    totals.forEach((tb, k) => {
      const taken: Box[] = tb ? [tb] : [];
      let cost = k * 3;
      const spots = mks.map((mk, i) => {
        const cands = lineSpots(mk, i > 0);
        const ok = cands.map((b) => free(b, taken));
        let j = cands.findIndex((b, q) => ok[q] && under(b) === '');
        if (j < 0) j = ok.indexOf(true);
        if (j < 0) { cost += 1000; return null; }
        cost += j + (under(cands[j]) ? 100 : 0);
        taken.push(cands[j]);
        return cands[j];
      });
      if (!best || cost < best.cost) best = { tb, spots, cost };
    });
    if (totalLabel) {
      if (!best?.tb) ctx.hidden.push(`${spec.categories[c]}: total covered by a line`);
      else { placed.push(best.tb); text(ctx, `label:total:c${c}`, { ...best.tb, text: totalLabel, bold: true, role: 'total' }); }
    }
    mks.forEach((mk, i) => {
      const spot = best?.spots[i];
      if (!spot) { ctx.hidden.push(`${spec.series[mk.s].name} / ${spec.categories[c]}: no room for label`); return; }
      placed.push(spot);
      if (c === nCat - 1) lastBoxes.push(spot);
      const fill = under(spot);
      text(ctx, `label:value:s${mk.s}:c${c}`, { ...spot, text: ctx.fmt(raw(mk.s, c)!), role: 'value', ...(fill ? { color: textOn(fill) } : {}) });
    });
  }

  if (namesRight) {
    const items = spec.series.map((_, s) => {
      if (!isLine(s)) return lastSegs.find((it) => it.s === s) ?? null;
      const mk = [...markers].reverse().find((it) => it.s === s);
      return mk ? { s, y: mk.y + r } : null;
    }).filter((i): i is { s: number; y: number } => i !== null).sort((a, b) => a.y - b.y);
    const edge = Math.max(cx(nCat - 1) + colW / 2, ...lastBoxes.map((b) => b.x + b.w), ...markers.filter((mk) => mk.c === nCat - 1).map((mk) => mk.x + mk.w));
    seriesLabelsRight(ctx, items, edge, frame.y + legendH, bottom);
  }
}

const grow = (b: Box, d: number): Box => ({ x: b.x - d, y: b.y - d, w: b.w + 2 * d, h: b.h + 2 * d });
