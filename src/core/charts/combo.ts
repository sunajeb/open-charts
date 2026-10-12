import type { Ctx } from '../common.ts';
import { axisFormatter, categoryHeight, categoryLabels, legend, num, seriesLabelWidth, seriesLabelsRight, text, valueAxis, valueTicks } from '../common.ts';
import { niceTicks, roundToTotal, textOn } from '../numbers.ts';
import { overlaps, segmentHitsBox, type Box } from '../place.ts';

/**
 * Combination chart: stacked (default) or clustered columns, plus every series marked `as: 'line'`
 * drawn as a line with markers over the same value axis. Column labels that a line or marker would
 * cross move within their segment or are hidden; line labels go above, below or beside their point,
 * never over another series' column, label or marker (hidden and reported when nothing is free).
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
  const ticks = o.valueAxis ? niceTicks(dmin, dmax, valueTicks(ctx)) : null;
  if (ticks) [dmin, dmax] = [ticks[0], ticks[ticks.length - 1]];
  const axisFmt = axisFormatter(ctx, ticks);
  const axisW = ticks ? Math.max(...ticks.map((t) => m.width(axisFmt(t), fs))) + pad * 1.5 : 0;
  const showTotals = !clustered && !!o.totals && colS.length > 0;

  // Reserve space: legend, then room above the tallest mark for totals, bar labels and line labels.
  // Lines default to dark greys so they read over any column colour; a white halo separates them from the bars.
  const lineColors = [style.text, style.muted, style.axis];
  const color = (s: number) => (!isLine(s) || spec.series[s].color ? ctx.color(s) : lineColors[lineS.indexOf(s) % lineColors.length]);
  const legendH = legend(ctx, (s) => (isLine(s) ? 'line' : 'box'), color);
  const aboveH = Math.max(showTotals ? lh + pad * 0.3 : 0, clustered ? lh + pad * 0.3 : 0, lineS.length ? lh + pad * 0.5 : 0);
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

  // Columns.
  const bars: Box[] = [];
  const placed: Box[] = [];
  const consistent = !clustered && !o.format;
  const rounded = cats.map((c) => (consistent ? roundToTotal(colS.map((s) => val(s, c) / ctx.scale.div), ctx.decimals) : null));
  const lastSegs: { s: number; y: number }[] = [];
  for (const c of cats) {
    if (clustered) {
      const barW = colW / Math.max(1, colS.length);
      colS.forEach((s, k) => {
        const v = raw(s, c);
        if (v == null) return;
        const x = cx(c) - colW / 2 + k * barW;
        const y1 = yOf(Math.max(0, v)), y2 = yOf(Math.min(0, v));
        const box = { x, y: y1, w: barW, h: Math.max(0.01, y2 - y1) };
        bars.push(box);
        ctx.add(`bar:s${s}:c${c}`, { kind: 'rect', ...box, fill: ctx.color(s), stroke: style.background, strokeWidth: 0.5 });
        const label = ctx.fmt(v);
        const tw = m.width(label, fs) + 2;
        const lb = { x: x + barW / 2 - tw / 2, y: v >= 0 ? y1 - lh - pad * 0.15 : y2 + pad * 0.15, w: tw, h: lh };
        if (tw > barW + pad * 0.4) ctx.hidden.push(`${spec.series[s].name} / ${spec.categories[c]}: label too wide`);
        else if (crossed(lb)) ctx.hidden.push(`${spec.series[s].name} / ${spec.categories[c]}: label covered by a line`);
        else { placed.push(lb); text(ctx, `label:value:s${s}:c${c}`, { ...lb, text: label, role: 'value' }); }
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
      bars.push(box);
      ctx.add(`bar:s${s}:c${c}`, { kind: 'rect', ...box, fill, stroke: style.background, strokeWidth: 0.75 });
      if (c === nCat - 1) lastSegs.push({ s, y: y1 + h / 2 });
      const label = rounded[c] ? ctx.fmt(rounded[c]![k], { scaled: true }) : ctx.fmt(v);
      const tw = m.width(label, fs) + 2;
      if (tw > colW || lh * 0.9 > h) { ctx.hidden.push(`${spec.series[s].name} / ${spec.categories[c]}: segment too small`); return; }
      // Centred, else as high or as low as the segment allows, whichever the line leaves free.
      const slack = Math.max(0, (h - lh) / 2);
      const spot = [0, -slack, slack].map((dy) => ({ x: cx(c) - tw / 2, y: y1 + h / 2 - lh / 2 + dy, w: tw, h: lh })).find((b2) => !crossed(b2));
      if (!spot) { ctx.hidden.push(`${spec.series[s].name} / ${spec.categories[c]}: label covered by a line`); return; }
      placed.push(spot);
      text(ctx, `label:value:s${s}:c${c}`, { ...spot, text: label, color: textOn(fill), role: 'value' });
    });
    if (showTotals) {
      const label = ctx.fmt(pos[c] + neg[c]);
      const tw = m.width(label, fs, true) + 2;
      // On top of the column; raised in small steps when a line crosses that spot.
      const tb = [0, 1, 2, 3, 4].map((k) => ({ x: cx(c) - tw / 2, y: yOf(up) - lh - pad * 0.2 - k * lh * 0.5, w: tw, h: lh }))
        .find((b2) => b2.y >= frame.y + legendH && !crossed(b2));
      if (!tb) ctx.hidden.push(`${spec.categories[c]}: total covered by a line`);
      else { placed.push(tb); text(ctx, `label:total:c${c}`, { ...tb, text: label, bold: true, role: 'total' }); }
    }
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

  // Line labels: above, below, then beside the point; every candidate must be clear of everything.
  const namesRight = o.seriesLabels === 'right';
  const free = (b: Box) =>
    b.y >= frame.y + legendH && b.y + b.h <= bottom && b.x >= frame.x && b.x + b.w <= px + pw + (namesRight ? 0 : rightW) &&
    !placed.some((p) => overlaps(p, b)) && !bars.some((p) => overlaps(p, b, 0)) && !crossed(b);
  const lastBoxes: Box[] = [];
  const labelAll = nCat * lineS.length <= 24;
  for (const mk of markers) {
    const { s, c } = mk;
    if (!labelAll && c !== nCat - 1) continue;
    const label = ctx.fmt(raw(s, c)!);
    const w = m.width(label, fs) + 2;
    const yc = mk.y + r;
    const vertical = [{ x: cx(c) - w / 2, y: yc - r - pad * 0.3 - lh, w, h: lh }, { x: cx(c) - w / 2, y: yc + r + pad * 0.3, w, h: lh }];
    const side = (dir: 1 | -1, dy: number) => ({ x: cx(c) + dir * (r + pad * 0.4) - (dir < 0 ? w : 0), y: yc - lh / 2 + dy, w, h: lh });
    const sides = namesRight && c === nCat - 1 ? [side(-1, -lh / 2), side(-1, lh / 2)] : [side(1, -lh / 2), side(-1, -lh / 2), side(1, lh / 2), side(-1, lh / 2)];
    // In the gap beside the column, above or below the line, when the point sits on a column.
    const gap = (dir: 1 | -1, dy: number) => ({ x: dir > 0 ? cx(c) + colW / 2 + pad * 0.3 : cx(c) - colW / 2 - pad * 0.3 - w, y: yc + dy, w, h: lh });
    const gaps = namesRight && c === nCat - 1 ? [gap(-1, -lh - pad * 0.3), gap(-1, pad * 0.3)] : [gap(1, -lh - pad * 0.3), gap(-1, -lh - pad * 0.3), gap(1, pad * 0.3), gap(-1, pad * 0.3)];
    const spot = [...vertical, ...sides, ...gaps].find(free);
    if (!spot) { ctx.hidden.push(`${spec.series[s].name} / ${spec.categories[c]}: no room for label`); continue; }
    placed.push(spot);
    if (c === nCat - 1) lastBoxes.push(spot);
    text(ctx, `label:value:s${s}:c${c}`, { ...spot, text: label, role: 'value' });
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
