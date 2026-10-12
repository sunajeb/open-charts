import type { Ctx } from '../common.ts';
import {
  arrowBubble, arrowRowHeight, axisFormatter, bracketArrow, bracketArrowH, cagrLabel, categoryHeight, categoryLabels,
  fitCategories, legend, num, reportFit, seriesLabelWidth, seriesLabelsRight, text, valueAxis, type ArrowKind, valueTicks,
} from '../common.ts';
import { niceTicks, roundToTotal, textOn } from '../numbers.ts';
import { cagrYears } from '../dates.ts';
import { placePending, type PendingLabel, type Side } from '../labels.ts';
import type { Box } from '../place.ts';

/** Numbers shared by both orientations: stacks, shares, rounded labels and the value domain. */
function model(ctx: Ctx) {
  const { spec, o } = ctx;
  const nCat = spec.categories.length;
  const nSer = spec.series.length;
  const percent = spec.type === 'percent';
  // Values entered as percentages (the datasheet's % toggle): each value is its own share and a
  // column is drawn at its own total, which may be more or less than 100%.
  const pdata = percent && !!o.percentData;
  const clustered = spec.type === 'clustered';
  const raw = (s: number, c: number) => num(spec.series[s].values[c]);
  const val = (s: number, c: number) => raw(s, c) ?? 0;
  const cats = [...Array(nCat).keys()];
  const sers = [...Array(nSer).keys()];
  const pos = cats.map((c) => sers.reduce((a, s) => a + Math.max(0, val(s, c)), 0));
  const neg = cats.map((c) => sers.reduce((a, s) => a + Math.min(0, val(s, c)), 0));
  const total = cats.map((c) => pos[c] + neg[c]);
  const shown = (s: number, c: number) => (pdata ? Math.max(0, val(s, c)) : percent ? (pos[c] ? (Math.max(0, val(s, c)) / pos[c]) * 100 : 0) : val(s, c));
  if (percent && neg.some((v) => v < 0)) ctx.hidden.push('negative values ignored in 100% chart');
  // 100% labels are rounded together so every column reads exactly 100%.
  const pctDecimals = o.decimals ?? (pdata ? ctx.decimals : 0);
  const shares = cats.map((c) => roundToTotal(sers.map((s) => shown(s, c)), pctDecimals));
  // Stacked labels are rounded together so the segments add up to the total shown above the column.
  const consistent = !clustered && !percent && !o.format;
  const rounded = cats.map((c) => (consistent ? roundToTotal(sers.map((s) => val(s, c) / ctx.scale.div), ctx.decimals) : null));
  const segLabel = (s: number, c: number) =>
    percent ? ctx.pct(shares[c][s], pctDecimals) : rounded[c] ? ctx.fmt(rounded[c]![s], { scaled: true }) : ctx.fmt(val(s, c));

  let [dmin, dmax] = percent ? [0, Math.max(100, ...(pdata ? pos : []))]
    : clustered ? [Math.min(0, ...cats.flatMap((c) => sers.map((s) => val(s, c)))), Math.max(0, ...cats.flatMap((c) => sers.map((s) => val(s, c))))]
    : [Math.min(0, ...neg), Math.max(0, ...pos)];
  if (dmax === dmin) dmax = dmin + 1;
  const ticks = o.valueAxis ? niceTicks(dmin, dmax, valueTicks(ctx)) : null;
  if (ticks) [dmin, dmax] = [ticks[0], ticks[ticks.length - 1]];
  const axisFmt = axisFormatter(ctx, ticks, percent);
  const showTotals = !clustered && !!o.totals;
  const totalsFor = (c: number) => (percent ? pos[c] : total[c]);
  const arrows = ([['cagr', o.cagr], ['diff', o.difference]] as const).filter(([, r]) => r && !clustered);
  return { pdata, nCat, nSer, percent, clustered, raw, val, cats, sers, pos, neg, total, shown, shares, pctDecimals, segLabel, dmin, dmax, ticks, axisFmt, showTotals, totalsFor, arrows };
}

type Model = ReturnType<typeof model>;

/** Label for one growth/difference arrow, or null with the reason recorded. */
function arrowLabel(ctx: Ctx, md: Model, kind: ArrowKind, ref: { from: number; to: number; years?: number }): string | null {
  const { from: a, to: b } = ref;
  if (a < 0 || b >= md.nCat || a >= b) { ctx.hidden.push(`${kind}: invalid category range`); return null; }
  const label = kind === 'cagr'
    ? cagrLabel(ctx, md.totalsFor(a), md.totalsFor(b), cagrYears(ctx.spec.categories, a, b, ref.years).years)
    : ctx.fmt(md.totalsFor(b) - md.totalsFor(a), { signed: true });
  if (!label) ctx.hidden.push(`${kind}: needs positive totals`);
  return label;
}

/** Stacked, 100% and clustered column (or, with orientation "horizontal", bar) charts. */
export function layoutColumns(ctx: Ctx) {
  const md = model(ctx);
  if (ctx.o.orientation === 'horizontal') layoutBars(ctx, md);
  else layoutVertical(ctx, md);
}

function layoutVertical(ctx: Ctx, md: Model) {
  const { spec, style, frame, o, fs, lh, pad, m } = ctx;
  const { nCat, nSer, percent, clustered, raw, val, cats, sers, total, shown, segLabel, ticks, axisFmt, showTotals } = md;
  let { dmin, dmax } = md;

  // Reserve space: legend and arrows on top, totals above columns, axis on the left, series labels right, categories below.
  const legendH = legend(ctx);
  const arrowsH = md.arrows.length * arrowRowHeight(ctx);
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
  const py = frame.y + legendH + arrowsH + Math.max(totalsH + barLabelH, ticks ? lh / 2 : 0); // room for the top tick label
  const ph = frame.y + frame.h - catH - (clustered && dmin < 0 ? barLabelH : 0) - py;
  if (ph < lh * 2) ctx.hidden.push('chart frame too small');
  const yOf = (v: number) => py + ph - ((v - dmin) / (dmax - dmin)) * ph;
  const cx = (c: number) => px + slot * (c + 0.5);

  if (ticks) valueAxis(ctx, ticks, yOf, px, axisW, axisFmt);

  // Totals wider than a slot would collide with their neighbours: all go, not every other one.
  const totalText = (c: number) => (md.pdata ? ctx.pct(total[c], md.pctDecimals) : ctx.fmt(total[c]));
  const totalsFit = !showTotals || cats.every((c) => m.width(totalText(c), fs, true) + 2 <= slot);
  if (!totalsFit) ctx.hidden.push('column totals wider than their slots: hidden');

  const tops: number[] = [];
  const lastSegs: { s: number; y: number }[] = [];
  const pending: PendingLabel[] = [];

  for (const c of cats) {
    if (clustered) {
      const barW = colW / nSer;
      let top = yOf(0);
      for (const s of sers) {
        const v = raw(s, c);
        if (v == null) continue; // missing data is an empty slot, not a zero bar (B5)
        const x = cx(c) - colW / 2 + s * barW;
        const y1 = yOf(Math.max(0, v)), y2 = yOf(Math.min(0, v));
        ctx.add(`bar:s${s}:c${c}`, { kind: 'rect', x, y: y1, w: barW, h: Math.max(0.01, y2 - y1), fill: ctx.color(s), stroke: style.background, strokeWidth: 0.5 });
        const label = ctx.fmt(v);
        const tw = m.width(label, fs);
        if (tw <= barW + pad * 0.4) {
          const ly = v >= 0 ? y1 - lh - pad * 0.15 : y2 + pad * 0.15;
          text(ctx, `label:value:s${s}:c${c}`, { x: x + barW / 2 - tw / 2 - 1, y: ly, w: tw + 2, h: lh, text: label, role: 'value' });
        } else ctx.hidden.push(`${spec.series[s].name} / ${spec.categories[c]}: label too wide`);
        top = Math.min(top, y1);
      }
      tops.push(top);
      continue;
    }

    let up = 0, down = 0;
    const x = cx(c) - colW / 2;
    const colBars: Box[] = [];
    for (const s of sers) {
      const v = percent ? Math.max(0, shown(s, c)) : val(s, c);
      if (!v) {
        // Keep the series label anchored at the stack edge even when this segment is empty.
        if (c === nCat - 1) lastSegs.push({ s, y: yOf(up) });
        continue;
      }
      const [a, b] = v > 0 ? [up, up + v] : [down + v, down];
      if (v > 0) up += v; else down += v;
      const y1 = yOf(b), y2 = yOf(a), h = y2 - y1;
      const fill = ctx.color(s);
      const seg = { kind: 'rect', x, y: y1, w: colW, h, fill, stroke: style.background, strokeWidth: 0.75 } as const;
      ctx.add(`bar:s${s}:c${c}`, seg);
      colBars.push(seg);
      const label = segLabel(s, c);
      const tw = m.width(label, fs);
      if (tw + 2 <= colW && lh * 0.9 <= h) {
        text(ctx, `label:value:s${s}:c${c}`, { x: x + colW / 2 - tw / 2 - 1, y: y1 + h / 2 - lh / 2, w: tw + 2, h: lh, text: label, color: textOn(fill), role: 'value' });
      } else {
        // Beside the column (never into the series names right of the last one), or above / below the stack.
        const sides: Side[] = c === nCat - 1 && o.seriesLabels === 'right' ? ['left'] : ['right', 'left'];
        pending.push({ s, c, text: label, seg, sides: [...sides, v > 0 ? 'above' : 'below'], boxCover: colBars, reason: `${spec.series[s].name} / ${spec.categories[c]}: segment too small` });
      }
      if (c === nCat - 1) lastSegs.push({ s, y: y1 + h / 2 });
    }
    let top = yOf(percent && !md.pdata ? 100 : up);
    if (showTotals && totalsFit) {
      const label = totalText(c);
      const tw = m.width(label, fs, true);
      top -= lh + pad * 0.2;
      text(ctx, `label:total:c${c}`, { x: cx(c) - tw / 2 - 1, y: top, w: tw + 2, h: lh, text: label, bold: true, role: 'total' });
    }
    tops.push(top);
  }

  // Baseline.
  ctx.add('axis', { kind: 'line', x1: px, y1: yOf(0), x2: px + pw, y2: yOf(0), color: style.axis, width: 0.75 });
  categoryLabels(ctx, cats.map(cx), slot, Math.max(yOf(0), yOf(dmin)) + (clustered && dmin < 0 ? barLabelH : 0));

  if (o.seriesLabels === 'right' && !clustered) {
    seriesLabelsRight(ctx, lastSegs.reverse(), cx(nCat - 1) + colW / 2, py - totalsH, py + ph);
  }

  // Growth and difference arrows, one row each, above the totals.
  let row = 0;
  for (const [kind, ref] of md.arrows) {
    const label = arrowLabel(ctx, md, kind, ref!);
    if (!label) continue;
    const { from: a, to: b } = ref!;
    const yLine = frame.y + legendH + row * arrowRowHeight(ctx) + arrowRowHeight(ctx) / 2;
    bracketArrow(ctx, kind, cx(a), tops[a] - pad * 0.3, cx(b), tops[b] - pad * 0.3, yLine, label);
    row++;
  }

  // Segment labels that did not fit go where they can, once everything else is drawn.
  placePending(ctx, pending);
}

/** Horizontal bars: categories down the left, values to the right, totals and arrows beyond the bar ends. */
function layoutBars(ctx: Ctx, md: Model) {
  const { spec, style, frame, o, fs, lh, pad, m } = ctx;
  const { nCat, nSer, percent, clustered, raw, val, cats, sers, total, shown, segLabel, ticks, axisFmt, showTotals } = md;
  const { dmin, dmax } = md;
  if (o.seriesLabels === 'right') ctx.o.seriesLabels = 'legend';

  const legendH = legend(ctx);
  const axisH = ticks ? lh + pad * 0.6 : 0;
  // Category column: as wide as the longest label, at most 30% of the frame; labels wrap inside it.
  const catW = Math.min(frame.w * 0.3, Math.max(...spec.categories.map((c) => m.width(c, fs))) + 2);
  const fit = fitCategories(ctx, catW);
  reportFit(ctx, fit);
  const totalLabels = cats.map((c) => (md.pdata ? ctx.pct(total[c], md.pctDecimals) : ctx.fmt(total[c])));
  const totalsW = showTotals ? Math.max(...totalLabels.map((t) => m.width(t, fs, true))) + pad * 0.8 : 0;
  const valueLabels = clustered ? cats.flatMap((c) => sers.map((s) => raw(s, c)).filter((v): v is number => v != null).map((v) => ({ v, w: m.width(ctx.fmt(v), fs) }))) : [];
  const barLabelR = Math.max(0, ...valueLabels.filter((l) => l.v >= 0).map((l) => l.w + pad * 0.6));
  const barLabelL = Math.max(0, ...valueLabels.filter((l) => l.v < 0).map((l) => l.w + pad * 0.6));
  const arrowLabels = md.arrows.map(([kind, ref]) => ({ kind, ref: ref!, label: arrowLabel(ctx, md, kind, ref!) }));
  const arrowW = arrowLabels.map((a) => (a.label ? arrowBubble(ctx, a.label).w + pad : 0));
  const arrowsW = arrowW.reduce((a, w) => a + w, 0) + (arrowW.some(Boolean) ? pad : 0);

  const px = frame.x + catW + pad + barLabelL;
  const pw = frame.x + frame.w - totalsW - barLabelR - arrowsW - px;
  const py = frame.y + legendH;
  const ph = frame.y + frame.h - axisH - py;
  if (pw < lh * 2 || ph < lh * nCat * 0.5) ctx.hidden.push('chart frame too small');
  const slot = ph / nCat;
  const barT = slot / (1 + o.gap);
  const xOf = (v: number) => px + ((v - dmin) / (dmax - dmin)) * pw;
  const cy = (c: number) => py + slot * (c + 0.5);

  // Value axis along the bottom.
  if (ticks) {
    const yA = py + ph;
    ctx.add('axis', { kind: 'line', x1: xOf(ticks[0]), y1: yA, x2: xOf(ticks[ticks.length - 1]), y2: yA, color: style.axis, width: 0.75 });
    ticks.forEach((t, k) => {
      const label = axisFmt(t);
      const w = m.width(label, fs) + 2;
      ctx.add(`axis:tick:${k}`, { kind: 'line', x1: xOf(t), y1: yA, x2: xOf(t), y2: yA + 3, color: style.axis, width: 0.75 });
      text(ctx, `axis:tick:${k}`, { x: xOf(t) - w / 2, y: yA + pad * 0.5, w, h: lh, text: label, color: style.muted, role: 'axis' });
    });
  }

  const ends: number[] = [];
  const pending: PendingLabel[] = [];
  for (const c of cats) {
    if (clustered) {
      const barH = barT / nSer;
      let end = xOf(0);
      for (const s of sers) {
        const v = raw(s, c);
        if (v == null) continue;
        const y = cy(c) - barT / 2 + s * barH;
        const x1 = xOf(Math.min(0, v)), x2 = xOf(Math.max(0, v));
        ctx.add(`bar:s${s}:c${c}`, { kind: 'rect', x: x1, y, w: Math.max(0.01, x2 - x1), h: barH, fill: ctx.color(s), stroke: style.background, strokeWidth: 0.5 });
        const label = ctx.fmt(v);
        const tw = m.width(label, fs) + 2;
        if (lh * 0.9 <= barH + pad * 0.4) {
          const lx = v >= 0 ? x2 + pad * 0.3 : x1 - pad * 0.3 - tw;
          text(ctx, `label:value:s${s}:c${c}`, { x: lx, y: y + barH / 2 - lh / 2, w: tw, h: lh, text: label, role: 'value' });
        } else ctx.hidden.push(`${spec.series[s].name} / ${spec.categories[c]}: bar too thin for label`);
        end = Math.max(end, x2);
      }
      ends.push(end);
      continue;
    }

    let right = 0, left = 0;
    const y = cy(c) - barT / 2;
    const colBars: Box[] = [];
    for (const s of sers) {
      const v = percent ? Math.max(0, shown(s, c)) : val(s, c);
      if (!v) continue;
      const [a, b] = v > 0 ? [right, right + v] : [left + v, left];
      if (v > 0) right += v; else left += v;
      const x1 = xOf(a), x2 = xOf(b), w = x2 - x1;
      const fill = ctx.color(s);
      const seg = { kind: 'rect', x: x1, y, w, h: barT, fill, stroke: style.background, strokeWidth: 0.75 } as const;
      ctx.add(`bar:s${s}:c${c}`, seg);
      colBars.push(seg);
      const label = segLabel(s, c);
      const tw = m.width(label, fs);
      if (tw + 2 <= w && lh * 0.9 <= barT) {
        text(ctx, `label:value:s${s}:c${c}`, { x: x1 + w / 2 - tw / 2 - 1, y: cy(c) - lh / 2, w: tw + 2, h: lh, text: label, color: textOn(fill), role: 'value' });
      } else {
        // In the gap above or below the bar, or beyond its end.
        pending.push({ s, c, text: label, seg, sides: ['above', 'below', v > 0 ? 'right' : 'left'], boxCover: colBars, reason: `${spec.series[s].name} / ${spec.categories[c]}: segment too small` });
      }
    }
    let end = xOf(percent && !md.pdata ? 100 : right);
    if (showTotals) {
      const tw = m.width(totalLabels[c], fs, true) + 2;
      text(ctx, `label:total:c${c}`, { x: end + pad * 0.4, y: cy(c) - lh / 2, w: tw, h: lh, text: totalLabels[c], bold: true, align: 'left', role: 'total' });
      end += pad * 0.4 + tw;
    }
    ends.push(end);
  }

  // Baseline at zero, then category labels right-aligned in their column.
  ctx.add('axis', { kind: 'line', x1: xOf(0), y1: py, x2: xOf(0), y2: py + ph, color: style.axis, width: 0.75 });
  for (const c of cats) {
    if (fit.hidden.has(c)) continue;
    const lines = fit.lines[c];
    const h = lines.length * fit.lh;
    text(ctx, `label:category:c${c}`, { x: frame.x, y: cy(c) - h / 2, w: catW, h, text: lines.join('\n'), align: 'right', size: fit.size, role: 'category' });
  }

  // Arrows in columns to the right, innermost first.
  let xCol = frame.x + frame.w - arrowsW + pad;
  arrowLabels.forEach(({ kind, ref, label }, i) => {
    if (!label) return;
    const xLine = xCol + arrowW[i] / 2;
    bracketArrowH(ctx, kind, ends[ref.from] + pad * 0.3, cy(ref.from), ends[ref.to] + pad * 0.3, cy(ref.to), xLine, label);
    xCol += arrowW[i];
  });

  placePending(ctx, pending);
}
