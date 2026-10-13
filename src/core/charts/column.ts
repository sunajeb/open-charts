import type { Ctx } from '../common.ts';
import {
  arrowBubble, arrowRowHeight, axisFormatter, bracketArrow, bracketArrowH, cagrLabel, categoryHeight, categoryLabels, fitCategories, legend, num, reportFit, seriesLabelWidth, seriesLabelsRight, text, valueAxis, type ArrowKind, type Heads, valueTicks, axisWidth, sharedScale,
} from '../common.ts';
import { niceTicks, roundToTotal, textOn } from '../numbers.ts';
import { cagrYears } from '../dates.ts';
import { placePending, type PendingLabel, type Side } from '../labels.ts';
import type { Box } from '../place.ts';
import type { DiffMode, LevelEnd } from '../types.ts';
import {
  axisDistance, breakSymbol, brokenTicks, categoryBreakSymbol, diffLabel, drawIndicator, drawLevelArrow, drawValueLine,
  indicatorSize, resolveBreak, slotLayout, slotUnits, spacedTicks, valueLineLabelW, valueLineText, type Break,
} from '../decorations.ts';
import { arrowBubble as bubbleSize, labelSpec } from '../common.ts';

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
  // The datasheet's "100%=" row: the value a column's shares are taken of (default: its sum).
  const base = (c: number): number | null => (percent ? o.percentBase?.[c] ?? null : null);
  const hasBase = percent && !!o.percentBase?.some((b) => b != null);
  const shown = (s: number, c: number) => {
    if (pdata) return Math.max(0, val(s, c));
    if (!percent) return val(s, c);
    const b = base(c) ?? pos[c];
    return b ? (Math.max(0, val(s, c)) / b) * 100 : 0;
  };
  if (percent && neg.some((v) => v < 0)) ctx.hidden.push('negative values ignored in 100% chart');
  // 100% labels are rounded together so every column reads exactly 100%.
  const pctDecimals = o.decimals ?? (pdata ? ctx.decimals : 0);
  const shares = cats.map((c) => roundToTotal(sers.map((s) => shown(s, c)), pctDecimals));
  // Stacked labels are rounded together so the segments add up to the total shown above the column.
  const consistent = !clustered && !percent && !o.format;
  const rounded = cats.map((c) => (consistent ? roundToTotal(sers.map((s) => val(s, c) / ctx.scale.div), ctx.decimals) : null));
  const segLabel = (s: number, c: number) =>
    percent ? ctx.pct(shares[c][s], pctDecimals) : rounded[c] ? ctx.fmt(rounded[c]![s], { scaled: true }) : ctx.fmt(val(s, c));

  // Top edge of every segment (the "levels" difference arrows connect), in axis units.
  const stackTop = cats.map((c) => (percent ? sers.reduce((a, s) => a + Math.max(0, shown(s, c)), 0) : pos[c]));
  const levelOf = (c: number, s?: number): number => {
    if (clustered) return raw(s ?? 0, c) ?? 0;
    if (s === undefined) return stackTop[c];
    let up = 0, down = 0;
    for (const k of sers) {
      const v = percent ? Math.max(0, shown(k, c)) : val(k, c);
      if (v > 0) up += v; else down += v;
      if (k === s) return v < 0 ? down : up;
    }
    return up;
  };
  const boundaries = clustered ? cats.flatMap((c) => sers.map((s) => val(s, c))) : cats.flatMap((c) => sers.map((s) => levelOf(c, s)));

  let [dmin, dmax] = percent ? [0, Math.max(100, ...(pdata ? pos : hasBase ? stackTop : []))]
    : clustered ? [Math.min(0, ...cats.flatMap((c) => sers.map((s) => val(s, c)))), Math.max(0, ...cats.flatMap((c) => sers.map((s) => val(s, c))))]
    : [Math.min(0, ...neg), Math.max(0, ...pos)];
  // Value lines (manual: snap to the arithmetic mean of the column totals; a clustered chart's columns are
  // its bars). A line at a fixed value beyond the data widens the axis to show it.
  const means = clustered ? cats.flatMap((c) => sers.map((s) => raw(s, c)).filter((v): v is number => v != null)) : total;
  const vlValues = (o.valueLines ?? []).map((vl) => (vl.value === 'mean' ? means.reduce((a, x) => a + x, 0) / Math.max(1, means.length) : vl.value));
  if (!percent) [dmin, dmax] = [Math.min(dmin, ...vlValues), Math.max(dmax, ...vlValues)];
  if (dmax === dmin) dmax = dmin + 1;
  // A value axis break compresses a range no segment ends in (manual: Value axis breaks).
  const brk: Break | null = resolveBreak(ctx, o.axisBreak, dmin, dmax, [0, ...boundaries]);
  const count = valueTicks(ctx);
  const ticks = !o.valueAxis ? null : brk ? brokenTicks(dmin, dmax, brk, count) : niceTicks(dmin, dmax, count);
  if (ticks) [dmin, dmax] = [Math.min(dmin, ticks[0]), Math.max(dmax, ticks[ticks.length - 1])];
  const axisFmt = axisFormatter(ctx, ticks, percent);
  const showTotals = !clustered && !!o.totals;
  const totalsFor = (c: number) => (percent ? (pdata ? base(c) ?? pos[c] : pos[c]) : total[c]);
  const totalText = (c: number) => (pdata ? (base(c) != null ? ctx.fmt(base(c)!) : ctx.pct(total[c], pctDecimals)) : ctx.fmt(total[c]));
  const arrows = ([['cagr', o.cagr], ['diff', o.difference]] as const).filter(([, r]) => r && !clustered);
  const valueLines = (o.valueLines ?? []).flatMap((vl, k) => {
    const v = vlValues[k];
    if (brk && v > brk.from && v < brk.to) { ctx.hidden.push(`value line ${k + 1}: inside the axis break`); return []; }
    if (v < dmin || v > dmax) { ctx.hidden.push(`value line ${k + 1}: outside the value axis`); return []; }
    return [{ k, v, side: vl.position ?? 'right', label: valueLineText(ctx, vl, v) }];
  });
  return {
    pdata, nCat, nSer, percent, clustered, raw, val, cats, sers, pos, neg, total, shown, shares, pctDecimals, segLabel,
    dmin, dmax, ticks, axisFmt, showTotals, totalsFor, totalText, arrows, levelOf, brk, valueLines, hasBase,
  };
}

type Model = ReturnType<typeof model>;

/** Arrow heads of a total difference arrow per mode (relative → to, reverse → from, absolute ↔ both). */
const headsOf = (kind: ArrowKind, ref: { mode?: DiffMode }): Heads =>
  kind === 'cagr' ? 'end' : ref.mode === 'relative' ? 'end' : ref.mode === 'reverse' ? 'start' : 'both';

/** Value of one end of a level arrow, in axis units (a value line end: the line's value). */
function endValue(md: Model, end: LevelEnd): number {
  return 'valueLine' in end ? md.valueLines.find((v) => v.k === end.valueLine)?.v ?? NaN : md.levelOf(end.category, end.series);
}

/** Level arrows with their labels (relative mode by default); arrows whose label is undefined are reported. */
function levelArrows(ctx: Ctx, md: Model) {
  return (ctx.o.levelArrows ?? []).flatMap((a, k) => {
    const mode = a.mode ?? 'relative';
    const [va, vb] = [endValue(md, a.from), endValue(md, a.to)];
    if (Number.isNaN(va) || Number.isNaN(vb)) { ctx.hidden.push(`level arrow ${k + 1}: its value line is not drawn`); return []; }
    const label = diffLabel(ctx, va, vb, mode, md.percent);
    if (!label) { ctx.hidden.push(`level arrow ${k + 1}: a relative difference needs a positive base`); return []; }
    const cats = [a.from, a.to].flatMap((e) => ('category' in e ? [e.category] : []));
    return [{ k, a, mode, label, later: Math.max(...cats), earlier: Math.min(...cats) }];
  });
}

/** Where the 100% indicator goes: the option, else right — left when series names take the right, none if a value axis is there too. */
function indicatorSide(ctx: Ctx, md: Model): 'left' | 'right' | null {
  if (!md.percent) return null;
  const opt = ctx.o.percentIndicator;
  if (opt === 'none') return null;
  if (opt) return opt;
  // A value axis already reads 100% on its side (bars: along the bottom).
  if (ctx.o.orientation === 'horizontal') return md.ticks ? null : 'right';
  if (ctx.o.seriesLabels !== 'right') return 'right';
  return md.ticks ? null : 'left';
}

/** Label for one growth/difference arrow, or null with the reason recorded. */
function arrowLabel(ctx: Ctx, md: Model, kind: ArrowKind, ref: { from: number; to: number; years?: number }): string | null {
  const { from: a, to: b } = ref;
  if (a < 0 || b >= md.nCat || a >= b) { ctx.hidden.push(`${kind}: invalid category range`); return null; }
  const label = kind === 'cagr'
    ? cagrLabel(ctx, md.totalsFor(a), md.totalsFor(b), cagrYears(ctx.spec.categories, a, b, ref.years).years)
    : diffLabel(ctx, md.totalsFor(a), md.totalsFor(b), (ref as { mode?: DiffMode }).mode ?? 'absolute');
  if (!label) ctx.hidden.push(`${kind}: needs positive totals`);
  return label;
}

/** Stacked, 100% and clustered column (or, with orientation "horizontal", bar) charts. */
export function layoutColumns(ctx: Ctx) {
  if (ctx.o.orientation === 'horizontal') { layoutBars(ctx, model(ctx)); return; }
  if (layoutVertical(ctx, model(ctx)) !== 'labels-too-wide' || ctx.o.valueAxis) return;
  // Clustered bars too narrow for their value labels even at the smallest size: a value axis carries the
  // numbers instead (a row of labels with gaps reads as missing data).
  ctx.shapes.length = 0;
  ctx.roles.length = 0;
  ctx.hidden.length = 0;
  ctx.o.valueAxis = true;
  layoutVertical(ctx, model(ctx));
}

function layoutVertical(ctx: Ctx, md: Model): 'labels-too-wide' | void {
  const { spec, style, frame, o, fs, lh, pad, m } = ctx;
  const { nCat, nSer, percent, clustered, raw, val, cats, sers, shown, segLabel, showTotals, brk } = md;
  let { dmin, dmax, ticks, axisFmt } = md;

  // Reserve space: legend and arrows on top, totals above columns, axis on the left, series labels right, categories below.
  const legendH = legend(ctx);
  const arrowsH = md.arrows.length * arrowRowHeight(ctx);
  const totalsH = showTotals ? lh + pad * 0.3 : 0;
  const barLabelH = clustered ? lh + pad * 0.3 : 0;
  // A broken axis is not linear: such a chart keeps its own scale (no same-scale group).
  const axisW = brk ? (ticks ? Math.max(...ticks.map((t) => m.width(axisFmt(t), fs))) + pad * 1.5 : 0) : axisWidth(ctx, dmin, ticks, axisFmt, percent);
  // Decorations beside the plot: value line labels and the 100% indicator in a column on their side;
  // a level arrow beside the last column needs room for its bubble before the series names.
  const indicator = indicatorSide(ctx, md);
  const sideW = (side: 'left' | 'right') => Math.max(0,
    ...md.valueLines.filter((v) => v.side === side).map((v) => valueLineLabelW(ctx, v.label) + pad * 0.3),
    indicator === side ? indicatorSize(ctx, false) : 0);
  const leftW = sideW('left'), decoRightW = sideW('right');
  const levels = levelArrows(ctx, md);
  const levelW = Math.max(0, ...levels.filter((l) => l.later === nCat - 1 && (l.a.position !== 'between' || l.earlier === l.later))
    .map((l) => arrowBubble(ctx, l.label).w + pad * 1.6));

  let rightW = 0, slot = 0, colW = 0, catH = 0;
  for (let i = 0; i < 3; i++) {
    slot = (frame.w - leftW - axisW - rightW) / slotUnits(nCat, o.categoryGaps);
    colW = slot / (1 + o.gap);
    catH = categoryHeight(ctx, slot);
    const names = clustered ? 0 : seriesLabelWidth(ctx, 0);
    rightW = Math.max(0, levelW + names + decoRightW - (slot - colW) / 2);
  }
  const px = frame.x + leftW + axisW;
  const pw = frame.w - leftW - axisW - rightW;
  // Room for the top tick label, and for a decoration label centred on the top of the scale.
  const py = frame.y + legendH + arrowsH + Math.max(totalsH + barLabelH, ticks || indicator || md.valueLines.length ? lh / 2 : 0);
  const ph = frame.y + frame.h - catH - (clustered && dmin < 0 ? barLabelH : 0) - py;
  if (!brk) {
    ({ dmin, dmax, ticks } = sharedScale(ctx, dmin, dmax, ph, ticks, [px, py + ph, px, py]));
    if (ticks) axisFmt = axisFormatter(ctx, ticks, percent);
  }
  if (ph < lh * 2) ctx.hidden.push('chart frame too small');
  const dist = axisDistance(dmin, dmax, brk, ph);
  const yOf = (v: number) => py + ph - dist(v);
  // Slots of the width settled above (the series column may still have moved by a hair).
  const cx = slotLayout(nCat, o.categoryGaps, px, slot * slotUnits(nCat, o.categoryGaps)).centre;

  if (ticks) valueAxis(ctx, brk ? spacedTicks(ticks, yOf, lh * 1.05) : ticks, yOf, px, axisW, axisFmt, frame.x + leftW);

  // Totals wider than a slot would collide with their neighbours: all go, not every other one.
  const totalText = md.totalText;
  const totalsFit = !showTotals || cats.every((c) => m.width(totalText(c), fs, true) + 2 <= slot);
  if (!totalsFit) ctx.hidden.push('column totals wider than their slots: hidden');

  // The compressed range of a break, in slide coordinates (labels keep out of it).
  const wc = brk ? (yOf(brk.to) + yOf(brk.from)) / 2 : 0;
  const band = brk ? { top: wc - brk.gap / 2, bottom: wc + brk.gap / 2 } : null;
  const spans = (a: number, b: number) => !!brk && a <= brk.from && b >= brk.to;
  const broken: { s: number; c: number; x: number; w: number }[] = [];
  // Bar labels (clustered) share one size: the largest, down to the theme minimum, at which every label
  // fits over its bar. If none does, they are all left out: labels on some bars only would look like gaps.
  // A label may be a little wider than its bar only where the neighbouring bar does not reach it.
  const barW = colW / nSer;
  const barLabels = clustered ? cats.flatMap((c) => sers.flatMap((s) => { const v = raw(s, c); return v == null ? [] : [{ s, c, v, text: ctx.fmt(v) }]; })) : [];
  const clearOf = (v: number, n: number | null) => n == null || (v >= 0 ? n < 0 || yOf(n) >= yOf(v) - pad * 0.15 : n >= 0 || yOf(n) <= yOf(v) + pad * 0.15);
  const barLabelsFit = (size: number) => barLabels.every(({ s, c, v, text: t }) => {
    const tw = m.width(t, size) + 2;
    return tw <= barW || (tw <= barW + pad * 0.4 && clearOf(v, s > 0 ? raw(s - 1, c) : null) && clearOf(v, s < nSer - 1 ? raw(s + 1, c) : null));
  });
  let barFs = fs;
  while (barFs - 1 >= ctx.minFs && !barLabelsFit(barFs)) barFs--;
  const showBarLabels = barLabelsFit(barFs);
  if (clustered && !showBarLabels) {
    if (!o.valueAxis) return 'labels-too-wide';
    ctx.hidden.push(`value labels wider than their bars (${barLabels.length} hidden): the value axis shows the values`);
  } else if (clustered && barFs < fs) ctx.hidden.push(`value labels shrunk to ${barFs} pt to fit their bars`);
  const barLh = m.lineHeight(barFs);

  const tops: number[] = [];
  const lastSegs: { s: number; y: number }[] = [];
  const pending: PendingLabel[] = [];

  for (const c of cats) {
    if (clustered) {
      let top = yOf(0);
      for (const s of sers) {
        const v = raw(s, c);
        if (v == null) continue; // missing data is an empty slot, not a zero bar (B5)
        const x = cx(c) - colW / 2 + s * barW;
        const y1 = yOf(Math.max(0, v)), y2 = yOf(Math.min(0, v));
        ctx.add(`bar:s${s}:c${c}`, { kind: 'rect', x, y: y1, w: barW, h: Math.max(0.01, y2 - y1), fill: ctx.color(s), stroke: style.background, strokeWidth: 0.5 });
        if (spans(Math.min(0, v), Math.max(0, v))) broken.push({ s, c, x, w: barW });
        if (showBarLabels) {
          const label = ctx.fmt(v);
          const tw = m.width(label, barFs);
          const ly = v >= 0 ? y1 - barLh - pad * 0.15 : y2 + pad * 0.15;
          text(ctx, `label:value:s${s}:c${c}`, { x: x + barW / 2 - tw / 2 - 1, y: ly, w: tw + 2, h: barLh, text: label, role: 'value', size: barFs });
        }
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
      // A broken segment is labelled in the larger of its two visible parts.
      let [ly1, ly2] = [y1, y2];
      if (band && spans(a, b)) {
        broken.push({ s, c, x, w: colW });
        [ly1, ly2] = band.top - y1 >= y2 - band.bottom ? [y1, band.top] : [band.bottom, y2];
      }
      const room = ly2 - ly1;
      const label = segLabel(s, c);
      // Measured as drawn: an override may change its text, size or direction.
      const L = labelSpec(ctx, `label:value:s${s}:c${c}`, label);
      if (L.w + 2 <= colW && L.h - lh * 0.1 <= room) {
        text(ctx, `label:value:s${s}:c${c}`, { x: x + colW / 2 - L.w / 2 - 1, y: ly1 + room / 2 - L.h / 2, w: L.w + 2, h: L.h, text: label, color: textOn(fill), role: 'value' });
      } else {
        // Beside the column (never into the series names right of the last one), or above / below the stack.
        const sides: Side[] = c === nCat - 1 && o.seriesLabels === 'right' ? ['left'] : ['right', 'left'];
        pending.push({ s, c, text: label, seg, sides: [...sides, v > 0 ? 'above' : 'below'], boxCover: colBars, reason: `${spec.series[s].name} / ${spec.categories[c]}: segment too small` });
      }
      if (c === nCat - 1) lastSegs.push({ s, y: y1 + h / 2 });
    }
    let top = yOf(percent && !md.pdata && !md.hasBase ? 100 : up);
    if (showTotals && totalsFit) {
      const label = totalText(c);
      const tw = m.width(label, fs, true);
      top -= lh + pad * 0.2;
      text(ctx, `label:total:c${c}`, { x: cx(c) - tw / 2 - 1, y: top, w: tw + 2, h: lh, text: label, bold: true, role: 'total' });
    }
    tops.push(top);
  }

  // Baseline, break symbols on the broken bars and the axis, category axis breaks.
  const yBase = yOf(0);
  ctx.add('axis', { kind: 'line', x1: px, y1: yBase, x2: px + pw, y2: yBase, color: style.axis, width: 0.75 });
  if (brk) {
    for (const b of broken) breakSymbol(ctx, `break:s${b.s}:c${b.c}`, false, b.x - 2, b.x + b.w + 2, wc, brk);
    if (ticks) breakSymbol(ctx, 'break:axis', false, px - 4, px + 4, wc, brk);
  }
  for (const i of new Set(o.categoryBreaks ?? [])) {
    if (i >= 0 && i < nCat - 1) categoryBreakSymbol(ctx, `break:cat:${i}`, false, (cx(i) + cx(i + 1)) / 2, yBase);
  }
  categoryLabels(ctx, cats.map(cx), slot, Math.max(yBase, yOf(dmin)) + (clustered && dmin < 0 ? barLabelH : 0));

  const lastRight = cx(nCat - 1) + colW / 2;
  if (o.seriesLabels === 'right' && !clustered) {
    seriesLabelsRight(ctx, lastSegs.reverse(), lastRight + levelW, py - totalsH, py + ph);
  }

  // Growth and difference arrows, one row each, above the totals.
  let row = 0;
  for (const [kind, ref] of md.arrows) {
    const label = arrowLabel(ctx, md, kind, ref!);
    if (!label) continue;
    const { from: a, to: b } = ref!;
    const yLine = frame.y + legendH + row * arrowRowHeight(ctx) + arrowRowHeight(ctx) / 2;
    bracketArrow(ctx, kind, cx(a), tops[a] - pad * 0.3, cx(b), tops[b] - pad * 0.3, yLine, label, headsOf(kind, ref!));
    row++;
  }

  // Segment labels that did not fit go where they can, once everything else is drawn.
  placePending(ctx, pending);

  // Decorations last, around the labels: value lines, level arrows, the 100% indicator.
  const area = { x: frame.x, y: frame.y + legendH, w: frame.w, h: frame.h - legendH };
  const rightEnd = frame.x + frame.w - decoRightW;
  for (const vl of md.valueLines) {
    drawValueLine(ctx, {
      horizontal: false, w: yOf(vl.v), side: vl.side, label: vl.label, role: `valueline:${vl.k}`, bounds: area,
      u1: vl.side === 'left' ? frame.x + leftW : px, u2: vl.side === 'left' ? px + pw : rightEnd, reason: `value line ${vl.k + 1}`,
    });
  }
  const colL = (c: number) => cx(c) - colW / 2, colR = (c: number) => cx(c) + colW / 2;
  for (const l of levels) {
    const between = l.a.position === 'between' && l.earlier !== l.later;
    const base = between ? (colR(l.later - 1) + colL(l.later)) / 2 : colR(l.later) + pad * 0.8;
    const guides = [l.a.from, l.a.to].flatMap((e) => ('category' in e
      ? [{ u: between && e.category === l.later ? colL(e.category) : colR(e.category), w: yOf(endValue(md, e)) }]
      : []));
    drawLevelArrow(ctx, {
      horizontal: false, us: between ? [base] : [base, base + pad * 0.5, base + pad], guides, mode: l.mode,
      wa: yOf(endValue(md, l.a.from)), wb: yOf(endValue(md, l.a.to)),
      label: l.label, role: `arrow:level:${l.k}`, bounds: area, reason: `level arrow ${l.k + 1}`,
    });
  }
  if (indicator) {
    const names = o.seriesLabels === 'right' && !clustered;
    const edge = indicator === 'left' ? px - axisW : names ? rightEnd - pad * 0.4 : lastRight + levelW;
    drawIndicator(ctx, { horizontal: false, w: yOf(100), edge, side: indicator, bounds: area });
  }
}

/** Horizontal bars: categories down the left, values to the right, totals and arrows beyond the bar ends. */
function layoutBars(ctx: Ctx, md: Model) {
  const { spec, style, frame, o, fs, lh, pad, m } = ctx;
  const { nCat, nSer, percent, clustered, raw, val, cats, sers, shown, segLabel, showTotals, brk } = md;
  let { dmin, dmax, ticks, axisFmt } = md;
  if (o.seriesLabels === 'right') ctx.o.seriesLabels = 'legend';

  const legendH = legend(ctx);
  const axisH = ticks ? lh + pad * 0.6 : 0;
  // Category column: as wide as the longest label, at most 30% of the frame; labels wrap inside it.
  const catW = Math.min(frame.w * 0.3, Math.max(...spec.categories.map((c) => m.width(c, fs))) + 2);
  const fit = fitCategories(ctx, catW);
  reportFit(ctx, fit);
  const totalLabels = cats.map(md.totalText);
  const totalsW = showTotals ? Math.max(...totalLabels.map((t) => m.width(t, fs, true))) + pad * 0.8 : 0;
  const valueLabels = clustered ? cats.flatMap((c) => sers.map((s) => raw(s, c)).filter((v): v is number => v != null).map((v) => ({ v, w: m.width(ctx.fmt(v), fs) }))) : [];
  const barLabelR = Math.max(0, ...valueLabels.filter((l) => l.v >= 0).map((l) => l.w + pad * 0.6));
  const barLabelL = Math.max(0, ...valueLabels.filter((l) => l.v < 0).map((l) => l.w + pad * 0.6));
  const arrowLabels = md.arrows.map(([kind, ref]) => ({ kind, ref: ref!, label: arrowLabel(ctx, md, kind, ref!) }));
  const arrowW = arrowLabels.map((a) => (a.label ? arrowBubble(ctx, a.label).w + pad : 0));
  const arrowsW = arrowW.reduce((a, w) => a + w, 0) + (arrowW.some(Boolean) ? pad : 0);
  // Rotated column decorations: value line labels below the plot ("right") or above it ("left"), the
  // 100% indicator below the last bar, a level arrow below the later bar.
  const indicator = indicatorSide(ctx, md);
  const vlH = (side: 'left' | 'right') => (md.valueLines.some((v) => v.side === side) ? lh + fs * 0.6 + pad * 0.6 : 0);
  const topH = Math.max(vlH('left'), indicator === 'left' ? indicatorSize(ctx, true) : 0);
  const bottomH = vlH('right') + (indicator === 'right' && ticks ? indicatorSize(ctx, true) : 0);
  const levels = levelArrows(ctx, md);
  const levelH = Math.max(0, ...levels.filter((l) => l.later === nCat - 1 && (l.a.position !== 'between' || l.earlier === l.later))
    .map((l) => arrowBubble(ctx, l.label).h + pad * 1.6), indicator === 'right' && !ticks ? indicatorSize(ctx, true) : 0);

  const px = frame.x + catW + pad + barLabelL;
  const pw = frame.x + frame.w - totalsW - barLabelR - arrowsW - px;
  const py = frame.y + legendH + topH;
  const yA = frame.y + frame.h - axisH - bottomH;
  let slot = 0, barT = 0, ph = 0;
  for (let i = 0; i < 2; i++) {
    ph = yA - py - Math.max(0, levelH - (slot - barT) / 2);
    slot = ph / slotUnits(nCat, o.categoryGaps);
    barT = slot / (1 + o.gap);
  }
  if (pw < lh * 2 || ph < lh * nCat * 0.5) ctx.hidden.push('chart frame too small');
  if (!brk) {
    ({ dmin, dmax, ticks } = sharedScale(ctx, dmin, dmax, pw, ticks, [px, yA, px + pw, yA]));
    if (ticks) axisFmt = axisFormatter(ctx, ticks, percent);
  }
  const dist = axisDistance(dmin, dmax, brk, pw);
  const xOf = (v: number) => px + dist(v);
  const cy = slotLayout(nCat, o.categoryGaps, py, ph).centre;

  // Value axis along the bottom.
  if (ticks) {
    const tickW = Math.max(...ticks.map((t) => m.width(axisFmt(t), fs))) + pad;
    const shownTicks = brk ? spacedTicks(ticks, xOf, tickW) : ticks;
    ctx.add('axis', { kind: 'line', x1: xOf(ticks[0]), y1: yA, x2: xOf(ticks[ticks.length - 1]), y2: yA, color: style.axis, width: 0.75 });
    shownTicks.forEach((t, k) => {
      const label = axisFmt(t);
      const w = m.width(label, fs) + 2;
      ctx.add(`axis:tick:${k}`, { kind: 'line', x1: xOf(t), y1: yA, x2: xOf(t), y2: yA + 3, color: style.axis, width: 0.75 });
      text(ctx, `axis:tick:${k}`, { x: xOf(t) - w / 2, y: yA + pad * 0.5, w, h: lh, text: label, color: style.muted, role: 'axis' });
    });
  }

  const wc = brk ? (xOf(brk.to) + xOf(brk.from)) / 2 : 0;
  const band = brk ? { left: wc - brk.gap / 2, right: wc + brk.gap / 2 } : null;
  const spans = (a: number, b: number) => !!brk && a <= brk.from && b >= brk.to;
  const broken: { s: number; c: number; y: number; h: number }[] = [];
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
        if (spans(Math.min(0, v), Math.max(0, v))) broken.push({ s, c, y, h: barH });
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
      let [lx1, lx2] = [x1, x2];
      if (band && spans(a, b)) {
        broken.push({ s, c, y, h: barT });
        [lx1, lx2] = band.left - x1 >= x2 - band.right ? [x1, band.left] : [band.right, x2];
      }
      const room = lx2 - lx1;
      const label = segLabel(s, c);
      const L = labelSpec(ctx, `label:value:s${s}:c${c}`, label);
      if (L.w + 2 <= room && L.h - lh * 0.1 <= barT) {
        text(ctx, `label:value:s${s}:c${c}`, { x: lx1 + room / 2 - L.w / 2 - 1, y: cy(c) - L.h / 2, w: L.w + 2, h: L.h, text: label, color: textOn(fill), role: 'value' });
      } else {
        // In the gap above or below the bar, or beyond its end.
        pending.push({ s, c, text: label, seg, sides: ['above', 'below', v > 0 ? 'right' : 'left'], boxCover: colBars, reason: `${spec.series[s].name} / ${spec.categories[c]}: segment too small` });
      }
    }
    let end = xOf(percent && !md.pdata && !md.hasBase ? 100 : right);
    if (showTotals) {
      const tw = m.width(totalLabels[c], fs, true) + 2;
      text(ctx, `label:total:c${c}`, { x: end + pad * 0.4, y: cy(c) - lh / 2, w: tw, h: lh, text: totalLabels[c], bold: true, align: 'left', role: 'total' });
      end += pad * 0.4 + tw;
    }
    ends.push(end);
  }

  // Baseline at zero, break symbols, then category labels right-aligned in their column.
  const xBase = xOf(0);
  ctx.add('axis', { kind: 'line', x1: xBase, y1: py, x2: xBase, y2: py + ph, color: style.axis, width: 0.75 });
  if (brk) {
    for (const b of broken) breakSymbol(ctx, `break:s${b.s}:c${b.c}`, true, b.y - 2, b.y + b.h + 2, wc, brk);
    if (ticks) breakSymbol(ctx, 'break:axis', true, yA - 4, yA + 4, wc, brk);
  }
  for (const i of new Set(o.categoryBreaks ?? [])) {
    if (i >= 0 && i < nCat - 1) categoryBreakSymbol(ctx, `break:cat:${i}`, true, (cy(i) + cy(i + 1)) / 2, xBase);
  }
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
    bracketArrowH(ctx, kind, ends[ref.from] + pad * 0.3, cy(ref.from), ends[ref.to] + pad * 0.3, cy(ref.to), xLine, label, headsOf(kind, ref));
    xCol += arrowW[i];
  });

  placePending(ctx, pending);

  // Decorations last, around the labels.
  const area = { x: frame.x, y: frame.y + legendH, w: frame.w, h: frame.h - legendH };
  for (const vl of md.valueLines) {
    drawValueLine(ctx, {
      horizontal: true, w: xOf(vl.v), side: vl.side, label: vl.label, role: `valueline:${vl.k}`, bounds: area,
      u1: py, u2: vl.side === 'right' ? yA + axisH : py + ph, reason: `value line ${vl.k + 1}`,
    });
  }
  const barT0 = (c: number) => cy(c) - barT / 2, barB = (c: number) => cy(c) + barT / 2;
  for (const l of levels) {
    const between = l.a.position === 'between' && l.earlier !== l.later;
    const base = between ? (barB(l.later - 1) + barT0(l.later)) / 2 : barB(l.later) + pad * 0.8;
    const guides = [l.a.from, l.a.to].flatMap((e) => ('category' in e
      ? [{ u: between && e.category === l.later ? barT0(e.category) : barB(e.category), w: xOf(endValue(md, e)) }]
      : []));
    drawLevelArrow(ctx, {
      horizontal: true, us: between ? [base] : [base, base + pad * 0.5], guides, mode: l.mode,
      wa: xOf(endValue(md, l.a.from)), wb: xOf(endValue(md, l.a.to)),
      label: l.label, role: `arrow:level:${l.k}`, bounds: area, reason: `level arrow ${l.k + 1}`,
    });
  }
  if (indicator) {
    const edge = indicator === 'left' ? barT0(0) : ticks ? yA + axisH : barB(nCat - 1);
    drawIndicator(ctx, { horizontal: true, w: xOf(100), edge, side: indicator, bounds: area });
  }
}
