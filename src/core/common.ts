import type { ChartOptions, ChartSpec, Frame, Shape, Style, TextShape } from './types.ts';
import type { Measurer } from './text.ts';
import { wrap } from './text.ts';
import { formatNumber, inferDecimals } from './numbers.ts';
import { spread } from './place.ts';

export interface Ctx {
  spec: ChartSpec;
  style: Style;
  frame: Frame;
  m: Measurer;
  o: ChartOptions & { gap: number; seriesLabels: 'right' | 'legend' | 'none' };
  fs: number;
  lh: number;
  pad: number;
  shapes: Shape[];
  hidden: string[];
  decimals: number;
  color(i: number): string;
  fmt(v: number, unit?: string, signed?: boolean): string;
}

export function makeCtx(spec: ChartSpec, style: Style, frame: Frame, m: Measurer): Ctx {
  const defaults = {
    gap: 0.5,
    seriesLabels: spec.type === 'clustered' ? 'legend' : 'right',
    totals: spec.type === 'stacked',
  } as const;
  const o = { ...defaults, ...spec.options };
  const all = spec.series.flatMap((s) => s.values).filter((v): v is number => v != null);
  const decimals = o.decimals ?? inferDecimals(all);
  const fs = style.fontSize;
  return {
    spec, style, frame, m, o, fs,
    lh: m.lineHeight(fs),
    pad: fs * 0.4,
    shapes: [],
    hidden: [],
    decimals,
    color: (i) => (spec.series[i]?.color ?? style.palette[i % style.palette.length]).replace('#', ''),
    fmt: (v, unit = o.unit ?? '', signed = false) => formatNumber(v, decimals, unit, signed),
  };
}

export function text(ctx: Ctx, t: Omit<TextShape, 'kind' | 'size' | 'color' | 'align' | 'valign'> & Partial<TextShape>) {
  ctx.shapes.push({ kind: 'text', size: ctx.fs, color: ctx.style.text, align: 'center', valign: 'middle', ...t });
}

/** Height needed for category labels wrapped into slots of the given width. */
export function categoryHeight(ctx: Ctx, slot: number): number {
  const lines = Math.max(1, ...ctx.spec.categories.map((c) => wrap(c, slot - 2, ctx.fs, ctx.m).length));
  return lines * ctx.lh + ctx.pad * 0.6;
}

export function categoryLabels(ctx: Ctx, centres: number[], slot: number, top: number) {
  ctx.spec.categories.forEach((c, i) => {
    const lines = wrap(c, slot - 2, ctx.fs, ctx.m);
    text(ctx, {
      x: centres[i] - slot / 2, y: top + ctx.pad * 0.4, w: slot, h: lines.length * ctx.lh,
      text: lines.join('\n'), valign: 'top', role: 'category',
    });
  });
}

/** Width a right-hand series label column needs, after borrowing the free space beside the last column. */
export function seriesLabelWidth(ctx: Ctx, borrow: number): number {
  if (ctx.o.seriesLabels !== 'right') return 0;
  const w = Math.max(0, ...ctx.spec.series.map((s) => ctx.m.width(s.name, ctx.fs)));
  return Math.max(0, w + ctx.pad * 1.5 - borrow);
}

/** Series names to the right of the last column, aligned to their segments and spread apart. */
export function seriesLabelsRight(ctx: Ctx, items: { name: string; y: number }[], x: number, top: number, bottom: number) {
  const heights = items.map(() => ctx.lh);
  const ys = spread(items.map((i) => i.y), heights, top, bottom, 0);
  items.forEach((it, k) => {
    const w = ctx.m.width(it.name, ctx.fs) + 2;
    const lx = x + ctx.pad;
    if (Math.abs(ys[k] - it.y) > 1) {
      ctx.shapes.push({ kind: 'line', x1: x, y1: it.y, x2: lx - 1, y2: ys[k], color: ctx.style.muted, width: 0.5 });
    }
    text(ctx, { x: lx, y: ys[k] - ctx.lh / 2, w, h: ctx.lh, text: it.name, align: 'left', role: 'series' });
  });
}

/** One-row legend: colour swatch + name, laid left to right from the frame's top-left. Returns its height. */
export function legend(ctx: Ctx, swatch: 'box' | 'line' = 'box'): number {
  if (ctx.o.seriesLabels !== 'legend') return 0;
  const { frame, fs, lh, pad } = ctx;
  let x = frame.x;
  const y = frame.y;
  const sw = fs * 0.75;
  ctx.spec.series.forEach((s, i) => {
    if (swatch === 'box') {
      ctx.shapes.push({ kind: 'rect', x, y: y + (lh - sw) / 2, w: sw, h: sw, fill: ctx.color(i) });
    } else {
      ctx.shapes.push({ kind: 'line', x1: x, y1: y + lh / 2, x2: x + sw * 1.6, y2: y + lh / 2, color: ctx.color(i), width: 2 });
    }
    x += (swatch === 'box' ? sw : sw * 1.6) + pad * 0.6;
    const w = ctx.m.width(s.name, fs) + 2;
    text(ctx, { x, y, w, h: lh, text: s.name, align: 'left', role: 'series' });
    x += w + pad * 2;
  });
  return lh + pad;
}

/** Bracket arrow from point A up to a shared line, across, and down to B, with a label bubble in the middle. */
export function bracketArrow(ctx: Ctx, xa: number, ya: number, xb: number, yb: number, yLine: number, label: string) {
  const c = ctx.style.axis;
  ctx.shapes.push(
    { kind: 'line', x1: xa, y1: ya, x2: xa, y2: yLine, color: c, width: 0.75 },
    { kind: 'line', x1: xa, y1: yLine, x2: xb, y2: yLine, color: c, width: 0.75 },
    { kind: 'line', x1: xb, y1: yLine, x2: xb, y2: yb, color: c, width: 0.75, arrowEnd: true },
  );
  const w = ctx.m.width(label, ctx.fs, true) + ctx.pad * 2;
  const h = ctx.lh + ctx.pad * 0.3;
  text(ctx, {
    x: (xa + xb) / 2 - w / 2, y: yLine - h / 2, w, h, text: label, bold: true,
    fill: ctx.style.background, stroke: c, rounded: true, role: 'decoration',
  });
}

/** Height one row of bracket arrows needs. */
export const arrowRowHeight = (ctx: Ctx) => ctx.lh + ctx.pad * 1.6;

export function cagrLabel(ctx: Ctx, a: number, b: number, years: number): string | null {
  if (a <= 0 || b <= 0 || years <= 0) return null;
  const r = ((b / a) ** (1 / years) - 1) * 100;
  return formatNumber(r, 1, '%', true);
}

/** Left value axis: vertical line, tick marks and right-aligned tick labels. */
export function valueAxis(ctx: Ctx, ticks: number[], yOf: (v: number) => number, px: number, axisW: number, fmt: (v: number) => string) {
  const { style, frame, lh, pad } = ctx;
  ctx.shapes.push({ kind: 'line', x1: px, y1: yOf(ticks[0]), x2: px, y2: yOf(ticks[ticks.length - 1]), color: style.axis, width: 0.75 });
  for (const t of ticks) {
    ctx.shapes.push({ kind: 'line', x1: px - 3, y1: yOf(t), x2: px, y2: yOf(t), color: style.axis, width: 0.75 });
    text(ctx, { x: frame.x, y: yOf(t) - lh / 2, w: axisW - pad, h: lh, text: fmt(t), align: 'right', color: style.muted, role: 'axis' });
  }
}
