import type { ChartOptions, ChartSpec, Frame, Shape, Style, TextShape } from './types.ts';
import type { Measurer } from './text.ts';
import { fits, wrap } from './text.ts';
import { formatValue, inferDecimals, pickScale, type NumberStyle } from './numbers.ts';
import { spread } from './place.ts';
import { parseCategoryDate } from './dates.ts';
import { TooSmall } from './minsize.ts';
import { optionDefaults } from './spec.ts';

export interface Ctx {
  spec: ChartSpec;
  style: Style;
  frame: Frame;
  m: Measurer;
  o: ChartOptions & { gap: number; seriesLabels: 'right' | 'legend' | 'none' };
  fs: number;
  /** Smallest font size a label role may shrink to. */
  minFs: number;
  lh: number;
  pad: number;
  shapes: Shape[];
  /** Parallel to `shapes`: the role of each shape (see Scene.roles). */
  roles: string[];
  hidden: string[];
  /** Decimals for value labels, in scaled units. */
  decimals: number;
  /** Shared magnitude: values are divided by `div` and labelled with `suffix` (k / m / bn). */
  scale: { div: number; suffix: string };
  color(i: number): string;
  /** Append shapes that share one role. */
  add(role: string, ...shapes: Shape[]): void;
  /** Format a data value (unscaled) with the chart's locale, unit, magnitude and format. */
  fmt(v: number, opts?: Partial<NumberStyle> & { scaled?: boolean }): string;
  /** Format a percentage (already ×100) with the chart's locale. */
  pct(v: number, decimals: number, signed?: boolean): string;
}

/** Numeric values of a series cell (null and markers such as "e" are not numbers). */
export const num = (v: unknown): number | null => (typeof v === 'number' && Number.isFinite(v) ? v : null);

export function makeCtx(spec: ChartSpec, style: Style, frame: Frame, m: Measurer): Ctx {
  const horizontal = spec.options?.orientation === 'horizontal';
  // seriesLabels is only a starting point here: chart types refine it from the data when it is unset.
  const defaults = {
    gap: 0.5,
    seriesLabels: spec.type === 'clustered' || horizontal ? 'legend' : 'right',
    ...optionDefaults(spec.type),
  } as const;
  const o = { ...defaults, ...spec.options };
  const all = spec.series.flatMap((s) => s.values.map(num)).filter((v): v is number => v != null);
  const scale = pickScale(all, o.scale);
  const decimals = o.decimals ?? inferDecimals(all.map((v) => v / scale.div));
  const fs = style.fontSize;
  const shapes: Shape[] = [];
  const roles: string[] = [];
  return {
    spec, style, frame, m, o, fs,
    minFs: Math.min(fs, 8),
    lh: m.lineHeight(fs),
    pad: fs * 0.4,
    shapes,
    roles,
    hidden: [],
    decimals,
    scale,
    color: (i) => (spec.series[i]?.color ?? style.palette[i % style.palette.length]).replace('#', ''),
    add(role, ...list) { for (const s of list) { shapes.push(s); roles.push(role); } },
    fmt: (v, opts = {}) => formatValue(opts.scaled ? v : v / scale.div, {
      locale: o.locale, decimals, format: o.format, unit: o.unit, magnitude: scale.suffix, ...opts,
    }),
    pct: (v, d, signed = false) => formatValue(v, { locale: o.locale, decimals: d, unit: '%', signed }),
  };
}

type TextInput = Omit<TextShape, 'kind' | 'size' | 'color' | 'align' | 'valign'> & Partial<TextShape>;

export function text(ctx: Ctx, role: string, t: TextInput) {
  const size = t.size ?? ctx.fs;
  ctx.add(role, {
    kind: 'text', size, color: ctx.style.text, align: 'center', valign: 'middle',
    lineHeight: ctx.m.lineHeight(size), ...t,
  });
}

// ---- Category labels: wrap, shrink the whole row one step at a time, then hide (B8) ------------

export interface CategoryFit {
  lines: string[][]; size: number; lh: number; height: number; hidden: Set<number>;
  /** Thinned time axis: every `span`-th label is shown, in `span` slots' width (1 = every label). */
  span: number;
}

/**
 * Wrap every category into its slot; shrink all of them together if any does not fit; hide what still
 * does not. With `thin`, a time axis (every category reads as a date) that still does not fit shows
 * every 2nd, 3rd, ... label instead, the way a consultant thins a crowded year axis.
 */
export function fitCategories(
  ctx: Ctx, slot: number | number[], maxLines = 3,
  /** Allow thinning a time axis, from `minSpan` slots per label, at most at `maxSize` pt. */
  thin: false | { minSpan?: number; maxSize?: number } = false,
): CategoryFit {
  const width = (i: number) => (Array.isArray(slot) ? slot[i] : slot) - 2;
  const cats = ctx.spec.categories;
  const shrinkFit = (w: (i: number) => number, only: (i: number) => boolean, lineCap = maxLines, from = ctx.fs) => {
    let size = from;
    let lines: string[][] = [];
    for (;;) {
      lines = cats.map((c, i) => (only(i) ? wrap(c, w(i), size, ctx.m, lineCap) : ['']));
      const ok = lines.every((l, i) => !only(i) || fits(l, w(i), size, ctx.m));
      if (ok || size - 1 < ctx.minFs) return { lines, size, ok };
      size -= 1;
    }
  };
  let { lines, size } = shrinkFit(width, () => true);
  const hidden = new Set<number>();
  lines.forEach((l, i) => { if (!fits(l, width(i), size, ctx.m)) hidden.add(i); });
  let span = 1;
  if (thin && hidden.size && !Array.isArray(slot) && cats.length > 2 && cats.every((c) => parseCategoryDate(c))) {
    // Thinned labels sit on one line, so a sparser axis is never taller.
    for (let k = thin.minSpan ?? 2; k <= Math.ceil(cats.length / 2); k++) {
      const t = shrinkFit(() => slot * k - 2, (i) => i % k === 0, 1, Math.min(ctx.fs, thin.maxSize ?? ctx.fs));
      if (!t.ok) continue;
      ({ lines, size } = t);
      span = k;
      hidden.clear();
      cats.forEach((_, i) => { if (i % k) hidden.add(i); });
      break;
    }
  }
  const lh = ctx.m.lineHeight(size);
  const n = Math.max(1, ...lines.filter((_, i) => !hidden.has(i)).map((l) => l.length));
  return { lines, size, lh, height: n * lh + ctx.pad * 0.6, hidden, span };
}

/** Height needed for category labels wrapped into slots of the given width. */
export function categoryHeight(ctx: Ctx, slot: number): number {
  return fitCategories(ctx, slot, 3, {}).height;
}

/** Widest frame a chart is expected to get (the default chart area of a 16:9 slide). */
const MAX_CHART_W = 864;

/**
 * Category names (not a thinned time axis) that cannot all be shown in uniform slots: if a chart of at
 * most the default slide width would show them, throw TooSmall with that width (the placeholder asks
 * for it) rather than draw an axis with gaps. Beyond that, labels stay hidden and reported.
 */
export function askForRoom(ctx: Ctx, fit: CategoryFit, slot: number, maxLines = 3) {
  if (fit.span !== 1 || !fit.hidden.size) return;
  const cats = ctx.spec.categories;
  // Narrowest slot in which every name wraps into `maxLines` lines at the minimum size.
  const ok = (w: number) => cats.every((c) => fits(wrap(c, w - 2, ctx.minFs, ctx.m, maxLines), w - 2, ctx.minFs, ctx.m));
  let lo = slot, hi = Math.max(slot, ...cats.map((c) => ctx.m.width(c, ctx.minFs) + 4));
  for (let k = 0; k < 12; k++) { const mid = (lo + hi) / 2; if (ok(mid)) hi = mid; else lo = mid; }
  const need = ctx.frame.w + cats.length * (hi - slot) + 1;
  if (need <= MAX_CHART_W) throw new TooSmall({ w: need }, 'category labels need wider slots');
}

export function categoryLabels(ctx: Ctx, centres: number[], slot: number, top: number) {
  const { frame } = ctx;
  // A thinned label is as wide as its text (up to `span` slots), kept inside the frame; if keeping the
  // end labels inside pushes them into their neighbours, thin further.
  const boxes = (fit: CategoryFit) => ctx.spec.categories.map((_, i) => {
    if (fit.hidden.has(i)) return null;
    const w = fit.span > 1 ? Math.max(slot, ...fit.lines[i].map((l) => ctx.m.width(l, fit.size) + 2)) : slot;
    const x = centres[i] - w / 2;
    return { x: x < frame.x - 0.01 ? frame.x : x + w > frame.x + frame.w + 0.01 ? frame.x + frame.w - w : x, w };
  });
  const clash = (bs: ({ x: number; w: number } | null)[]) => {
    const shown = bs.filter((b): b is { x: number; w: number } => !!b);
    return shown.some((b, k) => k > 0 && shown[k - 1].x + shown[k - 1].w > b.x + 0.5);
  };
  let fit = fitCategories(ctx, slot, 3, {});
  let bs = boxes(fit);
  while (fit.span > 1 && clash(bs)) {
    const next = fitCategories(ctx, slot, 3, { minSpan: fit.span + 1, maxSize: fit.size });
    if (next.span <= fit.span) break;
    fit = next;
    bs = boxes(fit);
  }
  askForRoom(ctx, fit, slot);
  reportFit(ctx, fit);
  ctx.spec.categories.forEach((_, i) => {
    const b = bs[i];
    if (!b) return;
    const lines = fit.lines[i];
    text(ctx, `label:category:c${i}`, {
      x: b.x, y: top + ctx.pad * 0.4, w: b.w, h: lines.length * fit.lh,
      text: lines.join('\n'), valign: 'top', role: 'category', size: fit.size,
    });
  });
}

export function reportFit(ctx: Ctx, fit: CategoryFit) {
  if (fit.size < ctx.fs) ctx.hidden.push(`category labels shrunk to ${fit.size} pt to fit`);
  if (fit.span > 1) {
    ctx.hidden.push(`category labels thinned to every ${fit.span}${fit.span === 2 ? 'nd' : fit.span === 3 ? 'rd' : 'th'} (${fit.hidden.size} hidden): slots too narrow`);
    return;
  }
  for (const i of fit.hidden) ctx.hidden.push(`${ctx.spec.categories[i]}: category label too wide for its slot`);
}

// ---- Series identification --------------------------------------------------------------------

/** Width a right-hand series label column needs, after borrowing the free space beside the last column. */
export function seriesLabelWidth(ctx: Ctx, borrow: number): number {
  if (ctx.o.seriesLabels !== 'right') return 0;
  const w = Math.max(0, ...ctx.spec.series.map((s) => ctx.m.width(s.name, ctx.fs)));
  return Math.max(0, w + ctx.pad * 1.5 - borrow);
}

export interface SeriesItem { s: number; y: number; name?: string }

/** Series names to the right of the last column, aligned to their segments and spread apart. */
export function seriesLabelsRight(ctx: Ctx, items: SeriesItem[], x: number, top: number, bottom: number) {
  const heights = items.map(() => ctx.lh);
  const ys = spread(items.map((i) => i.y), heights, top, bottom, 0);
  items.forEach((it, k) => {
    const name = it.name ?? ctx.spec.series[it.s].name;
    const w = ctx.m.width(name, ctx.fs) + 2;
    const lx = x + ctx.pad;
    const role = `label:series:s${it.s}`;
    if (Math.abs(ys[k] - it.y) > 1) {
      ctx.add(role, { kind: 'line', x1: x, y1: it.y, x2: lx - 1, y2: ys[k], color: ctx.style.muted, width: 0.5 });
    }
    text(ctx, role, { x: lx, y: ys[k] - ctx.lh / 2, w, h: ctx.lh, text: name, align: 'left', role: 'series' });
  });
}

/** Legend swatch: a square for areas, a short line for lines, a circle for point markers. */
export type Swatch = 'box' | 'line' | 'circle';

/**
 * Legend: colour swatch + name, laid left to right from the frame's top-left, wrapping onto as many
 * rows as the frame width needs (B9). Returns its height.
 */
export function legend(
  ctx: Ctx, swatchKind: Swatch | ((i: number) => Swatch) = 'box', colorOf: (s: number) => string = ctx.color,
  /** Entries other than the series (pie and doughnut list their categories): names and role key. */
  entries: { names: string[]; key: 's' | 'c' } = { names: ctx.spec.series.map((s) => s.name), key: 's' },
  /** Leave entries out (Mekko XY: the width series has no swatch). */
  include: (i: number) => boolean = () => true,
): number {
  // Names beside the plot would take more than a third of a narrow frame (a chart resized small):
  // a legend on top keeps every name and leaves the width to the data.
  if (ctx.o.seriesLabels === 'right' && seriesLabelWidth(ctx, 0) > ctx.frame.w / 3) ctx.o.seriesLabels = 'legend';
  if (ctx.o.seriesLabels !== 'legend') return 0;
  const { frame, fs, lh, pad } = ctx;
  const sw = fs * 0.75;
  let x = frame.x;
  let y = frame.y;
  entries.names.forEach((name, i) => {
    if (!include(i)) return;
    const swatch = typeof swatchKind === 'function' ? swatchKind(i) : swatchKind;
    const swW = swatch === 'line' ? sw * 1.6 : sw;
    const w = ctx.m.width(name, fs) + 2;
    const itemW = swW + pad * 0.6 + w;
    if (x > frame.x && x + itemW > frame.x + frame.w) { x = frame.x; y += lh; }
    const role = `legend:${entries.key}${i}`;
    if (swatch !== 'line') {
      ctx.add(role, { kind: 'rect', x, y: y + (lh - sw) / 2, w: sw, h: sw, fill: colorOf(i), ...(swatch === 'circle' ? { geom: 'ellipse' as const } : {}) });
    } else {
      ctx.add(role, { kind: 'line', x1: x, y1: y + lh / 2, x2: x + swW, y2: y + lh / 2, color: colorOf(i), width: 2 });
    }
    text(ctx, role, { x: x + swW + pad * 0.6, y, w, h: lh, text: name, align: 'left', role: 'series' });
    x += itemW + pad * 2;
  });
  return y - frame.y + lh + pad;
}

// ---- Decorations ------------------------------------------------------------------------------

export type ArrowKind = 'cagr' | 'diff';
/** Which ends of a bracket arrow carry a head: the "to" end, the "from" end, or both. */
export type Heads = 'end' | 'start' | 'both';

/** Bracket arrow from point A up to a shared line, across, and down to B, with a label bubble in the middle. */
export function bracketArrow(ctx: Ctx, kind: ArrowKind, xa: number, ya: number, xb: number, yb: number, yLine: number, label: string, heads: Heads = 'end') {
  const c = ctx.style.axis;
  const role = `arrow:${kind}`;
  ctx.add(role,
    heads === 'end' ? { kind: 'line', x1: xa, y1: ya, x2: xa, y2: yLine, color: c, width: 0.75 } : { kind: 'line', x1: xa, y1: yLine, x2: xa, y2: ya, color: c, width: 0.75, arrowEnd: true },
    { kind: 'line', x1: xa, y1: yLine, x2: xb, y2: yLine, color: c, width: 0.75 },
    { kind: 'line', x1: xb, y1: yLine, x2: xb, y2: yb, color: c, width: 0.75, ...(heads === 'start' ? {} : { arrowEnd: true }) },
  );
  const w = ctx.m.width(label, ctx.fs, true) + ctx.pad * 2;
  const h = ctx.lh + ctx.pad * 0.3;
  text(ctx, role, {
    x: (xa + xb) / 2 - w / 2, y: yLine - h / 2, w, h, text: label, bold: true,
    fill: ctx.style.background, stroke: c, rounded: true, role: 'decoration',
  });
}

/** Sideways bracket for horizontal bars: from A right to a shared vertical line, down, and back left to B. */
export function bracketArrowH(ctx: Ctx, kind: ArrowKind, xa: number, ya: number, xb: number, yb: number, xLine: number, label: string, heads: Heads = 'end') {
  const c = ctx.style.axis;
  const role = `arrow:${kind}`;
  ctx.add(role,
    heads === 'end' ? { kind: 'line', x1: xa, y1: ya, x2: xLine, y2: ya, color: c, width: 0.75 } : { kind: 'line', x1: xLine, y1: ya, x2: xa, y2: ya, color: c, width: 0.75, arrowEnd: true },
    { kind: 'line', x1: xLine, y1: ya, x2: xLine, y2: yb, color: c, width: 0.75 },
    { kind: 'line', x1: xLine, y1: yb, x2: xb, y2: yb, color: c, width: 0.75, ...(heads === 'start' ? {} : { arrowEnd: true }) },
  );
  const { w, h } = arrowBubble(ctx, label);
  text(ctx, role, {
    x: xLine - w / 2, y: (ya + yb) / 2 - h / 2, w, h, text: label, bold: true,
    fill: ctx.style.background, stroke: c, rounded: true, role: 'decoration',
  });
}

export const arrowBubble = (ctx: Ctx, label: string) => ({ w: ctx.m.width(label, ctx.fs, true) + ctx.pad * 2, h: ctx.lh + ctx.pad * 0.3 });

/** Height one row of bracket arrows needs. */
export const arrowRowHeight = (ctx: Ctx) => ctx.lh + ctx.pad * 1.6;

/** CAGR in percent, formatted, or null when it is undefined (non-positive endpoints or span). */
export function cagrLabel(ctx: Ctx, a: number, b: number, years: number): string | null {
  if (a <= 0 || b <= 0 || years <= 0) return null;
  const r = ((b / a) ** (1 / years) - 1) * 100;
  return ctx.pct(r, 1, true);
}

/**
 * How many ticks a value axis gets (3-5): a short axis (a chart resized small) gets fewer, so tick
 * labels never crowd. Vertical axes by the frame height, horizontal ones by its width.
 */
export function valueTicks(ctx: Ctx, horizontal = ctx.o.orientation === 'horizontal', length?: number): number {
  const len = length ?? (horizontal ? ctx.frame.w : ctx.frame.h - ctx.lh * 3);
  return Math.max(3, Math.min(5, Math.floor(len / (horizontal ? ctx.fs * 5 : ctx.lh * 2.5))));
}

/** Axis tick labels: unit, magnitude and locale of the chart, decimals inferred from the ticks. */
export function axisFormatter(ctx: Ctx, ticks: number[] | null, percent = false): (v: number) => string {
  const div = percent ? 1 : ctx.scale.div;
  const d = ticks ? inferDecimals(ticks.map((t) => t / div)) : 0;
  return (v) => (percent ? ctx.pct(v, d) : formatValue(v / div, {
    locale: ctx.o.locale, decimals: d, unit: ctx.o.unit, magnitude: ctx.scale.suffix,
  }));
}

/** Left value axis: vertical line, tick marks and right-aligned tick labels. */
export function valueAxis(ctx: Ctx, ticks: number[], yOf: (v: number) => number, px: number, axisW: number, fmt: (v: number) => string, left = ctx.frame.x) {
  const { style, lh, pad } = ctx;
  ctx.add('axis', { kind: 'line', x1: px, y1: yOf(ticks[0]), x2: px, y2: yOf(ticks[ticks.length - 1]), color: style.axis, width: 0.75 });
  ticks.forEach((t, k) => {
    ctx.add(`axis:tick:${k}`, { kind: 'line', x1: px - 3, y1: yOf(t), x2: px, y2: yOf(t), color: style.axis, width: 0.75 });
    text(ctx, `axis:tick:${k}`, { x: left, y: yOf(t) - lh / 2, w: axisW - pad, h: lh, text: fmt(t), align: 'right', color: style.muted, role: 'axis' });
  });
}
