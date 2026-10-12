// Chart decorations shared by the chart types (reference manual: Chart annotations, Axes): category
// gaps and category axis breaks, value axis breaks, level difference arrows, value lines and the 100%
// indicator. Geometry is given in "u/w" coordinates — u along the category axis, w along the value
// axis — so one implementation serves columns (u = x, w = y) and horizontal bars (u = y, w = x).
// Every decoration is drawn after the labels it must not touch: its lines step around text, and its
// labels go only where they collide with nothing; what finds no room is left out and reported.
import type { Ctx } from './common.ts';
import { arrowBubble, text } from './common.ts';
import { niceTicks } from './numbers.ts';
import { overlaps, segmentHitsBox, type Box } from './place.ts';
import type { AxisBreakSpec, DiffMode, PathCmd, Shape, TextShape, ValueLineSpec } from './types.ts';

// ---- Category slots with gaps ------------------------------------------------------------------

export interface Slots {
  /** Width of one category slot (column + gap). */
  slot: number;
  /** Number of slots: categories plus inserted gaps. */
  units: number;
  /** Centre of category c. */
  centre: (c: number) => number;
}

/**
 * Category slots along a length. A category gap after category i (manual: "as wide as one column
 * width plus one gap width") is one more slot; the chart keeps its size, so columns get narrower.
 */
export function slotLayout(n: number, gaps: readonly number[] | undefined, start: number, len: number): Slots {
  const g = (gaps ?? []).filter((i) => Number.isInteger(i) && i >= 0 && i < n - 1);
  const units = n + g.length;
  const slot = len / units;
  return { slot, units, centre: (c) => start + slot * (c + 0.5 + g.filter((i) => i < c).length) };
}

/** Number of slots for n categories with the chart's gaps (for fixed-point width loops). */
export const slotUnits = (n: number, gaps: readonly number[] | undefined) =>
  n + (gaps ?? []).filter((i) => Number.isInteger(i) && i >= 0 && i < n - 1).length;

// ---- Value axis breaks -------------------------------------------------------------------------

export interface Break { from: number; to: number; style: 'straight' | 'wiggle'; /** Points the compressed range takes. */ gap: number }

/**
 * The break to draw, or null with the reason recorded. Manual: a break applies to every segment that
 * shares the compressed range, so no segment may end inside it, and it must lie within the data.
 */
export function resolveBreak(ctx: Ctx, spec: AxisBreakSpec | null | undefined, lo: number, hi: number, levels: number[]): Break | null {
  if (!spec) return null;
  const { from, to } = spec;
  if (!(from > lo && to < hi)) { ctx.hidden.push(`axis break ${from}–${to} ignored: outside the data range ${lo}–${hi}`); return null; }
  const inside = levels.find((v) => v > from + 1e-9 && v < to - 1e-9);
  if (inside != null) { ctx.hidden.push(`axis break ${from}–${to} ignored: a segment ends inside it (at ${inside})`); return null; }
  return { from, to, style: spec.style ?? 'straight', gap: Math.max(ctx.pad * 2, ctx.lh * 0.8) };
}

/** Distance of v from the dmin end of an axis `len` points long, with the break's range compressed. */
export function axisDistance(dmin: number, dmax: number, brk: Break | null, len: number): (v: number) => number {
  const span = dmax - dmin || 1;
  if (!brk || len <= brk.gap * 2) return (v) => ((v - dmin) / span) * len;
  const k = (len - brk.gap) / (span - (brk.to - brk.from));
  return (v) => (v <= brk.from ? (v - dmin) * k
    : v >= brk.to ? (brk.from - dmin) * k + brk.gap + (v - brk.to) * k
    : (brk.from - dmin) * k + (brk.gap * (v - brk.from)) / (brk.to - brk.from));
}

/** Ticks of a broken axis: nice ticks below the break and nice ticks above it, by their share of the length. */
export function brokenTicks(dmin: number, dmax: number, brk: Break, count: number): number[] {
  const below = brk.from - dmin, above = dmax - brk.to;
  const n = (part: number) => Math.max(3, Math.round((count * part) / (below + above)) + 1);
  const lower = niceTicks(dmin, brk.from, n(below)).filter((t) => t <= brk.from + 1e-9);
  const upper = niceTicks(brk.to, dmax, n(above)).filter((t) => t >= brk.to - 1e-9);
  return [...lower, ...upper];
}

/** Keep tick labels at least `minGap` apart along the axis (those next to a break would collide). */
export function spacedTicks(ticks: number[], pos: (v: number) => number, minGap: number): number[] {
  const out: number[] = [];
  for (const t of ticks) if (!out.length || Math.abs(pos(t) - pos(out[out.length - 1])) >= minGap) out.push(t);
  return out;
}

const pt = (horizontal: boolean, u: number, w: number) => (horizontal ? { x: w, y: u } : { x: u, y: w });

/**
 * Break symbol across [u1, u2] at value-axis position wc (the centre of the compressed range): a band
 * in the background colour that cuts the bar or axis, edged by two straight slanted lines or two wiggles.
 */
export function breakSymbol(ctx: Ctx, role: string, horizontal: boolean, u1: number, u2: number, wc: number, brk: Break) {
  const half = brk.gap * 0.18, slant = brk.gap * 0.2;
  const edge = (side: 1 | -1): { u: number; w: number }[] => {
    if (brk.style === 'straight') return [{ u: u1, w: wc + side * half + slant }, { u: u2, w: wc + side * half - slant }];
    const n = Math.max(4, Math.round((u2 - u1) / 2.5));
    const wave = Math.max(4, brk.gap * 0.9);
    return Array.from({ length: n + 1 }, (_, i) => {
      const u = u1 + ((u2 - u1) * i) / n;
      return { u, w: wc + side * half + Math.sin(((u - u1) / wave) * 2 * Math.PI) * brk.gap * 0.12 };
    });
  };
  const top = edge(-1), bottom = edge(1);
  const cmds = (ps: { u: number; w: number }[], close: boolean): PathCmd[] => [
    ...ps.map((p, i) => ({ op: i ? 'L' as const : 'M' as const, ...pt(horizontal, p.u, p.w) })),
    ...(close ? [{ op: 'Z' as const }] : []),
  ];
  ctx.add(role,
    { kind: 'path', commands: cmds([...top, ...[...bottom].reverse()], true), fill: ctx.style.background },
    { kind: 'path', commands: cmds(top, false), stroke: ctx.style.axis, strokeWidth: 0.75 },
    { kind: 'path', commands: cmds(bottom, false), stroke: ctx.style.axis, strokeWidth: 0.75 });
}

/** Category axis break (manual: "a break in the continuity of the category axis"): a symbol on the baseline. */
export function categoryBreakSymbol(ctx: Ctx, role: string, horizontal: boolean, u: number, wBase: number) {
  const h = ctx.lh * 0.4, d = Math.max(2, ctx.pad * 0.5), s = h * 0.45;
  const ends = (du: number) => [pt(horizontal, u + du - s, wBase + h), pt(horizontal, u + du + s, wBase - h)];
  const [a1, a2] = ends(-d / 2), [b1, b2] = ends(d / 2);
  ctx.add(role,
    { kind: 'path', commands: [{ op: 'M', ...a1 }, { op: 'L', ...a2 }, { op: 'L', ...b2 }, { op: 'L', ...b1 }, { op: 'Z' }], fill: ctx.style.background },
    { kind: 'path', commands: [{ op: 'M', ...a1 }, { op: 'L', ...a2 }], stroke: ctx.style.axis, strokeWidth: 0.75 },
    { kind: 'path', commands: [{ op: 'M', ...b1 }, { op: 'L', ...b2 }], stroke: ctx.style.axis, strokeWidth: 0.75 });
}

// ---- Lines that step around text ---------------------------------------------------------------

const textBoxes = (ctx: Ctx, except?: string): Box[] =>
  ctx.shapes.flatMap((s, i) => (s.kind === 'text' && ctx.roles[i] !== except ? [s as TextShape] : []));

/**
 * An axis-parallel line drawn in pieces that leave a small gap around every text box it would cross,
 * so a decoration line never runs through a label.
 */
export function lineAround(ctx: Ctx, role: string, x1: number, y1: number, x2: number, y2: number, look: { color: string; width: number; dash?: boolean }, avoid: Box[] = textBoxes(ctx, role)) {
  const horiz = Math.abs(y2 - y1) < 1e-6;
  const [a, b] = horiz ? [Math.min(x1, x2), Math.max(x1, x2)] : [Math.min(y1, y2), Math.max(y1, y2)];
  const at = horiz ? y1 : x1;
  const g = 1.5;
  const blocked = avoid
    .filter((t) => (horiz ? t.y - 0.5 < at && at < t.y + t.h + 0.5 : t.x - 0.5 < at && at < t.x + t.w + 0.5))
    .map((t) => (horiz ? [t.x - g, t.x + t.w + g] : [t.y - g, t.y + t.h + g]))
    .sort((p, q) => p[0] - q[0]);
  let cur = a;
  const pieces: [number, number][] = [];
  for (const [s, e] of blocked) {
    if (e <= cur) continue;
    if (s > cur) pieces.push([cur, Math.min(s, b)]);
    cur = Math.max(cur, e);
    if (cur >= b) break;
  }
  if (cur < b) pieces.push([cur, b]);
  for (const [s, e] of pieces) {
    if (e - s < 1) continue;
    ctx.add(role, horiz
      ? { kind: 'line', x1: s, y1: at, x2: e, y2: at, ...look }
      : { kind: 'line', x1: at, y1: s, x2: at, y2: e, ...look });
  }
}

// ---- Difference labels -------------------------------------------------------------------------

/**
 * Difference arrow label in one of the manual's three modes: relative (→, change from `a` to `b` in
 * percent), absolute (↔, b − a) and reverse (←, change from `b` to `a` in percent). Null when a relative
 * difference has no positive base.
 */
export function diffLabel(ctx: Ctx, a: number, b: number, mode: DiffMode, percentPoints = false, decimals = 0): string | null {
  if (mode === 'absolute') return percentPoints ? ctx.pct(b - a, decimals, true) : ctx.fmt(b - a, { signed: true });
  const [base, other] = mode === 'relative' ? [a, b] : [b, a];
  if (!(base > 0)) return null;
  return ctx.pct(((other - base) / base) * 100, 0, true);
}

/** Arrow shafts from wa to wb at u: heads per mode (relative at b, reverse at a, absolute at both). */
function shafts(horizontal: boolean, u: number, wa: number, wb: number, mode: DiffMode): [number, number, number, number, boolean][] {
  const A = pt(horizontal, u, wa), B = pt(horizontal, u, wb), M = pt(horizontal, u, (wa + wb) / 2);
  if (mode === 'relative') return [[A.x, A.y, B.x, B.y, true]];
  if (mode === 'reverse') return [[B.x, B.y, A.x, A.y, true]];
  return [[M.x, M.y, A.x, A.y, true], [M.x, M.y, B.x, B.y, true]];
}

// ---- Level difference arrows -------------------------------------------------------------------

export interface LevelArrowGeom {
  horizontal: boolean;
  /** Candidate category-axis positions for the shaft, in order of preference. */
  us: number[];
  /** Value-axis positions of the two levels. */
  wa: number;
  wb: number;
  /** Dotted guide lines from a column edge / point (u) to the shaft at the level (w). */
  guides: { u: number; w: number }[];
  mode: DiffMode;
  label: string;
  role: string;
  /** Where the label may go. */
  bounds: Box;
  reason: string;
}

/** Marks, texts and lines a decoration label must stay off. */
function obstacles(ctx: Ctx): { boxes: Box[]; lines: [number, number, number, number][] } {
  const boxes: Box[] = [], lines: [number, number, number, number][] = [];
  ctx.shapes.forEach((s, i) => {
    const r = ctx.roles[i];
    if (s.kind === 'text') boxes.push(s);
    else if (s.kind === 'rect' && /^(bar|marker):/.test(r)) boxes.push(s);
    else if (s.kind === 'line' && /^(line:|arrow:|valueline:)/.test(r)) lines.push([s.x1, s.y1, s.x2, s.y2]);
  });
  return { boxes, lines };
}

const inside = (b: Box, f: Box) => b.x >= f.x - 0.5 && b.y >= f.y - 0.5 && b.x + b.w <= f.x + f.w + 0.5 && b.y + b.h <= f.y + f.h + 0.5;
const shrinkBox = (b: Box, d: number): Box => ({ x: b.x + d, y: b.y + d, w: Math.max(0, b.w - 2 * d), h: Math.max(0, b.h - 2 * d) });

/**
 * Manual (Level difference arrows): the arrow shows the difference between two levels; "if the arrow
 * is so small that the bubble would cover it up, the bubble is automatically placed next to the arrow".
 * The first shaft position whose shaft crosses no label and whose bubble finds a free spot wins.
 */
export function drawLevelArrow(ctx: Ctx, g: LevelArrowGeom): boolean {
  const { horizontal: hz, wa, wb } = g;
  if (Math.abs(wb - wa) < 2) { ctx.hidden.push(`${g.reason}: both ends at the same level`); return false; }
  const { w: bw, h: bh } = arrowBubble(ctx, g.label);
  const along = hz ? bw : bh; // bubble extent along the shaft
  const across = hz ? bh : bw;
  const { boxes, lines } = obstacles(ctx);
  const texts = boxes.filter((b) => (b as Shape).kind === 'text');
  const lo = Math.min(wa, wb), hi = Math.max(wa, wb), mid = (wa + wb) / 2;
  for (const u of g.us) {
    // The shaft may not cross a label.
    const A = pt(hz, u, lo), B = pt(hz, u, hi);
    if (texts.some((t) => segmentHitsBox(A.x, A.y, B.x, B.y, shrinkBox(t, 1)))) continue;
    const gap = ctx.pad * 0.4;
    const spots: Box[] = [];
    const box = (uc: number, wc = mid) => { const c = pt(hz, uc, wc); return { x: c.x - bw / 2, y: c.y - bh / 2, w: bw, h: bh }; };
    // On the shaft (centred, else slid along it), else beside it.
    const slack = (hi - lo - along - ctx.pad) / 2;
    if (slack >= 0) for (const f of [0, -0.5, 0.5, -1, 1]) spots.push(box(u, mid + f * slack));
    spots.push(box(u + gap + across / 2), box(u - gap - across / 2));
    const spot = spots.find((b) => inside(b, g.bounds)
      && !boxes.some((o) => overlaps(b, o, 0.5))
      && !lines.some(([x1, y1, x2, y2]) => segmentHitsBox(x1, y1, x2, y2, shrinkBox(b, 0.5))));
    if (!spot) continue;
    const c = ctx.style.axis;
    for (const gd of g.guides) {
      const P = pt(hz, gd.u, gd.w), Q = pt(hz, u, gd.w);
      if (Math.abs(gd.u - u) > 1) lineAround(ctx, g.role, P.x, P.y, Q.x, Q.y, { color: c, width: 0.5, dash: true }, texts);
    }
    for (const [x1, y1, x2, y2, head] of shafts(hz, u, wa, wb, g.mode)) {
      ctx.add(g.role, { kind: 'line', x1, y1, x2, y2, color: c, width: 1, ...(head ? { arrowEnd: true } : {}) });
    }
    text(ctx, g.role, { ...spot, text: g.label, bold: true, fill: ctx.style.background, stroke: c, rounded: true, role: 'decoration' });
    return true;
  }
  ctx.hidden.push(`${g.reason}: no room for the arrow and its label`);
  return false;
}

// ---- Value lines -------------------------------------------------------------------------------

/** Text of a value line label: the custom label with {value} filled in, else "Ø value" for a mean. */
export function valueLineText(ctx: Ctx, vl: ValueLineSpec, v: number): string {
  const formatted = ctx.fmt(v);
  if (vl.label != null) return vl.label.includes('{value}') ? vl.label.replace(/\{value\}/g, formatted) : vl.label;
  return vl.value === 'mean' ? `Ø ${formatted}` : formatted;
}

/** Width a value line label column needs (marker + text). */
export const valueLineLabelW = (ctx: Ctx, label: string) => ctx.fs * 0.6 + ctx.pad * 0.6 + ctx.m.width(label, ctx.fs) + 4;

/**
 * Value line (manual: "a line parallel to the X-axis to visualize a certain value ... the label placed
 * to the left or to the right of the chart"): dashed across [u1, u2], label beyond the end on its side
 * with a small triangle pointing at the line. Returns false (and reports) when the label has no room.
 */
export function drawValueLine(ctx: Ctx, o: { horizontal: boolean; w: number; u1: number; u2: number; side: 'left' | 'right'; label: string; role: string; bounds: Box; reason: string }): boolean {
  const { horizontal: hz, w, side } = o;
  const tri = ctx.fs * 0.6, gap = ctx.pad * 0.6;
  const tw = ctx.m.width(o.label, ctx.fs) + 2;
  const lh = ctx.lh;
  // Columns: the label sits beyond the left / right end; bars: below (right) / above (left) the plot.
  const end = side === 'right' ? o.u2 : o.u1;
  const dir = side === 'right' ? 1 : -1;
  let labelBox: Box, triCmds: PathCmd[];
  if (!hz) {
    const tx = dir > 0 ? end + 1 + tri + gap * 0.6 : end - 1 - tri - gap * 0.6 - tw;
    labelBox = { x: tx, y: w - lh / 2, w: tw, h: lh };
    const tip = end + dir * 1, base = end + dir * (1 + tri);
    triCmds = [{ op: 'M', x: tip, y: w }, { op: 'L', x: base, y: w - tri / 2 }, { op: 'L', x: base, y: w + tri / 2 }, { op: 'Z' }];
  } else {
    const ty = dir > 0 ? end + tri + gap * 0.3 : end - tri - gap * 0.3 - lh;
    labelBox = { x: w - tw / 2, y: ty, w: tw, h: lh };
    const tip = end + dir * 1, base = end + dir * (1 + tri);
    triCmds = [{ op: 'M', x: w, y: tip }, { op: 'L', x: w - tri / 2, y: base }, { op: 'L', x: w + tri / 2, y: base }, { op: 'Z' }];
  }
  const { boxes } = obstacles(ctx);
  if (!inside(labelBox, o.bounds) || boxes.some((b) => overlaps(labelBox, b, 0.5))) {
    ctx.hidden.push(`${o.reason}: no room for its label`);
    return false;
  }
  const P = pt(hz, o.u1, w), Q = pt(hz, o.u2, w);
  lineAround(ctx, o.role, P.x, P.y, Q.x, Q.y, { color: ctx.style.text, width: 0.75, dash: true });
  ctx.add(o.role, { kind: 'path', commands: triCmds, fill: ctx.style.text });
  text(ctx, o.role, { ...labelBox, text: o.label, align: hz ? 'center' : dir > 0 ? 'left' : 'right', role: 'decoration' });
  return true;
}

// ---- 100% indicator ----------------------------------------------------------------------------

/** Width (columns) or height (bars) the 100% indicator takes beside the plot. */
export function indicatorSize(ctx: Ctx, horizontal: boolean): number {
  const t = ctx.m.width(ctx.pct(100, 0), ctx.fs) + 2;
  return horizontal ? ctx.pad * 2 + ctx.lh + ctx.pad * 0.4 : ctx.pad * 2.4 + t + ctx.pad * 0.4;
}

/**
 * The 100% indicator (manual: "In 100% stacked charts ... you can display a 100% indicator ... to the
 * left or to the right of your chart"): a short arrow pointing at the 100% level from beside the chart,
 * with "100%". `edge` is the category-axis position of the chart's outer edge on that side.
 */
export function drawIndicator(ctx: Ctx, o: { horizontal: boolean; w: number; edge: number; side: 'left' | 'right'; bounds: Box }): boolean {
  const { horizontal: hz, w, edge, side } = o;
  const dir = side === 'right' ? 1 : -1;
  const label = ctx.pct(100, 0);
  const tw = ctx.m.width(label, ctx.fs) + 2, lh = ctx.lh;
  const a0 = edge + dir * ctx.pad * 0.4, a1 = a0 + dir * ctx.pad * 2;
  const tb: Box = hz
    ? { x: w - tw / 2, y: dir > 0 ? a1 : a1 - lh, w: tw, h: lh }
    : { x: dir > 0 ? a1 + 1 : a1 - 1 - tw, y: w - lh / 2, w: tw, h: lh };
  const { boxes } = obstacles(ctx);
  if (!inside(tb, o.bounds) || boxes.some((b) => overlaps(tb, b, 0.5))) {
    ctx.hidden.push('100% indicator: no room beside the chart');
    return false;
  }
  const P = pt(hz, a1, w), Q = pt(hz, a0, w);
  ctx.add('indicator', { kind: 'line', x1: P.x, y1: P.y, x2: Q.x, y2: Q.y, color: ctx.style.text, width: 0.75, arrowEnd: true });
  text(ctx, 'indicator', { ...tb, text: label, align: hz ? 'center' : dir > 0 ? 'left' : 'right', role: 'decoration' });
  return true;
}
