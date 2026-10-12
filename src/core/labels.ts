import type { RectShape, Shape } from './types.ts';
import type { Ctx } from './common.ts';
import { text } from './common.ts';
import { textOn } from './numbers.ts';
import { overlaps, segmentHitsBox, type Box } from './place.ts';
import { pathBounds } from './path.ts';

/**
 * Label placement as constrained candidate selection (the reference engine's model): every label
 * has a fixed set of allowed positions with a cost; the placer picks the cheapest valid one, where
 * valid means inside the bounds, clear of every mark and every other label, and with a leader that
 * crosses nothing. Labels without any valid candidate are left out and reported by the caller.
 *
 * The segment-label rules come from the manual (Chart labels → Automatic label placement):
 * - a label larger than its segment sits on it with "a colored rectangle underneath" (boxed);
 * - "if there is not enough space inside a segment, place the label outside the segment and add a
 *   line that points to the related segment" (outside, with a leader);
 * - "if two labels are too close together, offset one to the left and the other to the right".
 * The style file's `avoidSegmentLabelBoxing` default picks boxed or outside "depending on the label
 * size": boxed while the label overflows its segment by at most half its own height and stays within
 * the column's width (a box poking out of the column's sides would cover its neighbours), else outside.
 */

export type Side = 'left' | 'right' | 'above' | 'below';

export interface Leader { x1: number; y1: number; x2: number; y2: number }

export interface Candidate {
  box: Box;
  cost: number;
  side: Side | 'boxed';
  /** Line from the mark to the label, for labels placed away from their mark. */
  leader?: Leader;
  /** Obstacles this candidate may cover (a boxed label covers its own column). */
  cover?: Set<Box>;
}

export interface LabelRequest {
  /** The mark the label belongs to: it may touch it, and its leader starts on it. */
  anchor: Box;
  candidates: Candidate[];
}

const shrink = (b: Box, d: number): Box => ({ x: b.x + d, y: b.y + d, w: Math.max(0, b.w - 2 * d), h: Math.max(0, b.h - 2 * d) });
const inside = (b: Box, f: Box) => b.x >= f.x - 0.5 && b.y >= f.y - 0.5 && b.x + b.w <= f.x + f.w + 0.5 && b.y + b.h <= f.y + f.h + 0.5;
const hits = (l: Leader, b: Box) => segmentHitsBox(l.x1, l.y1, l.x2, l.y2, shrink(b, 0.5));

/** Proper crossing of two segments (shared end points do not count). */
function crosses(a: Leader, b: Leader): boolean {
  const d = (px: number, py: number, qx: number, qy: number, rx: number, ry: number) => (qx - px) * (ry - py) - (qy - py) * (rx - px);
  const d1 = d(a.x1, a.y1, a.x2, a.y2, b.x1, b.y1), d2 = d(a.x1, a.y1, a.x2, a.y2, b.x2, b.y2);
  const d3 = d(b.x1, b.y1, b.x2, b.y2, a.x1, a.y1), d4 = d(b.x1, b.y1, b.x2, b.y2, a.x2, a.y2);
  return d1 * d2 < 0 && d3 * d4 < 0;
}

/** Air added to a label's text width (points); a boxed label gives it up to stay within its column. */
export const LABEL_PAD = 2;

/** Labels keep this much air between each other (points). */
const GAP = 1;

/**
 * Pick one candidate per label (index into its `candidates`, or null when none is valid). Greedy in
 * order of fewest valid positions, then local improvement: every label moves to its cheapest valid
 * candidate given the others, and a label left out may evict a single blocker to that blocker's
 * next valid position. Candidates are tried cheapest first; ties keep their listed order.
 */
export function placeLabels(reqs: LabelRequest[], obstacles: Box[], bounds: Box): (number | null)[] {
  const sorted = reqs.map((r) => r.candidates.map((_, k) => k).sort((a, b) => r.candidates[a].cost - r.candidates[b].cost));
  const statik = reqs.map((r) => r.candidates.map((c) => {
    if (!inside(c.box, bounds)) return false;
    for (const o of obstacles) {
      if (o === r.anchor || c.cover?.has(o)) continue;
      if (overlaps(c.box, o, 0.5)) return false;
      if (c.leader && hits(c.leader, o)) return false;
    }
    return true;
  }));
  const choice: (number | null)[] = reqs.map(() => null);
  const conflict = (i: number, k: number, j: number, q: number) => {
    const a = reqs[i].candidates[k], b = reqs[j].candidates[q];
    if (overlaps(a.box, b.box, -GAP)) return true;
    if (a.leader && (hits(a.leader, b.box) || (b.leader && crosses(a.leader, b.leader)))) return true;
    return !!b.leader && hits(b.leader, a.box);
  };
  const blockers = (i: number, k: number) => {
    const out: number[] = [];
    choice.forEach((q, j) => { if (j !== i && q != null && conflict(i, k, j, q)) out.push(j); });
    return out;
  };
  const best = (i: number) => sorted[i].find((k) => statik[i][k] && !blockers(i, k).length) ?? null;

  const order = reqs.map((_, i) => i).sort((a, b) => statik[a].filter(Boolean).length - statik[b].filter(Boolean).length || a - b);
  for (const i of order) choice[i] = best(i);

  for (let round = 0; round < 4; round++) {
    let changed = false;
    for (const i of order) {
      const cur = choice[i];
      choice[i] = null;
      choice[i] = best(i);
      if (choice[i] !== cur) changed = true;
      if (choice[i] != null) continue;
      // Left out: try moving exactly one placed label out of the way.
      for (const k of sorted[i]) {
        if (!statik[i][k]) continue;
        const bl = blockers(i, k);
        if (bl.length !== 1) continue;
        const j = bl[0], old = choice[j];
        choice[j] = null;
        choice[i] = k;
        const alt = best(j);
        if (alt != null) { choice[j] = alt; changed = true; break; }
        choice[i] = null;
        choice[j] = old;
      }
    }
    if (!changed) break;
  }
  return choice;
}

export interface AroundOptions {
  /** Allowed outside positions, in order of preference. */
  sides: Side[];
  /** Distance between the mark and an outside label (the leader's length when level). */
  lead: number;
  /** How many label sizes an outside label may move along its side to make room (stagger). */
  stagger?: number;
  /** Marks a boxed label may cover; omit to disallow boxing. */
  boxCover?: Box[];
}

/**
 * Candidates for a label of size w×h belonging to `a`: boxed on the mark (when its text is no wider
 * than the mark — the box is trimmed to the mark's width — and overflows its height by at most half),
 * then beside it on each allowed side, level with the mark's centre or staggered one, two, … label
 * sizes along that side, each with a leader from the mark's edge.
 */
export function segmentCandidates(a: Box, w: number, h: number, o: AroundOptions): Candidate[] {
  const out: Candidate[] = [];
  const mx = a.x + a.w / 2, my = a.y + a.h / 2;
  if (o.boxCover && w - LABEL_PAD <= a.w && a.h >= h / 2) {
    const bw = Math.min(w, a.w);
    out.push({ box: { x: mx - bw / 2, y: my - h / 2, w: bw, h }, cost: 1, side: 'boxed', cover: new Set([a, ...o.boxCover]) });
  }
  const n = o.stagger ?? 3;
  o.sides.forEach((side, rank) => {
    for (let s = 0; s <= 2 * n; s++) {
      const k = s % 2 ? -(s + 1) / 2 : s / 2; // 0, -1, 1, -2, 2, …: up / left before down / right
      const cost = 10 + rank + Math.abs(k) * 4;
      if (side === 'left' || side === 'right') {
        const cy = my + k * (h + GAP);
        const x = side === 'right' ? a.x + a.w + o.lead : a.x - o.lead - w;
        const ex = side === 'right' ? a.x + a.w : a.x;
        out.push(lead({ box: { x, y: cy - h / 2, w, h }, cost, side, leader: { x1: ex, y1: my, x2: side === 'right' ? x - 0.5 : x + w + 0.5, y2: cy } }, h));
      } else {
        const cx = mx + k * (w + GAP * 2);
        const y = side === 'above' ? a.y - o.lead - h : a.y + a.h + o.lead;
        const ey = side === 'above' ? a.y : a.y + a.h;
        // The leader leaves the mark's edge as close below / above the label as the mark allows.
        const ex = a.w > 2 ? Math.min(Math.max(cx, a.x + 1), a.x + a.w - 1) : mx;
        out.push(lead({ box: { x: cx - w / 2, y, w, h }, cost, side, leader: { x1: ex, y1: ey, x2: cx, y2: side === 'above' ? y + h + 0.5 : y - 0.5 } }, h));
      }
    }
  });
  return out;
}

/**
 * A leader should run away from its mark, not along it: the part of its run that goes sideways
 * (across, beyond its run away from the mark) costs as much per label height as one stagger step.
 * A long, nearly flat leader thus loses to a closer spot or a box.
 */
function lead(c: Candidate, h: number): Candidate {
  const l = c.leader!, dx = Math.abs(l.x2 - l.x1), dy = Math.abs(l.y2 - l.y1);
  const [away, across] = c.side === 'left' || c.side === 'right' ? [dx, dy] : [dy, dx];
  return { ...c, cost: c.cost + (4 * Math.max(0, across - away)) / h };
}

/** Bounding box of any scene shape; rects and texts are their own box, so identity is kept. */
export const shapeBox = (s: Shape): Box =>
  s.kind === 'line'
    ? { x: Math.min(s.x1, s.x2), y: Math.min(s.y1, s.y2), w: Math.abs(s.x2 - s.x1), h: Math.abs(s.y2 - s.y1) }
    : s.kind === 'path' ? pathBounds(s) : s;

/** A segment label that did not fit inside its segment, waiting for a position. */
export interface PendingLabel {
  s: number;
  c: number;
  text: string;
  /** The segment: the label's anchor. */
  seg: RectShape;
  sides: Side[];
  /** Marks a boxed label may cover (its own column); omit to disallow boxing. */
  boxCover?: Box[];
  /** Reason recorded when no position is valid. */
  reason: string;
  bold?: boolean;
  /** Text role; default `label:value:s<i>:c<j>`. */
  role?: string;
}

/**
 * Place pending segment labels against everything already drawn (marks, labels, lines) within
 * `bounds`, then draw them: boxed labels on their segment's colour, outside labels with a thin
 * leader (`leader:s<i>:c<j>`). Labels without a valid position are reported in `ctx.hidden`.
 */
export function placePending(ctx: Ctx, items: PendingLabel[], bounds: Box = ctx.frame) {
  if (!items.length) return;
  const obstacles = ctx.shapes.map(shapeBox);
  // A boxed label never covers a segment whose own label points at it from outside.
  const anchors = new Set<Box>(items.map((it) => it.seg));
  const reqs: LabelRequest[] = items.map((it) => ({
    anchor: it.seg,
    candidates: segmentCandidates(it.seg, ctx.m.width(it.text, ctx.fs, it.bold) + LABEL_PAD, ctx.lh, {
      sides: it.sides, lead: ctx.pad * 0.8, boxCover: it.boxCover?.filter((b) => !anchors.has(b)),
    }),
  }));
  const choice = placeLabels(reqs, obstacles, bounds);
  items.forEach((it, i) => {
    const k = choice[i];
    if (k == null) { ctx.hidden.push(it.reason); return; }
    const c = reqs[i].candidates[k];
    if (c.leader) ctx.add(`leader:s${it.s}:c${it.c}`, { kind: 'line', ...c.leader, color: ctx.style.muted, width: 0.5 });
    const boxed = c.side === 'boxed';
    text(ctx, it.role ?? `label:value:s${it.s}:c${it.c}`, {
      ...c.box, text: it.text, role: it.bold ? 'total' : 'value', ...(it.bold ? { bold: true } : {}),
      align: c.side === 'right' ? 'left' : c.side === 'left' ? 'right' : 'center',
      ...(boxed ? { color: textOn(it.seg.fill), fill: it.seg.fill } : {}),
    });
  });
}
