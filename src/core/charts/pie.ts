import type { Ctx } from '../common.ts';
import { legend, num, text } from '../common.ts';
import { roundToTotal, textOn } from '../numbers.ts';
import { overlaps, segmentHitsBox, type Box } from '../place.ts';
import type { PathCmd } from '../types.ts';

/** Doughnut hole as a share of the outer radius. */
const HOLE = 0.55;

type Pt = { x: number; y: number };

/**
 * Pie and doughnut charts: one series of values across the categories (further series are ignored).
 *
 * - Slices run clockwise from 12 o'clock in category order, each a `path` with an arc; colours
 *   follow the palette by category (the last slice never repeats the first slice's colour).
 * - Inside labels carry the slice's share of the total, rounded together (largest remainder) so the
 *   labels add up to exactly 100% — or its value, or both (`options.labelContent`). They sit as
 *   close to the rim as they fit; two that would touch are offset towards the centre.
 * - Outside labels carry the category names, left or right of the pie depending on the slice's side
 *   (default, `seriesLabels: 'right'`); or the names go in a legend (`'legend'`) or nowhere (`'none'`).
 * - An inside label larger than its slice: with names outside it sits on a rectangle in the slice's
 *   colour, still at the rim (the manual's rule) — unless the rectangle would be wider than the slice;
 *   then, and without names outside, it becomes an outside label (with the name, on a leader).
 * - Outside labels hug the rim next to their slice. Where they crowd they stack in a column, pushed
 *   away from the equator (upwards in the top half, downwards in the bottom half), each joined to
 *   its slice by a leader from the mid-angle rim point; leaders never enter the pie, cross each
 *   other or cross a label. If that needs room, the pie shrinks a little and moves away from the crowd.
 * - Doughnut with `options.totals: true` shows the bold total in the hole.
 */
export function layoutPie(ctx: Ctx) {
  const { spec, style, frame, o, fs, lh, pad, m } = ctx;
  const doughnut = spec.type === 'doughnut';
  const cats = spec.categories.map((_, c) => c);
  const raw = cats.map((c) => num(spec.series[0].values[c]));
  if (raw.some((v) => v != null && v < 0)) ctx.hidden.push('negative values ignored in pie chart');
  const vals = raw.map((v) => Math.max(0, v ?? 0));
  const total = vals.reduce((a, v) => a + v, 0);
  if (total <= 0) { ctx.hidden.push('pie chart needs positive values'); return; }
  const decimals = o.decimals ?? 0;
  const shares = roundToTotal(vals.map((v) => (v / total) * 100), decimals);
  const pctText = cats.map((c) => ctx.pct(shares[c], decimals));
  const content = o.labelContent ?? 'percent';
  const valueText = cats.map((c) =>
    (content === 'value' ? ctx.fmt(vals[c]) : content === 'both' ? `${ctx.fmt(vals[c])} (${pctText[c]})` : pctText[c]));
  const live = cats.filter((c) => vals[c] > 0);
  for (const c of cats) if (!vals[c]) ctx.hidden.push(`${spec.categories[c]}: zero value, no slice`);
  const pal = style.palette.map((p) => p.replace('#', ''));
  // Beyond the palette, colours repeat as lighter tints, so no two slices (or legend entries) match.
  const colorOf = (c: number) => tint(pal[c % pal.length], style.background, Math.min(0.8, Math.floor(c / pal.length) * 0.45));

  // Narrow frame (resized in the add-in): names beside the pie would squeeze it to a dot; a legend on
  // top keeps every name and leaves the width to the pie.
  if (o.seriesLabels === 'right') {
    const nameW = Math.max(...live.map((c) => m.width(spec.categories[c], fs))) + 2;
    if (frame.w / 2 - nameW - pad * 3 < 0.6 * (frame.h / 2 - lh)) o.seriesLabels = 'legend';
  }
  const names = o.seriesLabels === 'right';
  /** Outside label text: the name (its value inside), or the value alone when names are elsewhere. */
  const outsideText = (c: number, insideOk: boolean) =>
    (!names ? valueText[c] : insideOk ? spec.categories[c] : `${spec.categories[c]} ${valueText[c]}`);
  const legendH = legend(ctx, 'box', colorOf, { names: spec.categories, key: 'c' });
  const gap = pad * 1.5;
  const boxPad = pad * 0.6;
  const outsideW = Math.max(...live.map((c) => m.width(names ? spec.categories[c] : valueText[c], fs))) + 2;
  const top = frame.y + legendH + lh;
  const bottom = frame.y + frame.h - lh;
  const rMax = Math.max(lh, Math.min((bottom - top) / 2, frame.w / 2 - outsideW - gap * 2));
  const cx = frame.x + frame.w / 2;

  // Slices, clockwise from 12 o'clock (angles in degrees, 0 = 3 o'clock, clockwise positive).
  const start: number[] = [], sweep: number[] = [];
  let acc = 0;
  for (const c of cats) {
    start.push(-90 + (acc / total) * 360);
    sweep.push((vals[c] / total) * 360);
    acc += vals[c];
  }
  const full = live.length === 1;
  const rad = (a: number) => (a * Math.PI) / 180;
  const midOf = (c: number) => rad(start[c] + sweep[c] / 2);

  /** Label placement for a given radius and centre, without emitting anything. */
  function plan(r: number, cy: number) {
    const hole = doughnut ? r * HOLE : 0;
    const inSlice = (c: number, x: number, y: number) => {
      const d = Math.hypot(x - cx, y - cy);
      if (d > r - 2 || d < hole + 2) return false;
      if (full) return true;
      const rel = ((((Math.atan2(y - cy, x - cx) * 180) / Math.PI - start[c]) % 360) + 360) % 360;
      if (rel > sweep[c]) return false;
      const edge = (a: number) => (a >= 90 ? d : d * Math.sin(rad(a)));
      return edge(rel) >= 2 && edge(sweep[c] - rel) >= 2;
    };
    const inRing = (x: number, y: number) => { const d = Math.hypot(x - cx, y - cy); return d <= r - 1 && d >= hole + 1; };
    const corners = (b: Box) => [[b.x, b.y], [b.x + b.w, b.y], [b.x, b.y + b.h], [b.x + b.w, b.y + b.h]];
    const placed: Box[] = [];
    const inside = new Map<number, Box>();
    const boxed = new Set<number>();
    /** Walk from the rim towards the centre along the slice's mid-angle; first spot that `ok` accepts. */
    const radial = (c: number, w: number, ok: (b: Box) => boolean) => {
      const halfDiag = Math.hypot(w, lh) / 2;
      for (let d = r - halfDiag - pad * 0.5; d >= hole + halfDiag * 0.5; d -= 2) {
        const box = { x: cx + d * Math.cos(midOf(c)) - w / 2, y: cy + d * Math.sin(midOf(c)) - lh / 2, w, h: lh };
        if (ok(box) && !placed.some((p) => overlaps(p, box, 0))) return box;
      }
      return null;
    };
    for (const c of live) {
      const box = radial(c, m.width(valueText[c], fs) + 2, (b) => corners(b).every(([x, y]) => inSlice(c, x, y)));
      if (box) { placed.push(box); inside.set(c, box); }
    }
    // With names outside, a label larger than its slice stays inside on a rectangle in the slice colour,
    // as long as the rectangle is no wider than the slice where it sits; else it joins the name outside.
    const narrower = (c: number, b: Box) => {
      if (full || sweep[c] >= 180) return true;
      const t = midOf(c), d = Math.hypot(b.x + b.w / 2 - cx, b.y + b.h / 2 - cy);
      return b.w * Math.abs(Math.sin(t)) + b.h * Math.abs(Math.cos(t)) <= 2 * d * Math.sin(rad(sweep[c]) / 2);
    };
    if (names) {
      for (const c of live) {
        if (inside.has(c)) continue;
        const box = radial(c, m.width(valueText[c], fs) + 2 + boxPad * 2, (b) => narrower(c, b) && corners(b).every(([x, y]) => inRing(x, y)));
        if (box) { placed.push(box); inside.set(c, box); boxed.add(c); }
      }
    }

    type Out = { c: number; box: Box; side: number; moved: boolean; anchor: Pt; rim: Pt; leader: boolean };
    const outside: Out[] = [];
    const lost: number[] = [];
    const lo = frame.y + legendH, hi = frame.y + frame.h;
    for (const side of [1, -1]) {
      const items = live
        .filter((c) => (names || !inside.has(c)) && (Math.cos(midOf(c)) >= -1e-9 ? 1 : -1) === side)
        .map((c) => ({
          c,
          anchor: { x: cx + (r + gap) * Math.cos(midOf(c)), y: cy + (r + gap) * Math.sin(midOf(c)) },
          rim: { x: cx + r * Math.cos(midOf(c)), y: cy + r * Math.sin(midOf(c)) },
          w: m.width(outsideText(c, inside.has(c)), fs) + 2,
        }))
        .sort((a, b) => a.anchor.y - b.anchor.y);
      const ys = outward(items.map((it) => it.anchor.y), lh, lo, hi, cy);
      // Distance from the centre line to the label's near edge: next to the slice, clear of the
      // circle across the label's whole height.
      const nears = items.map((it, k) => {
        const y = ys[k] - lh / 2;
        const dy = y <= cy && y + lh >= cy ? 0 : Math.min(Math.abs(y - cy), Math.abs(y + lh - cy));
        const chord = dy < r + gap ? Math.sqrt(Math.max(0, (r + gap) ** 2 - dy * dy)) : 0;
        return Math.max(side * (it.anchor.x - cx), chord);
      });
      // A crowd stacked edge to edge forms a column, aligned on its outermost member.
      for (let a = 0; a < items.length;) {
        let b = a;
        while (b + 1 < items.length && ys[b + 1] - ys[b] <= lh + 1e-6) b++;
        if (b > a) { const n = Math.max(...nears.slice(a, b + 1)); for (let k = a; k <= b; k++) nears[k] = n; }
        a = b + 1;
      }
      items.forEach((it, k) => {
        const near = nears[k];
        const box = { x: side > 0 ? cx + near : cx - near - it.w, y: ys[k] - lh / 2, w: it.w, h: lh };
        if (box.x < frame.x - 0.5 || box.x + it.w > frame.x + frame.w + 0.5 || box.y < lo - 0.5 || box.y + lh > hi + 0.5
          || placed.some((p) => overlaps(p, box, 0))) { lost.push(it.c); return; }
        placed.push(box);
        const moved = Math.hypot(ys[k] - it.anchor.y, near - side * (it.anchor.x - cx)) > lh * 0.4;
        outside.push({ c: it.c, box, side, moved, anchor: it.anchor, rim: it.rim, leader: moved || !inside.has(it.c) });
      });
    }

    // Leaders: from the rim at the slice's mid-angle to the near edge of the label.
    const leaders = outside.filter((x) => x.leader).map((x) => ({
      c: x.c, p: x.rim, q: { x: x.side > 0 ? x.box.x - 1 : x.box.x + x.box.w + 1, y: x.box.y + lh / 2 },
    }));
    let bad = 0;
    for (const l of leaders) {
      // A segment from a point on the circle stays outside it iff it heads away from the centre.
      if ((l.q.x - l.p.x) * (l.p.x - cx) + (l.q.y - l.p.y) * (l.p.y - cy) < 0) bad++;
      const own = outside.find((x) => x.c === l.c)!.box;
      bad += placed.filter((b) => b !== own && segmentHitsBox(l.p.x, l.p.y, l.q.x, l.q.y, b)).length;
    }
    leaders.forEach((a, i) => leaders.slice(i + 1).forEach((b) => { if (crosses(a.p, a.q, b.p, b.q)) bad++; }));

    const drift = Math.max(0, ...outside.map((x) => Math.hypot(x.box.y + lh / 2 - x.anchor.y, x.side > 0 ? x.box.x - x.anchor.x : x.anchor.x - x.box.x - x.box.w)));
    const crowdedTop = outside.filter((x) => x.moved).reduce((a, x) => a + Math.sign(cy - x.anchor.y), 0) >= 0;
    const score = lost.length * 1000 + bad * 100 + drift;
    return { r, cy, hole, inside, boxed, outside, leaders, lost, bad, drift, crowdedTop, score };
  }

  // Crowded labels (small slices side by side) need room next to their slices: shrink the pie, at
  // most to 70%, and move it away from the crowded side, until they sit close to their slices.
  const mid = (top + bottom) / 2;
  let p = plan(rMax, mid);
  for (let r = rMax * 0.96; (p.drift > lh * 2 || p.lost.length || p.bad) && r >= rMax * 0.7 - 1e-9; r *= 0.96) {
    const shift = (bottom - top) / 2 - r;
    for (const dir of p.crowdedTop ? [1, -1] : [-1, 1]) {
      const next = plan(r, mid + dir * shift);
      if (next.score < p.score - 1) p = next;
    }
  }
  // Centre the pie and its labels together in the space below the legend.
  const blockTop = Math.min(p.cy - p.r, ...p.outside.map((x) => x.box.y));
  const blockBottom = Math.max(p.cy + p.r, ...p.outside.map((x) => x.box.y + lh));
  const offset = (frame.y + legendH + frame.y + frame.h) / 2 - (blockTop + blockBottom) / 2;
  if (Math.abs(offset) > 1) {
    const moved = plan(p.r, p.cy + offset);
    if (moved.lost.length <= p.lost.length && moved.bad <= p.bad && moved.drift <= p.drift + 1) p = moved;
  }
  const { r, cy, hole } = p;

  for (const c of live) {
    const commands: PathCmd[] = full
      ? [{ op: 'A', cx, cy, r, start: -90, sweep: 360 }, ...(doughnut ? [{ op: 'A', cx, cy, r: hole, start: 270, sweep: -360 } as PathCmd] : []), { op: 'Z' }]
      : doughnut
        ? [{ op: 'A', cx, cy, r, start: start[c], sweep: sweep[c] }, { op: 'A', cx, cy, r: hole, start: start[c] + sweep[c], sweep: -sweep[c] }, { op: 'Z' }]
        : [{ op: 'M', x: cx, y: cy }, { op: 'A', cx, cy, r, start: start[c], sweep: sweep[c] }, { op: 'Z' }];
    ctx.add(`wedge:s0:c${c}`, { kind: 'path', commands, fill: colorOf(c), ...(full ? {} : { stroke: style.background, strokeWidth: 1 }) });
  }
  for (const [c, box] of p.inside) {
    text(ctx, `label:value:s0:c${c}`, { ...box, text: valueText[c], color: textOn(colorOf(c)), role: 'value', ...(p.boxed.has(c) ? { fill: colorOf(c) } : {}) });
  }
  if (doughnut && o.totals) {
    const t = ctx.fmt(total);
    const w = m.width(t, fs, true) + 2;
    if (w < hole * 1.7) text(ctx, 'label:total', { x: cx - w / 2, y: cy - lh / 2, w, h: lh, text: t, bold: true, role: 'total' });
    else ctx.hidden.push('total too wide for the doughnut hole');
  }
  for (const c of p.lost) ctx.hidden.push(`${spec.categories[c]}: no room for label`);
  for (const l of p.leaders) {
    ctx.add(`leader:c${l.c}`, { kind: 'line', x1: l.p.x, y1: l.p.y, x2: l.q.x, y2: l.q.y, color: style.muted, width: 0.5 });
  }
  for (const { c, box, side } of p.outside) {
    const t = outsideText(c, p.inside.has(c));
    text(ctx, names ? `label:category:c${c}` : `label:value:s0:c${c}`, { ...box, text: t, align: side > 0 ? 'left' : 'right', role: names ? 'category' : 'value' });
  }
}

/**
 * Stack labels of height `h` on one side of the pie, anchors sorted top to bottom. Labels move away
 * from the equator `mid` — upwards above it, downwards below it — so a leader from the slice always
 * heads out of the pie; then the stack is kept within [min, max] (pushed back towards the equator).
 * Returns the label centres.
 */
function outward(anchors: number[], h: number, min: number, max: number, mid: number): number[] {
  const y = anchors.slice();
  const n = y.length;
  let split = anchors.findIndex((a) => a >= mid);
  if (split < 0) split = n;
  for (let k = split - 2; k >= 0; k--) y[k] = Math.min(y[k], y[k + 1] - h);
  for (let k = split; k < n; k++) if (k > 0) y[k] = Math.max(y[k], y[k - 1] + h);
  for (let k = 0; k < n; k++) y[k] = Math.max(y[k], k ? y[k - 1] + h : min + h / 2);
  for (let k = n - 1; k >= 0; k--) y[k] = Math.min(y[k], k < n - 1 ? y[k + 1] - h : max - h / 2);
  return y;
}

/** True if segments ab and cd properly intersect. */
function crosses(a: Pt, b: Pt, c: Pt, d: Pt): boolean {
  const o = (p: Pt, q: Pt, r: Pt) => Math.sign((q.x - p.x) * (r.y - p.y) - (q.y - p.y) * (r.x - p.x));
  return o(a, b, c) * o(a, b, d) < 0 && o(c, d, a) * o(c, d, b) < 0;
}

/** Mix a hex colour towards another by `t` (0 = unchanged, 1 = the other colour). */
function tint(hex: string, toward: string, t: number): string {
  if (!t) return hex;
  const ch = (h: string, i: number) => parseInt(h.replace('#', '').slice(i, i + 2), 16);
  return [0, 2, 4].map((i) => Math.round(ch(hex, i) + (ch(toward, i) - ch(hex, i)) * t).toString(16).padStart(2, '0')).join('').toUpperCase();
}
