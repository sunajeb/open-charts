import type { Ctx } from '../common.ts';
import { legend, num, text } from '../common.ts';
import { roundToTotal, textOn } from '../numbers.ts';
import { overlaps, type Box } from '../place.ts';
import type { PathCmd } from '../types.ts';

/** Doughnut hole as a share of the outer radius. */
const HOLE = 0.55;

/**
 * Pie and doughnut charts: one series of values across the categories (further series are ignored).
 *
 * - Slices run clockwise from 12 o'clock in category order, each a `path` with an arc; colours
 *   follow the palette by category (the last slice never repeats the first slice's colour).
 * - Every slice carries its share of the total, rounded together (largest remainder) so the labels
 *   add up to exactly 100%. The share sits inside the slice, as close to the rim as it fits; when it
 *   does not fit it moves outside, after the category name, with a leader line.
 * - Category names sit outside the pie, left or right depending on the slice's side, spread apart
 *   vertically (default, `seriesLabels: 'right'`), or in a legend (`'legend'`), or nowhere
 *   (`'none'`). A label pushed away from its slice gets a leader line; when labels crowd (small
 *   slices side by side) the pie shrinks a little so they can sit next to their slices.
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
  const live = cats.filter((c) => vals[c] > 0);
  for (const c of cats) if (!vals[c]) ctx.hidden.push(`${spec.categories[c]}: zero value, no slice`);
  const pal = style.palette.map((p) => p.replace('#', ''));
  // Beyond the palette, colours repeat as lighter tints, so no two slices (or legend entries) match.
  const colorOf = (c: number) => tint(pal[c % pal.length], style.background, Math.min(0.8, Math.floor(c / pal.length) * 0.45));

  const names = o.seriesLabels === 'right';
  const outsideText = (c: number, insideOk: boolean) =>
    (!names ? pctText[c] : insideOk ? spec.categories[c] : `${spec.categories[c]} ${pctText[c]}`);
  const legendH = legend(ctx, 'box', colorOf, { names: spec.categories, key: 'c' });
  const gap = pad * 1.5;
  const outsideW = Math.max(...live.map((c) => m.width(outsideText(c, false), fs))) + 2;
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

  /** Label placement for a given radius, without emitting anything. */
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
    const placed: Box[] = [];
    const inside = new Map<number, Box>();
    for (const c of live) {
      const w = m.width(pctText[c], fs) + 2;
      const halfDiag = Math.hypot(w, lh) / 2;
      for (let d = r - halfDiag - pad * 0.5; d >= hole + halfDiag * 0.5; d -= 2) {
        const box = { x: cx + d * Math.cos(midOf(c)) - w / 2, y: cy + d * Math.sin(midOf(c)) - lh / 2, w, h: lh };
        const corners = [[box.x, box.y], [box.x + w, box.y], [box.x, box.y + lh], [box.x + w, box.y + lh]];
        if (corners.every(([x, y]) => inSlice(c, x, y)) && !placed.some((p) => overlaps(p, box, 0))) {
          placed.push(box);
          inside.set(c, box);
          break;
        }
      }
    }
    const outside: { c: number; box: Box; side: number; moved: boolean; anchor: { x: number; y: number } }[] = [];
    const lost: number[] = [];
    let drift = 0;
    for (const side of [1, -1]) {
      const items = live.filter((c) => (names || !inside.has(c)) && (Math.cos(midOf(c)) >= 0 ? 1 : -1) === side);
      const anchor = items.map((c) => ({ x: cx + (r + gap) * Math.cos(midOf(c)), y: cy + (r + gap) * Math.sin(midOf(c)) }));
      const ys = stack(anchor.map((a) => a.y), lh, frame.y + legendH, frame.y + frame.h);
      items.forEach((c, k) => {
        const w = m.width(outsideText(c, inside.has(c)), fs) + 2;
        const y = ys[k] - lh / 2;
        // Keep clear of the circle across the label's whole height.
        const dy = y <= cy && y + lh >= cy ? 0 : Math.min(Math.abs(y - cy), Math.abs(y + lh - cy));
        const chord = dy < r + gap ? Math.sqrt(Math.max(0, (r + gap) ** 2 - dy * dy)) : 0;
        const near = Math.max(side * (anchor[k].x - cx), chord);
        const box = { x: side > 0 ? cx + near : cx - near - w, y, w, h: lh };
        if (box.x < frame.x - 0.5 || box.x + w > frame.x + frame.w + 0.5 || placed.some((p) => overlaps(p, box, 0))) { lost.push(c); return; }
        placed.push(box);
        const shift = Math.hypot(ys[k] - anchor[k].y, near - side * (anchor[k].x - cx));
        drift = Math.max(drift, shift);
        outside.push({ c, box, side, moved: shift > lh * 0.4, anchor: anchor[k] });
      });
    }
    const crowdedTop = outside.filter((x) => x.moved).reduce((a, x) => a + Math.sign(cy - x.anchor.y), 0) >= 0;
    return { r, cy, hole, inside, outside, lost, drift, crowdedTop };
  }

  // Crowded labels (small slices side by side) need room next to their slices: shrink the pie, at
  // most to 80%, and move it away from the crowded side, until they sit close to their slices.
  const mid = (top + bottom) / 2;
  let p = plan(rMax, mid);
  for (let r = rMax * 0.96; (p.drift > lh || p.lost.length) && r >= rMax * 0.8 - 1e-9; r *= 0.96) {
    const shift = (bottom - top) / 2 - r;
    const next = plan(r, mid + (p.crowdedTop ? shift : -shift));
    if (next.lost.length < p.lost.length || (next.lost.length === p.lost.length && next.drift < p.drift - 1)) p = next;
  }
  // Centre the pie and its labels together in the space below the legend.
  const blockTop = Math.min(p.cy - p.r, ...p.outside.map((x) => x.box.y));
  const blockBottom = Math.max(p.cy + p.r, ...p.outside.map((x) => x.box.y + lh));
  const offset = (frame.y + legendH + frame.y + frame.h) / 2 - (blockTop + blockBottom) / 2;
  if (Math.abs(offset) > 1) {
    const moved = plan(p.r, p.cy + offset);
    if (moved.lost.length <= p.lost.length && moved.drift <= p.drift + 1) p = moved;
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
    text(ctx, `label:value:s0:c${c}`, { ...box, text: pctText[c], color: textOn(colorOf(c)), role: 'value' });
  }
  if (doughnut && o.totals) {
    const t = ctx.fmt(total);
    const w = m.width(t, fs, true) + 2;
    if (w < hole * 1.7) text(ctx, 'label:total', { x: cx - w / 2, y: cy - lh / 2, w, h: lh, text: t, bold: true, role: 'total' });
    else ctx.hidden.push('total too wide for the doughnut hole');
  }
  for (const c of p.lost) ctx.hidden.push(`${spec.categories[c]}: no room for label`);
  for (const { c, box, side, moved } of p.outside) {
    const share = !p.inside.has(c);
    if (moved || share) {
      // Leader from the slice (inside it for a small slice whose share sits outside) to the label.
      const from = share ? Math.max(hole + 2, r - Math.min(r * 0.1, pad * 1.5)) : r + 1;
      const q = { x: side > 0 ? box.x - 1 : box.x + box.w + 1, y: box.y + lh / 2 };
      ctx.add(`leader:c${c}`, { kind: 'line', x1: cx + from * Math.cos(midOf(c)), y1: cy + from * Math.sin(midOf(c)), x2: q.x, y2: q.y, color: style.muted, width: 0.5 });
    }
    const t = outsideText(c, !share);
    text(ctx, names ? `label:category:c${c}` : `label:value:s0:c${c}`, { ...box, text: t, align: side > 0 ? 'left' : 'right', role: names ? 'category' : 'value' });
  }
}

/**
 * Stack labels of equal height along y, each as close to its anchor as possible: overlapping labels
 * form a cluster centred on the mean of its anchors (so a crowd grows both ways), kept within [min, max].
 */
function stack(anchors: number[], h: number, min: number, max: number): number[] {
  const order = anchors.map((_, i) => i).sort((a, b) => anchors[a] - anchors[b]);
  type Cluster = { ids: number[]; top: number };
  const clusters: Cluster[] = [];
  const place = (c: Cluster) => {
    const mean = c.ids.reduce((a, i) => a + anchors[i], 0) / c.ids.length;
    c.top = Math.min(Math.max(mean - (c.ids.length * h) / 2, min), max - c.ids.length * h);
  };
  for (const i of order) {
    const c = { ids: [i], top: 0 };
    place(c);
    clusters.push(c);
    // Merge with the cluster above while they overlap.
    while (clusters.length > 1) {
      const a = clusters[clusters.length - 2], b = clusters[clusters.length - 1];
      if (a.top + a.ids.length * h <= b.top) break;
      a.ids.push(...b.ids);
      clusters.pop();
      place(a);
    }
  }
  const out = anchors.slice();
  for (const c of clusters) c.ids.forEach((i, k) => { out[i] = c.top + h * k + h / 2; });
  return out;
}

/** Mix a hex colour towards another by `t` (0 = unchanged, 1 = the other colour). */
function tint(hex: string, toward: string, t: number): string {
  if (!t) return hex;
  const ch = (h: string, i: number) => parseInt(h.replace('#', '').slice(i, i + 2), 16);
  return [0, 2, 4].map((i) => Math.round(ch(hex, i) + (ch(toward, i) - ch(hex, i)) * t).toString(16).padStart(2, '0')).join('').toUpperCase();
}
