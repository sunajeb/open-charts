import type { Ctx } from '../common.ts';
import { fitCategories, reportFit, text } from '../common.ts';
import { parseCategoryDate } from '../dates.ts';
import { overlaps, type Box } from '../place.ts';

/**
 * Gantt chart (timeline). Categories are the activity rows; the first series holds start dates, the
 * second end dates (bars), every further series a kind of milestone (diamonds with a date label).
 * Dates are read with parseCategoryDate ("2026-03-15", "Mar 2026", "Q2 2026"); as in the reference
 * product a date means the beginning of that day, so an activity ending "2026-03-16" includes the 15th.
 *
 * - Calendar scale: the finest of week / month / quarter / year whose labels fit becomes the lower
 *   tier, the next larger unit the upper tier. Larger units get solid separators, smaller dashed ones.
 * - The date range fits the data, snapped to whole units of the lower tier.
 * - Row labels sit in a column on the left (wrapped, shrunk, then hidden and reported).
 * - `options.shades`: shaded periods behind every row, labelled under the chart.
 */
export function layoutGantt(ctx: Ctx) {
  const { spec, style, frame, o, fs, lh, pad, m } = ctx;
  const nRow = spec.categories.length;
  const rows = [...Array(nRow).keys()];
  const day = (v: unknown): number | null => {
    if (typeof v !== 'string') return null;
    const d = parseCategoryDate(v);
    return d ? Date.UTC(d.y, d.m - 1, d.d) / DAY : null;
  };
  const start = rows.map((r) => day(spec.series[0]?.values[r]));
  const end = rows.map((r) => day(spec.series[1]?.values[r]));
  const msSeries = spec.series.map((_, i) => i).slice(2);
  const ms = msSeries.flatMap((s) => rows.map((r) => ({ s, r, d: day(spec.series[s].values[r]) }))).filter((x): x is { s: number; r: number; d: number } => x.d != null);
  const shades = (o.shades ?? []).map((sh) => ({ from: day(sh.from)!, to: day(sh.to)!, label: sh.label ?? '' }));
  const all = [...start, ...end, ...ms.map((x) => x.d), ...ms.map((x) => x.d + 1), ...shades.flatMap((s) => [s.from, s.to])].filter((d): d is number => d != null);
  if (!all.length) { ctx.hidden.push('Gantt chart has no dates'); return; }
  let lo = Math.min(...all), hi = Math.max(...all);

  // Row label column.
  const catW = Math.min(frame.w * 0.3, Math.max(...spec.categories.map((c) => m.width(c, fs))) + 2);
  const fit = fitCategories(ctx, catW + 2, 2);
  reportFit(ctx, fit);
  const px = frame.x + catW + pad * 1.5;
  const pw = frame.x + frame.w - px;

  // Calendar scale: the finest unit whose labels fit (and that does not draw a forest of lines).
  const tierH = lh + pad * 0.6;
  const diaMax = fs * 1.1;
  const pickScale = () => {
    for (const u of ['week', 'month', 'quarter', 'year'] as Unit[]) {
      const a = floorTo(u, lo), b = ceilTo(u, hi);
      const n = units(u, a, b).length;
      const w = (pw * avgDays[u]) / (b - a);
      if (u === 'year' || (n <= 60 && w >= fineLabelW(u) + pad * 0.6)) return u;
    }
    return 'year' as Unit;
  };
  // Milestones near either end need room for the diamond (and on the right for the date label).
  let fine = pickScale();
  for (let k = 0; k < 2; k++) {
    const ppd = pw / (ceilTo(fine, hi) - floorTo(fine, lo));
    const span = ceilTo(fine, hi) - floorTo(fine, lo);
    for (const x of ms) {
      lo = Math.min(lo, x.d - diaMax / 2 / ppd);
      hi = Math.max(hi, x.d + (diaMax / 2 + pad * 0.3 + m.width(dateLabel(x.d, o.locale, span), fs) + 2) / ppd);
    }
    fine = pickScale();
  }
  const t0 = floorTo(fine, lo), t1 = ceilTo(fine, hi);
  const coarse = UP[fine];
  const xOf = (d: number) => px + ((d - t0) / (t1 - t0)) * pw;

  function fineLabelW(u: Unit) {
    return Math.max(...SAMPLE[u].map((t) => m.width(t, fs))) + 2;
  }

  // Shade labels under the chart; rows fill the rest, but not taller than three lines each.
  const shadeH = shades.some((s) => s.label) ? lh + pad * 0.6 : 0;
  // Milestone kinds need a key once there is more than one.
  const showKey = msSeries.length > 1 && o.seriesLabels !== 'none';
  const keyH = showKey ? lh + pad * 0.6 : 0;
  const headH = tierH * (coarse ? 2 : 1);
  const top = frame.y + headH;
  const rowH = Math.min((frame.y + frame.h - shadeH - keyH - top) / Math.max(1, nRow), Math.max(lh * 2.4, fit.lh * 2 + pad));
  if (rowH < lh) ctx.hidden.push('chart frame too small');
  const bottom = top + rowH * nRow;
  const yMid = (r: number) => top + rowH * (r + 0.5);

  // Shades behind everything.
  shades.forEach((sh, k) => {
    const x1 = Math.max(px, xOf(sh.from)), x2 = Math.min(px + pw, xOf(sh.to));
    if (x2 <= x1) return;
    ctx.add(`shade:${k}`, { kind: 'rect', x: x1, y: top, w: x2 - x1, h: bottom - top, fill: SHADE });
    if (!sh.label) return;
    const w = m.width(sh.label, fs) + 2;
    const lx = Math.min(Math.max((x1 + x2) / 2 - w / 2, px), px + pw - w);
    text(ctx, `shade:${k}`, { x: lx, y: bottom + pad * 0.4, w, h: lh, text: sh.label, color: style.muted, role: 'decoration' });
  });

  // Header tiers, separators and the rule under the header.
  const tiers: [Unit, number][] = coarse ? [[coarse, frame.y], [fine, frame.y + tierH]] : [[fine, frame.y]];
  let tick = 0, grid = 0;
  for (const [u, y] of tiers) {
    const isFine = u === fine;
    const cells = units(u, floorTo(u, t0), t1);
    cells.forEach(([a, b]) => {
      const x1 = Math.max(px, xOf(a)), x2 = Math.min(px + pw, xOf(b));
      if (x2 - x1 < 1) return;
      const room = x2 - x1 - pad * 0.4;
      const label = unitLabel(u, a, !isFine && !!coarse).find((t) => m.width(t, fs) + 2 <= room) ?? null;
      const role = `axis:tick:${tick++}`;
      if (label) {
        const w = m.width(label, fs) + 2;
        const lx = isFine ? (x1 + x2) / 2 - w / 2 : x1 + pad * 0.3;
        text(ctx, role, { x: lx, y: y + pad * 0.3, w, h: lh, text: label, color: isFine ? style.muted : style.text, bold: !isFine && !!coarse, align: isFine ? 'center' : 'left', role: 'axis' });
      }
      // Separator at the unit's start: solid for the upper tier (through rows and header), dashed for the lower.
      if (a > t0) {
        if (isFine && coarse && a === floorTo(coarse, a)) return; // drawn by the upper tier
        if (isFine) {
          ctx.add(`gridline:${grid++}`, { kind: 'line', x1: xOf(a), y1: y, x2: xOf(a), y2: bottom, color: SEP_FINE, width: 0.5, dash: true });
        } else {
          // Through its own tier and the rows; it skips the lower tier when it falls inside one of its cells (a week across two months).
          const inside = coarse && floorTo(fine, a) !== a;
          const color = style.axis;
          if (inside) {
            ctx.add(`gridline:${grid++}`, { kind: 'line', x1: xOf(a), y1: frame.y, x2: xOf(a), y2: frame.y + tierH, color, width: 0.75 });
            ctx.add(`gridline:${grid++}`, { kind: 'line', x1: xOf(a), y1: top, x2: xOf(a), y2: bottom, color, width: 0.75 });
          } else ctx.add(`gridline:${grid++}`, { kind: 'line', x1: xOf(a), y1: frame.y, x2: xOf(a), y2: bottom, color, width: 0.75 });
        }
      }
    });
  }
  ctx.add('axis', { kind: 'line', x1: frame.x, y1: top, x2: px + pw, y2: top, color: style.text, width: 0.75 });
  if (coarse) ctx.add('axis', { kind: 'line', x1: px, y1: frame.y + tierH, x2: px + pw, y2: frame.y + tierH, color: style.axis, width: 0.5 });

  // Rows: label, separator, bar, milestones.
  const barH = Math.min(rowH * 0.5, lh * 1.1);
  const dia = Math.min(rowH * 0.6, diaMax);
  const msColors = [style.text, style.muted, ...style.palette];
  const msColor = (s: number) => (spec.series[s].color ?? msColors[(s - 2) % msColors.length]).replace('#', '');
  const items: Box[][] = rows.map(() => []);
  for (const r of rows) {
    if (!fit.hidden.has(r)) {
      const lines = fit.lines[r];
      const h = lines.length * fit.lh;
      text(ctx, `label:category:c${r}`, { x: frame.x, y: yMid(r) - h / 2, w: catW + 2, h, text: lines.join('\n'), align: 'left', size: fit.size, role: 'category' });
    }
    if (r < nRow - 1) ctx.add(`gridline:${grid++}`, { kind: 'line', x1: frame.x, y1: top + rowH * (r + 1), x2: px + pw, y2: top + rowH * (r + 1), color: SEP_FINE, width: 0.5 });
    const a = start[r], b = end[r];
    if (a != null && b != null && b > a) {
      const box = { x: xOf(a), y: yMid(r) - barH / 2, w: xOf(b) - xOf(a), h: barH };
      items[r].push(box);
      ctx.add(`bar:s0:c${r}`, { kind: 'rect', ...box, fill: ctx.color(0), stroke: style.background, strokeWidth: 0.5 });
    }
  }
  for (const x of ms) {
    const box = { x: xOf(x.d) - dia / 2, y: yMid(x.r) - dia / 2, w: dia, h: dia };
    items[x.r].push(box);
    ctx.add(`marker:s${x.s}:c${x.r}`, { kind: 'rect', ...box, fill: msColor(x.s), stroke: style.background, strokeWidth: 0.75, geom: 'diamond' });
  }

  if (showKey) {
    let x = frame.x;
    const y = bottom + shadeH + pad * 0.4;
    const ks = Math.min(dia, fs * 0.9);
    for (const s of msSeries) {
      const name = spec.series[s].name;
      const w = m.width(name, fs) + 2;
      if (x + ks + pad * 0.5 + w > frame.x + frame.w) { ctx.hidden.push(`${name}: no room in milestone key`); continue; }
      ctx.add(`legend:s${s}`, { kind: 'rect', x, y: y + (lh - ks) / 2, w: ks, h: ks, fill: msColor(s), geom: 'diamond' });
      text(ctx, `legend:s${s}`, { x: x + ks + pad * 0.5, y, w, h: lh, text: name, align: 'left', role: 'series' });
      x += ks + pad * 0.5 + w + pad * 2;
    }
  }

  // Milestone date labels: right of the diamond, else left; never over a bar, diamond or other label.
  const placed: Box[] = [];
  for (const x of ms) {
    const label = dateLabel(x.d, o.locale, t1 - t0);
    const w = m.width(label, fs) + 2;
    const yc = yMid(x.r);
    const right = { x: xOf(x.d) + dia / 2 + pad * 0.3, y: yc - lh / 2, w, h: lh };
    const left = { x: xOf(x.d) - dia / 2 - pad * 0.3 - w, y: yc - lh / 2, w, h: lh };
    const spot = [right, left].find((b) => b.x >= px && b.x + b.w <= frame.x + frame.w &&
      !items[x.r].some((it) => overlaps(it, b, 0)) && !placed.some((p) => overlaps(p, b)));
    if (!spot) { ctx.hidden.push(`${spec.series[x.s].name} / ${spec.categories[x.r]}: no room for date label`); continue; }
    placed.push(spot);
    text(ctx, `label:value:s${x.s}:c${x.r}`, { ...spot, text: label, color: style.text, role: 'value' });
  }
}

type Unit = 'week' | 'month' | 'quarter' | 'year';
const DAY = 86_400_000;
const SHADE = 'EDEDED';
const SEP_FINE = 'D9D9D9';
const UP: Record<Unit, Unit | null> = { week: 'month', month: 'year', quarter: 'year', year: null };
const avgDays: Record<Unit, number> = { week: 7, month: 30.44, quarter: 91.3, year: 365.25 };
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const SAMPLE: Record<Unit, string[]> = { week: ['28'], month: ['May', 'Mar'], quarter: ['Q4'], year: ['2026'] };

const ymd = (d: number) => { const t = new Date(d * DAY); return { y: t.getUTCFullYear(), m: t.getUTCMonth(), d: t.getUTCDate(), dow: t.getUTCDay() }; };
const utc = (y: number, m: number, d = 1) => Date.UTC(y, m, d) / DAY;

function floorTo(u: Unit, d: number): number {
  const t = ymd(d);
  if (u === 'week') return d - ((t.dow + 6) % 7); // weeks start on Monday
  if (u === 'month') return utc(t.y, t.m);
  if (u === 'quarter') return utc(t.y, t.m - (t.m % 3));
  return utc(t.y, 0);
}

function nextOf(u: Unit, d: number): number {
  const t = ymd(d);
  if (u === 'week') return d + 7;
  if (u === 'month') return utc(t.y, t.m + 1);
  if (u === 'quarter') return utc(t.y, t.m + 3);
  return utc(t.y + 1, 0);
}

const ceilTo = (u: Unit, d: number) => (floorTo(u, d) === d ? d : nextOf(u, floorTo(u, d)));

function units(u: Unit, a: number, b: number): [number, number][] {
  const out: [number, number][] = [];
  for (let d = floorTo(u, a); d < b && out.length < 500; d = nextOf(u, d)) out.push([d, nextOf(u, d)]);
  return out;
}

/** Labels for one calendar unit, longest first; the upper tier spells out the year on months. */
function unitLabel(u: Unit, d: number, upper: boolean): string[] {
  const t = ymd(d);
  const q = `Q${Math.floor(t.m / 3) + 1}`;
  if (u === 'year') return [String(t.y), `'${String(t.y).slice(2)}`];
  if (u === 'quarter') return [q];
  if (u === 'month') return upper ? [`${MON[t.m]} ${t.y}`, MON[t.m], MON[t.m][0]] : [MON[t.m], MON[t.m][0]];
  return [String(t.d)];
}

/** Milestone date: day and month; the year joins in once the chart spans more than a year, the day drops after three. */
function dateLabel(d: number, locale: string | undefined, span: number): string {
  const t = ymd(d);
  const us = !locale || locale === 'en-US';
  if (span > 3 * 366) return `${MON[t.m]} ${t.y}`;
  if (span > 366) return us ? `${MON[t.m]} ${t.d}, ${t.y}` : `${t.d} ${MON[t.m]} ${t.y}`;
  return us ? `${MON[t.m]} ${t.d}` : `${t.d} ${MON[t.m]}`;
}
