import type { Ctx } from '../common.ts';
import { askForRoom, fitCategories, num, reportFit, text } from '../common.ts';
import { wrap, fits } from '../text.ts';

/**
 * Harvey balls: a grid with one ball per value — rows are the series (names on the left), columns
 * the categories (headers on top). A ball is an ellipse outline with a filled wedge from 12 o'clock,
 * clockwise.
 *
 * Values are quarters 0–4 (the default four segments) or, when every value lies in 0–1 and at least
 * one is fractional, a completion fraction 0–1. Out-of-range values are clamped (and reported);
 * empty cells stay empty. Balls use the series colour, else the first palette colour.
 */
export function layoutHarvey(ctx: Ctx) {
  const { spec, style, frame, fs, lh, pad, m } = ctx;
  const nCat = spec.categories.length;
  const nSer = spec.series.length;
  const all = spec.series.flatMap((s) => s.values.map(num)).filter((v): v is number => v != null);
  const fraction = all.length > 0 && all.every((v) => v >= 0 && v <= 1) && all.some((v) => !Number.isInteger(v));
  const max = fraction ? 1 : 4;
  if (all.some((v) => v < 0 || v > max)) ctx.hidden.push(`Harvey ball values outside 0–${max} clamped`);
  const fill = (s: number, c: number) => {
    const v = num(spec.series[s].values[c]);
    return v == null ? null : Math.min(1, Math.max(0, v / max));
  };

  // Name column: as wide as the longest series name, at most 35% of the frame.
  const named = spec.series.some((s) => s.name.trim());
  const nameW = named ? Math.min(frame.w * 0.35, Math.max(...spec.series.map((s) => m.width(s.name, fs))) + pad * 2) : 0;
  const colW = (frame.w - nameW) / nCat;
  const fit = fitCategories(ctx, colW);
  askForRoom(ctx, fit, colW);
  reportFit(ctx, fit);
  const headH = fit.height + pad * 0.4;
  const d0 = fs * 2;
  const rowH = Math.min((frame.h - headH) / nSer, d0 * 1.8);
  const d = Math.max(4, Math.min(d0, rowH * 0.7, colW * 0.6));
  if (rowH < Math.max(lh, 8)) ctx.hidden.push('chart frame too small');
  const x0 = frame.x + nameW;
  // The table keeps its natural row height (balls sized to the text, not stretched to the frame)
  // and sits centred vertically in the frame.
  const y0 = frame.y + Math.max(0, (frame.h - headH - rowH * nSer) / 2) + headH;
  const cx = (c: number) => x0 + colW * (c + 0.5);
  const cy = (s: number) => y0 + rowH * (s + 0.5);

  // Column headers, bottom-aligned on the rule under them.
  spec.categories.forEach((_, c) => {
    if (fit.hidden.has(c)) return;
    const lines = fit.lines[c];
    const h = lines.length * fit.lh;
    text(ctx, `label:category:c${c}`, { x: cx(c) - colW / 2, y: y0 - pad * 0.4 - h, w: colW, h, text: lines.join('\n'), valign: 'bottom', size: fit.size, role: 'category' });
  });
  ctx.add('gridline:0', { kind: 'line', x1: frame.x, y1: y0, x2: frame.x + frame.w, y2: y0, color: style.axis, width: 0.75 });

  spec.series.forEach((ser, s) => {
    // Light rule under every row, the last one closing the table.
    ctx.add(`gridline:${s + 1}`, { kind: 'line', x1: frame.x, y1: y0 + rowH * (s + 1), x2: frame.x + frame.w, y2: y0 + rowH * (s + 1), color: 'D9D9D9', width: 0.5 });
    if (named) {
      const lines = wrap(ser.name, nameW - pad, fs, m, Math.max(1, Math.floor(rowH / lh)));
      if (fits(lines, nameW - pad, fs, m) && lines.length * lh <= rowH + 0.5) {
        text(ctx, `label:series:s${s}`, { x: frame.x, y: cy(s) - (lines.length * lh) / 2, w: nameW - pad, h: lines.length * lh, text: lines.join('\n'), align: 'left', role: 'series' });
      } else ctx.hidden.push(`${ser.name}: series name too long for its row`);
    }
    const color = (ser.color ?? style.palette[0]).replace('#', '');
    spec.categories.forEach((_, c) => {
      const f = fill(s, c);
      if (f == null) return;
      const role = `ball:s${s}:c${c}`;
      const r = d / 2;
      const box = { x: cx(c) - r, y: cy(s) - r, w: d, h: d };
      if (f >= 1) {
        ctx.add(role, { kind: 'rect', ...box, fill: color, stroke: color, strokeWidth: 1, geom: 'ellipse' });
        return;
      }
      ctx.add(role, { kind: 'rect', ...box, fill: style.background, stroke: color, strokeWidth: 1, geom: 'ellipse' });
      if (f > 0) {
        ctx.add(role, {
          kind: 'path', fill: color,
          commands: [{ op: 'M', x: cx(c), y: cy(s) }, { op: 'A', cx: cx(c), cy: cy(s), r, start: -90, sweep: 360 * f }, { op: 'Z' }],
        });
      }
    });
  });
}
