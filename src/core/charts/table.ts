import type { Ctx } from '../common.ts';
import { text } from '../common.ts';
import { inferDecimals } from '../numbers.ts';
import { wrap } from '../text.ts';
import { parseCategoryDate } from '../dates.ts';
import { TooSmall } from '../minsize.ts';

/**
 * Table of native shapes: a header row with the categories, one row per series (its name in the
 * first column), numbers formatted by the numbers service (locale, unit, magnitude, format), text
 * cells shown as they are. Decimals are inferred per metric: per row when the headers are periods
 * (a financial summary), per column otherwise (a comparison table).
 *
 * Columns are fitted to their content: headers and row names wrap (up to three lines) before the
 * font shrinks; the table then stretches its data columns to the frame width. Rows that still do
 * not fit the frame height are dropped and reported.
 */
export function layoutTable(ctx: Ctx) {
  const { spec, style, frame, o, m } = ctx;
  const nCol = spec.categories.length;
  const rows = spec.series.map((_, i) => i);
  const cols = [...Array(nCol).keys()];
  const cellText = (s: number, c: number): string => {
    const v = spec.series[s].values[c];
    if (typeof v === 'string') return v;
    if (typeof v !== 'number' || !Number.isFinite(v)) return '';
    return ctx.fmt(v, { decimals: byRow ? rowDec[s] : colDec[c] });
  };
  const isNum = (v: unknown): v is number => typeof v === 'number' && Number.isFinite(v);
  const dec = (vs: unknown[]) => o.decimals ?? inferDecimals(vs.filter(isNum).map((v) => v / ctx.scale.div));
  const byRow = spec.categories.filter((c) => parseCategoryDate(c)).length * 2 >= nCol;
  const rowDec = rows.map((s) => dec(spec.series[s].values));
  const colDec = cols.map((c) => dec(rows.map((s) => spec.series[s].values[c])));
  const cells = rows.map((s) => cols.map((c) => cellText(s, c)));
  // Text columns ("Preferred", "Exit") read left-aligned; numbers, and text that reads as a number ("11.1%"), right.
  const numeric = (t: string) => /^[(+\-−–]?[$€£¥]?\s?\d/.test(t);
  const align = cols.map((c) => (rows.some((s) => cells[s][c] && !numeric(cells[s][c])) ? 'left' : 'right') as 'left' | 'right');

  // Fit: try the full size, then smaller sizes; at each size wrap headers / row names as needed.
  let fs = ctx.fs;
  let lay = fit(fs);
  while (!lay.ok && fs - 1 >= ctx.minFs) lay = fit(--fs);
  // Columns that do not fit side by side even at the smallest size: the table cannot be drawn here.
  if (lay.need > frame.w + 0.01) throw new TooSmall({ w: lay.need }, `${nCol} columns need ${Math.ceil(lay.need)} pt at ${fs} pt`);
  if (fs < ctx.fs) ctx.hidden.push(`table text shrunk to ${fs} pt to fit`);

  function fit(size: number) {
    const padX = size * 0.5;
    const cellW = cols.map((c) => Math.max(0, ...rows.map((s) => m.width(cells[s][c], size))) + 2);
    const headerMin = cols.map((c) => longestWord(spec.categories[c], size, true));
    const labelW = Math.max(0, ...rows.map((s) => m.width(spec.series[s].name, size, true))) + 2;
    const labelMin = Math.max(0, ...rows.map((s) => longestWord(spec.series[s].name, size, true)));
    const headerFull = cols.map((c) => m.width(spec.categories[c], size, true) + 2);
    // Natural widths; when too wide, cap every column at one shared width (never below its longest word or
    // number), the largest cap that fits; then let the row-name column wrap down to 30% of the frame.
    const natural = cols.map((c) => Math.max(cellW[c], headerFull[c]) + 2 * padX);
    const minCol = cols.map((c) => Math.max(cellW[c], headerMin[c]) + 2 * padX);
    let first = labelW + 2 * padX;
    const firstMin = Math.min(first, Math.max(labelMin + 2 * padX, frame.w * 0.3));
    const widths = (cap: number) => cols.map((c) => Math.max(minCol[c], Math.min(natural[c], cap)));
    const sum = (ws: number[]) => ws.reduce((a, w) => a + w, 0);
    let colW = natural;
    if (first + sum(colW) > frame.w) {
      let lo = 0, hi = Math.max(0, ...natural);
      for (let k = 0; k < 30; k++) { const mid = (lo + hi) / 2; if (firstMin + sum(widths(mid)) <= frame.w) lo = mid; else hi = mid; }
      colW = widths(lo);
      first = Math.max(firstMin, Math.min(first, frame.w - sum(colW)));
    }
    const total = () => first + sum(colW);
    const ok = total() <= frame.w + 0.01;
    // Stretch the data columns to the frame width.
    const extra = Math.max(0, frame.w - total());
    for (const c of cols) colW[c] += extra / Math.max(1, nCol);
    const lh = m.lineHeight(size);
    const headLines = cols.map((c) => wrap(spec.categories[c], colW[c] - 2 * padX, size, m, 3, true));
    const rowLines = rows.map((s) => wrap(spec.series[s].name, first - 2 * padX, size, m, 3, true));
    const padY = size * 0.35;
    const headH = Math.max(1, ...headLines.map((l) => l.length)) * lh + 2 * padY;
    const rowH = rows.map((s) => rowLines[s].length * lh + 2 * padY);
    const height = headH + rowH.reduce((a, h) => a + h, 0);
    const fitsLines = headLines.every((l, c) => l.every((t) => m.width(t, size, true) <= colW[c] - 2 * padX + 0.01))
      && rowLines.every((l) => l.every((t) => m.width(t, size, true) <= first - 2 * padX + 0.01));
    const need = firstMin + sum(widths(0));
    return { ok: ok && fitsLines && height <= frame.h, need, size, padX, padY, lh, colW, first, headLines, rowLines, headH, rowH };
  }

  function longestWord(t: string, size: number, bold: boolean) {
    return Math.max(0, ...t.split(/\s+/).map((w) => m.width(w, size, bold))) + 2;
  }

  const { size, padX, padY, lh, colW, first, headLines, rowLines, headH, rowH } = lay;
  const xs = [frame.x + first];
  for (const c of cols) xs.push(xs[c] + colW[c]);
  const right = xs[nCol];

  // Header: categories, bold, bottom-aligned, right-aligned like the numbers beneath them.
  let y = frame.y;
  cols.forEach((c) => {
    const lines = headLines[c];
    text(ctx, `label:category:c${c}`, {
      x: xs[c] + padX, y: y + headH - padY - lines.length * lh, w: colW[c] - 2 * padX, h: lines.length * lh,
      text: lines.join('\n'), bold: true, align: align[c], size, valign: 'bottom', role: 'category',
    });
  });
  y += headH;
  ctx.add('axis', { kind: 'line', x1: frame.x, y1: y, x2: right, y2: y, color: style.text, width: 1 });

  let k = 0;
  for (const s of rows) {
    if (y + rowH[s] > frame.y + frame.h + 0.01) { ctx.hidden.push(`${spec.series[s].name}: table row does not fit`); continue; }
    const lines = rowLines[s];
    text(ctx, `label:series:s${s}`, {
      x: frame.x + padX, y: y + padY, w: first - 2 * padX, h: lines.length * lh,
      text: lines.join('\n'), align: 'left', valign: 'top', size, role: 'series',
    });
    for (const c of cols) {
      const t = cells[s][c];
      if (!t) continue;
      if (m.width(t, size) > colW[c] - 2 * padX + 0.01) { ctx.hidden.push(`${spec.series[s].name} / ${spec.categories[c]}: cell too wide`); continue; }
      text(ctx, `label:value:s${s}:c${c}`, {
        x: xs[c] + padX, y: y + padY, w: colW[c] - 2 * padX, h: lh, text: t, align: align[c], valign: 'top', size, role: 'value',
      });
    }
    y += rowH[s];
    const last = s === rows[rows.length - 1];
    ctx.add(last ? 'axis' : `gridline:${k++}`, { kind: 'line', x1: frame.x, y1: y, x2: right, y2: y, color: last ? style.text : style.palette[style.palette.length - 1] ?? style.axis, width: last ? 0.75 : 0.5 });
  }
}
