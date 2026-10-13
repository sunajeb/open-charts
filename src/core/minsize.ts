// Minimum chart sizes and the "too small" placeholder. The add-in re-lays a chart out on every resize;
// below the minimum a chart would be unreadable (bars a few points wide, every label hidden), so the
// engine draws a placeholder instead: the frame outline and a message saying how large to make it.
import type { ChartSpec, Frame, Scene, Style } from './types.ts';
import type { Measurer } from './text.ts';
import { fits, wrap } from './text.ts';
import { sideways } from './direction.ts';

export interface Size { w: number; h: number }

/** Thrown by a chart type whose content cannot fit the frame at any font size the theme allows. */
export class TooSmall extends Error {
  need: Partial<Size>;
  constructor(need: Partial<Size>, why: string) {
    super(why);
    this.need = need;
  }
}

/**
 * Smallest frame (pt) a chart can be laid out in. Documented per type; category charts grow with
 * their categories (each needs a slot of ~14 pt for a bar and its gap):
 *
 * | type                                              | width               | height              |
 * |---------------------------------------------------|---------------------|---------------------|
 * | stacked, clustered, percent, waterfall, combo,    | max(200, 40 + 14·n) | 120                 |
 * | line, area, area100 (n = categories)              |                     |                     |
 * | the same, direction left or right (bars, profile) | 200                 | max(120, 30 + 12·n) |
 * | mekko, mekkoxy                                    | max(240, 24·n)      | 140                 |
 * | mekko, direction left or right                    | 200                 | max(160, 24·n)      |
 * | pie, doughnut                                     | 200                 | 120                 |
 * | scatter, bubble                                   | 200                 | 140                 |
 * | harvey (r = rows)                                 | max(160, 60 + 24·n) | max(80, 24 + 16·r)  |
 * | table                                             | 120 (+ content)     | 50                  |
 * | gantt (n = activity rows)                         | 320                 | max(120, 40 + 14·n) |
 *
 * Tables and Gantt charts also check their content (columns at the minimum font size) and fall back to
 * the placeholder when it cannot fit.
 */
export function minimumSize(spec: ChartSpec): Size {
  const n = spec.categories.length;
  switch (spec.type) {
    case 'stacked': case 'clustered': case 'percent': case 'waterfall': case 'combo': case 'line': case 'area': case 'area100':
      if (sideways(spec.options ?? {})) return { w: 200, h: Math.max(120, 30 + 12 * n) };
      return { w: Math.max(200, 40 + 14 * n), h: 120 };
    case 'mekko':
      if (sideways(spec.options ?? {})) return { w: 200, h: Math.max(160, 24 * n) };
      return { w: Math.max(240, 24 * n), h: 140 };
    case 'mekkoxy':
      return { w: Math.max(240, 24 * n), h: 140 };
    case 'pie': case 'doughnut':
      return { w: 200, h: 120 };
    case 'scatter': case 'bubble':
      return { w: 200, h: 140 };
    case 'harvey':
      return { w: Math.max(160, 60 + 24 * n), h: Math.max(80, 24 + 16 * spec.series.length) };
    case 'table':
      return { w: 120, h: 50 };
    case 'gantt':
      return { w: 320, h: Math.max(120, 40 + 14 * n) };
    default:
      return { w: 200, h: 120 };
  }
}

/** Frame outline and a centred message, as large as the frame allows (never below the theme minimum). */
export function placeholder(spec: ChartSpec, frame: Frame, need: Size, m: Measurer, style: Style, why?: string): Scene {
  const minFs = Math.min(style.fontSize, 8);
  const shapes: Scene['shapes'] = [{ kind: 'rect', x: frame.x, y: frame.y, w: frame.w, h: frame.h, fill: style.background, stroke: style.axis, strokeWidth: 0.75 }];
  const roles = ['placeholder'];
  const dims = `${Math.ceil(need.w)} × ${Math.ceil(need.h)} pt`;
  const messages = [`Chart too small: make it at least ${dims.replace(/ /g, '\u00a0')}`, 'Chart too small', 'Too small'];
  const inner = { w: frame.w - 8, h: frame.h - 6 };
  outer: for (const msg of messages) {
    for (let size = Math.min(style.fontSize, 10); size >= minFs; size--) {
      const lines = wrap(msg, inner.w, size, m, 3);
      const lh = m.lineHeight(size);
      if (!fits(lines, inner.w, size, m) || lines.length * lh > inner.h) continue;
      const w = Math.max(...lines.map((l) => m.width(l, size))) + 2;
      const h = lines.length * lh;
      shapes.push({
        kind: 'text', x: frame.x + (frame.w - w) / 2, y: frame.y + (frame.h - h) / 2, w, h, text: lines.join('\n'),
        size, color: style.muted, align: 'center', valign: 'middle', lineHeight: lh,
      });
      roles.push('placeholder');
      break outer;
    }
  }
  const reason = `chart too small: ${spec.type} needs at least ${dims}, frame is ${Math.round(frame.w)} × ${Math.round(frame.h)} pt${why ? ` (${why})` : ''}`;
  return { frame, shapes, roles, hidden: [reason] };
}
