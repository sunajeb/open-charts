// Chart directions (options.direction, manual: Rotate and flip elements). A chart type lays itself out
// "up" (columns) or "right" (bars, profile lines) natively; "down" and "left" are those layouts
// mirrored: columns hang from a baseline at the top, bars grow leftwards from a baseline on the right.
// A mirror is an isometry, so every fit and collision decision of the native layout still holds; text
// is mirrored as a box (never as glyphs) with its alignment flipped, so it keeps hugging what it labels.
import type { ChartOptions, Frame, PathCmd, Shape } from './types.ts';

export type Direction = NonNullable<ChartOptions['direction']>;

/** Bars rather than columns: the categories run down the side and the values across. */
export const sideways = (o: Pick<ChartOptions, 'direction'>): boolean => o.direction === 'left' || o.direction === 'right';

/** The native direction a layout draws before mirroring: up for down, right for left. */
export const nativeDirection = (d: Direction | undefined): Direction => (d === 'down' ? 'up' : d === 'left' ? 'right' : d ?? 'up');

/**
 * Mirror the shapes of a scene in place: across the frame's vertical centre line (`x`, bars growing
 * left) or its horizontal one (`y`, hanging columns). The legend keeps its reading order: it moves as
 * one block (x: stays in its full-width band on top; y: its band goes to the bottom).
 */
export function mirrorShapes(shapes: Shape[], roles: string[], frame: Frame, axis: 'x' | 'y') {
  const sum = axis === 'x' ? 2 * frame.x + frame.w : 2 * frame.y + frame.h;
  const legend = shapes.map((_, i) => roles[i]?.startsWith('legend:') ?? false);
  // The legend block's band, to move it as a whole along y.
  let dy = 0;
  if (axis === 'y' && legend.some(Boolean)) {
    const ys = shapes.flatMap((s, i) => (legend[i] ? extentY(s) : []));
    const [y0, y1] = [Math.min(...ys), Math.max(...ys)];
    dy = sum - y1 - y0;
  }
  shapes.forEach((s, i) => {
    if (legend[i]) { if (axis === 'y') translateY(s, dy); return; }
    mirror(s, axis, sum);
  });
}

function extentY(s: Shape): number[] {
  if (s.kind === 'line') return [s.y1, s.y2];
  if (s.kind === 'path') return s.commands.flatMap((c) => (c.op === 'Z' ? [] : c.op === 'A' ? [c.cy - c.r, c.cy + c.r] : [c.y]));
  return [s.y, s.y + s.h];
}

function translateY(s: Shape, dy: number) {
  if (s.kind === 'line') { s.y1 += dy; s.y2 += dy; return; }
  if (s.kind === 'path') { s.commands = s.commands.map((c) => (c.op === 'Z' ? c : c.op === 'A' ? { ...c, cy: c.cy + dy } : { ...c, y: c.y + dy })); return; }
  s.y += dy;
}

const FLIP_ALIGN = { left: 'right', right: 'left', center: 'center' } as const;
const FLIP_VALIGN = { top: 'bottom', bottom: 'top', middle: 'middle' } as const;

function mirror(s: Shape, axis: 'x' | 'y', sum: number) {
  const x = axis === 'x';
  switch (s.kind) {
    case 'line':
      if (x) { s.x1 = sum - s.x1; s.x2 = sum - s.x2; } else { s.y1 = sum - s.y1; s.y2 = sum - s.y2; }
      return;
    case 'path':
      s.commands = s.commands.map((c) => mirrorCmd(c, axis, sum));
      return;
    case 'text':
      if (x) { s.x = sum - s.x - s.w; s.align = FLIP_ALIGN[s.align]; } else { s.y = sum - s.y - s.h; s.valign = FLIP_VALIGN[s.valign]; }
      return;
    case 'rect':
      if (x) s.x = sum - s.x - s.w; else s.y = sum - s.y - s.h;
      // An up-pointing triangle would point the wrong way after a vertical mirror: no layout that can
      // be mirrored draws one (checked by test/direction.test.ts).
      return;
  }
}

function mirrorCmd(c: PathCmd, axis: 'x' | 'y', sum: number): PathCmd {
  if (c.op === 'Z') return c;
  if (c.op === 'A') {
    // Angles are clockwise from 3 o'clock: a mirror reverses the sweep and reflects the start.
    return axis === 'x'
      ? { ...c, cx: sum - c.cx, start: 180 - c.start, sweep: -c.sweep }
      : { ...c, cy: sum - c.cy, start: -c.start, sweep: -c.sweep };
  }
  return axis === 'x' ? { ...c, x: sum - c.x } : { ...c, y: sum - c.y };
}
