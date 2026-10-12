import type { PathCmd, PathShape, RectShape } from './types.ts';

export interface Pt { x: number; y: number }

/** Largest angle (degrees) one flattened arc step may cover; keeps chords within ~0.05 pt for r ≤ 300 pt. */
const STEP = 2;

/**
 * Flattens a path into closed/open sub-paths of points (arcs become short chords).
 * Every renderer and host draws these same points, so PowerPoint, SVG and VBA always agree.
 */
export function flattenPath(cmds: PathCmd[]): { points: Pt[]; closed: boolean }[] {
  const out: { points: Pt[]; closed: boolean }[] = [];
  let cur: Pt[] = [];
  const flush = (closed: boolean) => {
    if (cur.length > 1) out.push({ points: cur, closed });
    cur = [];
  };
  for (const c of cmds) {
    if (c.op === 'M') { flush(false); cur.push({ x: c.x, y: c.y }); }
    else if (c.op === 'L') cur.push({ x: c.x, y: c.y });
    else if (c.op === 'A') {
      const n = Math.max(1, Math.ceil(Math.abs(c.sweep) / STEP));
      for (let i = 0; i <= n; i++) {
        const a = ((c.start + (c.sweep * i) / n) * Math.PI) / 180;
        cur.push({ x: c.cx + c.r * Math.cos(a), y: c.cy + c.r * Math.sin(a) });
      }
    } else flush(true);
  }
  flush(false);
  return out;
}

export function pathBounds(s: PathShape): { x: number; y: number; w: number; h: number } {
  const pts = flattenPath(s.commands).flatMap((p) => p.points);
  if (!pts.length) return { x: 0, y: 0, w: 0, h: 0 };
  const xs = pts.map((p) => p.x), ys = pts.map((p) => p.y);
  const x = Math.min(...xs), y = Math.min(...ys);
  return { x, y, w: Math.max(...xs) - x, h: Math.max(...ys) - y };
}

/** Effective preset of a rect shape (honours the legacy `ellipse` flag). */
export const rectGeom = (s: RectShape) => s.geom ?? (s.ellipse ? 'ellipse' : 'rect');
