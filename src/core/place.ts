/**
 * Spread labels along one axis so none overlap, keeping each as close to its
 * preferred centre as possible and preserving order. Returns new centres.
 */
export function spread(centres: number[], heights: number[], min: number, max: number, gap = 1): number[] {
  const idx = centres.map((_, i) => i).sort((a, b) => centres[a] - centres[b]);
  const pos = centres.slice();
  // Pass 1: push down from the top.
  let edge = min;
  for (const i of idx) {
    const top = Math.max(pos[i] - heights[i] / 2, edge);
    pos[i] = top + heights[i] / 2;
    edge = top + heights[i] + gap;
  }
  // Pass 2: push up from the bottom if we ran past the end.
  edge = max;
  for (const i of [...idx].reverse()) {
    const bottom = Math.min(pos[i] + heights[i] / 2, edge);
    pos[i] = bottom - heights[i] / 2;
    edge = bottom - heights[i] - gap;
  }
  return pos;
}

export interface Box { x: number; y: number; w: number; h: number }

export const overlaps = (a: Box, b: Box, tol = 0.5) =>
  a.x + tol < b.x + b.w && b.x + tol < a.x + a.w && a.y + tol < b.y + b.h && b.y + tol < a.y + a.h;
