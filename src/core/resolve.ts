// Last line of defence after a chart type has laid itself out: whatever text still collides (with the
// frame, another label, a foreign mark, or an arrow) is hidden — never shrunk below the theme minimum —
// and reported in `hidden`. Chart types fit their labels first; at the default frame this is a no-op,
// and it only steps in at sizes a chart type did not plan for (resizing in the add-in). It runs the same
// tests as the quality gate (src/core/quality.ts) with the same tolerance, so a resolved scene passes it.
import type { Frame, LineShape, Shape, TextShape } from './types.ts';
import type { Measurer } from './text.ts';
import { bounds, coversMark, ownedBy, roleOwner, shrink, textNeed } from './quality.ts';
import { overlaps, segmentHitsBox } from './place.ts';

const TOL = 1;

/** Higher keeps its place first. Arrows and their bubbles are decorations: dropped before any label. */
function priority(role: string): number {
  if (role === 'placeholder') return 100;
  if (role.startsWith('label:category')) return 90;
  if (role.startsWith('legend:')) return 85;
  if (role.startsWith('label:series')) return 80;
  if (role === 'axis') return 75; // axis titles
  if (role.startsWith('label:total')) return 70;
  if (role.startsWith('axis:tick') || role.startsWith('axis:secondary:tick')) return 60;
  if (role.startsWith('shade:')) return 55;
  if (role.startsWith('label:value')) return 50;
  if (role.startsWith('arrow:') || role.startsWith('valueline:') || role === 'indicator') return 10;
  return 40;
}

/** Roles whose shapes stand or fall together (swatch + name, leader + name, bracket + bubble). */
const grouped = (role: string) => /^(arrow:|valueline:|indicator$|legend:|label:series:)/.test(role);

const inFrame = (s: Shape, f: Frame) => {
  const b = bounds(s);
  return b.x >= f.x - TOL && b.y >= f.y - TOL && b.x + b.w <= f.x + f.w + TOL && b.y + b.h <= f.y + f.h + TOL;
};

/**
 * Remove colliding text (and the shapes that only exist for it) from a laid-out scene, in place.
 * Returns the reasons, one per removed label, for `Scene.hidden`.
 */
export function resolveCollisions(shapes: Shape[], roles: string[], frame: Frame, m: Measurer): string[] {
  const units = new Map<string, number[]>();
  shapes.forEach((s, i) => {
    const role = roles[i];
    if (s.kind !== 'text' && !grouped(role)) return;
    const key = grouped(role) ? role : `#${i}`;
    if (!units.has(key)) units.set(key, []);
    units.get(key)!.push(i);
  });
  // A grouped role without text (e.g. a stray leader) is not a unit.
  for (const [k, idx] of units) if (!idx.some((i) => shapes[i].kind === 'text')) units.delete(k);

  // A label placed by hand (pinned) keeps its place first; everything else yields to it.
  const rank = (idx: number[]) => (idx.some((i) => (shapes[i] as TextShape).pinned) ? 1000 : priority(roles[idx[0]]));
  const order = [...units.entries()].sort((a, b) => rank(b[1]) - rank(a[1]) || a[1][0] - b[1][0]);
  const marks = shapes.map((s, i) => i).filter((i) => shapes[i].kind === 'rect' && /^(bar|marker):/.test(roles[i]));
  const dataLines = shapes.map((s, i) => i).filter((i) => shapes[i].kind === 'line' && roles[i].startsWith('line:'));
  const lineOwners = new Set(dataLines.map((i) => roleOwner(roles[i])));
  const kept: number[] = [];
  const drop = new Set<number>();
  const reasons: string[] = [];

  const problem = (idx: number[]): string | null => {
    for (const i of idx) {
      const s = shapes[i];
      if (!inFrame(s, frame)) return 'outside the frame';
      if (s.kind === 'line' && /^(arrow:|valueline:|indicator$)/.test(roles[i])) {
        const ln = s as LineShape;
        const hit = kept.find((k) => roles[k] !== roles[i] && segmentHitsBox(ln.x1, ln.y1, ln.x2, ln.y2, shrink(shapes[k] as TextShape, TOL)));
        if (hit != null) return `arrow crosses "${(shapes[hit] as TextShape).text}"`;
      }
      if (s.kind !== 'text') continue;
      const hit = kept.find((k) => overlaps(s, shapes[k] as TextShape, TOL));
      if (hit != null) return `overlaps "${(shapes[hit] as TextShape).text}"`;
      const mk = marks.find((k) => coversMark(roles[i], s, roles[k], bounds(shapes[k]), lineOwners, TOL));
      if (mk != null) return `covers ${roles[mk]}`;
      const inner = shrink(s, TOL);
      const ln = dataLines.find((k) => {
        const l = shapes[k] as LineShape;
        return !ownedBy(roles[i], roles[k]) && segmentHitsBox(l.x1, l.y1, l.x2, l.y2, inner);
      });
      if (ln != null) return `crosses ${roles[ln]}`;
      const need = textNeed(s, m);
      if (need.w > s.w + TOL || need.h > s.h + TOL) return 'does not fit its box';
    }
    return null;
  };

  for (const [, idx] of order) {
    const why = idx.some((i) => (shapes[i] as TextShape).pinned) ? null : problem(idx);
    if (!why) { for (const i of idx) if (shapes[i].kind === 'text') kept.push(i); continue; }
    for (const i of idx) drop.add(i);
    const t = idx.map((i) => shapes[i]).find((s): s is TextShape => s.kind === 'text')!;
    reasons.push(`"${t.text.replace(/\n/g, ' ')}" (${roles[idx[0]]}) hidden: ${why}`);
  }
  if (!drop.size) return reasons;

  // Pie / doughnut leaders belong to the labels of their slice: drop them when no label is left.
  const labelled = new Set(shapes.flatMap((s, i) => (s.kind === 'text' && !drop.has(i) ? [roles[i].match(/:c(\d+)$/)?.[1]] : [])));
  shapes.forEach((_, i) => { const c = roles[i].match(/^leader:c(\d+)$/)?.[1]; if (c != null && !labelled.has(c)) drop.add(i); });
  // A segment label's leader goes with its label.
  const keptRoles = new Set(shapes.flatMap((s, i) => (s.kind === 'text' && !drop.has(i) ? [roles[i]] : [])));
  shapes.forEach((_, i) => {
    const mm = roles[i].match(/^leader:(s\d+):(c\d+)$/);
    if (mm && !keptRoles.has(`label:value:${mm[1]}:${mm[2]}`) && !keptRoles.has(`label:total:${mm[2]}`)) drop.add(i);
  });

  let w = 0;
  for (let i = 0; i < shapes.length; i++) {
    if (drop.has(i)) continue;
    shapes[w] = shapes[i];
    roles[w] = roles[i];
    w++;
  }
  shapes.length = w;
  roles.length = w;
  return reasons;
}
