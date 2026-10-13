// Path B planning for the Office.js add-in: a laid-out Scene becomes a list of plain drawing
// operations that map 1:1 onto PowerPoint JS calls (addGeometricShape / addLine / addTextBox).
// Pure, so it is unit-tested in Node; src/officejs/powerpoint.ts executes the operations.
//
// What PowerPoint JS cannot draw (docs/plan/03 §2.1): freeform geometry, arrowheads, line flips.
// The planner maps what it can onto presets — rotated lines, rotated triangles for arrowheads,
// Pie / BlockArc / Donut presets with adjustments for wedges — when the host has the API sets for
// it (HostCaps). Anything else makes the scene "Path A only": the host then inserts the slide our
// OOXML writer produced (insertSlidesFromBase64), which draws everything exactly.
import type { LineShape, PathCmd, PathShape, RectShape, Scene, Shape, TextShape } from '../core/types.ts';
import { colorHex } from '../core/theme.ts';
import { rectGeom } from '../core/path.ts';
import { shapeName } from '../render/pptx.ts';

/** What the running PowerPoint can do, from Office.context.requirements (PowerPointApi sets). */
export interface HostCaps {
  /** 1.4: add shapes, lines, text boxes; fill, line and text formatting. Path B floor. */
  shapes: boolean;
  /** 1.5: getSelectedShapes / setSelectedShapes (edit flow). */
  selection: boolean;
  /** 1.8: addGroup, parentGroup, zOrderPosition, setZOrder, exportAsBase64, Slide.index. */
  groups: boolean;
  /** 1.10: Shape.rotation, Shape.adjustments, altTextDescription, Slide.themeColorScheme. */
  rotation: boolean;
  /** 1.2: insertSlidesFromBase64 (Path A). */
  insertSlides: boolean;
}

export const NO_CAPS: HostCaps = { shapes: false, selection: false, groups: false, rotation: false, insertSlides: false };

export function hostCaps(isSet: (version: string) => boolean): HostCaps {
  return { insertSlides: isSet('1.2'), shapes: isSet('1.4'), selection: isSet('1.5'), groups: isSet('1.8'), rotation: isSet('1.10') };
}

export interface Box { left: number; top: number; width: number; height: number }
/** `dash`: PowerPoint JS ShapeLineDashStyle ('Dash' for the writer's prstDash dash, 'SystemDot' for sysDot). */
export interface LineStyle { color: string; weight: number; dash?: 'Dash' | 'SystemDot' }
export interface TextStyle {
  text: string; font: string; size: number; bold: boolean; color: string;
  align: 'Left' | 'Center' | 'Right'; valign: 'Top' | 'Middle' | 'Bottom';
}

/** One PowerPoint shape to create. Colours are "#RRGGBB"; `role` becomes the child's OC_ROLE tag. */
export type Op =
  | { kind: 'geom'; geom: string; box: Box; fill: string | null; line: LineStyle | null; rotation?: number; adjustments?: number[]; text?: TextStyle; role?: string; name: string }
  | { kind: 'line'; box: Box; line: LineStyle; rotation?: number; role?: string; name: string }
  | { kind: 'text'; box: Box; fill: string | null; line: LineStyle | null; rotation?: number; text: TextStyle; role?: string; name: string };

export type Plan = { path: 'B'; ops: Op[] } | { path: 'A'; reason: string };

const r2 = (v: number) => Math.round(v * 1000) / 1000;
const box = (x: number, y: number, w: number, h: number): Box => ({ left: r2(x), top: r2(y), width: r2(w), height: r2(h) });
const norm = (deg: number) => r2(((deg % 360) + 360) % 360);
const dashOf = (d?: boolean | 'dot') => (d === 'dot' ? { dash: 'SystemDot' as const } : d ? { dash: 'Dash' as const } : {});
const noPattern = (s: { pattern?: string }, role?: string) => { if (s.pattern) throw new Unsupported(`"${role ?? 'shape'}" has a pattern fill, which PowerPoint JS cannot set`); };

/**
 * A scene colour ("1F3A5F", "1F3A5F@accent1") as "#RRGGBB". PowerPoint JS sets RGB only, so theme
 * slots are drawn as the colour they resolved to — the host lays the chart out against the slide's
 * own theme colours (Slide.themeColorScheme, 1.10) so these are the deck's colours.
 */
export const hexOf = (c: string): string => `#${colorHex(c).toUpperCase()}`;

class Unsupported extends Error {}

/**
 * Plans Path B for a scene. Returns Path A (with the first reason) when any shape needs an API the
 * host lacks or geometry no preset draws.
 */
export function planScene(scene: Scene, caps: HostCaps, font: string): Plan {
  if (!caps.shapes) return { path: 'A', reason: 'PowerPointApi 1.4 (adding shapes) is not available' };
  if (!caps.groups) return { path: 'A', reason: 'PowerPointApi 1.8 (grouping shapes) is not available' };
  const hex = hexOf;
  const ops: Op[] = [];
  try {
    scene.shapes.forEach((s, i) => {
      const role = scene.roles?.[i] || undefined;
      const name = shapeName(s, i + 2, role);
      ops.push(...shapeOps(s, caps, font, hex, role, name));
    });
  } catch (e) {
    if (e instanceof Unsupported) return { path: 'A', reason: e.message };
    throw e;
  }
  return { path: 'B', ops };
}

type Hex = (c: string) => string;

function shapeOps(s: Shape, caps: HostCaps, font: string, hex: Hex, role: string | undefined, name: string): Op[] {
  if (s.kind === 'rect') return [rectOp(s, hex, role, name)];
  if (s.kind === 'line') return lineOps(s, caps, hex, role, name);
  if (s.kind === 'text') return [textOp(s, caps, font, hex, role, name)];
  return [pathOp(s, caps, hex, role, name)];
}

const GEOM: Record<string, string> = { rect: 'Rectangle', ellipse: 'Ellipse', diamond: 'Diamond', roundRect: 'RoundRectangle', triangle: 'Triangle' };

function rectOp(s: RectShape, hex: Hex, role: string | undefined, name: string): Op {
  noPattern(s, role);
  return {
    kind: 'geom', geom: GEOM[rectGeom(s)], box: box(s.x, s.y, s.w, s.h), fill: hex(s.fill),
    line: s.stroke ? { color: hex(s.stroke), weight: s.strokeWidth ?? 0.75, ...dashOf(s.dash) } : null, role, name,
  };
}

function lineOps(s: LineShape, caps: HostCaps, hex: Hex, role: string | undefined, name: string): Op[] {
  const line: LineStyle = { color: hex(s.color), weight: s.width, ...dashOf(s.dash) };
  const dx = s.x2 - s.x1, dy = s.y2 - s.y1;
  const out: Op[] = [];
  if (dx * dy >= 0) {
    // Axis-aligned or top-left → bottom-right: the bounding box draws it as the writer does.
    out.push({ kind: 'line', box: box(Math.min(s.x1, s.x2), Math.min(s.y1, s.y2), Math.abs(dx), Math.abs(dy)), line, role, name });
  } else {
    // The other diagonal needs flipH, which PowerPoint JS cannot set: a horizontal line of the same
    // length through the midpoint, rotated onto the segment.
    if (!caps.rotation) throw new Unsupported('a rising diagonal line needs Shape.rotation (PowerPointApi 1.10)');
    const len = Math.hypot(dx, dy);
    const mx = (s.x1 + s.x2) / 2, my = (s.y1 + s.y2) / 2;
    out.push({ kind: 'line', box: box(mx - len / 2, my, len, 0), line, rotation: norm((Math.atan2(dy, dx) * 180) / Math.PI), role, name });
  }
  if (s.arrowEnd) out.push(arrowHead(s, caps, line.color, role, name));
  if (s.arrowStart) out.push(arrowHead({ ...s, x1: s.x2, y1: s.y2, x2: s.x1, y2: s.y1 }, caps, line.color, role, `${name} start`));
  return out;
}

/**
 * The writer's `tailEnd type="triangle" w="med" len="med"`: a filled isosceles triangle three line
 * widths wide and long, its tip on the line's end point. The Triangle preset points up; rotating it
 * points it along the line.
 */
function arrowHead(s: LineShape, caps: HostCaps, color: string, role: string | undefined, name: string): Op {
  const size = 3 * s.width;
  const ang = (Math.atan2(s.y2 - s.y1, s.x2 - s.x1) * 180) / Math.PI;
  const rotation = norm(ang + 90);
  if (rotation !== 0 && !caps.rotation) throw new Unsupported('an arrowhead needs Shape.rotation (PowerPointApi 1.10)');
  const len = Math.hypot(s.x2 - s.x1, s.y2 - s.y1) || 1;
  const cx = s.x2 - ((s.x2 - s.x1) / len) * (size / 2), cy = s.y2 - ((s.y2 - s.y1) / len) * (size / 2);
  return { kind: 'geom', geom: 'Triangle', box: box(cx - size / 2, cy - size / 2, size, size), fill: color, line: null, ...(rotation ? { rotation } : {}), role, name: `${name} head` };
}

const ALIGN = { left: 'Left', center: 'Center', right: 'Right' } as const;
const VALIGN = { top: 'Top', middle: 'Middle', bottom: 'Bottom' } as const;

function textOp(t: TextShape, caps: HostCaps, chartFont: string, hex: Hex, role: string | undefined, name: string): Op {
  // "+mj-lt" / "+mn-lt" (the theme's fonts) are DrawingML names; PowerPoint JS takes a family name.
  if (t.font?.startsWith('+')) throw new Unsupported(`"${role ?? 'label'}" uses a theme font, which PowerPoint JS cannot name`);
  const text: TextStyle = { text: t.text, font: t.font ?? chartFont, size: t.size, bold: !!t.bold, color: hex(t.color), align: ALIGN[t.align], valign: VALIGN[t.valign] };
  const fill = t.fill ? hex(t.fill) : null;
  const line = t.stroke ? { color: hex(t.stroke), weight: t.strokeWidth ?? 0.75 } : null;
  let b = box(t.x, t.y, t.w, t.h);
  let rotation: number | undefined;
  if (t.rotate) {
    // Turned text (vert270 / vert in the writer): the unturned box, centred on the footprint, rotated.
    if (!caps.rotation) throw new Unsupported('turned text needs Shape.rotation (PowerPointApi 1.10)');
    b = box(t.x + t.w / 2 - t.h / 2, t.y + t.h / 2 - t.w / 2, t.h, t.w);
    rotation = t.rotate === -90 ? 270 : 90;
  }
  if (t.rounded) {
    // CAGR bubble: the writer's roundRect has adj 50000 (fully round ends); the adjustment needs 1.10,
    // without it the preset's default corner is used (docs/plan/03 §4.1 fallback).
    return { kind: 'geom', geom: 'RoundRectangle', box: b, fill, line, ...(caps.rotation ? { adjustments: [0.5] } : {}), ...(rotation ? { rotation } : {}), text, role, name };
  }
  return { kind: 'text', box: b, fill, line, ...(rotation ? { rotation } : {}), text, role, name };
}

const near = (a: number, b: number, eps = 0.01) => Math.abs(a - b) <= eps;
type Arc = Extract<PathCmd, { op: 'A' }>;
const isArc = (c: PathCmd | undefined): c is Arc => c?.op === 'A';

/** Start/end angle of an arc drawn clockwise (PowerPoint's pie/blockArc presets sweep clockwise). */
function clockwise(a: Arc): [number, number] {
  const start = a.sweep >= 0 ? a.start : a.start + a.sweep;
  return [norm(start), norm(start + Math.abs(a.sweep))];
}

/**
 * Freeform shapes PowerPoint JS has a preset for: full circles (Ellipse), pie wedges (Pie), full rings
 * (Donut), ring segments (BlockArc), and isosceles triangles with an axis-aligned base (Triangle,
 * rotated). Adjustment values follow the VBA/OOXML conventions: angles in degrees clockwise from
 * 3 o'clock, ring thickness as a fraction of the shape's size.
 */
function pathOp(s: PathShape, caps: HostCaps, hex: Hex, role: string | undefined, name: string): Op {
  noPattern(s, role);
  const fill = s.fill ? hex(s.fill) : null;
  const line: LineStyle | null = s.stroke ? { color: hex(s.stroke), weight: s.strokeWidth ?? 0.75, ...dashOf(s.dash) } : null;
  const c = s.commands;
  const base = { kind: 'geom' as const, fill, line, role, name };
  const circle = (a: Arc) => box(a.cx - a.r, a.cy - a.r, 2 * a.r, 2 * a.r);
  const needAdj = (what: string) => { if (!caps.rotation) throw new Unsupported(`${what} needs Shape.adjustments (PowerPointApi 1.10)`); };

  // Full circle: A 360, Z.
  if (c.length === 2 && isArc(c[0]) && c[1].op === 'Z' && Math.abs(c[0].sweep) >= 360 - 1e-9) {
    return { ...base, geom: 'Ellipse', box: circle(c[0]) };
  }
  // Pie wedge: M centre, A, Z.
  if (c.length === 3 && c[0].op === 'M' && isArc(c[1]) && c[2].op === 'Z' && near(c[0].x, c[1].cx) && near(c[0].y, c[1].cy)) {
    const a = c[1];
    if (Math.abs(a.sweep) >= 360 - 1e-9) return { ...base, geom: 'Ellipse', box: circle(a) };
    needAdj('a pie wedge');
    return { ...base, geom: 'Pie', box: circle(a), adjustments: clockwise(a) };
  }
  // Ring: A outer, A inner (reverse), Z — concentric.
  if (c.length === 3 && isArc(c[0]) && isArc(c[1]) && c[2].op === 'Z' && near(c[0].cx, c[1].cx) && near(c[0].cy, c[1].cy) && c[1].r < c[0].r) {
    const [o, i] = [c[0], c[1]];
    needAdj('a doughnut segment');
    const thick = r2((o.r - i.r) / (2 * o.r));
    if (Math.abs(o.sweep) >= 360 - 1e-9) return { ...base, geom: 'Donut', box: circle(o), adjustments: [thick] };
    return { ...base, geom: 'BlockArc', box: circle(o), adjustments: [...clockwise(o), thick] };
  }
  // Triangle: M, L, L, Z with an axis-aligned base and the tip on its perpendicular bisector.
  if (c.length === 4 && c[0].op === 'M' && c[1].op === 'L' && c[2].op === 'L' && c[3].op === 'Z') {
    const pts = [c[0], c[1], c[2]] as { x: number; y: number }[];
    for (let k = 0; k < 3; k++) {
      const tip = pts[k], p = pts[(k + 1) % 3], q = pts[(k + 2) % 3];
      const horiz = near(p.y, q.y), vert = near(p.x, q.x);
      if (!horiz && !vert) continue;
      const mx = (p.x + q.x) / 2, my = (p.y + q.y) / 2;
      if (horiz ? !near(tip.x, mx) : !near(tip.y, my)) continue;
      const baseLen = Math.hypot(q.x - p.x, q.y - p.y), height = Math.hypot(tip.x - mx, tip.y - my);
      // Unrotated the preset's tip points up; rotation turns it about the box centre.
      const rotation = horiz ? (tip.y < my ? 0 : 180) : (tip.x > mx ? 90 : 270);
      if (rotation && !caps.rotation) throw new Unsupported('a triangle that does not point up needs Shape.rotation (PowerPointApi 1.10)');
      const cx = (tip.x + mx) / 2, cy = (tip.y + my) / 2;
      return { ...base, geom: 'Triangle', box: box(cx - baseLen / 2, cy - height / 2, baseLen, height), ...(rotation ? { rotation } : {}) };
    }
  }
  throw new Unsupported(`"${role ?? 'shape'}" is a freeform PowerPoint JS cannot draw`);
}
