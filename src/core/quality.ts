import type { LineShape, Scene, Shape, TextShape } from './types.ts';
import type { Measurer } from './text.ts';
import { overlaps, segmentHitsBox, type Box } from './place.ts';
import { pathBounds } from './path.ts';

export interface Issue {
  kind: 'overlap' | 'out-of-frame' | 'text-mark' | 'arrow-label' | 'overflow' | 'hidden';
  detail: string;
}

export interface CheckOptions {
  /** Points of slack before two boxes count as touching / a shape counts as outside. Default 1. */
  tol?: number;
  /** Report every hidden or shrunk label as an issue (B13), so tests and `--strict` can fail on it. */
  strict?: boolean;
  /** Hidden-label reasons that strict mode accepts (substring or pattern), e.g. for declared stress fixtures. */
  allowHidden?: (string | RegExp)[];
  /** With a measurer the gate also checks that glyphs fit inside their text box (B8). */
  measurer?: Measurer;
}

const bounds = (s: Shape): Box =>
  s.kind === 'line'
    ? { x: Math.min(s.x1, s.x2), y: Math.min(s.y1, s.y2), w: Math.abs(s.x2 - s.x1), h: Math.abs(s.y2 - s.y1) }
    : s.kind === 'path' ? pathBounds(s) : { x: s.x, y: s.y, w: s.w, h: s.h };

const shrink = (b: Box, d: number): Box => ({ x: b.x + d, y: b.y + d, w: Math.max(0, b.w - 2 * d), h: Math.max(0, b.h - 2 * d) });

/** "bar:s1:c2" / "label:value:s1:c2" → "s1:c2"; "line:s1" → "s1". */
const owner = (role: string) => role.match(/s\d+(:c\d+)?$/)?.[0] ?? null;

/** "label:value:s1:c2" / "bar:s0:c2" → "c2". */
const column = (role: string) => role.match(/:(c\d+)$/)?.[1] ?? null;

function ownedBy(textRole: string, markRole: string): boolean {
  const t = owner(textRole), mk = owner(markRole);
  if (!t || !mk) return false;
  return t === mk || (markRole.startsWith('line:') && t.startsWith(`${mk}:`));
}

/**
 * Layout quality gate v2 (B12): text×text and text×frame, plus text over a mark of another owner
 * (bars, markers, data lines), any shape outside the frame, arrows crossing labels and — with a
 * measurer — glyphs overflowing their box. In strict mode every hidden label is an issue too (B13).
 * Accepts a bare number as the tolerance for compatibility.
 */
export function checkScene(scene: Scene, options: number | CheckOptions = {}): Issue[] {
  const opts = typeof options === 'number' ? { tol: options } : options;
  const tol = opts.tol ?? 1;
  const issues: Issue[] = [];
  const f = scene.frame;
  const roles = scene.roles ?? scene.shapes.map(() => '');
  const items = scene.shapes.map((s, i) => ({ s, role: roles[i] ?? '' }));
  const texts = items.filter((it): it is { s: TextShape; role: string } => it.s.kind === 'text');
  const name = (it: { s: Shape; role: string }) => (it.s.kind === 'text' ? `"${it.s.text}"` : `${it.s.kind} ${it.role || '?'}`);

  for (const it of items) {
    const b = bounds(it.s);
    if (b.x < f.x - tol || b.y < f.y - tol || b.x + b.w > f.x + f.w + tol || b.y + b.h > f.y + f.h + tol) {
      issues.push({ kind: 'out-of-frame', detail: name(it) });
    }
  }

  for (let i = 0; i < texts.length; i++) {
    for (let j = i + 1; j < texts.length; j++) {
      if (overlaps(texts[i].s, texts[j].s, tol)) {
        issues.push({ kind: 'overlap', detail: `"${texts[i].s.text}" × "${texts[j].s.text}"` });
      }
    }
  }

  const marks = items.filter((it) => /^(bar|marker):/.test(it.role) && it.s.kind === 'rect');
  const dataLines = items.filter((it): it is { s: LineShape; role: string } => it.role.startsWith('line:') && it.s.kind === 'line');
  const arrows = items.filter((it): it is { s: LineShape; role: string } => it.role.startsWith('arrow:') && it.s.kind === 'line');
  // A line's point label may sit wholly inside a column segment (combination charts, contrasting colour).
  const lineOwners = new Set(dataLines.map((ln) => owner(ln.role)));
  const within = (a: Box, b: Box) => a.x >= b.x - tol && a.y >= b.y - tol && a.x + a.w <= b.x + b.w + tol && a.y + a.h <= b.y + b.h + tol;
  for (const t of texts) {
    const inner = shrink(t.s, tol);
    const lineLabel = lineOwners.has(owner(t.role)?.split(':')[0] ?? null);
    for (const mk of marks) {
      if (lineLabel && mk.role.startsWith('bar:') && within(t.s, bounds(mk.s))) continue;
      // A boxed segment label (drawn on a rectangle of its segment's colour) may overlap its own column.
      if (t.s.fill && t.role.startsWith('label:value:') && mk.role.startsWith('bar:') && column(t.role) === column(mk.role)) continue;
      if (!ownedBy(t.role, mk.role) && overlaps(t.s, bounds(mk.s), tol)) {
        issues.push({ kind: 'text-mark', detail: `"${t.s.text}" × ${mk.role}` });
      }
    }
    for (const ln of dataLines) {
      if (!ownedBy(t.role, ln.role) && segmentHitsBox(ln.s.x1, ln.s.y1, ln.s.x2, ln.s.y2, inner)) {
        issues.push({ kind: 'text-mark', detail: `"${t.s.text}" × ${ln.role}` });
      }
    }
    for (const ar of arrows) {
      if (ar.role !== t.role && segmentHitsBox(ar.s.x1, ar.s.y1, ar.s.x2, ar.s.y2, inner)) {
        issues.push({ kind: 'arrow-label', detail: `${ar.role} × "${t.s.text}"` });
      }
    }
  }

  if (opts.measurer) {
    const m = opts.measurer;
    for (const { s } of texts) {
      const lines = s.text.split('\n');
      const w = Math.max(...lines.map((l) => m.width(l, s.size, s.bold)));
      const h = lines.length * (s.lineHeight ?? m.lineHeight(s.size));
      if (w > s.w + tol || h > s.h + tol) {
        issues.push({ kind: 'overflow', detail: `"${s.text}" needs ${w.toFixed(1)}×${h.toFixed(1)} pt, box is ${s.w.toFixed(1)}×${s.h.toFixed(1)}` });
      }
    }
  }

  if (opts.strict) {
    const allow = opts.allowHidden ?? [];
    for (const h of scene.hidden) {
      if (!allow.some((a) => (typeof a === 'string' ? h.includes(a) : a.test(h)))) issues.push({ kind: 'hidden', detail: h });
    }
  }
  return issues;
}
