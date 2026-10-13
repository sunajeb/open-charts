// The line format the engine hands to the add-in's VBA: one record per line, tab-separated
// fields, text escaped so a field never contains a tab or newline. VBA reads it with Split(), so
// every property it needs to draw a shape natively is spelled out — VBA never interprets the scene.
//
//   OC <tab> 1                         header + format version
//   G  <tab> name <tab> altText        the chart group
//   F  <tab> x,y,w,h                   the frame: VBA adds an invisible rectangle there, so the
//                                      group's bounds are the frame and a resize is visible as such
//   T  <tab> key <tab> value           a group tag (tag contract v1, OC_SPEC chunked)
//   S  <tab> kind <tab> ...            a shape, columns below (SHAPE_COLUMNS)
//   M  <tab> message                   tell the user something (no error)
//   E  <tab> message                   error: nothing is drawn
//   I  <tab> id                        the edit session opened for this chart (edit)
//   N                                  nothing to do (relayout: the chart is already up to date)
//
// Shape kinds: rect | ellipse | roundRect | diamond | triangle | line | text | path.
// Geometry: rect-like and text use x y w h; line uses x1 y1 x2 y2 in the same four columns; path
// uses the bounding box there and the outline in `points`. Text with a box behind it (fill or
// line set) is drawn as an autoshape of `textGeom` (rect | roundRect) carrying the text. Numbers use "." and at most 2 decimals,
// so VBA's locale-independent Val() reads them. Colours are RRGGBB, or a theme colour name
// (accent1…accent6, dk1 lt1 dk2 lt2, tx1 bg1 tx2 bg2, hlink folHlink) that VBA sets as ObjectThemeColor,
// optionally with a tint / shade ("accent1-lighter40", "accent1-darker25": TintAndShade ±0.4); "" = none.
// Per-element overrides (core/overrides.ts) arrive as: `font` per shape (a font name, or the theme's
// "+mj-lt" / "+mn-lt"), `dash` 0 | 1 | dot, `rotate` -90 | 90 | "" (text direction: TextFrame
// Orientation upward / downward), `arrowStart` 0 | 1, and `pattern`: an OOXML preset pattern
// (dkUpDiag, narVert, …) VBA draws with Fill.Patterned in the fill colour on the background.
import type { PathCmd, Shape, TextShape } from '../core/types.ts';
import { flattenPath } from '../core/path.ts';
import { colorHex, colorSlot } from '../core/theme.ts';

export const FORMAT_VERSION = '2';

export const SHAPE_COLUMNS = [
  'kind', 'x', 'y', 'w', 'h', 'fill', 'lineColor', 'lineWidth', 'dash', 'arrowEnd',
  'text', 'font', 'size', 'bold', 'color', 'align', 'valign', 'wrap', 'autosize',
  'role', 'name', 'points', 'textGeom', 'rotate', 'arrowStart', 'pattern',
] as const;

export type ShapeRecord = Record<(typeof SHAPE_COLUMNS)[number], string>;

export const escapeField = (s: string) =>
  s.replace(/\\/g, '\\\\').replace(/\t/g, '\\t').replace(/\r/g, '\\r').replace(/\n/g, '\\n');

export const unescapeField = (s: string) =>
  s.replace(/\\(.)/g, (_, c: string) => ({ t: '\t', n: '\n', r: '\r' } as Record<string, string>)[c] ?? c);

export const num = (v: number) => (Number.isFinite(v) ? String(+v.toFixed(2)) : '0');
const hex = (c?: string) => {
  if (!c) return '';
  const theme = c.split('@')[1];
  return colorSlot(c) ?? (theme && /-(lighter|darker)\d+$|^(hlink|folHlink)$/.test(theme) ? theme : colorHex(c).toUpperCase());
};
const flag = (b?: boolean) => (b ? '1' : '0');
const dash = (d?: boolean | 'dot') => (d === 'dot' ? 'dot' : flag(!!d));

export const record = (...fields: string[]) => fields.map(escapeField).join('\t');

/** Bounding box of a shape as PowerPoint will report it (lines: their end points). */
export function shapeBounds(s: Shape): { x0: number; y0: number; x1: number; y1: number } {
  const any = s as Shape & { kind: string; commands?: PathCmd[] };
  if (any.kind === 'line') {
    const l = s as Extract<Shape, { kind: 'line' }>;
    return { x0: Math.min(l.x1, l.x2), y0: Math.min(l.y1, l.y2), x1: Math.max(l.x1, l.x2), y1: Math.max(l.y1, l.y2) };
  }
  if (any.kind === 'path') {
    const pts = flattenPath(any.commands ?? []).flatMap((p) => p.points);
    if (!pts.length) return { x0: 0, y0: 0, x1: 0, y1: 0 };
    const xs = pts.map((p) => p.x), ys = pts.map((p) => p.y);
    return { x0: Math.min(...xs), y0: Math.min(...ys), x1: Math.max(...xs), y1: Math.max(...ys) };
  }
  const r = s as { x: number; y: number; w: number; h: number };
  return { x0: r.x, y0: r.y, x1: r.x + r.w, y1: r.y + r.h };
}

const ALIGN = { left: 'l', center: 'c', right: 'r' } as const;
const VALIGN = { top: 't', middle: 'm', bottom: 'b' } as const;

/** One S record. Unknown shape kinds throw, so a newer engine never silently drops shapes. */
export function shapeRecord(s: Shape, font: string, role = '', name = ''): string {
  const any = s as Shape & Record<string, unknown>;
  const base: ShapeRecord = Object.fromEntries(SHAPE_COLUMNS.map((c) => [c, ''])) as ShapeRecord;
  Object.assign(base, { dash: '0', arrowEnd: '0', bold: '0', wrap: '0', autosize: '0', arrowStart: '0', role, name });
  switch (any.kind) {
    case 'rect': {
      const r = s as Extract<Shape, { kind: 'rect' }> & { geom?: string };
      const kind = r.geom && r.geom !== 'rect' ? r.geom : r.ellipse ? 'ellipse' : 'rect';
      Object.assign(base, {
        kind, x: num(r.x), y: num(r.y), w: num(r.w), h: num(r.h), fill: hex(r.fill),
        lineColor: hex(r.stroke), lineWidth: r.stroke ? num(r.strokeWidth ?? 0.75) : '', dash: dash(r.dash), pattern: r.pattern ?? '',
      });
      break;
    }
    case 'line': {
      const l = s as Extract<Shape, { kind: 'line' }>;
      Object.assign(base, {
        kind: 'line', x: num(l.x1), y: num(l.y1), w: num(l.x2), h: num(l.y2),
        lineColor: hex(l.color), lineWidth: num(l.width), dash: dash(l.dash), arrowEnd: flag(l.arrowEnd), arrowStart: flag(l.arrowStart),
      });
      break;
    }
    case 'text': {
      const t = s as TextShape;
      Object.assign(base, {
        kind: 'text', x: num(t.x), y: num(t.y), w: num(t.w), h: num(t.h),
        fill: hex(t.fill), lineColor: hex(t.stroke), lineWidth: t.stroke ? num(t.strokeWidth ?? 0.75) : '',
        text: t.text, font: t.font ?? font, size: num(t.size), bold: flag(t.bold), color: hex(t.color),
        align: ALIGN[t.align] ?? 'l', valign: VALIGN[t.valign] ?? 't', textGeom: t.rounded ? 'roundRect' : '',
        rotate: t.rotate ? String(t.rotate) : '',
      });
      break;
    }
    case 'path': {
      const p = any as unknown as { commands: PathCmd[]; fill?: string; stroke?: string; strokeWidth?: number; dash?: boolean | 'dot' };
      const subs = flattenPath(p.commands ?? []); // the same points every renderer draws
      const b = shapeBounds(s);
      Object.assign(base, {
        kind: 'path', x: num(b.x0), y: num(b.y0), w: num(b.x1 - b.x0), h: num(b.y1 - b.y0),
        fill: hex(p.fill), lineColor: hex(p.stroke), lineWidth: p.stroke ? num(p.strokeWidth ?? 0.75) : '',
        dash: dash(p.dash as boolean | 'dot' | undefined), pattern: (any.pattern as string | undefined) ?? '',
        // "x,y x,y ...", sub-paths separated by "|", a trailing "z" marks a closed sub-path.
        points: subs.map((sp) => sp.points.map((p) => `${num(p.x)},${num(p.y)}`).join(' ') + (sp.closed ? ' z' : '')).join('|'),
      });
      break;
    }
    default:
      throw new Error(`The add-in cannot draw shape kind "${String(any.kind)}" yet`);
  }
  return record('S', ...SHAPE_COLUMNS.map((c) => base[c]));
}

/** Parses a response back into records (tests, and the e2e script's sanity checks). */
export function parseResponse(text: string) {
  const out = { version: '', frame: '', nothing: false, group: undefined as undefined | { name: string; alt: string }, tags: [] as [string, string][], shapes: [] as ShapeRecord[], messages: [] as string[], errors: [] as string[] };
  for (const line of text.split('\n')) {
    if (!line) continue;
    const f = line.split('\t').map(unescapeField);
    if (f[0] === 'OC') out.version = f[1];
    else if (f[0] === 'G') out.group = { name: f[1], alt: f[2] };
    else if (f[0] === 'T') out.tags.push([f[1], f[2]]);
    else if (f[0] === 'S') out.shapes.push(Object.fromEntries(SHAPE_COLUMNS.map((c, i) => [c, f[i + 1] ?? ''])) as ShapeRecord);
    else if (f[0] === 'M') out.messages.push(f[1]);
    else if (f[0] === 'E') out.errors.push(f[1]);
    else if (f[0] === 'F') out.frame = f[1];
    else if (f[0] === 'N') out.nothing = true;
  }
  return out;
}
