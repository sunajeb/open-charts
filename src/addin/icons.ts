// The chart gallery's icons, drawn once as a handful of flat shapes on a 48 × 48 grid: rendered as
// inline SVG for the editor's gallery and rasterised (anti-aliased, zero dependencies) to the PNGs the
// PowerPoint ribbon shows (addin/icons/*.png, written by tools/icons.ts and checked by a test).
// Flat bright blues on a light-blue baseline read on light and dark ribbons alike.
import { readFileSync } from 'node:fs';
import { deflateSync } from 'node:zlib';
import { parse } from '../../vendor/opentype.module.js';

export type IconShape =
  | { k: 'rect'; x: number; y: number; w: number; h: number; c: string }
  | { k: 'poly'; pts: number[]; c: string }
  | { k: 'circle'; cx: number; cy: number; r: number; c: string }
  /** Ring sector (r0 = 0: a pie wedge); angles in degrees clockwise from 12 o'clock. */
  | { k: 'sector'; cx: number; cy: number; r0: number; r1: number; a0: number; a1: number; c: string }
  /** Stroked polyline with round joins and caps. */
  | { k: 'line'; pts: number[]; w: number; c: string }
  /** Filled outlines (text glyphs), non-zero winding; each polygon is x0, y0, x1, y1, … */
  | { k: 'glyphs'; polys: number[][]; c: string };

/** The palette: three bright blues, a pale one, the light-blue baseline and the combination line's green. */
export const ICON_COLORS = { c1: '1D5BD8', c2: '3B82F6', c3: '6FA8FF', c4: 'A9CCFF', axis: 'B7D3FF', green: '22B35E' } as const;
const { c1, c2, c3, c4, axis, green } = ICON_COLORS;

const SIZE = 48;
const base = (x1 = 4, x2 = 44, y = 42): IconShape => ({ k: 'rect', x: x1, y, w: x2 - x1, h: 1.6, c: axis });
const yAxis = (x = 4, y1 = 6, y2 = 43.6): IconShape => ({ k: 'rect', x, y: y1, w: 1.6, h: y2 - y1, c: axis });
const col = (x: number, w: number, segs: [number, string][], bottom = 42): IconShape[] => {
  let y = bottom;
  return segs.map(([h, c]) => { y -= h; return { k: 'rect', x, y, w, h: h - 0.6, c } as IconShape; });
};
const shift = (shapes: IconShape[], dy: number): IconShape[] => shapes.map((s) => (s.k === 'rect' ? { ...s, y: s.y + dy } : s));

/** Icon of each gallery tile (and the direction arrows), in the "up" direction. */
export const ICONS: Record<string, IconShape[]> = {
  stacked: [
    ...col(7, 9, [[10, c1], [7, c2], [5, c3]]), ...col(19.5, 9, [[14, c1], [8, c2], [6, c3]]), ...col(32, 9, [[18, c1], [10, c2], [7, c3]]),
    base(),
  ],
  percent: [
    ...col(7, 9, [[16, c1], [12, c2], [8, c3]]), ...col(19.5, 9, [[12, c1], [14, c2], [10, c3]]), ...col(32, 9, [[20, c1], [9, c2], [7, c3]]),
    base(),
  ],
  clustered: [
    ...col(5, 5.5, [[18, c1]]), ...col(10.5, 5.5, [[12, c3]]), ...col(18.5, 5.5, [[26, c1]]), ...col(24, 5.5, [[19, c3]]),
    ...col(32, 5.5, [[22, c1]]), ...col(37.5, 5.5, [[30, c3]]),
    base(),
  ],
  waterfall: [
    ...col(5, 6.5, [[10, c3]]), ...shift(col(12.5, 6.5, [[7, c3]]), -10), ...shift(col(20, 6.5, [[8, c3]]), -17),
    ...shift(col(27.5, 6.5, [[6, c3]]), -25), ...col(35, 7, [[31, c1]]),
    base(),
  ],
  waterfallDown: [
    ...col(5, 7, [[31, c1]]), ...shift(col(13, 6.5, [[8, c3]]), -23), ...shift(col(20.5, 6.5, [[7, c3]]), -16),
    ...shift(col(28, 6.5, [[6, c3]]), -10), ...col(35.5, 6.5, [[10, c3]]),
    base(),
  ],
  mekko: [
    ...col(5, 15, [[16, c1], [12, c2], [8, c3]]), ...col(20.6, 10, [[10, c1], [16, c2], [10, c3]]), ...col(31.2, 12.8, [[22, c1], [8, c2], [6, c3]]),
    base(),
  ],
  combo: [
    ...col(7, 9, [[14, c2]]), ...col(19.5, 9, [[20, c2]]), ...col(32, 9, [[17, c2]]),
    { k: 'line', pts: [11.5, 22, 24, 13, 36.5, 9], w: 2.2, c: green },
    { k: 'circle', cx: 11.5, cy: 22, r: 2.3, c: green }, { k: 'circle', cx: 24, cy: 13, r: 2.3, c: green }, { k: 'circle', cx: 36.5, cy: 9, r: 2.3, c: green },
    base(),
  ],
  line: [
    { k: 'line', pts: [6, 30, 15, 24, 24, 27, 33, 15, 42, 10], w: 2.4, c: c1 },
    { k: 'line', pts: [6, 37, 15, 34, 24, 31, 33, 31, 42, 24], w: 2.4, c: c3 },
    base(),
  ],
  area: [
    { k: 'poly', pts: [4, 42, 4, 22, 17, 17, 30, 19, 44, 9, 44, 42], c: c3 },
    { k: 'poly', pts: [4, 42, 4, 29, 17, 26, 30, 27, 44, 20, 44, 42], c: c2 },
    { k: 'poly', pts: [4, 42, 4, 35, 17, 33, 30, 34, 44, 30, 44, 42], c: c1 },
    base(),
  ],
  area100: [
    { k: 'rect', x: 4, y: 6, w: 40, h: 36, c: c3 },
    { k: 'poly', pts: [4, 42, 4, 20, 17, 15, 30, 21, 44, 13, 44, 42], c: c2 },
    { k: 'poly', pts: [4, 42, 4, 31, 17, 27, 30, 32, 44, 26, 44, 42], c: c1 },
    base(),
  ],
  mekkoUnit: [
    ...col(5, 14, [[14, c1], [9, c2], [6, c3]]), ...col(19.6, 11, [[8, c1], [10, c2], [5, c3]]), ...col(31.2, 12.8, [[18, c1], [10, c2], [7, c3]]),
    yAxis(3), base(3),
  ],
  pie: [
    { k: 'sector', cx: 24, cy: 24, r0: 0, r1: 17, a0: 0, a1: 140, c: c1 },
    { k: 'sector', cx: 24, cy: 24, r0: 0, r1: 17, a0: 141.5, a1: 250, c: c2 },
    { k: 'sector', cx: 24, cy: 24, r0: 0, r1: 17, a0: 251.5, a1: 358.5, c: c3 },
  ],
  doughnut: [
    { k: 'sector', cx: 24, cy: 24, r0: 9.5, r1: 17, a0: 0, a1: 150, c: c1 },
    { k: 'sector', cx: 24, cy: 24, r0: 9.5, r1: 17, a0: 152, a1: 255, c: c2 },
    { k: 'sector', cx: 24, cy: 24, r0: 9.5, r1: 17, a0: 257, a1: 358, c: c3 },
  ],
  scatter: [
    yAxis(), base(4),
    ...[[12, 33, c1], [17, 26, c1], [23, 29, c1], [28, 19, c3], [35, 13, c3], [39, 21, c3], [31, 30, c1]].map(([cx, cy, c]) => ({ k: 'circle', cx, cy, r: 2.6, c }) as IconShape),
  ],
  bubble: [
    yAxis(), base(4),
    { k: 'circle', cx: 15, cy: 30, r: 6, c: c3 }, { k: 'circle', cx: 29, cy: 18, r: 9, c: c1 }, { k: 'circle', cx: 38, cy: 32, r: 4.5, c: c2 },
  ],
  gantt: [
    { k: 'rect', x: 4, y: 6, w: 40, h: 1.6, c: axis },
    { k: 'rect', x: 6, y: 11, w: 16, h: 6, c: c1 }, { k: 'rect', x: 16, y: 21, w: 18, h: 6, c: c2 }, { k: 'rect', x: 28, y: 31, w: 14, h: 6, c: c3 },
    { k: 'poly', pts: [12, 30, 15.5, 34, 12, 38, 8.5, 34], c: c1 },
  ],
  mekkoxy: [
    ...col(5, 12, [[8, c1], [4, c3]]), ...col(17.6, 8, [[13, c1], [4, c3]]), ...col(26.2, 7, [[19, c1], [5, c3]]), ...col(33.8, 10, [[26, c1], [6, c3]]),
    base(),
  ],
  harvey: [
    { k: 'circle', cx: 10, cy: 24, r: 7, c: c4 }, { k: 'sector', cx: 10, cy: 24, r0: 0, r1: 7, a0: 0, a1: 90, c: c1 },
    { k: 'circle', cx: 24, cy: 24, r: 7, c: c4 }, { k: 'sector', cx: 24, cy: 24, r0: 0, r1: 7, a0: 0, a1: 180, c: c1 },
    { k: 'circle', cx: 38, cy: 24, r: 7, c: c1 },
  ],
  table: [
    { k: 'rect', x: 4, y: 9, w: 40, h: 7, c: c1 },
    ...[16.6, 24.2, 31.8].map((y) => ({ k: 'rect', x: 4, y, w: 40, h: 6.4, c: y === 24.2 ? c4 : 'DCE9FF' }) as IconShape),
    { k: 'rect', x: 4, y: 38.4, w: 40, h: 1.6, c: axis },
  ],
  // The direction buttons next to the ribbon gallery: one triangle each.
  dirUp: [{ k: 'poly', pts: [24, 10, 40, 36, 8, 36], c: c1 }],
  dirDown: [{ k: 'poly', pts: [8, 12, 40, 12, 24, 38], c: c1 }],
  dirLeft: [{ k: 'poly', pts: [10, 24, 36, 8, 36, 40], c: c1 }],
  dirRight: [{ k: 'poly', pts: [12, 8, 38, 24, 12, 40], c: c1 }],
};

// ---- SVG ---------------------------------------------------------------------------------------

const n = (v: number) => String(Math.round(v * 100) / 100);
const polar = (cx: number, cy: number, r: number, a: number) => [cx + r * Math.sin((a * Math.PI) / 180), cy - r * Math.cos((a * Math.PI) / 180)];

function sectorPath(s: Extract<IconShape, { k: 'sector' }>): string {
  const large = s.a1 - s.a0 > 180 ? 1 : 0;
  const [x1, y1] = polar(s.cx, s.cy, s.r1, s.a0), [x2, y2] = polar(s.cx, s.cy, s.r1, s.a1);
  if (!s.r0) return `M${n(s.cx)} ${n(s.cy)}L${n(x1)} ${n(y1)}A${n(s.r1)} ${n(s.r1)} 0 ${large} 1 ${n(x2)} ${n(y2)}Z`;
  const [x3, y3] = polar(s.cx, s.cy, s.r0, s.a1), [x4, y4] = polar(s.cx, s.cy, s.r0, s.a0);
  return `M${n(x1)} ${n(y1)}A${n(s.r1)} ${n(s.r1)} 0 ${large} 1 ${n(x2)} ${n(y2)}L${n(x3)} ${n(y3)}A${n(s.r0)} ${n(s.r0)} 0 ${large} 0 ${n(x4)} ${n(y4)}Z`;
}

/** One icon as an inline SVG (no ids, so many can share a page). */
export function iconSvg(name: string, cls = 'gal-ico'): string {
  const body = (ICONS[name] ?? []).map((s) => {
    switch (s.k) {
      case 'rect': return `<rect x="${n(s.x)}" y="${n(s.y)}" width="${n(s.w)}" height="${n(s.h)}" fill="#${s.c}"/>`;
      case 'poly': return `<polygon points="${s.pts.map(n).join(' ')}" fill="#${s.c}"/>`;
      case 'circle': return `<circle cx="${n(s.cx)}" cy="${n(s.cy)}" r="${n(s.r)}" fill="#${s.c}"/>`;
      case 'sector': return `<path d="${sectorPath(s)}" fill="#${s.c}"/>`;
      case 'line': return `<polyline points="${s.pts.map(n).join(' ')}" fill="none" stroke="#${s.c}" stroke-width="${n(s.w)}" stroke-linejoin="round" stroke-linecap="round"/>`;
      case 'glyphs': return `<path d="${s.polys.map((pl) => `M${pl.map(n).join(' ')}Z`).join('')}" fill="#${s.c}"/>`;
    }
  }).join('');
  return `<svg class="${cls}" viewBox="0 0 ${SIZE} ${SIZE}" aria-hidden="true">${body}</svg>`;
}

// ---- Raster (PNG) --------------------------------------------------------------------------------

const segDist2 = (px: number, py: number, ax: number, ay: number, bx: number, by: number) => {
  const dx = bx - ax, dy = by - ay;
  const t = Math.max(0, Math.min(1, ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy || 1)));
  const qx = ax + t * dx - px, qy = ay + t * dy - py;
  return qx * qx + qy * qy;
};

/** Whether icon point (x, y) is inside a shape. */
function inside(s: IconShape, x: number, y: number): boolean {
  switch (s.k) {
    case 'rect': return x >= s.x && x < s.x + s.w && y >= s.y && y < s.y + s.h;
    case 'circle': return (x - s.cx) ** 2 + (y - s.cy) ** 2 <= s.r * s.r;
    case 'sector': {
      const d2 = (x - s.cx) ** 2 + (y - s.cy) ** 2;
      if (d2 > s.r1 * s.r1 || d2 < s.r0 * s.r0) return false;
      const a = ((Math.atan2(x - s.cx, s.cy - y) * 180) / Math.PI + 360) % 360;
      return a >= s.a0 && a <= s.a1;
    }
    case 'poly': {
      let on = false;
      for (let i = 0, j = s.pts.length - 2; i < s.pts.length; j = i, i += 2) {
        const [xi, yi, xj, yj] = [s.pts[i], s.pts[i + 1], s.pts[j], s.pts[j + 1]];
        if ((yi > y) !== (yj > y) && x < ((xj - xi) * (y - yi)) / (yj - yi) + xi) on = !on;
      }
      return on;
    }
    case 'line': {
      const r2 = (s.w / 2) ** 2;
      for (let i = 0; i + 3 < s.pts.length; i += 2) if (segDist2(x, y, s.pts[i], s.pts[i + 1], s.pts[i + 2], s.pts[i + 3]) <= r2) return true;
      return false;
    }
    case 'glyphs': {
      let wind = 0;
      for (const pl of s.polys) {
        for (let i = 0, j = pl.length - 2; i < pl.length; j = i, i += 2) {
          const [x0, y0, x1, y1] = [pl[j], pl[j + 1], pl[i], pl[i + 1]];
          if (y0 <= y ? y1 > y && (x1 - x0) * (y - y0) - (x - x0) * (y1 - y0) > 0 : y1 <= y && (x1 - x0) * (y - y0) - (x - x0) * (y1 - y0) < 0) wind += y0 <= y ? 1 : -1;
        }
      }
      return wind !== 0;
    }
  }
}

/** RGBA pixels of an icon at `px` × `px`, transparent background, 4 × 4 supersampled. */
export function rasterize(name: string, px: number): Uint8Array {
  return rasterShapes(ICONS[name] ?? [], px, px, SIZE / px);
}

/** Bounding box [x0, y0, x1, y1] of a shape (a little generous). */
function bounds(s: IconShape): [number, number, number, number] {
  const of = (xs: number[], ys: number[], d = 0): [number, number, number, number] => [Math.min(...xs) - d, Math.min(...ys) - d, Math.max(...xs) + d, Math.max(...ys) + d];
  const split = (pts: number[]) => [pts.filter((_, i) => !(i % 2)), pts.filter((_, i) => i % 2)];
  switch (s.k) {
    case 'rect': return [s.x, s.y, s.x + s.w, s.y + s.h];
    case 'circle': return [s.cx - s.r, s.cy - s.r, s.cx + s.r, s.cy + s.r];
    case 'sector': return [s.cx - s.r1, s.cy - s.r1, s.cx + s.r1, s.cy + s.r1];
    case 'poly': { const [xs, ys] = split(s.pts); return of(xs, ys); }
    case 'line': { const [xs, ys] = split(s.pts); return of(xs, ys, s.w); }
    case 'glyphs': { const all = s.polys.flat(); const [xs, ys] = split(all); return xs.length ? of(xs, ys) : [0, 0, -1, -1]; }
  }
}

/** RGBA pixels of shapes drawn on a `w` × `h` pixel canvas, `scale` icon units per pixel. */
function rasterShapes(shapes: IconShape[], w: number, h: number, scale: number): Uint8Array {
  const rgb = shapes.map((s) => [parseInt(s.c.slice(0, 2), 16), parseInt(s.c.slice(2, 4), 16), parseInt(s.c.slice(4, 6), 16)]);
  const box = shapes.map(bounds);
  const out = new Uint8Array(w * h * 4);
  const SS = 4;
  for (let py = 0; py < h; py++) {
    for (let qx = 0; qx < w; qx++) {
      let r = 0, g = 0, b = 0, a = 0;
      for (let sy = 0; sy < SS; sy++) {
        for (let sx = 0; sx < SS; sx++) {
          const x = (qx + (sx + 0.5) / SS) * scale, y = (py + (sy + 0.5) / SS) * scale;
          // The topmost shape covering the sample wins (shapes are opaque, drawn in order).
          for (let k = shapes.length - 1; k >= 0; k--) {
            const bb = box[k];
            if (x < bb[0] || x > bb[2] || y < bb[1] || y > bb[3] || !inside(shapes[k], x, y)) continue;
            r += rgb[k][0]; g += rgb[k][1]; b += rgb[k][2]; a++;
            break;
          }
        }
      }
      const o = (py * w + qx) * 4;
      if (a) { out[o] = Math.round(r / a); out[o + 1] = Math.round(g / a); out[o + 2] = Math.round(b / a); out[o + 3] = Math.round((255 * a) / (SS * SS)); }
    }
  }
  return out;
}

const CRC = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) { let c = i; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[i] = c >>> 0; }
  return t;
})();
const crc32 = (buf: Uint8Array) => { let c = 0xffffffff; for (const b of buf) c = CRC[(c ^ b) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
function chunk(type: string, data: Uint8Array): Buffer {
  const head = Buffer.alloc(8);
  head.writeUInt32BE(data.length, 0); head.write(type, 4, 'ascii');
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(Buffer.concat([head.subarray(4), data])), 0);
  return Buffer.concat([head, data, crc]);
}

function png(rgba: Uint8Array, w: number, h: number): Buffer {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 6;
  const raw = Buffer.alloc((w * 4 + 1) * h);
  for (let y = 0; y < h; y++) Buffer.from(rgba.buffer, y * w * 4, w * 4).copy(raw, y * (w * 4 + 1) + 1);
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw, { level: 9 })), chunk('IEND', Buffer.alloc(0))]);
}

/** An icon as a PNG file (8-bit RGBA). Deterministic: the same icon gives the same bytes. */
export function iconPng(name: string, px: number): Buffer {
  return png(rasterize(name, px), px, px);
}

// ---- Gallery tiles for the ribbon: icon over its label, in one image ------------------------------
// PowerPoint for Mac draws a gallery item's label beside its image, which makes uneven, wide rows;
// a tile image with the label under the icon gives the reference's grid (the item keeps its label
// for screen readers and its tooltip).

/** Tile size in points: icon 40 pt on top, two label lines below. */
export const TILE = { w: 64, h: 62, icon: 40, font: 9, line: 10.5 } as const;
/** Label colour: a mid grey that reads on PowerPoint's light and dark ribbons. */
export const TILE_LABEL = '7D8492';

type Font = { getPath(text: string, x: number, y: number, size: number): { commands: { type: string; x?: number; y?: number; x1?: number; y1?: number; x2?: number; y2?: number }[] }; getAdvanceWidth(text: string, size: number): number };
let font: Font | undefined;
const labelFont = (): Font => {
  if (!font) {
    const bytes = readFileSync(new URL('../../vendor/fonts/LiberationSans-Regular.ttf', import.meta.url));
    font = parse(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)) as unknown as Font;
  }
  return font;
};

/** Text as filled outlines, one shape per character, centred on `cx` with its baseline at `y`, in icon units. */
function textShapes(text: string, cx: number, y: number, size: number, c: string): IconShape[] {
  const f = labelFont();
  let x = cx - f.getAdvanceWidth(text, size) / 2;
  const out: IconShape[] = [];
  for (const ch of text) {
    out.push(glyphShape(f, ch, x, y, size, c));
    x += f.getAdvanceWidth(ch, size);
  }
  return out;
}

/** One character's outline (curves flattened). */
function glyphShape(f: Font, text: string, x: number, y: number, size: number, c: string): IconShape {
  const polys: number[][] = [];
  let cur: number[] = [], px = 0, py = 0;
  for (const cmd of f.getPath(text, x, y, size).commands) {
    if (cmd.type === 'M') { if (cur.length > 4) polys.push(cur); cur = [cmd.x!, cmd.y!]; }
    else if (cmd.type === 'L') cur.push(cmd.x!, cmd.y!);
    else if (cmd.type === 'Q' || cmd.type === 'C') {
      for (let k = 1; k <= 8; k++) {
        const t = k / 8, u = 1 - t;
        if (cmd.type === 'Q') cur.push(u * u * px + 2 * u * t * cmd.x1! + t * t * cmd.x!, u * u * py + 2 * u * t * cmd.y1! + t * t * cmd.y!);
        else cur.push(u ** 3 * px + 3 * u * u * t * cmd.x1! + 3 * u * t * t * cmd.x2! + t ** 3 * cmd.x!, u ** 3 * py + 3 * u * u * t * cmd.y1! + 3 * u * t * t * cmd.y2! + t ** 3 * cmd.y!);
      }
    } else if (cmd.type === 'Z') { if (cur.length > 4) polys.push(cur); cur = []; }
    if (cmd.x !== undefined) { px = cmd.x; py = cmd.y!; }
  }
  if (cur.length > 4) polys.push(cur);
  return { k: 'glyphs', polys, c };
}

/** A gallery tile's shapes: the icon (scaled into the top) and its label (lines split at "\n"). */
export function tileShapes(icon: string, label: string): IconShape[] {
  const k = TILE.icon / SIZE, dx = (TILE.w - TILE.icon) / 2;
  const scaled = (ICONS[icon] ?? []).map((s): IconShape => {
    const p = (pts: number[]) => pts.map((v, i) => (i % 2 ? v * k : v * k + dx));
    switch (s.k) {
      case 'rect': return { ...s, x: s.x * k + dx, y: s.y * k, w: s.w * k, h: s.h * k };
      case 'circle': return { ...s, cx: s.cx * k + dx, cy: s.cy * k, r: s.r * k };
      case 'sector': return { ...s, cx: s.cx * k + dx, cy: s.cy * k, r0: s.r0 * k, r1: s.r1 * k };
      case 'poly': return { ...s, pts: p(s.pts) };
      case 'line': return { ...s, pts: p(s.pts), w: s.w * k };
      case 'glyphs': return { ...s, polys: s.polys.map(p) };
    }
  });
  const lines = label.split('\n');
  const text = lines.flatMap((l, i) => textShapes(l, TILE.w / 2, TILE.icon + 2 + TILE.line * (i + 1) - 2.5, TILE.font, TILE_LABEL));
  return [...scaled, ...text];
}

/** A gallery tile as a PNG at `scale` pixels per point (2: Retina). */
export function tilePng(icon: string, label: string, scale = 2): Buffer {
  const w = TILE.w * scale, h = TILE.h * scale;
  return png(rasterShapes(tileShapes(icon, label), w, h, 1 / scale), w, h);
}

/** File name of an icon PNG: <icon>-<size>.png. */
export const iconFile = (name: string, px: number) => `${name}-${px}.png`;
