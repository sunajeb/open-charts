// The chart gallery's icons, drawn once as a handful of flat shapes on a 48 × 48 grid: rendered as
// inline SVG for the editor's gallery and rasterised (anti-aliased, zero dependencies) to the PNGs the
// PowerPoint ribbon shows (addin/icons/*.png, written by tools/icons.ts and checked by a test).
// Flat bright blues on a light-blue baseline read on light and dark ribbons alike.
import { deflateSync } from 'node:zlib';

export type IconShape =
  | { k: 'rect'; x: number; y: number; w: number; h: number; c: string }
  | { k: 'poly'; pts: number[]; c: string }
  | { k: 'circle'; cx: number; cy: number; r: number; c: string }
  /** Ring sector (r0 = 0: a pie wedge); angles in degrees clockwise from 12 o'clock. */
  | { k: 'sector'; cx: number; cy: number; r0: number; r1: number; a0: number; a1: number; c: string }
  /** Stroked polyline with round joins and caps. */
  | { k: 'line'; pts: number[]; w: number; c: string };

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
  }
}

/** RGBA pixels of an icon at `px` × `px`, transparent background, 4 × 4 supersampled. */
export function rasterize(name: string, px: number): Uint8Array {
  const shapes = ICONS[name] ?? [];
  const rgb = shapes.map((s) => [parseInt(s.c.slice(0, 2), 16), parseInt(s.c.slice(2, 4), 16), parseInt(s.c.slice(4, 6), 16)]);
  const out = new Uint8Array(px * px * 4);
  const SS = 4, scale = SIZE / px;
  for (let py = 0; py < px; py++) {
    for (let qx = 0; qx < px; qx++) {
      let r = 0, g = 0, b = 0, a = 0;
      for (let sy = 0; sy < SS; sy++) {
        for (let sx = 0; sx < SS; sx++) {
          const x = (qx + (sx + 0.5) / SS) * scale, y = (py + (sy + 0.5) / SS) * scale;
          // The topmost shape covering the sample wins (shapes are opaque, drawn in order).
          for (let k = shapes.length - 1; k >= 0; k--) {
            if (!inside(shapes[k], x, y)) continue;
            r += rgb[k][0]; g += rgb[k][1]; b += rgb[k][2]; a++;
            break;
          }
        }
      }
      const o = (py * px + qx) * 4;
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

/** An icon as a PNG file (8-bit RGBA). Deterministic: the same icon gives the same bytes. */
export function iconPng(name: string, px: number): Buffer {
  const rgba = rasterize(name, px);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(px, 0); ihdr.writeUInt32BE(px, 4); ihdr[8] = 8; ihdr[9] = 6;
  const raw = Buffer.alloc((px * 4 + 1) * px);
  for (let y = 0; y < px; y++) Buffer.from(rgba.buffer, y * px * 4, px * 4).copy(raw, y * (px * 4 + 1) + 1);
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw, { level: 9 })), chunk('IEND', Buffer.alloc(0))]);
}

/** Names of the PNG files the ribbon uses: <icon>-<size>.png. */
export const ICON_SIZES = [32, 64] as const;
export const iconFile = (name: string, px: number) => `${name}-${px}.png`;
