// A small reader for OpenType / TrueType font files (sfnt): what text measurement and font lookup
// need, without parsing outlines. Works on TrueType (glyf) and CFF / CFF2 fonts alike, on single
// fonts and collections (.ttc / .otc), and on variable fonts (advance widths at any `wght` from
// HVAR). Pure: bytes come through a `read(offset, length)` callback, so the Node side can read only
// the tables it needs from a 50 MB collection (src/fonts.ts) and the browser can pass a whole file.
import type { FontMetrics } from './text.ts';

/** Reads `length` bytes at `offset` (fewer at the end of the file). */
export type ReadAt = (offset: number, length: number) => Uint8Array;

export const bytesReader = (bytes: ArrayBuffer | Uint8Array): ReadAt => {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  return (offset, length) => u8.subarray(offset, Math.min(u8.length, offset + length));
};

const view = (b: Uint8Array) => new DataView(b.buffer, b.byteOffset, b.byteLength);
const tag = (d: DataView, o: number) => String.fromCharCode(d.getUint8(o), d.getUint8(o + 1), d.getUint8(o + 2), d.getUint8(o + 3));

/** What the font index stores about one face of a font file. */
export interface FaceInfo {
  /** Face number inside a collection (0 for a single font). */
  index: number;
  /** Family names (name ID 1, every language): what PowerPoint and GDI call the font. */
  families: string[];
  /** Typographic family names (name ID 16), when the font has them. */
  typoFamilies: string[];
  /** Full names (name ID 4), e.g. "Arial Bold". */
  fullNames: string[];
  /** Subfamily (name ID 2, English when present), e.g. "Bold Italic". */
  subfamily: string;
  /** OS/2 usWeightClass (400 regular, 700 bold). */
  weight: number;
  bold: boolean;
  italic: boolean;
  /** The weight axis of a variable font. */
  wght?: { min: number; def: number; max: number };
  outlines: 'truetype' | 'cff' | 'cff2';
}

interface Table { offset: number; length: number }
type Tables = Map<string, Table>;

/** Offsets of the table directories of every face in the file (one for a plain .ttf / .otf). */
export function faceOffsets(read: ReadAt): number[] {
  const h = view(read(0, 12));
  if (h.byteLength < 12) throw new Error('not a font file');
  const t = tag(h, 0);
  if (t === 'ttcf') {
    const n = h.getUint32(8);
    if (n > 10000) throw new Error('not a font file');
    const d = view(read(12, n * 4));
    return Array.from({ length: n }, (_, i) => d.getUint32(i * 4));
  }
  if (t === 'OTTO' || t === 'true' || h.getUint32(0) === 0x00010000) return [0];
  throw new Error('not an OpenType / TrueType font');
}

function tablesAt(read: ReadAt, at: number): Tables {
  const h = view(read(at, 12));
  const n = h.getUint16(4);
  const d = view(read(at + 12, n * 16));
  const out: Tables = new Map();
  for (let i = 0; i < n; i++) out.set(tag(d, i * 16), { offset: d.getUint32(i * 16 + 8), length: d.getUint32(i * 16 + 12) });
  return out;
}

const table = (read: ReadAt, t: Tables, name: string): DataView | null => {
  const e = t.get(name);
  if (!e) return null;
  const b = read(e.offset, e.length);
  return b.byteLength ? view(b) : null;
};

function utf16be(d: DataView, o: number, len: number): string {
  let s = '';
  for (let i = 0; i + 1 < len; i += 2) s += String.fromCharCode(d.getUint16(o + i));
  return s;
}
function latin1(d: DataView, o: number, len: number): string {
  let s = '';
  for (let i = 0; i < len; i++) s += String.fromCharCode(d.getUint8(o + i));
  return s;
}

/** Name records by name ID, English first. */
function names(d: DataView | null): Map<number, string[]> {
  const out = new Map<number, { s: string; en: boolean }[]>();
  if (!d) return new Map();
  const count = d.getUint16(2), strings = d.getUint16(4);
  for (let i = 0; i < count; i++) {
    const r = 6 + i * 12;
    if (r + 12 > d.byteLength) break;
    const platform = d.getUint16(r), encoding = d.getUint16(r + 2), lang = d.getUint16(r + 4), id = d.getUint16(r + 6);
    if (![1, 2, 4, 16, 17].includes(id)) continue;
    const len = d.getUint16(r + 8), off = strings + d.getUint16(r + 10);
    if (off + len > d.byteLength) continue;
    let s: string;
    if (platform === 0 || (platform === 3 && (encoding === 0 || encoding === 1 || encoding === 10))) s = utf16be(d, off, len);
    else if (platform === 1 && encoding === 0) s = latin1(d, off, len);
    else continue;
    s = s.replace(/\0/g, '').trim();
    if (!s) continue;
    const en = (platform === 3 && lang === 0x409) || (platform === 1 && lang === 0) || platform === 0;
    const list = out.get(id) ?? [];
    if (!list.some((x) => x.s === s)) list.push({ s, en });
    out.set(id, list);
  }
  return new Map([...out].map(([id, l]) => [id, [...l.filter((x) => x.en), ...l.filter((x) => !x.en)].map((x) => x.s)]));
}

const fixed = (d: DataView, o: number) => d.getInt32(o) / 65536;

function wghtAxis(d: DataView | null): FaceInfo['wght'] {
  if (!d) return undefined;
  const axesAt = d.getUint16(4), count = d.getUint16(8), size = d.getUint16(10);
  for (let i = 0; i < count; i++) {
    const o = axesAt + i * size;
    if (tag(d, o) === 'wght') return { min: fixed(d, o + 4), def: fixed(d, o + 8), max: fixed(d, o + 12) };
  }
  return undefined;
}

/** Names, style and weight of every face in a font file. */
export function readFaces(read: ReadAt): FaceInfo[] {
  return faceOffsets(read).map((at, index) => {
    const t = tablesAt(read, at);
    const n = names(table(read, t, 'name'));
    const os2 = table(read, t, 'OS/2'), head = table(read, t, 'head');
    const fsSel = os2 && os2.byteLength >= 64 ? os2.getUint16(62) : 0;
    const mac = head && head.byteLength >= 46 ? head.getUint16(44) : 0;
    const weight = os2 && os2.byteLength >= 6 ? os2.getUint16(4) : (mac & 1 ? 700 : 400);
    const subfamily = n.get(2)?.[0] ?? '';
    return {
      index, families: n.get(1) ?? [], typoFamilies: n.get(16) ?? [], fullNames: n.get(4) ?? [], subfamily, weight,
      bold: os2 ? !!(fsSel & 0x20) : !!(mac & 1),
      italic: os2 ? !!(fsSel & 0x01) : !!(mac & 2),
      ...(t.has('fvar') ? { wght: wghtAxis(table(read, t, 'fvar')) } : {}),
      outlines: t.has('CFF2') ? 'cff2' : t.has('CFF ') ? 'cff' : 'truetype',
    };
  });
}

// ---- Metrics ------------------------------------------------------------------------------------

/**
 * Character → glyph map, chosen and unrolled exactly as opentype.js 1.3.4 does it (the last Unicode
 * subtable of format 4 or 12), so widths match the shaper's to the bit; other formats are skipped.
 */
function cmapCodes(d: DataView): Map<number, number> {
  const n = d.getUint16(2);
  const ok = (p: number, e: number) => (p === 3 && (e === 0 || e === 1 || e === 10)) || (p === 0 && e <= 4);
  for (let i = n - 1; i >= 0; i--) {
    const p = d.getUint16(4 + i * 8), e = d.getUint16(6 + i * 8), off = d.getUint32(8 + i * 8);
    if (!ok(p, e) || off + 2 > d.byteLength) continue;
    const format = d.getUint16(off);
    const map = new Map<number, number>();
    if (format === 12) {
      const groups = d.getUint32(off + 12);
      for (let g = 0; g < groups; g++) {
        const r = off + 16 + g * 12;
        const start = d.getUint32(r), end = d.getUint32(r + 4);
        let gid = d.getUint32(r + 8);
        if (end - start > 0x10ffff) continue;
        for (let c = start; c <= end; c++) map.set(c, gid++);
      }
      return map;
    }
    if (format === 4) {
      const seg = d.getUint16(off + 6) >> 1;
      const ends = off + 14, starts = off + 16 + seg * 2, deltas = off + 16 + seg * 4, ranges = off + 16 + seg * 6;
      for (let s = 0; s < seg - 1; s++) {
        const end = d.getUint16(ends + s * 2), start = d.getUint16(starts + s * 2);
        const delta = d.getInt16(deltas + s * 2), range = d.getUint16(ranges + s * 2);
        for (let c = start; c <= end; c++) {
          let gid: number;
          if (range) {
            const at = ranges + s * 2 + range + (c - start) * 2;
            gid = at + 2 <= d.byteLength ? d.getUint16(at) : 0;
            if (gid) gid = (gid + delta) & 0xffff;
          } else gid = (c + delta) & 0xffff;
          map.set(c, gid);
        }
      }
      return map;
    }
  }
  return new Map();
}

/**
 * Single-spaced line height as a multiple of the font size, the way PowerPoint (GDI / DirectWrite)
 * computes it: typo metrics when OS/2 asks for them (USE_TYPO_METRICS), else win ascent + descent
 * plus whatever extra leading hhea.lineGap adds on top. Arial: (1854 + 434 + 67) / 2048 ≈ 1.1499.
 */
function lineHeight(em: number, hhea: DataView | null, os2: DataView | null): number {
  const asc = hhea ? hhea.getInt16(4) : 0, desc = hhea ? hhea.getInt16(6) : 0, gap = hhea ? hhea.getInt16(8) : 0;
  if (os2 && os2.byteLength >= 78) {
    if (os2.getUint16(62) & 0x80) return (os2.getInt16(68) - os2.getInt16(70) + os2.getInt16(72)) / em;
    const winAsc = os2.getUint16(74);
    if (winAsc) {
      const win = winAsc + os2.getUint16(76);
      return (win + (hhea ? Math.max(0, gap - (win - (asc - desc))) : 0)) / em;
    }
  }
  if (hhea) return (asc - desc + gap) / em;
  return 1.2;
}

// ---- Variable fonts: advance widths at a weight (HVAR, avar) -------------------------------------

const f2dot14 = (d: DataView, o: number) => d.getInt16(o) / 16384;

/** User `wght` → normalised coordinate (-1…1), through avar when the font has one. */
function normalise(v: number, axis: NonNullable<FaceInfo['wght']>, avar: DataView | null, axisIndex: number): number {
  let n = v === axis.def ? 0 : v < axis.def ? Math.max(-1, (v - axis.def) / (axis.def - axis.min || 1)) : Math.min(1, (v - axis.def) / (axis.max - axis.def || 1));
  if (avar) {
    let o = 8;
    const count = avar.getUint16(6);
    for (let a = 0; a < count && a <= axisIndex; a++) {
      const pairs = avar.getUint16(o);
      if (a === axisIndex && pairs) {
        const map = Array.from({ length: pairs }, (_, k) => [f2dot14(avar, o + 2 + k * 4), f2dot14(avar, o + 4 + k * 4)]);
        for (let k = 1; k < map.length; k++) {
          const [x0, y0] = map[k - 1], [x1, y1] = map[k];
          if (n <= x1) { n = x1 === x0 ? y1 : y0 + (n - x0) * (y1 - y0) / (x1 - x0); break; }
        }
      }
      o += 2 + pairs * 4;
    }
  }
  return n;
}

/** Advance-width deltas (font units) for each glyph at normalised coordinates `coords`, from HVAR. */
function hvarDeltas(hvar: DataView, coords: number[], numGlyphs: number): Float64Array {
  const out = new Float64Array(numGlyphs);
  const store = hvar.getUint32(4), mapAt = hvar.getUint32(8);
  const regionsAt = store + hvar.getUint32(store + 2);
  const axisCount = hvar.getUint16(regionsAt), regionCount = hvar.getUint16(regionsAt + 2);
  const scalars = Array.from({ length: regionCount }, (_, r) => {
    let s = 1;
    for (let a = 0; a < axisCount; a++) {
      const o = regionsAt + 4 + (r * axisCount + a) * 6;
      const start = f2dot14(hvar, o), peak = f2dot14(hvar, o + 2), end = f2dot14(hvar, o + 4);
      const v = coords[a] ?? 0;
      if (peak === 0 || start > peak || peak > end || (start < 0 && end > 0)) continue;
      if (v < start || v > end) return 0;
      if (v === peak) continue;
      s *= v < peak ? (v - start) / (peak - start) : (end - v) / (end - peak);
    }
    return s;
  });
  const dataCount = hvar.getUint16(store + 6);
  const data = Array.from({ length: dataCount }, (_, i) => store + hvar.getUint32(store + 8 + i * 4));
  // Delta-set index map (outer, inner) per glyph; none = outer 0, inner = glyph id.
  let index = (g: number) => [0, g];
  if (mapAt) {
    const format = hvar.getUint8(mapAt), entry = hvar.getUint8(mapAt + 1);
    const count = format === 0 ? hvar.getUint16(mapAt + 2) : hvar.getUint32(mapAt + 2);
    const first = mapAt + (format === 0 ? 4 : 6), size = ((entry >> 4) & 3) + 1, innerBits = (entry & 0xf) + 1;
    index = (g: number) => {
      const o = first + Math.min(g, count - 1) * size;
      let v = 0;
      for (let b = 0; b < size; b++) v = (v << 8) | hvar.getUint8(o + b);
      return [v >>> innerBits, v & ((1 << innerBits) - 1)];
    };
  }
  for (let g = 0; g < numGlyphs; g++) {
    const [outer, inner] = index(g);
    const at = data[outer];
    if (at === undefined) continue;
    const items = hvar.getUint16(at), words = hvar.getUint16(at + 2), regions = hvar.getUint16(at + 4);
    if (inner >= items) continue;
    const long = !!(words & 0x8000), wordCount = words & 0x7fff;
    const big = long ? 4 : 2, small = long ? 2 : 1;
    const rowSize = wordCount * big + (regions - wordCount) * small;
    let o = at + 6 + regions * 2 + inner * rowSize, delta = 0;
    for (let k = 0; k < regions; k++) {
      const region = hvar.getUint16(at + 6 + k * 2);
      const w = k < wordCount;
      const v = w ? (long ? hvar.getInt32(o) : hvar.getInt16(o)) : (long ? hvar.getInt16(o) : hvar.getInt8(o));
      o += w ? big : small;
      delta += v * (scalars[region] ?? 0);
    }
    out[g] = delta;
  }
  return out;
}

/**
 * Measuring metrics of one face: advance width per code point, line height, family. `wght` picks
 * an instance of a variable font (advances from HVAR); without it, the default instance.
 */
export function faceMetrics(read: ReadAt, index = 0, wght?: number): FontMetrics {
  const at = faceOffsets(read)[index];
  if (at === undefined) throw new Error(`no face ${index} in this font file`);
  const t = tablesAt(read, at);
  const head = table(read, t, 'head'), hhea = table(read, t, 'hhea'), hmtx = table(read, t, 'hmtx');
  const maxp = table(read, t, 'maxp'), cmap = table(read, t, 'cmap'), os2 = table(read, t, 'OS/2');
  if (!head || !hhea || !hmtx || !maxp || !cmap) throw new Error('font file lacks the tables measuring needs');
  const em = head.getUint16(18) || 2048;
  const numGlyphs = maxp.getUint16(4), numH = Math.max(1, hhea.getUint16(34));
  const advOf = (g: number) => (g >= numGlyphs ? 0 : hmtx.getUint16(Math.min(g, numH - 1) * 4));
  let deltas: Float64Array | null = null;
  const axis = t.has('fvar') ? wghtAxis(table(read, t, 'fvar')) : undefined;
  const hvar = table(read, t, 'HVAR');
  if (wght !== undefined && axis && hvar) {
    const fvar = table(read, t, 'fvar')!;
    const count = fvar.getUint16(8), size = fvar.getUint16(10), axesAt = fvar.getUint16(4);
    const coords = Array.from({ length: count }, (_, i) => tag(fvar, axesAt + i * size) === 'wght'
      ? normalise(Math.min(axis.max, Math.max(axis.min, wght)), axis, table(read, t, 'avar'), i) : 0);
    if (coords.some((c) => c !== 0)) deltas = hvarDeltas(hvar, coords, numGlyphs);
  }
  const adv = (g: number) => (deltas ? Math.max(0, Math.round(advOf(g) + (deltas[g] ?? 0))) : advOf(g));
  const map = cmapCodes(cmap);
  const codes = [...map.keys()].sort((a, b) => a - b);
  const n = names(table(read, t, 'name'));
  return {
    v: 1, family: n.get(1)?.[0], unitsPerEm: em, lineHeight: lineHeight(em, hhea, os2),
    notdef: adv(0), codes, advances: codes.map((c) => adv(map.get(c)!)),
  };
}

/**
 * One face of a collection as a standalone font file (for the opentype.js shaper, which reads
 * single fonts only). A plain font file comes back unchanged.
 */
export function extractFace(bytes: ArrayBuffer, index: number): ArrayBuffer {
  const read = bytesReader(bytes);
  const offsets = faceOffsets(read);
  if (offsets.length === 1 && offsets[0] === 0) return bytes;
  const at = offsets[index];
  const src = new Uint8Array(bytes);
  const h = view(read(at, 12));
  const n = h.getUint16(4);
  const entries = [...tablesAt(read, at)];
  const dirSize = 12 + n * 16;
  const size = entries.reduce((s, [, e]) => s + ((e.length + 3) & ~3), dirSize);
  const out = new Uint8Array(size);
  const d = new DataView(out.buffer);
  out.set(src.subarray(at, at + 12), 0);
  let pos = dirSize;
  const dir = view(read(at + 12, n * 16));
  for (let i = 0; i < n; i++) {
    out.set(src.subarray(at + 12 + i * 16, at + 12 + i * 16 + 16), 12 + i * 16);
    const off = dir.getUint32(i * 16 + 8), len = dir.getUint32(i * 16 + 12);
    out.set(src.subarray(off, off + len), pos);
    d.setUint32(12 + i * 16 + 8, pos);
    pos += (len + 3) & ~3;
  }
  return out.buffer;
}
