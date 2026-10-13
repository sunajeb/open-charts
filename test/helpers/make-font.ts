// Tiny synthetic font files for the font resolver's tests: just the tables measuring and lookup read
// (head, hhea, maxp, hmtx, cmap, name, OS/2; fvar + HVAR for a variable font), no outlines. Made in
// the test, so no font file (and no licence) is committed.

export interface FaceSpec {
  family: string;
  subfamily?: string;
  /** Typographic family (name ID 16). */
  typoFamily?: string;
  fullName?: string;
  weight?: number;
  bold?: boolean;
  italic?: boolean;
  /** Advance width (units of a 1000 em) of each character; others are 500, .notdef 600. */
  advances?: Record<string, number>;
  /** A variable font: wght axis, and the advance delta at the axis maximum for every glyph. */
  variable?: { min: number; def: number; max: number; delta: number };
  /** OS/2 win ascent / descent and hhea line gap (units of 1000). */
  ascent?: number; descent?: number; lineGap?: number;
  /** Outline flavour tag only ('OTTO' + an empty CFF table): the reader never parses outlines. */
  cff?: boolean;
}

const EM = 1000;

class W {
  bytes: number[] = [];
  u8(v: number) { this.bytes.push(v & 255); return this; }
  u16(v: number) { return this.u8(v >> 8).u8(v); }
  i16(v: number) { return this.u16(v < 0 ? v + 0x10000 : v); }
  u32(v: number) { return this.u16(Math.floor(v / 0x10000)).u16(v & 0xffff); }
  fixed(v: number) { return this.u32(Math.round(v * 65536) >>> 0); }
  tag(t: string) { for (const c of t) this.u8(c.charCodeAt(0)); return this; }
  pad(n: number) { while (this.bytes.length < n) this.u8(0); return this; }
}

function tables(f: FaceSpec): Map<string, number[]> {
  const chars = Object.keys(f.advances ?? {}).map((c) => c.codePointAt(0)!).sort((a, b) => a - b);
  for (const c of 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 .,-%()'.split('').map((c) => c.codePointAt(0)!)) if (!chars.includes(c)) chars.push(c);
  chars.sort((a, b) => a - b);
  const adv = (c: number) => f.advances?.[String.fromCodePoint(c)] ?? 500;
  const numGlyphs = chars.length + 1;
  const out = new Map<string, number[]>();
  const asc = f.ascent ?? 900, desc = f.descent ?? 250, gap = f.lineGap ?? 0;
  const weight = f.weight ?? (f.bold ? 700 : 400);

  const head = new W().u32(0x00010000).u32(0).u32(0).u32(0x5f0f3cf5).u16(0).u16(EM).pad(36).i16(0).i16(-desc).i16(EM).i16(asc)
    .u16((f.bold ? 1 : 0) | (f.italic ? 2 : 0)).u16(8).i16(2).i16(0).i16(0);
  out.set('head', head.bytes);
  out.set('hhea', new W().u32(0x00010000).i16(asc).i16(-desc).i16(gap).u16(EM).pad(34).u16(numGlyphs).bytes);
  out.set('maxp', new W().u32(0x00005000).u16(numGlyphs).bytes);
  const hmtx = new W().u16(600).i16(0);
  for (const c of chars) hmtx.u16(adv(c)).i16(0);
  out.set('hmtx', hmtx.bytes);
  // cmap: format 4, one segment per character plus the final 0xFFFF one.
  const seg = chars.length + 1;
  const sub = new W().u16(4).u16(16 + seg * 8).u16(0).u16(seg * 2).u16(0).u16(0).u16(0);
  chars.forEach((c) => sub.u16(c)); sub.u16(0xffff);
  sub.u16(0);
  chars.forEach((c) => sub.u16(c)); sub.u16(0xffff);
  chars.forEach((c, i) => sub.u16(((i + 1 - c) % 0x10000 + 0x10000) % 0x10000)); sub.u16(1); // idDelta: glyph i + 1
  chars.forEach(() => sub.u16(0)); sub.u16(0);
  out.set('cmap', [...new W().u16(0).u16(1).u16(3).u16(1).u32(12).bytes, ...sub.bytes]);
  // name: Windows, English.
  const names: [number, string][] = [[1, f.family], [2, f.subfamily ?? (f.bold ? (f.italic ? 'Bold Italic' : 'Bold') : f.italic ? 'Italic' : 'Regular')],
    [4, f.fullName ?? `${f.family}${f.bold || f.italic ? ` ${f.subfamily ?? [f.bold ? 'Bold' : '', f.italic ? 'Italic' : ''].filter(Boolean).join(' ')}` : ''}`]];
  if (f.typoFamily) names.push([16, f.typoFamily]);
  const strs = names.map(([, s]) => [...s].flatMap((ch) => [ch.charCodeAt(0) >> 8, ch.charCodeAt(0) & 255]));
  const name = new W().u16(0).u16(names.length).u16(6 + names.length * 12);
  let off = 0;
  names.forEach(([id], i) => { name.u16(3).u16(1).u16(0x409).u16(id).u16(strs[i].length).u16(off); off += strs[i].length; });
  strs.forEach((s) => s.forEach((b) => name.u8(b)));
  out.set('name', name.bytes);
  const fsSel = (f.italic ? 1 : 0) | (f.bold ? 0x20 : 0) | (!f.bold && !f.italic ? 0x40 : 0);
  out.set('OS/2', new W().u16(1).i16(500).u16(weight).pad(62).u16(fsSel).u16(32).u16(0xffff).i16(asc).i16(-desc).i16(gap).u16(asc).u16(desc).bytes);
  if (f.variable) {
    const v = f.variable;
    out.set('fvar', new W().u16(1).u16(0).u16(16).u16(2).u16(1).u16(20).u16(0).u16(8).tag('wght').fixed(v.min).fixed(v.def).fixed(v.max).u16(0).u16(256).bytes);
    // HVAR: one region (wght 0 → 1, peak 1), every glyph's advance grows by `delta` at the maximum.
    const store = new W().u16(1).u32(12).u16(1).u32(12 + 10);
    store.u16(1).u16(1).i16(0).i16(16384).i16(16384); // region list: 1 axis, 1 region
    store.u16(numGlyphs).u16(1).u16(1).u16(0); // item data: rows, 1 word delta, region 0
    for (let g = 0; g < numGlyphs; g++) store.i16(v.delta);
    out.set('HVAR', [...new W().u16(1).u16(0).u32(20).u32(0).u32(0).u32(0).bytes, ...store.bytes]);
  }
  if (f.cff) out.set('CFF ', [1, 0, 4, 1]);
  return out;
}

/** Lays out faces as one file: a plain font for one face, a collection (TTC) for several. */
export function makeFont(...faces: FaceSpec[]): Uint8Array {
  const sets = faces.map(tables);
  const collection = faces.length > 1;
  const headerSize = collection ? 12 + faces.length * 4 : 0;
  const dirSizes = sets.map((t) => 12 + t.size * 16);
  let pos = headerSize + dirSizes.reduce((a, b) => a + b, 0);
  const out = new W();
  if (collection) {
    out.tag('ttcf').u32(0x00010000).u32(faces.length);
    let at = headerSize;
    for (const d of dirSizes) { out.u32(at); at += d; }
  }
  const bodies: number[][] = [];
  sets.forEach((t, i) => {
    const tags = [...t.keys()].sort();
    out.u32(faces[i].cff ? 0x4f54544f : 0x00010000).u16(tags.length).u16(0).u16(0).u16(0);
    for (const tg of tags) {
      const data = t.get(tg)!;
      out.tag(tg).u32(0).u32(pos).u32(data.length);
      const padded = [...data]; while (padded.length % 4) padded.push(0);
      bodies.push(padded);
      pos += padded.length;
    }
  });
  for (const b of bodies) out.bytes.push(...b);
  return new Uint8Array(out.bytes);
}
