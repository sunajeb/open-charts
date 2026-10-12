// Minimal MS-CFB (OLE compound file) writer and reader, version 3 (512-byte sectors).
// Enough for vbaProject.bin: nested storages, streams below 4096 bytes in the mini stream,
// larger ones in regular sectors, one or more FAT sectors listed in the header DIFAT (≤ 109).

export interface CfbStorage { [name: string]: Uint8Array | CfbStorage }

const SECTOR = 512, MINI = 64, CUTOFF = 4096;
const FREE = 0xffffffff, END = 0xfffffffe, FATSECT = 0xfffffffd, NOSTREAM = 0xffffffff;

interface Entry {
  name: string;
  type: 1 | 2 | 5; // storage, stream, root
  data?: Uint8Array;
  children: Entry[];
  left: number; right: number; child: number;
  start: number; size: number;
}

/** CFB sibling order: shorter names first, then case-insensitive code-unit order. */
export function compareNames(a: string, b: string): number {
  if (a.length !== b.length) return a.length - b.length;
  const ua = a.toUpperCase(), ub = b.toUpperCase();
  return ua < ub ? -1 : ua > ub ? 1 : 0;
}

export function writeCfb(root: CfbStorage): Uint8Array {
  const entries: Entry[] = [];
  const add = (name: string, type: Entry['type'], data?: Uint8Array): Entry => {
    const e: Entry = { name, type, data, children: [], left: NOSTREAM, right: NOSTREAM, child: NOSTREAM, start: END, size: 0 };
    entries.push(e);
    return e;
  };
  const walk = (parent: Entry, storage: CfbStorage) => {
    for (const [name, v] of Object.entries(storage)) {
      if (name.length > 31) throw new Error(`CFB name too long: ${name}`);
      const e = v instanceof Uint8Array ? add(name, 2, v) : add(name, 1);
      parent.children.push(e);
      if (!(v instanceof Uint8Array)) walk(e, v);
    }
  };
  const rootEntry = add('Root Entry', 5);
  walk(rootEntry, root);

  // Children of each storage form a binary search tree; a balanced tree coloured all black is valid.
  const index = new Map(entries.map((e, i) => [e, i]));
  const tree = (list: Entry[]): number => {
    if (!list.length) return NOSTREAM;
    const mid = list.length >> 1;
    const node = list[mid];
    node.left = tree(list.slice(0, mid));
    node.right = tree(list.slice(mid + 1));
    return index.get(node)!;
  };
  for (const e of entries) e.child = tree([...e.children].sort((a, b) => compareNames(a.name, b.name)));

  // Mini stream: every stream below the cutoff, 64-byte aligned.
  const miniFat: number[] = [];
  const miniParts: Uint8Array[] = [];
  const bigStreams: Entry[] = [];
  for (const e of entries) {
    if (e.type !== 2) continue;
    e.size = e.data!.length;
    if (e.size === 0) { e.start = END; continue; }
    if (e.size >= CUTOFF) { bigStreams.push(e); continue; }
    const n = Math.ceil(e.size / MINI);
    e.start = miniFat.length;
    for (let i = 0; i < n; i++) miniFat.push(i === n - 1 ? END : miniFat.length + 1);
    const padded = new Uint8Array(n * MINI);
    padded.set(e.data!);
    miniParts.push(padded);
  }
  const miniStream = concat(miniParts);

  // Sector plan: mini stream, big streams, mini FAT, directory, then FAT.
  const fat: number[] = [];
  const sectors: Uint8Array[] = [];
  const place = (data: Uint8Array): number => {
    if (!data.length) return END;
    const n = Math.ceil(data.length / SECTOR);
    const start = fat.length;
    for (let i = 0; i < n; i++) fat.push(i === n - 1 ? END : fat.length + 1);
    const padded = new Uint8Array(n * SECTOR);
    padded.set(data);
    sectors.push(padded);
    return start;
  };
  rootEntry.start = place(miniStream);
  rootEntry.size = miniStream.length;
  for (const e of bigStreams) e.start = place(e.data!);
  const miniFatBytes = u32s(miniFat, Math.ceil(miniFat.length / (SECTOR / 4)) * (SECTOR / 4), FREE);
  const miniFatStart = place(miniFatBytes);
  const miniFatSectors = miniFatBytes.length / SECTOR;

  const dirCount = Math.ceil(entries.length / 4) * 4;
  const dir = new Uint8Array(dirCount * 128);
  entries.forEach((e, i) => dirEntry(dir.subarray(i * 128, i * 128 + 128), e));
  for (let i = entries.length; i < dirCount; i++) {
    const v = new DataView(dir.buffer, i * 128, 128);
    v.setUint32(68, NOSTREAM, true); v.setUint32(72, NOSTREAM, true); v.setUint32(76, NOSTREAM, true);
  }
  const dirStart = place(dir);

  // FAT sectors also describe themselves.
  let fatSectors = 0;
  while (Math.ceil((fat.length + fatSectors) / (SECTOR / 4)) > fatSectors) fatSectors++;
  if (fatSectors > 109) throw new Error('CFB: file too large for header DIFAT');
  const fatStart = fat.length;
  for (let i = 0; i < fatSectors; i++) fat.push(FATSECT);
  sectors.push(u32s(fat, fatSectors * (SECTOR / 4), FREE));

  const header = new Uint8Array(SECTOR);
  const h = new DataView(header.buffer);
  header.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  h.setUint16(24, 0x003e, true);
  h.setUint16(26, 0x0003, true);
  h.setUint16(28, 0xfffe, true);
  h.setUint16(30, 9, true);
  h.setUint16(32, 6, true);
  h.setUint32(44, fatSectors, true);
  h.setUint32(48, dirStart, true);
  h.setUint32(56, CUTOFF, true);
  h.setUint32(60, miniFatSectors ? miniFatStart : END, true);
  h.setUint32(64, miniFatSectors, true);
  h.setUint32(68, END, true);
  for (let i = 0; i < 109; i++) h.setUint32(76 + i * 4, i < fatSectors ? fatStart + i : FREE, true);
  return concat([header, ...sectors]);
}

function dirEntry(b: Uint8Array, e: Entry) {
  const v = new DataView(b.buffer, b.byteOffset, 128);
  for (let i = 0; i < e.name.length; i++) v.setUint16(i * 2, e.name.charCodeAt(i), true);
  v.setUint16(64, (e.name.length + 1) * 2, true);
  b[66] = e.type;
  b[67] = 1; // black
  v.setUint32(68, e.left, true);
  v.setUint32(72, e.right, true);
  v.setUint32(76, e.child, true);
  v.setUint32(116, e.type === 1 ? 0 : e.start, true);
  v.setUint32(120, e.type === 1 ? 0 : e.size, true);
}

function u32s(values: number[], count: number, pad: number): Uint8Array {
  const out = new Uint8Array(count * 4);
  const v = new DataView(out.buffer);
  for (let i = 0; i < count; i++) v.setUint32(i * 4, i < values.length ? values[i] : pad, true);
  return out;
}

function concat(parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((a, p) => a + p.length, 0));
  let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
}

/** Reads a compound file back into nested storages (header DIFAT only, which is all writeCfb emits). */
export function readCfb(buf: Uint8Array): CfbStorage {
  const v = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  if (v.getUint32(0, true) !== 0xe011cfd0) throw new Error('Not a compound file');
  const ss = 1 << v.getUint16(30, true);
  const sector = (i: number) => buf.subarray(ss + i * ss, ss + (i + 1) * ss);
  const fat: number[] = [];
  for (let i = 0; i < Math.min(109, v.getUint32(44, true)); i++) {
    const s = sector(v.getUint32(76 + i * 4, true));
    const sv = new DataView(s.buffer, s.byteOffset, s.length);
    for (let j = 0; j < ss / 4; j++) fat.push(sv.getUint32(j * 4, true));
  }
  const chain = (start: number, table: number[], read: (i: number) => Uint8Array) => {
    const parts: Uint8Array[] = [];
    for (let s = start, guard = 0; s !== END && s !== FREE && guard < table.length + 1; s = table[s], guard++) parts.push(read(s));
    return concat(parts);
  };
  const dir = chain(v.getUint32(48, true), fat, sector);
  const miniFatBytes = v.getUint32(64, true) ? chain(v.getUint32(60, true), fat, sector) : new Uint8Array(0);
  const mfv = new DataView(miniFatBytes.buffer, miniFatBytes.byteOffset, miniFatBytes.length);
  const miniFat = Array.from({ length: miniFatBytes.length / 4 }, (_, i) => mfv.getUint32(i * 4, true));
  const dv = new DataView(dir.buffer, dir.byteOffset, dir.length);
  const at = (i: number) => ({
    name: String.fromCharCode(...Array.from({ length: Math.max(0, dv.getUint16(i * 128 + 64, true) / 2 - 1) }, (_, k) => dv.getUint16(i * 128 + k * 2, true))),
    type: dir[i * 128 + 66], left: dv.getUint32(i * 128 + 68, true), right: dv.getUint32(i * 128 + 72, true),
    child: dv.getUint32(i * 128 + 76, true), start: dv.getUint32(i * 128 + 116, true), size: dv.getUint32(i * 128 + 120, true),
  });
  const root = at(0);
  const mini = chain(root.start, fat, sector);
  const read = (e: ReturnType<typeof at>): Uint8Array => {
    if (e.size < 4096) return chain(e.start, miniFat, (i) => mini.subarray(i * 64, i * 64 + 64)).subarray(0, e.size);
    return chain(e.start, fat, sector).subarray(0, e.size);
  };
  const storage = (child: number): CfbStorage => {
    const out: CfbStorage = {};
    const visit = (i: number) => {
      if (i === NOSTREAM) return;
      const e = at(i);
      visit(e.left);
      out[e.name] = e.type === 2 ? read(e) : storage(e.child);
      visit(e.right);
    };
    visit(child);
    return out;
  };
  return storage(root.child);
}
