// Minimal ZIP writer — enough for OOXML packages, no dependencies. New parts are stored
// uncompressed; parts copied from another package (see src/ooxml/unzip.ts) keep their original
// compressed bytes, so anything we did not touch comes out byte-identical.

// CRC-32, slicing by 4: four 256-entry tables, four bytes per step (the bytewise loop was the
// single largest cost of writing a deck).
const CRC = (() => {
  const t = new Uint32Array(1024);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  for (let n = 0; n < 256; n++) for (let k = 1; k < 4; k++) t[k * 256 + n] = (t[(k - 1) * 256 + n] >>> 8) ^ t[t[(k - 1) * 256 + n] & 0xff];
  return t;
})();

export function crc32(buf: Uint8Array): number {
  let c = 0xffffffff;
  let i = 0;
  const n4 = buf.length & ~3;
  for (; i < n4; i += 4) {
    c ^= buf[i] | (buf[i + 1] << 8) | (buf[i + 2] << 16) | (buf[i + 3] << 24);
    c = CRC[768 + (c & 0xff)] ^ CRC[512 + ((c >>> 8) & 0xff)] ^ CRC[256 + ((c >>> 16) & 0xff)] ^ CRC[c >>> 24];
  }
  for (; i < buf.length; i++) c = CRC[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

/** A part copied verbatim from an existing zip: compression method, CRC and sizes as found. */
export interface RawEntry { method: number; crc: number; size: number; compressed: Uint8Array }

export type ZipInput = { name: string; data: string | Uint8Array } | { name: string; raw: RawEntry };

export function zip(files: ZipInput[]): Uint8Array {
  const enc = new TextEncoder();
  const chunks: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;

  for (const f of files) {
    const name = enc.encode(f.name);
    let data: Uint8Array, crc: number, size: number, method = 0;
    if ('raw' in f) {
      ({ method, crc, size } = f.raw);
      data = f.raw.compressed;
    } else {
      data = typeof f.data === 'string' ? enc.encode(f.data) : f.data;
      crc = crc32(data);
      size = data.length;
    }

    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true);
    local.setUint16(4, 20, true);
    local.setUint16(8, method, true);
    local.setUint32(14, crc, true);
    local.setUint32(18, data.length, true);
    local.setUint32(22, size, true);
    local.setUint16(26, name.length, true);
    chunks.push(new Uint8Array(local.buffer), name, data);

    const cen = new DataView(new ArrayBuffer(46));
    cen.setUint32(0, 0x02014b50, true);
    cen.setUint16(4, 20, true);
    cen.setUint16(6, 20, true);
    cen.setUint16(10, method, true);
    cen.setUint32(16, crc, true);
    cen.setUint32(20, data.length, true);
    cen.setUint32(24, size, true);
    cen.setUint16(28, name.length, true);
    cen.setUint32(42, offset, true);
    central.push(new Uint8Array(cen.buffer), name);

    offset += 30 + name.length + data.length;
  }

  const cenSize = central.reduce((a, c) => a + c.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, files.length, true);
  end.setUint16(10, files.length, true);
  end.setUint32(12, cenSize, true);
  end.setUint32(16, offset, true);

  const parts = [...chunks, ...central, new Uint8Array(end.buffer)];
  const out = new Uint8Array(parts.reduce((a, p) => a + p.length, 0));
  let pos = 0;
  for (const p of parts) { out.set(p, pos); pos += p.length; }
  return out;
}
