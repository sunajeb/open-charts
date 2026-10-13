// Minimal ZIP reader for OOXML packages (.pptx, .xlsx). Pure: the caller supplies raw-deflate
// decompression (node:zlib in unzip.ts; DecompressionStream in the browser).
// Reads the central directory, so streamed entries (data descriptors) work. Guards against
// path traversal and zip bombs; ZIP64 and encryption are rejected with a clear message.
import type { RawEntry } from '../render/zip.ts';

export interface ZipEntry {
  name: string;
  raw: RawEntry;
  /** Decompressed bytes (inflated on first call, then cached). */
  bytes(): Uint8Array;
  text(): string;
}

export interface UnzipLimits {
  /** Max total uncompressed bytes across all entries. Default 512 MB. */
  maxTotal?: number;
  /** Max uncompressed/compressed ratio per entry. Default 200. */
  maxRatio?: number;
}

const dec = new TextDecoder();

/** Inflates raw-deflate data whose uncompressed size is `size`. */
export type InflateRaw = (compressed: Uint8Array, size: number) => Uint8Array;

export function readZip(buf: Uint8Array, inflate: InflateRaw, limits: UnzipLimits = {}): Map<string, ZipEntry> {
  const maxTotal = limits.maxTotal ?? 512 * 1024 * 1024;
  const maxRatio = limits.maxRatio ?? 200;
  const v = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);

  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 22 - 0xffff); i--) {
    if (v.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('Not a zip file (no end-of-central-directory record)');
  const count = v.getUint16(eocd + 10, true);
  let p = v.getUint32(eocd + 16, true);
  if (count === 0xffff || p === 0xffffffff) throw new Error('ZIP64 packages are not supported');

  const out = new Map<string, ZipEntry>();
  let total = 0;
  for (let n = 0; n < count; n++) {
    if (p + 46 > buf.length || v.getUint32(p, true) !== 0x02014b50) throw new Error('Corrupt zip central directory');
    const flags = v.getUint16(p + 8, true);
    const method = v.getUint16(p + 10, true);
    const crc = v.getUint32(p + 16, true);
    const csize = v.getUint32(p + 20, true);
    const size = v.getUint32(p + 24, true);
    const nameLen = v.getUint16(p + 28, true);
    const extraLen = v.getUint16(p + 30, true);
    const commentLen = v.getUint16(p + 32, true);
    const localOff = v.getUint32(p + 42, true);
    const name = dec.decode(buf.subarray(p + 46, p + 46 + nameLen));
    p += 46 + nameLen + extraLen + commentLen;

    if (name.endsWith('/')) continue; // directory entry
    if (flags & 1) throw new Error(`Encrypted zip entry "${name}" is not supported`);
    if (name.startsWith('/') || name.includes('\\') || name.split('/').includes('..')) throw new Error(`Unsafe zip entry name "${name}"`);
    if (method !== 0 && method !== 8) throw new Error(`Unsupported compression method ${method} for "${name}"`);
    total += size;
    if (total > maxTotal) throw new Error('Zip expands beyond the size limit');
    if (csize > 0 && size / csize > maxRatio && size > 1 << 20) throw new Error(`Zip entry "${name}" exceeds the compression-ratio limit`);

    if (v.getUint32(localOff, true) !== 0x04034b50) throw new Error(`Corrupt local header for "${name}"`);
    const start = localOff + 30 + v.getUint16(localOff + 26, true) + v.getUint16(localOff + 28, true);
    const compressed = buf.subarray(start, start + csize);
    let cache: Uint8Array | undefined;
    const bytes = () => {
      if (!cache) {
        cache = method === 0 ? compressed : inflate(compressed, size);
        if (cache.length !== size) throw new Error(`Size mismatch in zip entry "${name}"`);
      }
      return cache;
    };
    out.set(name, { name, raw: { method, crc, size, compressed }, bytes, text: () => dec.decode(bytes()) });
  }
  return out;
}
