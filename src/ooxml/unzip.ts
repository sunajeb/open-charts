// Minimal ZIP reader for OOXML packages (.pptx, .xlsx). Zero dependencies: node:zlib for deflate.
// Reads the central directory, so streamed entries (data descriptors) work. Guards against
// path traversal and zip bombs; ZIP64 and encryption are rejected with a clear message.
// The reader itself is pure (zipread.ts) so the browser bundle can use it with its own inflate.
import { inflateRawSync } from 'node:zlib';
import { readZip, type UnzipLimits, type ZipEntry } from './zipread.ts';

export type { UnzipLimits, ZipEntry } from './zipread.ts';

export function unzip(buf: Uint8Array, limits: UnzipLimits = {}): Map<string, ZipEntry> {
  return readZip(buf, (compressed, size) => new Uint8Array(inflateRawSync(compressed, { maxOutputLength: Math.max(size, 1) })), limits);
}
