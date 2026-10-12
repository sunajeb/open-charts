// Zero-dependency PNG decode/encode (8-bit greyscale, RGB, RGBA; non-interlaced) and an image diff.
import { deflateSync, inflateSync } from 'node:zlib';

export interface Image { width: number; height: number; data: Uint8Array } // RGBA, row-major

const SIG = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

export function decodePng(buf: Buffer): Image {
  if (!buf.subarray(0, 8).equals(SIG)) throw new Error('not a PNG');
  let p = 8, width = 0, height = 0, depth = 0, colour = 0, interlace = 0;
  const idat: Buffer[] = [];
  while (p < buf.length) {
    const len = buf.readUInt32BE(p);
    const type = buf.toString('latin1', p + 4, p + 8);
    const body = buf.subarray(p + 8, p + 8 + len);
    if (type === 'IHDR') {
      width = body.readUInt32BE(0); height = body.readUInt32BE(4);
      depth = body[8]; colour = body[9]; interlace = body[12];
    } else if (type === 'IDAT') idat.push(body);
    else if (type === 'IEND') break;
    p += 12 + len;
  }
  const channels = { 0: 1, 2: 3, 6: 4 }[colour];
  if (depth !== 8 || !channels || interlace) throw new Error(`unsupported PNG (depth ${depth}, colour type ${colour}, interlace ${interlace})`);
  const raw = inflateSync(Buffer.concat(idat));
  const stride = width * channels;
  const px = new Uint8Array(stride * height);
  for (let y = 0; y < height; y++) {
    const filter = raw[y * (stride + 1)];
    const src = y * (stride + 1) + 1, dst = y * stride;
    for (let x = 0; x < stride; x++) {
      const a = x >= channels ? px[dst + x - channels] : 0;
      const b = y > 0 ? px[dst - stride + x] : 0;
      const c = x >= channels && y > 0 ? px[dst - stride + x - channels] : 0;
      let v = raw[src + x];
      if (filter === 1) v += a;
      else if (filter === 2) v += b;
      else if (filter === 3) v += (a + b) >> 1;
      else if (filter === 4) {
        const pa = Math.abs(b - c), pb = Math.abs(a - c), pc = Math.abs(a + b - 2 * c);
        v += pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
      } else if (filter !== 0) throw new Error(`bad PNG filter ${filter}`);
      px[dst + x] = v;
    }
  }
  const data = new Uint8Array(width * height * 4);
  for (let i = 0; i < width * height; i++) {
    const s = i * channels;
    data[i * 4] = px[s];
    data[i * 4 + 1] = channels === 1 ? px[s] : px[s + 1];
    data[i * 4 + 2] = channels === 1 ? px[s] : px[s + 2];
    data[i * 4 + 3] = channels === 4 ? px[s + 3] : 255;
  }
  return { width, height, data };
}

const CRC = new Uint32Array(256).map((_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});
function crc32(buf: Buffer): number {
  let c = 0xffffffff;
  for (const b of buf) c = CRC[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}
function chunk(type: string, body: Buffer): Buffer {
  const out = Buffer.alloc(12 + body.length);
  out.writeUInt32BE(body.length, 0);
  out.write(type, 4, 'latin1');
  body.copy(out, 8);
  out.writeUInt32BE(crc32(out.subarray(4, 8 + body.length)), 8 + body.length);
  return out;
}

export function encodePng(img: Image): Buffer {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(img.width, 0); ihdr.writeUInt32BE(img.height, 4);
  ihdr[8] = 8; ihdr[9] = 6;
  const stride = img.width * 4;
  const raw = Buffer.alloc((stride + 1) * img.height);
  for (let y = 0; y < img.height; y++) Buffer.from(img.data.buffer, img.data.byteOffset + y * stride, stride).copy(raw, y * (stride + 1) + 1);
  return Buffer.concat([SIG, chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

export interface DiffOptions {
  /** A pixel differs when any channel differs by more than this (0–255). Absorbs anti-aliasing noise. */
  pixel: number;
  /** Fail when more than this fraction of pixels differ. */
  fraction: number;
  /** Fail when any 8×8 block has a mean absolute difference above this (0–255). Catches small, solid changes. */
  block: number;
}
export const DEFAULT_DIFF: DiffOptions = { pixel: 48, fraction: 0.005, block: 24 };

export interface DiffResult { ok: boolean; reason: string; differing: number; fraction: number; worstBlock: number; diff?: Image }

/** Compare two images; returns a red-on-grey diff image when they differ. */
export function diffImages(a: Image, b: Image, o: DiffOptions = DEFAULT_DIFF): DiffResult {
  if (a.width !== b.width || a.height !== b.height) {
    return { ok: false, reason: `size ${b.width}×${b.height} ≠ golden ${a.width}×${a.height}`, differing: -1, fraction: 1, worstBlock: 255 };
  }
  const { width: w, height: h } = a;
  const diff = new Uint8Array(w * h * 4);
  const delta = new Float32Array(w * h);
  let differing = 0;
  for (let i = 0; i < w * h; i++) {
    let d = 0, sum = 0;
    for (let k = 0; k < 3; k++) { const x = Math.abs(a.data[i * 4 + k] - b.data[i * 4 + k]); d = Math.max(d, x); sum += x; }
    delta[i] = sum / 3;
    const grey = 255 - (255 - (a.data[i * 4] + a.data[i * 4 + 1] + a.data[i * 4 + 2]) / 3) * 0.25;
    if (d > o.pixel) { differing++; diff.set([255, 0, 0, 255], i * 4); } else diff.set([grey, grey, grey, 255], i * 4);
  }
  let worstBlock = 0;
  for (let by = 0; by + 8 <= h; by += 8) {
    for (let bx = 0; bx + 8 <= w; bx += 8) {
      let s = 0;
      for (let y = by; y < by + 8; y++) for (let x = bx; x < bx + 8; x++) s += delta[y * w + x];
      worstBlock = Math.max(worstBlock, s / 64);
    }
  }
  const fraction = differing / (w * h);
  const reasons = [];
  if (fraction > o.fraction) reasons.push(`${(fraction * 100).toFixed(2)}% pixels differ (limit ${(o.fraction * 100).toFixed(2)}%)`);
  if (worstBlock > o.block) reasons.push(`an 8×8 block differs by ${worstBlock.toFixed(1)}/255 on average (limit ${o.block})`);
  return { ok: !reasons.length, reason: reasons.join('; ') || 'match', differing, fraction, worstBlock, diff: differing ? { width: w, height: h, data: diff } : undefined };
}
