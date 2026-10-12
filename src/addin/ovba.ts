// MS-OVBA 2.4.1 "CompressedContainer": the LZ77 variant VBA uses for the dir stream and module source.
// compress() is greedy (longest match first); decompress() is the reference algorithm, used by tests
// and by anyone reading a vbaProject.bin back.

const CHUNK = 4096;

/** Copy-token field split for a token at `pos` bytes into the decompressed chunk (MS-OVBA 2.4.1.3.19.1). */
function split(pos: number) {
  let bitCount = 4;
  while (1 << bitCount < pos) bitCount++;
  const lengthMask = 0xffff >> bitCount;
  return { bitCount, lengthMask, maxLength: lengthMask + 3, maxOffset: 1 << bitCount };
}

function compressChunk(src: Uint8Array): number[] {
  const out: number[] = [];
  let pos = 0;
  while (pos < src.length) {
    const flagAt = out.length;
    out.push(0);
    let flags = 0;
    for (let bit = 0; bit < 8 && pos < src.length; bit++) {
      let bestLen = 0, bestOff = 0;
      if (pos > 0) {
        const { maxLength, maxOffset } = split(pos);
        const limit = Math.min(maxLength, src.length - pos);
        for (let cand = pos - 1; cand >= 0 && pos - cand <= maxOffset; cand--) {
          let len = 0;
          while (len < limit && src[cand + len] === src[pos + len]) len++;
          if (len > bestLen) { bestLen = len; bestOff = pos - cand; if (len === limit) break; }
        }
      }
      if (bestLen >= 3) {
        const { bitCount } = split(pos);
        const token = ((bestOff - 1) << (16 - bitCount)) | (bestLen - 3);
        out.push(token & 0xff, token >>> 8);
        flags |= 1 << bit;
        pos += bestLen;
      } else {
        out.push(src[pos++]);
      }
    }
    out[flagAt] = flags;
  }
  return out;
}

export function compress(data: Uint8Array): Uint8Array {
  const out: number[] = [0x01];
  for (let start = 0; start < data.length; start += CHUNK) {
    const chunk = data.subarray(start, start + CHUNK);
    const body = compressChunk(chunk);
    if (body.length <= CHUNK) {
      const header = 0xb000 | (body.length - 1); // flag=1, signature=0b011, size = total - 3
      out.push(header & 0xff, header >>> 8, ...body);
    } else if (chunk.length === CHUNK) {
      const header = 0x3000 | (CHUNK - 1); // raw chunk: exactly 4096 bytes
      out.push(header & 0xff, header >>> 8, ...chunk);
    } else {
      throw new Error('OVBA: incompressible final chunk'); // never happens for VBA source text
    }
  }
  return Uint8Array.from(out);
}

export function decompress(buf: Uint8Array): Uint8Array {
  if (buf[0] !== 0x01) throw new Error('OVBA: bad signature byte');
  const out: number[] = [];
  let p = 1;
  while (p < buf.length) {
    const header = buf[p] | (buf[p + 1] << 8);
    const size = (header & 0x0fff) + 3;
    if (((header >> 12) & 0x7) !== 0b011) throw new Error('OVBA: bad chunk signature');
    const end = Math.min(p + size, buf.length);
    const chunkStart = out.length;
    p += 2;
    if (!(header & 0x8000)) {
      for (let i = 0; i < CHUNK; i++) out.push(buf[p + i]);
      p += CHUNK;
      continue;
    }
    while (p < end) {
      const flags = buf[p++];
      for (let bit = 0; bit < 8 && p < end; bit++) {
        if (flags & (1 << bit)) {
          const token = buf[p] | (buf[p + 1] << 8);
          p += 2;
          const { bitCount, lengthMask } = split(out.length - chunkStart);
          const len = (token & lengthMask) + 3;
          const off = (token >>> (16 - bitCount)) + 1;
          for (let i = 0; i < len; i++) out.push(out[out.length - off]);
        } else {
          out.push(buf[p++]);
        }
      }
    }
  }
  return Uint8Array.from(out);
}
