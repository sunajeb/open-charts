import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import { decodePng, diffImages, encodePng, type Image } from './goldens/png.ts';

function solid(w: number, h: number, rgb: [number, number, number]): Image {
  const data = new Uint8Array(w * h * 4);
  for (let i = 0; i < w * h; i++) data.set([...rgb, 255], i * 4);
  return { width: w, height: h, data };
}

test('PNG encode → decode round-trips', () => {
  const img = solid(17, 9, [10, 200, 30]);
  img.data.set([255, 0, 0, 128], (4 * 17 + 3) * 4);
  const back = decodePng(encodePng(img));
  assert.equal(back.width, 17);
  assert.equal(back.height, 9);
  assert.deepEqual(back.data, img.data);
});

test('decodes the committed pdftoppm goldens (RGB, all filter types)', () => {
  const dir = new URL('./goldens/darwin/', import.meta.url);
  const files = readdirSync(dir).filter((f) => f.endsWith('.png'));
  assert.ok(files.length > 0);
  const img = decodePng(readFileSync(new URL(files[0], dir)));
  assert.ok(Math.abs(img.width - 800) <= 1, `width ${img.width}`); // 13.33 in × 60 dpi
  assert.ok(Math.abs(img.height - 450) <= 1, `height ${img.height}`);
  assert.deepEqual([...img.data.subarray(0, 4)], [255, 255, 255, 255]); // white slide corner
  assert.ok(diffImages(img, img).ok);
});

test('diff ignores anti-aliasing noise but catches a moved or recoloured block', () => {
  const a = solid(64, 64, [255, 255, 255]);
  const noisy = structuredClone(a);
  for (let i = 0; i < 64 * 64; i += 7) noisy.data[i * 4] = 230; // faint, scattered
  assert.ok(diffImages(a, noisy).ok);

  const moved = structuredClone(a);
  for (let y = 8; y < 16; y++) for (let x = 8; x < 16; x++) moved.data.set([31, 58, 95, 255], (y * 64 + x) * 4);
  const r = diffImages(a, moved);
  assert.equal(r.ok, false); // 64 px of 4096 = 1.6 %, and one solid block
  assert.equal(r.differing, 64);
  assert.ok(r.diff);

  assert.match(diffImages(a, solid(32, 64, [255, 255, 255])).reason, /size/);
});
