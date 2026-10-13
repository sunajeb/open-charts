// The gallery icons: one drawing (src/addin/icons.ts) behind the editor's inline SVG and the ribbon's
// PNGs; the committed PNGs (addin/icons) are exactly what tools/icons.ts writes.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import { ICONS, ICON_SIZES, iconFile, iconPng, iconSvg, rasterize } from '../src/addin/icons.ts';
import { CATALOGUE, iconOf } from '../src/addin/registry.ts';

const dir = new URL('../addin/icons/', import.meta.url);

test('icons: every catalogue tile and direction button has its own icon', () => {
  for (const e of CATALOGUE) assert.ok(ICONS[e.icon ?? e.id], e.id);
  assert.deepEqual(new Set(CATALOGUE.map(iconOf)).size, CATALOGUE.length, 'no two tiles share an icon');
  for (const d of ['dirUp', 'dirDown', 'dirLeft', 'dirRight']) assert.ok(ICONS[d], d);
});

test('icons: the committed PNGs are what tools/icons.ts writes (run it after changing an icon)', () => {
  const want = Object.keys(ICONS).flatMap((n) => ICON_SIZES.map((px) => iconFile(n, px))).sort();
  assert.deepEqual(readdirSync(dir).filter((f) => f.endsWith('.png')).sort(), want);
  for (const n of Object.keys(ICONS)) {
    for (const px of ICON_SIZES) assert.ok(readFileSync(new URL(iconFile(n, px), dir)).equals(iconPng(n, px)), iconFile(n, px));
  }
});

test('icons: rasterised anti-aliased on a transparent background; SVG uses the same shapes', () => {
  const px = rasterize('stacked', 32);
  const alpha = (x: number, y: number) => px[(y * 32 + x) * 4 + 3];
  assert.equal(alpha(0, 0), 0, 'transparent corner');
  assert.equal(alpha(6, 28), 255, 'solid inside the first column');
  const partial = [...Array(32 * 32).keys()].filter((i) => px[i * 4 + 3] > 0 && px[i * 4 + 3] < 255).length;
  assert.ok(partial > 0, 'edges are anti-aliased');
  const svg = iconSvg('pie');
  assert.match(svg, /^<svg class="gal-ico" viewBox="0 0 48 48" aria-hidden="true">(<path d="[^"]+" fill="#[0-9A-F]{6}"\/>){3}<\/svg>$/);
  assert.equal((iconSvg('stacked').match(/<rect/g) ?? []).length, ICONS.stacked.length);
});
