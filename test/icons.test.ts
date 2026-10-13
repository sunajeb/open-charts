// The gallery icons: one drawing (src/addin/icons.ts) behind the editor's inline SVG and the ribbon's
// PNG tiles (icon over label); the committed PNGs (addin/icons) are exactly what tools/icons.ts writes.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import { ICONS, iconSvg, rasterize, TILE, tilePng, tileShapes } from '../src/addin/icons.ts';
import { CATALOGUE, galleryItems, iconOf } from '../src/addin/registry.ts';
import { ribbonImages } from '../src/addin/package.ts';

const dir = new URL('../addin/icons/', import.meta.url);

test('icons: every catalogue tile and direction button has its own icon', () => {
  for (const e of CATALOGUE) assert.ok(ICONS[e.icon ?? e.id], e.id);
  assert.deepEqual(new Set(CATALOGUE.map(iconOf)).size, CATALOGUE.length, 'no two tiles share an icon');
  for (const d of ['dirUp', 'dirDown', 'dirLeft', 'dirRight']) assert.ok(ICONS[d], d);
});

test('icons: the committed PNGs are what tools/icons.ts writes (run it after changing an icon or the catalogue)', () => {
  const images = ribbonImages(galleryItems());
  assert.deepEqual(readdirSync(dir).filter((f) => f.endsWith('.png')).sort(), images.map((i) => i.file).sort());
  for (const i of images) assert.ok(readFileSync(new URL(i.file, dir)).equals(i.draw()), i.file);
});

test('icons: a ribbon tile is the icon over its label, at 2x', () => {
  const shapes = tileShapes('stacked', 'Stacked\nColumn/Bar');
  const glyphs = shapes.filter((s) => s.k === 'glyphs');
  assert.equal(glyphs.length, 'StackedColumn/Bar'.length, 'one outline per character');
  const iconBottom = Math.max(...shapes.filter((s) => s.k === 'rect').map((s) => (s as { y: number; h: number }).y + (s as { h: number }).h));
  assert.ok(iconBottom <= TILE.icon + 0.01, 'the icon fits the top');
  const png = tilePng('stacked', 'Stacked\nColumn/Bar');
  assert.equal(png.readUInt32BE(16), TILE.w * 2);
  assert.equal(png.readUInt32BE(20), TILE.h * 2);
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
