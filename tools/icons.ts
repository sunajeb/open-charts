// Writes the ribbon's images (addin/icons/*.png) from src/addin/icons.ts: a tile per gallery item
// (icon over label), the gallery button's icon and the direction arrows. Zero dependencies: shapes and
// label glyphs are rasterised in-process. test/icons.test.ts fails when the committed PNGs differ
// from what this writes, so run it after changing an icon, a label or the catalogue:
//   node --experimental-strip-types --no-warnings tools/icons.ts
import { mkdirSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { ICON_DIR, ribbonImages } from '../src/addin/package.ts';
import { galleryItems } from '../src/addin/registry.ts';

mkdirSync(ICON_DIR, { recursive: true });
const images = ribbonImages(galleryItems());
const keep = new Set(images.map((i) => i.file));
for (const f of readdirSync(ICON_DIR)) if (f.endsWith('.png') && !keep.has(f)) rmSync(new URL(f, ICON_DIR));
for (const i of images) writeFileSync(new URL(i.file, ICON_DIR), i.draw());
console.log(`${images.length} images written to addin/icons/`);
