// Writes the ribbon's gallery icons (addin/icons/<icon>-32.png and -64.png) from src/addin/icons.ts.
// Zero dependencies: the shapes are rasterised in-process. test/icons.test.ts fails when the
// committed PNGs differ from what this writes, so run it after changing an icon:
//   node --experimental-strip-types --no-warnings tools/icons.ts
import { mkdirSync, writeFileSync } from 'node:fs';
import { ICONS, ICON_SIZES, iconFile, iconPng } from '../src/addin/icons.ts';

const dir = new URL('../addin/icons/', import.meta.url);
mkdirSync(dir, { recursive: true });
for (const name of Object.keys(ICONS)) {
  for (const px of ICON_SIZES) writeFileSync(new URL(iconFile(name, px), dir), iconPng(name, px));
}
console.log(`${Object.keys(ICONS).length * ICON_SIZES.length} icons written to addin/icons/`);
