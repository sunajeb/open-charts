// Contact sheet of one chart at several sizes, for looking at resize behaviour.
//   node --experimental-strip-types --no-warnings scripts/resize-sheet.ts <case-regex> [out dir]
// Writes <out>/<case>.svg (frames outlined in light grey). Convert to PNG with any SVG viewer, e.g.
//   qlmanage -t -s 1600 -o out/resize out/resize/*.svg
import { mkdirSync, writeFileSync } from 'node:fs';
import { layoutChart, type Scene } from '../src/core/index.ts';
import { createMeasurer } from '../src/fonts.ts';
import { sceneToSvg } from '../src/render/svg.ts';
import { allCases } from '../test/resize-matrix.ts';

const m = createMeasurer();
const [pattern = '.', out = 'out/resize'] = process.argv.slice(2);
const SIZES: [number, number][] = [[200, 120], [240, 130], [360, 200], [520, 300], [864, 408], [200, 500], [900, 120], [360, 560]];
mkdirSync(out, { recursive: true });
for (const c of allCases().filter((c) => new RegExp(pattern).test(c.name))) {
  const scenes: Scene[] = [];
  const outlines: string[] = [];
  let x = 10, y = 10, rowH = 0;
  const W = 1500;
  for (const [w, h] of SIZES) {
    if (x + w > W) { x = 10; y += rowH + 24; rowH = 0; }
    const frame = { x, y: y + 14, w, h };
    scenes.push(layoutChart(c.spec, frame, m));
    outlines.push(`<rect x="${x}" y="${y + 14}" width="${w}" height="${h}" fill="none" stroke="#DDDDDD" stroke-width="0.5"/><text x="${x}" y="${y + 9}" font-size="8" fill="#999">${w}×${h}</text>`);
    x += w + 20;
    rowH = Math.max(rowH, h + 14);
  }
  const svg = sceneToSvg(scenes, { w: W, h: y + rowH + 20 }).replace('</svg>', `${outlines.join('')}</svg>`);
  writeFileSync(`${out}/${c.name.replace(/[^\w#-]+/g, '_')}.svg`, svg);
}
