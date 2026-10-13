// Golden rendering: examples/*.json and test/bench/*.json → .pptx (CLI) → PDF (LibreOffice) → PNG (pdftoppm) → diff against
// committed baselines in test/goldens/<platform>/. Font rendering differs per OS, so each platform keeps
// its own baselines.
//
//   npm run goldens                     compare; diffs go to out/goldens/diff/
//   npm run goldens -- --update         rewrite the baselines for this platform
//   npm run goldens -- --allow-missing  render and pass when this platform has no baselines yet (CI bootstrap)
import { execFileSync } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, renameSync, rmSync, writeFileSync } from 'node:fs';
import { basename, join } from 'node:path';
import { parseArgs } from 'node:util';
import { fileURLToPath } from 'node:url';
import { DEFAULT_DIFF, decodePng, diffImages, encodePng } from './png.ts';
import { unzip } from '../../src/ooxml/unzip.ts';

const root = fileURLToPath(new URL('../..', import.meta.url));
const { values: args } = parseArgs({
  options: {
    update: { type: 'boolean', default: false },
    'allow-missing': { type: 'boolean', default: false },
    platform: { type: 'string', default: process.platform },
    pixel: { type: 'string' },
    fraction: { type: 'string' },
    block: { type: 'string' },
  },
});
const limits = {
  pixel: args.pixel ? +args.pixel : DEFAULT_DIFF.pixel,
  fraction: args.fraction ? +args.fraction : DEFAULT_DIFF.fraction,
  block: args.block ? +args.block : DEFAULT_DIFF.block,
};
const goldenDir = join(root, 'test/goldens', args.platform!);
const out = join(root, 'out/goldens');
const actualDir = join(out, 'actual');
const diffDir = join(out, 'diff');
const soffice = process.env.SOFFICE ?? 'soffice';

rmSync(out, { recursive: true, force: true });
mkdirSync(actualDir, { recursive: true });
mkdirSync(diffDir, { recursive: true });

// 1. Build every example deck and every bench deck (test/bench/NN-*.json, pages "bench-NN-…") with the CLI.
const decks = [
  ...readdirSync(join(root, 'examples')).filter((f) => f.endsWith('.json')).sort().map((f) => ({ path: `examples/${f}`, name: basename(f, '.json') })),
  ...readdirSync(join(root, 'test/bench')).filter((f) => /^\d+-.*\.json$/.test(f)).sort().map((f) => ({ path: `test/bench/${f}`, name: `bench-${basename(f, '.json')}` })),
];
const expectedPages = new Map<string, number>();
for (const { path, name } of decks) {
  execFileSync(process.execPath, ['--experimental-strip-types', '--no-warnings', 'src/cli.ts', path, '-o', join(out, name)], { cwd: root, stdio: ['ignore', 'ignore', 'inherit'] });
  const spec = JSON.parse(readFileSync(join(root, path), 'utf8'));
  // Slides hidden in the slide show (agenda: a chapter with no content before its first subchapter)
  // are not in LibreOffice's PDF.
  const z = unzip(new Uint8Array(readFileSync(join(out, `${name}.pptx`))));
  const hidden = [...z.keys()].filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k) && /<p:sld [^>]*show="0"/.test(z.get(k)!.text())).length;
  expectedPages.set(name, (Array.isArray(spec.slides) ? spec.slides.length : 1) - hidden);
}
// Style files: the style showcase (one chart per type) under each examples/styles/*.json ("style-<name>").
const showcase = JSON.parse(readFileSync(join(root, 'examples/style-showcase.json'), 'utf8'));
for (const f of readdirSync(join(root, 'examples/styles')).filter((x) => x.endsWith('.json')).sort()) {
  const name = `style-${basename(f, '.json')}`;
  execFileSync(process.execPath, ['--experimental-strip-types', '--no-warnings', 'src/cli.ts', 'examples/style-showcase.json', '--style', `examples/styles/${f}`, '-o', join(out, name)], { cwd: root, stdio: ['ignore', 'ignore', 'inherit'] });
  expectedPages.set(name, showcase.slides.length);
}
// Template fill: the example template filled with the example data ("fill-template", 10 slides).
execFileSync(process.execPath, ['--experimental-strip-types', '--no-warnings', 'src/cli.ts', 'fill', 'examples/template/template.pptx', 'examples/template/data.json', '-o', join(out, 'fill-template.pptx')], { cwd: root, stdio: ['ignore', 'ignore', 'inherit'] });
expectedPages.set('fill-template', 10);

// 2. One LibreOffice run converts every deck (a private profile avoids clashing with a running instance).
const profile = `file://${join(out, 'lo-profile')}`;
execFileSync(soffice, [`-env:UserInstallation=${profile}`, '--headless', '--convert-to', 'pdf', '--outdir', out, ...[...expectedPages.keys()].map((n) => join(out, `${n}.pptx`))], { stdio: ['ignore', 'ignore', 'inherit'], timeout: 300_000 });

// 3. Rasterise at 60 dpi and normalise page names to <deck>-<page>.png.
const failures: string[] = [];
for (const [name, pages] of expectedPages) {
  const pdf = join(out, `${name}.pdf`);
  if (!existsSync(pdf)) { failures.push(`${name}: LibreOffice produced no PDF`); continue; }
  execFileSync('pdftoppm', ['-r', '60', '-png', pdf, join(out, `${name}-page`)]);
  const rendered = readdirSync(out).filter((f) => f.startsWith(`${name}-page-`) && f.endsWith('.png')).sort();
  if (rendered.length !== pages) failures.push(`${name}: ${rendered.length} pages rendered, deck has ${pages} slides`);
  for (const f of rendered) {
    const page = Number(f.slice(`${name}-page-`.length, -4));
    renameSync(join(out, f), join(actualDir, `${name}-${page}.png`));
  }
}

// 4. Compare (or update).
const actual = readdirSync(actualDir).filter((f) => f.endsWith('.png')).sort();
if (args.update) {
  rmSync(goldenDir, { recursive: true, force: true });
  mkdirSync(goldenDir, { recursive: true });
  for (const f of actual) copyFileSync(join(actualDir, f), join(goldenDir, f));
  console.log(`Updated ${actual.length} goldens in ${goldenDir}`);
} else if (!existsSync(goldenDir)) {
  const msg = `No goldens for platform "${args.platform}" (${goldenDir}). Rendered ${actual.length} pages to ${actualDir}; review them and commit with --update.`;
  if (args['allow-missing']) console.log(msg); else failures.push(msg);
} else {
  const golden = readdirSync(goldenDir).filter((f) => f.endsWith('.png')).sort();
  for (const f of golden.filter((g) => !actual.includes(g))) failures.push(`${f}: golden exists but the page was not rendered`);
  for (const f of actual) {
    if (!golden.includes(f)) { failures.push(`${f}: new page without a golden (run with --update)`); continue; }
    const r = diffImages(decodePng(readFileSync(join(goldenDir, f))), decodePng(readFileSync(join(actualDir, f))), limits);
    const stat = r.differing < 0 ? r.reason : `${(r.fraction * 100).toFixed(3)}% px, worst block ${r.worstBlock.toFixed(1)}`;
    console.log(`${r.ok ? 'ok  ' : 'FAIL'} ${f}  ${stat}`);
    if (r.diff) writeFileSync(join(diffDir, f), encodePng(r.diff));
    if (!r.ok) failures.push(`${f}: ${r.reason} — see ${join(diffDir, f)}`);
  }
}

if (failures.length) {
  console.error(`\n${failures.length} golden failure(s):\n${failures.map((f) => `  ${f}`).join('\n')}`);
  process.exit(1);
}
console.log(`\nGoldens OK (${actual.length} pages, platform ${args.platform}).`);
