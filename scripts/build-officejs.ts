// Builds the Office.js PowerPoint add-in into dist/officejs/ (npm run build:officejs):
//   taskpane.html  the datasheet editor (src/web/editor.html) in ?host=office mode, + office.js + taskpane.js
//   taskpane.js    the task pane: engine, writer, PowerPoint glue (IIFE, esbuild)
//   engine.js      the engine + writer as an ES module for other browser hosts and the parity test
//   commands.html  the ribbon's function file (the one button only opens the pane)
//   manifest.xml   PowerPoint XML manifest pointing at --base (default https://localhost:4175)
//   fonts/         Liberation Sans Regular + Bold (SIL OFL, licence alongside) for text measurement
//   assets/        ribbon and add-in icons
// esbuild bundles for the browser platform, so any Node built-in reaching the engine fails the build.
import { build } from 'esbuild';
import { copyFileSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { editorHtml } from '../src/server.ts';
import { DEFAULT_BASE, manifestXml } from '../src/officejs/manifest.ts';

const ROOT = fileURLToPath(new URL('../', import.meta.url));
export const OUT = join(ROOT, 'dist', 'officejs');
const OFFICE_JS = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';

export async function buildOfficeJs(opts: { base?: string; out?: string; quiet?: boolean } = {}) {
  const out = opts.out ?? OUT;
  const base = opts.base ?? DEFAULT_BASE;
  rmSync(out, { recursive: true, force: true });
  mkdirSync(join(out, 'fonts'), { recursive: true });
  mkdirSync(join(out, 'assets'), { recursive: true });
  // `fs` only appears in opentype.js's Font.download() (Node branch), which the engine never calls.
  const common = { bundle: true, platform: 'browser' as const, target: ['es2022', 'safari16'], logLevel: 'warning' as const, legalComments: 'none' as const, minify: true, external: ['fs'] };
  await build({ ...common, entryPoints: [join(ROOT, 'src/officejs/taskpane.ts')], outfile: join(out, 'taskpane.js'), format: 'iife' });
  await build({ ...common, entryPoints: [join(ROOT, 'src/officejs/browser.ts')], outfile: join(out, 'engine.js'), format: 'esm' });

  // The editor page as the local server serves it (option table injected), plus office.js and the
  // task pane bundle ahead of the editor's own script, so window.OC_HOST exists when it runs.
  const html = editorHtml();
  const at = html.indexOf('<script');
  const head = `<script src="${OFFICE_JS}"></script>\n<script src="taskpane.js"></script>\n`;
  writeFileSync(join(out, 'taskpane.html'), html.slice(0, at) + head + html.slice(at));
  writeFileSync(join(out, 'commands.html'), `<!doctype html>\n<html><head><meta charset="utf-8"><title>open-charts commands</title>\n<script src="${OFFICE_JS}"></script>\n<script>Office.onReady(() => {});</script></head><body></body></html>\n`);
  writeFileSync(join(out, 'support.html'), `<!doctype html>\n<html><head><meta charset="utf-8"><title>open-charts add-in</title></head><body style="font:14px system-ui;max-width:640px;margin:2em auto">\n<h1>open-charts for PowerPoint</h1><p>Open the pane with <b>Home ▸ open-charts</b>. With nothing selected, <b>Insert on slide</b> draws the chart on the current slide. Select an open-charts chart to edit its data, then <b>Apply to slide</b>; <b>Revert</b> puts back the chart as it was before the last Apply.</p><p>All layout runs in this pane; no data leaves your computer. See USER-GUIDE.md, section "Office.js add-in".</p></body></html>\n`);
  writeFileSync(join(out, 'manifest.xml'), manifestXml(base));
  for (const f of ['LiberationSans-Regular.ttf', 'LiberationSans-Bold.ttf', 'LiberationSans.LICENSE']) copyFileSync(join(ROOT, 'vendor/fonts', f), join(out, 'fonts', f));
  for (const n of [16, 32, 64, 80, 128]) copyFileSync(join(ROOT, `src/officejs/assets/icon-${n}.png`), join(out, 'assets', `icon-${n}.png`));
  if (!opts.quiet) console.log(`open-charts Office.js add-in built in ${out} (manifest for ${base})`);
  return out;
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const i = process.argv.indexOf('--base');
  await buildOfficeJs({ base: i > 0 ? process.argv[i + 1] : undefined });
}
