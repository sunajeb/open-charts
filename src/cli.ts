#!/usr/bin/env -S node --experimental-strip-types --no-warnings
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname } from 'node:path';
import { parseArgs } from 'node:util';
import type { DeckSpec } from './deck.ts';
import { ADDIN_USAGE } from './addin/usage.ts';
// Each command loads only the modules it needs (dynamic imports): every TypeScript file costs start-up time.

const USAGE = `Usage: open-charts <deck.json> [-o out/deck] [--font Arial.ttf --bold-font "Arial Bold.ttf"] [--svg] [--strict] [--report report.json]
       open-charts inspect <deck.pptx> [--json]
       open-charts import <deck.pptx> [-o deck.json]
       open-charts serve [--port 4173]
       open-charts addin build|install|uninstall|types   (PowerPoint add-in, see below)

Writes <out>.pptx with one native, editable chart per slide.
--svg also writes <out>-<n>.svg previews. Layout warnings are printed to stderr.
--strict exits with code 1 if any slide has a layout issue (the deck is still written).
--report writes the issues per slide as JSON.
inspect lists the charts in a .pptx: frames, roles, tag sizes, contract version.
import recovers the deck JSON from the charts' OC_SPEC tags.

${ADDIN_USAGE}`;

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    out: { type: 'string', short: 'o' },
    font: { type: 'string' },
    'bold-font': { type: 'string' },
    svg: { type: 'boolean', default: false },
    strict: { type: 'boolean', default: false },
    report: { type: 'string' },
    style: { type: 'string' },
    json: { type: 'boolean', default: false },
    help: { type: 'boolean', short: 'h' },
    port: { type: 'string', default: '4173' },
    'no-register': { type: 'boolean', default: false },
  },
});

const command = positionals[0];
if (values.help) {
  console.log(USAGE);
} else if (command === 'addin') {
  try {
    const { addinMain } = await import('./addin/cli.ts');
    await addinMain(positionals.slice(1), { out: values.out, register: !values['no-register'] });
  } catch (err) {
    fail((err as Error).message);
  }
} else if (command === 'serve') {
  const { serve } = await import('./server.ts');
  serve(Number(values.port), { font: values.font, bold: values['bold-font'] });
} else if (command === 'inspect' || command === 'import') {
  if (positionals.length !== 2) fail(USAGE);
  const file = positionals[1];
  try {
    const bytes = new Uint8Array(readFileSync(file));
    const { importDeck, readCharts } = await import('./ooxml/pptx-read.ts');
    if (command === 'inspect') {
      const { formatInspect } = await import('./inspect.ts');
      const slides = readCharts(bytes);
      console.log(values.json ? JSON.stringify(slides, null, 2) : formatInspect(file, slides));
    } else {
      const { deck, warnings } = importDeck(bytes);
      for (const w of warnings) console.error(`  ${w}`);
      if (!deck.slides.length) fail(`${file}: no open-charts charts found`);
      const json = `${JSON.stringify(deck, null, 2)}\n`;
      if (values.out) {
        mkdirSync(dirname(values.out), { recursive: true });
        writeFileSync(values.out, json);
        console.log(`Wrote ${values.out} (${deck.slides.length} chart${deck.slides.length === 1 ? '' : 's'})`);
      } else {
        process.stdout.write(json);
      }
    }
  } catch (err) {
    fail(`${file}: ${(err as Error).message}`);
  }
} else {
  await run();
}

/** The `--style` file, checked as untrusted input (size, keys, values); exits with every problem listed. */
async function loadStyle() {
  if (!values.style) return undefined;
  const { parseStyleFile } = await import('./core/style-file.ts');
  try {
    return parseStyleFile(readFileSync(values.style, 'utf8'), values.style);
  } catch (err) {
    fail((err as Error).message);
  }
}

function fail(msg: string): never {
  console.error(msg);
  process.exit(1);
}

async function run() {
  const [{ createMeasurer }, { buildDeck }, { buildReport }] = await Promise.all([import('./fonts.ts'), import('./deck.ts'), import('./report.ts')]);
  if (positionals.length !== 1) fail(USAGE);

  const input = JSON.parse(readFileSync(positionals[0], 'utf8'));
  const deck: DeckSpec = Array.isArray(input.slides) ? input : { slides: [{ chart: input }] };
  const out = values.out ?? positionals[0].replace(/\.json$/, '');
  mkdirSync(dirname(out), { recursive: true });

  const { pptx, slides } = buildDeck(deck, createMeasurer(values.font, values['bold-font']), await loadStyle());
  writeFileSync(`${out}.pptx`, pptx);
  console.log(`Wrote ${out}.pptx (${slides.length} slide${slides.length === 1 ? '' : 's'})`);

  slides.forEach((s, i) => {
    if (values.svg) writeFileSync(`${out}-${i + 1}.svg`, s.svg);
    for (const issue of s.issues) console.error(`  slide ${i + 1}: ${issue}`);
  });

  const report = buildReport(slides);
  if (values.report) {
    mkdirSync(dirname(values.report), { recursive: true });
    writeFileSync(values.report, `${JSON.stringify(report, null, 2)}\n`);
  }
  if (values.strict && !report.ok) fail(`--strict: ${report.issues} layout issue${report.issues === 1 ? '' : 's'}`);
}
