#!/usr/bin/env -S node --experimental-strip-types --no-warnings
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname } from 'node:path';
import { parseArgs } from 'node:util';
import { createMeasurer } from './core/text.ts';
import { buildDeck, type DeckSpec } from './deck.ts';

const USAGE = `Usage: open-charts <deck.json> [-o out/deck] [--font Arial.ttf --bold-font "Arial Bold.ttf"] [--svg]

Writes <out>.pptx with one native, editable chart per slide.
--svg also writes <out>-<n>.svg previews. Layout warnings are printed to stderr.`;

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    out: { type: 'string', short: 'o' },
    font: { type: 'string' },
    'bold-font': { type: 'string' },
    svg: { type: 'boolean', default: false },
    help: { type: 'boolean', short: 'h' },
  },
});

if (values.help || positionals.length !== 1) {
  console.log(USAGE);
  process.exit(values.help ? 0 : 1);
}

const input = JSON.parse(readFileSync(positionals[0], 'utf8'));
const deck: DeckSpec = Array.isArray(input.slides) ? input : { slides: [{ chart: input }] };
const out = values.out ?? positionals[0].replace(/\.json$/, '');
mkdirSync(dirname(out), { recursive: true });

const { pptx, slides } = buildDeck(deck, createMeasurer(values.font, values['bold-font']));
writeFileSync(`${out}.pptx`, pptx);
console.log(`Wrote ${out}.pptx (${slides.length} slide${slides.length === 1 ? '' : 's'})`);

slides.forEach((s, i) => {
  if (values.svg) writeFileSync(`${out}-${i + 1}.svg`, s.svg);
  for (const issue of s.issues) console.error(`  slide ${i + 1}: ${issue}`);
});
