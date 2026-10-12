// The JSON examples in docs/DATA-FORMAT.md must stay valid, and the reference must cover every
// chart type and option the validator knows.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { CHART_TYPES, OPTION_KEYS, STYLE_KEYS, migrate, migrateDeck, validateDeck, validateSpec } from '../src/core/index.ts';
import { buildDeck, type DeckSpec } from '../src/deck.ts';
import { createMeasurer } from '../src/fonts.ts';
import { optionsTable } from '../scripts/options-table.ts';

const DOC = readFileSync(new URL('../docs/DATA-FORMAT.md', import.meta.url), 'utf8');
const examples = [...DOC.matchAll(/```json\n([\s\S]*?)```/g)].map((m) => m[1]);

test('DATA-FORMAT.md has JSON examples', () => {
  assert.ok(examples.length >= 10, `found ${examples.length} examples`);
});

const m = createMeasurer();
examples.forEach((text, i) => {
  const first = text.split('\n').find((l) => l.includes('"type"'))?.trim() ?? '';
  test(`DATA-FORMAT.md example ${i + 1} is valid and lays out (${first})`, () => {
    const json = JSON.parse(text);
    const isDeck = Array.isArray(json.slides);
    const errors = isDeck ? validateDeck(migrateDeck(json)) : validateSpec(migrate(json));
    assert.deepEqual(errors, []);
    const deck: DeckSpec = isDeck ? json : { slides: [{ chart: json }] };
    const { slides } = buildDeck(deck, m);
    assert.equal(slides.length, deck.slides.length);
  });
});

test('DATA-FORMAT.md options table is the one generated from the applicability table and defaults', () => {
  assert.ok(DOC.includes(`${optionsTable()}\n`), 'run scripts/options-table.ts and paste its output into DATA-FORMAT.md');
});

test('DATA-FORMAT.md documents every chart type, option and style key', () => {
  for (const t of CHART_TYPES) assert.match(DOC, new RegExp(`^\\| \`${t}\` \\|`, 'm'), `chart type ${t}`);
  for (const k of OPTION_KEYS) assert.match(DOC, new RegExp(`^#### \`${k}\`$`, 'm'), `option ${k}`);
  for (const k of STYLE_KEYS) assert.match(DOC, new RegExp(`^\\| \`${k}\` \\|`, 'm'), `style key ${k}`);
});
