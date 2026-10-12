import { test } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { breakPieces, wrap, fits } from '../src/core/text.ts';
import { createMeasurer } from '../src/fonts.ts';

const m = createMeasurer();

test('src/core has no file-system, network or DOM access', () => {
  const files: string[] = [];
  const walk = (d: string) => readdirSync(d, { withFileTypes: true }).forEach((e) =>
    e.isDirectory() ? walk(join(d, e.name)) : e.name.endsWith('.ts') && files.push(join(d, e.name)));
  walk('src/core');
  const banned = [
    /from\s+['"](node:)?(fs|fs\/promises|path|http|https|net|child_process|os|url|worker_threads)['"]/,
    /require\(/, /\bfetch\(/, /\bXMLHttpRequest\b/, /\bdocument\./, /\bwindow\./, /\bprocess\./, /\bimport\(/,
  ];
  for (const f of files) {
    const src = readFileSync(f, 'utf8');
    for (const re of banned) assert.ok(!re.test(src), `${f} matches ${re}`);
  }
  assert.ok(files.length >= 8);
});

test('B7: line height comes from the font tables, not a constant', () => {
  // Arial: (winAscent 1854 + winDescent 434 + extra lineGap 67) / 2048.
  assert.ok(Math.abs(m.lineHeight(10) - 10 * 2355 / 2048) < 1e-9);
  const georgia = '/System/Library/Fonts/Supplemental/Georgia.ttf';
  if (existsSync(georgia)) {
    const g = createMeasurer(georgia);
    assert.ok(Math.abs(g.lineHeight(10) - 10 * (1878 + 449) / 2048) < 1e-9);
    assert.equal(g.family, 'Georgia');
  }
});

test('line breaking: spaces, hyphens, slashes and CJK', () => {
  assert.deepEqual(breakPieces('North America').map((p) => p.s), ['North', 'America']);
  assert.deepEqual(breakPieces('Mid-market/SMB').map((p) => p.s), ['Mid-', 'market/', 'SMB']);
  assert.deepEqual(breakPieces('東京都').map((p) => p.s), ['東', '京', '都']);
  assert.deepEqual(breakPieces('Q1 (est.)').map((p) => p.s), ['Q1', '(est.)']);
  const lines = wrap('東京都港区六本木', 30, 12, m);
  assert.ok(lines.length > 1 && fits(lines, 30, 12, m), lines.join('|'));
  assert.deepEqual(wrap('Mid-market', m.width('Mid-', 12) + 1, 12, m), ['Mid-', 'market']);
  assert.equal(fits(wrap('Supercalifragilistic', 30, 12, m), 30, 12, m), false);
});
