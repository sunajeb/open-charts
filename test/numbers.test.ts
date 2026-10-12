import { test } from 'node:test';
import assert from 'node:assert/strict';
import { applyFormat, formatValue, inferDecimals, pickScale, roundToTotal } from '../src/core/numbers.ts';

const NB = ' ', NNB = ' ', M = '−';

test('B14: Excel format strings', () => {
  assert.equal(applyFormat(1234.56, '#,##0.0'), '1,234.6');
  assert.equal(applyFormat(1234.56, '$#,##0'), '$1,235');
  assert.equal(applyFormat(0.423, '0%'), '42%');
  assert.equal(applyFormat(0.4235, '0.0%'), '42.4%');
  assert.equal(applyFormat(-1234, '#,##0;(#,##0)'), '(1,234)');
  assert.equal(applyFormat(1234, '#,##0;(#,##0)'), '1,234');
  assert.equal(applyFormat(0, '#,##0;(#,##0);"–"'), '–');
  assert.equal(applyFormat(-5, '#,##0'), `${M}5`);
  assert.equal(applyFormat(1234567, '#,##0,"k"'), '1,235k');
  assert.equal(applyFormat(1234567, '#,##0.0,,"m"'), '1.2m');
  assert.equal(applyFormat(0.5, '#.00'), '.50');
  assert.equal(applyFormat(3.1, '0.0#'), '3.1');
  assert.equal(applyFormat(3.14159, '0.0#'), '3.14');
  assert.equal(applyFormat(7, '000'), '007');
  assert.equal(applyFormat(42, '[Red]0\\ "units"'), '42 units');
  assert.equal(applyFormat(-0.004, '0.0'), '0.0', 'no minus on a value that rounds to zero');
});

test('B14: locales', () => {
  assert.equal(formatValue(1234567.891, { decimals: 2, locale: 'de-DE' }), '1.234.567,89');
  assert.equal(formatValue(1234567.891, { decimals: 1, locale: 'fr-FR' }), `1${NNB}234${NNB}567,9`);
  assert.equal(formatValue(1234567.891, { decimals: 0, locale: 'en-IN' }), '12,34,568');
  assert.equal(formatValue(123456789, { decimals: 0, locale: 'en-IN' }), '12,34,56,789');
  assert.equal(formatValue(1234.5, { decimals: 1, locale: 'en-GB' }), '1,234.5');
  assert.equal(formatValue(42, { unit: '%', locale: 'de-DE' }), `42${NB}%`);
  assert.equal(formatValue(42, { unit: '%', locale: 'en-US' }), '42%');
  assert.equal(formatValue(0.42, { format: '0%', locale: 'de-DE' }), `42${NB}%`);
  assert.equal(formatValue(1234.5, { format: '#,##0.0', locale: 'de-DE' }), '1.234,5');
  assert.equal(formatValue(1, { decimals: 0, locale: 'de' }), '1', 'language-only falls back to de-DE');
});

test('B14: units as prefix or suffix, with optional space, and signs', () => {
  assert.equal(formatValue(42, { unit: { text: '$', position: 'prefix' } }), '$42');
  assert.equal(formatValue(-42, { unit: { text: '$', position: 'prefix' } }), `${M}$42`);
  assert.equal(formatValue(42, { unit: { text: '€', space: true } }), `42${NB}€`);
  assert.equal(formatValue(42, { unit: { text: '€', position: 'prefix', space: true } }), `€${NB}42`);
  assert.equal(formatValue(42, { unit: 'm' }), '42m');
  assert.equal(formatValue(4.2, { unit: '%', decimals: 1, signed: true }), '+4.2%');
  assert.equal(formatValue(-1234, { format: '#,##0;(#,##0)', unit: 'm' }), '(1,234)m');
  assert.equal(formatValue(5, { format: '0', signed: true }), '+5');
});

test('B14: magnitude scaling is shared per chart', () => {
  assert.deepEqual(pickScale([1200, 3_400_000], 'auto'), { div: 1e6, suffix: 'm' });
  assert.deepEqual(pickScale([12_000_000_000], 'auto'), { div: 1e9, suffix: 'bn' });
  assert.deepEqual(pickScale([999], 'auto'), { div: 1, suffix: '' });
  assert.deepEqual(pickScale([5], 'k'), { div: 1e3, suffix: 'k' });
  assert.equal(formatValue(1.2, { decimals: 1, magnitude: 'bn', unit: { text: '$', position: 'prefix' } }), '$1.2bn');
  assert.equal(formatValue(1.2, { decimals: 1, magnitude: 'm', unit: { text: '€', space: true } }), `1.2m${NB}€`);
});

test('B14: inferDecimals keeps two significant digits for small values', () => {
  assert.equal(inferDecimals([0.123]), 2);
  assert.equal(inferDecimals([0.0123, 1]), 3);
  assert.equal(inferDecimals([5.678]), 1);
  assert.equal(inferDecimals([123.456]), 0);
  assert.equal(inferDecimals([1.25]), 2);
});

test('B14: consistent rounding works with negatives', () => {
  const r = roundToTotal([1.4, 1.4, 1.4, -0.6], 0);
  assert.equal(r.reduce((a, b) => a + b), Math.round(1.4 * 3 - 0.6));
});
