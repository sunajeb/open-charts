import { test } from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildReport, parseIssue } from '../src/report.ts';

const CLI = fileURLToPath(new URL('../src/cli.ts', import.meta.url));
const EXAMPLE = fileURLToPath(new URL('../examples/deck.json', import.meta.url));
const run = (...args: string[]) =>
  spawnSync(process.execPath, ['--experimental-strip-types', '--no-warnings', CLI, ...args], { encoding: 'utf8' });

const n = 40;
const crowded = {
  type: 'stacked',
  categories: Array.from({ length: n }, (_, i) => `Category number ${i}`),
  series: [{ name: 'A', values: Array.from({ length: n }, (_, i) => i + 0.123) }, { name: 'B', values: Array.from({ length: n }, () => 1.5) }],
};

test('report: issues are parsed into kind/detail per slide', () => {
  assert.deepEqual(parseIssue('overlap: a and b'), { kind: 'overlap', detail: 'a and b' });
  const r = buildReport([{ title: 'A', issues: [] }, { issues: ['hidden: label x'] }]);
  assert.deepEqual(r, { ok: false, issues: 1, slides: [{ slide: 1, title: 'A', issues: [] }, { slide: 2, issues: [{ kind: 'hidden', detail: 'label x' }] }] });
});

test('cli: build → inspect → import round trip, --report and --strict', () => {
  const dir = mkdtempSync(join(tmpdir(), 'oc-cli-'));
  try {
    const out = join(dir, 'deck');
    let r = run(EXAMPLE, '-o', out, '--report', join(dir, 'report.json'));
    assert.equal(r.status, 0, r.stderr);
    const report = JSON.parse(readFileSync(join(dir, 'report.json'), 'utf8'));
    assert.equal(report.slides.length, 4);
    assert.equal(report.issues, report.slides.reduce((a: number, s: { issues: unknown[] }) => a + s.issues.length, 0));

    writeFileSync(join(dir, 'clean.json'), JSON.stringify({ type: 'clustered', categories: ['2023', '2024', '2025'], series: [{ name: 'Revenue', values: [40, 55, 70] }] }));
    r = run(join(dir, 'clean.json'), '-o', join(dir, 'clean'), '--strict');
    assert.equal(r.status, 0, r.stderr);

    r = run('inspect', `${out}.pptx`);
    assert.equal(r.status, 0, r.stderr);
    assert.match(r.stdout, /4 slides, 4 charts \(tag contract 1\)/);
    assert.match(r.stdout, /frame 48,96,864,408/);
    r = run('inspect', `${out}.pptx`, '--json');
    assert.equal(JSON.parse(r.stdout)[0].charts[0].contract, '1');

    r = run('import', `${out}.pptx`, '-o', join(dir, 'back.json'));
    assert.equal(r.status, 0, r.stderr);
    assert.deepEqual(JSON.parse(readFileSync(join(dir, 'back.json'), 'utf8')), JSON.parse(readFileSync(EXAMPLE, 'utf8')));

    writeFileSync(join(dir, 'crowded.json'), JSON.stringify(crowded));
    r = run(join(dir, 'crowded.json'), '-o', join(dir, 'crowded'), '--report', join(dir, 'r2.json'));
    assert.equal(r.status, 0);
    const r2 = JSON.parse(readFileSync(join(dir, 'r2.json'), 'utf8'));
    assert.ok(r2.issues > 0, 'crowded chart should report issues');
    r = run(join(dir, 'crowded.json'), '-o', join(dir, 'crowded'), '--strict');
    assert.equal(r.status, 1);
    assert.match(r.stderr, /--strict: \d+ layout issue/);

    r = run('import', join(dir, 'crowded.json'));
    assert.equal(r.status, 1);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});
