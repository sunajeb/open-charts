// The live PowerPoint runner's logic (tools/live/), without PowerPoint: arguments, scenario
// selection, fail-fast running, per-call timeouts and dialogs, the lock, pre-flight and the report.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  acquireLock, formatDryRun, formatReport, LiveTimeout, parseArgs, preflight, Refused, reportOk, runScenarios, selectScenarios, Skip, writeReport,
  type Recorder, type Report, type ScenarioDef,
} from '../tools/live/core.ts';
import { DIALOG_SCRIPT, isTimeout, PowerPoint, type Exec } from '../tools/live/powerpoint.ts';
import { SCENARIOS } from '../tools/live/scenarios.ts';

const ROOT = fileURLToPath(new URL('../', import.meta.url));
const tmp = () => mkdtempSync(join(tmpdir(), 'oc-live-'));
const def = (name: string, extra: Partial<ScenarioDef> = {}): ScenarioDef => ({ name, description: name, budgetSec: 60, module: `./${name}.ts`, ...extra });
const REG = [def('install', { explicit: true }), def('a'), def('b'), def('c')];

test('arguments: scenarios, all, flags, decks', () => {
  const o = parseArgs(['b', 'a', '--keep-going', '--allow-open-docs', '--deck', 'x.pptx', '--deck', 'y.pptx', '--install', '--dry-run']);
  assert.deepEqual(o.scenarios, ['b', 'a']);
  assert.ok(o.keepGoing && o.allowOpenDocs && o.install && o.dryRun && !o.all && !o.list);
  assert.deepEqual(o.decks, ['x.pptx', 'y.pptx']);
  assert.ok(parseArgs(['all']).all);
  assert.ok(parseArgs(['--list']).list);
  assert.throws(() => parseArgs(['--fast']), Refused);
  assert.throws(() => parseArgs(['--deck']), Refused);
  assert.throws(() => parseArgs(['--deck', '--list']), Refused);
});

test('selection: registry order, once each; all leaves install out unless --install', () => {
  const names = (o: Parameters<typeof selectScenarios>[1]) => selectScenarios(REG, o).map((s) => s.name);
  assert.deepEqual(names({ scenarios: ['c', 'a', 'c'], all: false, install: false }), ['a', 'c']);
  assert.deepEqual(names({ scenarios: [], all: true, install: false }), ['a', 'b', 'c']);
  assert.deepEqual(names({ scenarios: [], all: true, install: true }), ['install', 'a', 'b', 'c']);
  assert.deepEqual(names({ scenarios: ['b'], all: false, install: true }), ['install', 'b']);
  assert.deepEqual(names({ scenarios: ['install'], all: false, install: false }), ['install']);
  assert.deepEqual(names({ scenarios: [], all: false, install: true }), ['install']);
  assert.throws(() => names({ scenarios: ['nope', 'a'], all: false, install: false }), /Unknown scenario: nope/);
  assert.throws(() => names({ scenarios: [], all: false, install: false }), /Name a scenario/);
});

test('the registered scenarios: unique names, short budgets, every module exports run', async () => {
  const names = SCENARIOS.map((s) => s.name);
  assert.equal(new Set(names).size, names.length);
  for (const n of ['install', 'insert-all', 'edit-apply', 'resize', 'chart-tools', 'directions', 'agenda', 'popup', 'handles', 'theme', 'fill', 'cli-compare']) assert.ok(names.includes(n), n);
  for (const s of SCENARIOS) {
    assert.ok(s.budgetSec <= (s.name === 'install' ? 180 : 120), `${s.name} budget ${s.budgetSec} s`);
    const mod = await import(new URL(s.module, new URL('../tools/live/', import.meta.url)).href);
    assert.equal(typeof mod.run, 'function', `${s.module} exports run`);
  }
  assert.deepEqual(SCENARIOS.filter((s) => s.explicit).map((s) => s.name), ['install']);
});

/** Runs fake scenarios: each body gets the recorder. */
async function runFake(bodies: Record<string, (r: Recorder) => unknown>, keepGoing = false) {
  const after: string[] = [], lines: string[] = [];
  const results = await runScenarios(Object.keys(bodies).map((n) => def(n)), keepGoing, {
    load: async (d) => (rec: Recorder) => bodies[d.name](rec),
    context: (_d, rec) => rec,
    after: (d) => { after.push(d.name); if (d.name === 'boom-after') throw new Error('close failed'); },
    log: (l) => lines.push(l),
  });
  return { results, after, lines };
}

test('running: pass, fail, skip; the first failure stops the run; clean-up after every scenario', async () => {
  const { results, after } = await runFake({
    ok: (r) => { r.check(true, 'fine'); r.time('insert', 5); r.time('insert', 7); },
    skipped: () => { throw new Skip('not here'); },
    bad: (r) => { r.check(true, 'one'); r.check(false, 'two'); },
    later: (r) => r.check(true, 'never'),
  });
  assert.deepEqual(results.map((s) => s.status), ['pass', 'skipped', 'fail', 'not-run']);
  assert.deepEqual(results[0].timings, { insert: [5, 7] });
  assert.deepEqual(after, ['ok', 'skipped', 'bad']);
  assert.equal(reportOk(results), false);
  assert.equal(reportOk(results.slice(0, 2)), true);
});

test('running: --keep-going runs the rest; a timeout keeps the dialog; no checks is a failure', async () => {
  const { results } = await runFake({
    hung: () => { throw new LiveTimeout('no answer from PowerPoint within 30 s: run VB macro', 'Microsoft PowerPoint | The file is damaged. Repair?'); },
    empty: () => {},
    'boom-after': (r) => r.check(true, 'ok'),
  }, true);
  assert.deepEqual(results.map((s) => s.status), ['fail', 'fail', 'pass']);
  assert.match(results[0].error!, /within 30 s/);
  assert.match(results[0].dialog!, /Repair\?/);
  assert.match(results[1].error!, /no checks/);
  assert.deepEqual(results[2].notes, ['clean-up: close failed']);
});

/** A fake osascript: `answer(script)` returns or throws; records every call. */
function fakePowerPoint(answer: (script: string) => string, o: { log?: string } = {}) {
  const calls: { file: string; args: string[]; timeoutMs: number }[] = [];
  const exec: Exec = (file, args, timeoutMs) => {
    calls.push({ file, args, timeoutMs });
    if (file === 'pgrep') return '123';
    return answer(args[1]);
  };
  const pp = new PowerPoint({ exec, log: o.log ?? join(tmp(), 'addin.log') });
  return { pp, calls };
}
const timeoutError = () => Object.assign(new Error('spawnSync osascript ETIMEDOUT'), { code: 'ETIMEDOUT' });

test('a call that times out reads and dismisses PowerPoint\'s dialog and throws LiveTimeout', () => {
  const { pp, calls } = fakePowerPoint((s) => {
    if (s === DIALOG_SCRIPT) return 'Microsoft PowerPoint | PowerPoint found a problem with content. | ';
    throw timeoutError();
  });
  assert.throws(() => pp.tell('make new presentation', 20), (e: unknown) => {
    assert.ok(e instanceof LiveTimeout);
    assert.match(e.message, /within 20 s: make new presentation/);
    assert.equal(e.dialog, 'Microsoft PowerPoint | PowerPoint found a problem with content.');
    return true;
  });
  const call = calls.find((c) => c.args[1].includes('make new presentation'))!;
  assert.match(call.args[1], /^with timeout of 20 seconds/);
  assert.equal(call.timeoutMs, 25_000, 'the process is killed 5 s after the AppleScript timeout');
  assert.ok(isTimeout(Object.assign(new Error('x'), { signal: 'SIGTERM' })));
  assert.ok(isTimeout(new Error('execution error: AppleEvent timed out. (-1712)')));
  assert.ok(!isTimeout(new Error('execution error: (-18)')));
});

test('every call is capped by the scenario budget; a used-up budget fails without calling PowerPoint', () => {
  const { pp, calls } = fakePowerPoint(() => 'ok');
  pp.deadline = Date.now() + 10_500;
  pp.tell('get name of active presentation', 30);
  assert.match(calls.at(-1)!.args[1], /^with timeout of 10 seconds/);
  pp.deadline = Date.now() + 500;
  const n = calls.length;
  assert.throws(() => pp.tell('get name of active presentation'), /used up its time budget/);
  assert.equal(calls.length, n);
});

test('macros: retried briefly while the add-in compiles, never after; a timeout is not retried', () => {
  const dir = tmp(), log = join(dir, 'addin.log');
  writeFileSync(log, 'old line\n');
  let fails = 2;
  const { pp, calls } = fakePowerPoint((s) => {
    if (s === DIALOG_SCRIPT) return '';
    if (s.includes('run VB macro') && fails-- > 0) throw new Error('execution error: Microsoft PowerPoint got an error: (-18)');
    writeFileSync(log, 'old line\n2026 RESULT chart_line OK (12 ms)\n');
    return '';
  }, { log });
  pp.prefix = 'P1';
  pp.retryMs = 1;
  const r = pp.run('chart_line');
  assert.equal(r.result, '2026 RESULT chart_line OK (12 ms)');
  assert.ok(pp.ready);
  assert.equal(calls.filter((c) => c.args[1]?.includes('run VB macro macro name "P1_Run" list of parameters {"chart_line"}')).length, 3);
  assert.match(calls.filter((c) => c.args[1]?.includes('run VB macro')).at(-1)!.args[1], /with timeout of 60 seconds/, 'the cold first macro may take 60 s');
  fails = 1;
  assert.throws(() => pp.macro('Run', 'chart_line'), /-18/, 'once ready, an error is not retried');

  const hung = fakePowerPoint((s) => { if (s === DIALOG_SCRIPT) return ''; throw timeoutError(); }).pp;
  hung.prefix = 'P1';
  assert.throws(() => hung.macro('Run', 'x'), LiveTimeout);
});

test('PowerPoint is left as found: only the runner\'s presentations are closed, quit only if the run started it', () => {
  const { pp, calls } = fakePowerPoint((s) => (s.includes('get name of active presentation') ? 'Presentation7' : ''));
  pp.deadline = Infinity;
  assert.equal(pp.newPresentation(), 'Presentation7');
  pp.closeOwned();
  const closes = calls.filter((c) => c.args[1]?.includes('close presentation'));
  assert.equal(closes.length, 1);
  assert.match(closes[0].args[1], /close presentation "Presentation7" saving no/);
  assert.equal(pp.quit(), false, 'wasRunning defaults to true: never quit before the runner looked');
  assert.ok(!calls.some((c) => /\bquit\b/.test(c.args[1] ?? '')));
});

test('pre-flight: refuses beside the owner\'s open presentations unless allowed', () => {
  assert.throws(() => preflight({ running: true, openDocs: ['Board deck.pptx'], allowOpenDocs: false }), (e: unknown) => e instanceof Refused && /Board deck\.pptx/.test(e.message));
  preflight({ running: true, openDocs: ['Board deck.pptx'], allowOpenDocs: true });
  preflight({ running: true, openDocs: [], allowOpenDocs: false });
  preflight({ running: false, openDocs: [], allowOpenDocs: false });
});

test('lock: one holder at a time; a dead holder\'s lock is taken over; a foreign lock is respected', () => {
  const dir = tmp(), lock = join(dir, 'ppt.lock');
  const release = acquireLock(lock, 111, () => true);
  assert.match(readFileSync(join(lock, 'owner'), 'utf8'), /^111 /);
  assert.throws(() => acquireLock(lock, 222, () => true), (e: unknown) => e instanceof Refused && /held \(111 /.test(e.message));
  const takeover = acquireLock(lock, 333, (pid) => pid !== 111);
  assert.match(readFileSync(join(lock, 'owner'), 'utf8'), /^333 /);
  release(); // not 111's any more: left alone
  assert.ok(existsSync(lock));
  takeover();
  assert.ok(!existsSync(lock));
  mkdirSync(lock); // another live tool's lock (no owner file)
  assert.throws(() => acquireLock(lock, 444, () => false), /held \(by another live tool or the owner\)/);
});

test('report: JSON and Markdown with timings, pass/fail per check, errors and dialogs', () => {
  const report: Report = {
    started: '2026-10-13T10:00:00.000Z', finished: '2026-10-13T10:02:00.000Z', ms: 120_000, ok: false, argv: ['insert-all', 'resize'],
    powerpoint: { wasRunning: false, openDocs: [] },
    scenarios: [
      { name: 'insert-all', status: 'pass', ms: 41_000, checks: [{ what: 'insert line', ok: true }], timings: { insert: [900, 1100, 1000] }, notes: [] },
      { name: 'resize', status: 'fail', ms: 30_000, checks: [{ what: 'line x0.5', ok: false }], timings: {}, notes: ['n1'], error: 'no answer from PowerPoint within 30 s', dialog: 'Repair?' },
      { name: 'theme', status: 'not-run', ms: 0, checks: [], timings: {}, notes: [] },
    ],
  };
  const md = formatReport(report);
  assert.match(md, /^# Live PowerPoint run: FAIL/);
  assert.match(md, /\| insert-all \| PASS \| 41\.0 s \| 1\/1 \|/);
  assert.match(md, /\| resize \| FAIL \| 30\.0 s \| 0\/1 \|/);
  assert.match(md, /\| theme \| NOT-RUN \|  \| 0\/0 \|/);
  assert.match(md, /- insert: median 1000 ms, max 1100 ms \(n=3\)/);
  assert.match(md, /- \*\*FAIL\*\* line x0\.5/);
  assert.match(md, /PowerPoint dialog \(dismissed\): Repair\?/);
  assert.ok(!md.includes('## theme'));
  const dir = tmp();
  const files = writeReport(join(dir, 'live'), report);
  assert.deepEqual(JSON.parse(readFileSync(files.json, 'utf8')), report);
  assert.equal(readFileSync(files.md, 'utf8'), md);
});

test('--dry-run and --list touch nothing: no lock, no PowerPoint, no report', () => {
  const report = join(ROOT, 'out', 'live', 'report.json');
  const before = existsSync(report) ? statSync(report).mtimeMs : 0;
  const run = (...args: string[]) => spawnSync(process.execPath, ['--experimental-strip-types', '--no-warnings', 'tools/live/main.ts', ...args], { cwd: ROOT, encoding: 'utf8', env: { ...process.env, PATH: '/nonexistent' } });
  const dry = run('--dry-run', '--install', 'all');
  assert.equal(dry.status, 0, dry.stderr);
  assert.match(dry.stdout, /Dry run: nothing is started/);
  assert.match(dry.stdout, / 1\. install /);
  assert.match(dry.stdout, /12\. cli-compare /);
  const list = run('--list');
  assert.equal(list.status, 0);
  assert.equal(list.stdout.trim().split('\n').length, SCENARIOS.length);
  assert.equal(run('nope').status, 2);
  assert.equal(run().status, 2);
  assert.equal(existsSync(report) ? statSync(report).mtimeMs : 0, before);
  assert.match(formatDryRun(selectScenarios(SCENARIOS, { scenarios: ['resize'], all: false, install: false }), parseArgs(['resize'])), /installed add-in/);
});

test('the old live scripts are gone; npm scripts point at the runner', () => {
  const pkg = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8'));
  assert.match(pkg.scripts.live, /tools\/live\/main\.ts$/);
  assert.match(pkg.scripts['e2e:powerpoint'], /tools\/live\/main\.ts --install all$/);
  for (const f of ['e2e-powerpoint.ts', 'e2e-agenda.ts', 'e2e-directions.ts', 'e2e-popup-live.ts', 'e2e-handles-live.ts', 'fill-powerpoint.ts']) {
    assert.ok(!existsSync(join(ROOT, 'tools', f)), `tools/${f} is replaced by a scenario`);
  }
});
