// The live runner's logic that does not touch PowerPoint: arguments, scenario selection, running
// scenarios one after another (fail fast), the lock, the pre-flight decision and the report.
// Unit-tested in test/live-runner.test.ts; the PowerPoint side is tools/live/powerpoint.ts.
import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

export const LOCK = '/private/tmp/open-charts-ppt.lock';

/** One scenario: short, independent, starts from a fresh presentation. */
export interface ScenarioDef {
  name: string;
  /** One line for --list. */
  description: string;
  /** Wall-clock budget; every PowerPoint call is capped by what is left of it. */
  budgetSec: number;
  /** Module (relative to tools/live/) exporting `run(ctx)`; loaded only when the scenario runs. */
  module: string;
  /** Not part of `all` (install runs only when asked for: `install` or --install). */
  explicit?: boolean;
}

export interface Check { what: string; ok: boolean }
export type Status = 'pass' | 'fail' | 'skipped' | 'not-run';
export interface ScenarioResult {
  name: string;
  status: Status;
  ms: number;
  checks: Check[];
  timings: Record<string, number[]>;
  notes: string[];
  error?: string;
  /** Text of a PowerPoint dialog found (and dismissed) when a call timed out. */
  dialog?: string;
}

export interface Report {
  started: string;
  finished: string;
  ms: number;
  ok: boolean;
  argv: string[];
  powerpoint: { wasRunning: boolean; openDocs: string[] };
  scenarios: ScenarioResult[];
}

/** A PowerPoint call that did not answer in time (or the scenario's budget ran out). */
export class LiveTimeout extends Error {
  dialog: string;
  constructor(message: string, dialog = '') {
    super(message);
    this.name = 'LiveTimeout';
    this.dialog = dialog;
  }
}
/** Thrown by a scenario that cannot run here (reported as skipped, not failed). */
export class Skip extends Error {}
/** The run must not start (lock held, the owner's documents open, usage). Exit code 2. */
export class Refused extends Error {}

// ---------- arguments ----------

export interface Options {
  scenarios: string[];
  all: boolean;
  install: boolean;
  list: boolean;
  dryRun: boolean;
  keepGoing: boolean;
  allowOpenDocs: boolean;
  help: boolean;
  /** Decks for the fill scenario (default: the example template, filled). */
  decks: string[];
}

export const USAGE = `Usage: npm run live -- [options] <scenario...|all>
  --list            list the scenarios and exit
  --dry-run         print what would run; touches nothing (no PowerPoint, no lock, no files)
  --install         install the add-in first (otherwise the installed one is used; it must be current)
  --keep-going      after a failed scenario, run the next ones anyway (default: stop)
  --allow-open-docs run although PowerPoint has other presentations open (they are left alone)
  --deck <file>     a deck for the fill scenario (repeatable; default: the example template, filled)
Report: out/live/report.json and out/live/report.md. Exit code 0 pass, 1 failure, 2 refused.`;

export function parseArgs(argv: string[]): Options {
  const o: Options = { scenarios: [], all: false, install: false, list: false, dryRun: false, keepGoing: false, allowOpenDocs: false, help: false, decks: [] };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    switch (a) {
      case '--list': o.list = true; break;
      case '--dry-run': o.dryRun = true; break;
      case '--install': o.install = true; break;
      case '--keep-going': o.keepGoing = true; break;
      case '--allow-open-docs': o.allowOpenDocs = true; break;
      case '-h': case '--help': o.help = true; break;
      case '--deck': {
        const d = argv[++i];
        if (!d || d.startsWith('--')) throw new Refused('--deck needs a file');
        o.decks.push(d);
        break;
      }
      default:
        if (a.startsWith('-')) throw new Refused(`Unknown option ${a}\n${USAGE}`);
        if (a === 'all') o.all = true;
        else o.scenarios.push(a);
    }
  }
  return o;
}

/** The scenarios to run, in registry order, each once. */
export function selectScenarios(registry: ScenarioDef[], o: Pick<Options, 'scenarios' | 'all' | 'install'>): ScenarioDef[] {
  const unknown = o.scenarios.filter((n) => !registry.some((s) => s.name === n));
  if (unknown.length) throw new Refused(`Unknown scenario${unknown.length > 1 ? 's' : ''}: ${unknown.join(', ')}. Known: ${registry.map((s) => s.name).join(', ')}`);
  const want = new Set(o.scenarios);
  if (o.install) want.add('install');
  const picked = registry.filter((s) => want.has(s.name) || (o.all && !s.explicit));
  if (!picked.length) throw new Refused(`Name a scenario, or all.\n${USAGE}`);
  return picked;
}

export function formatList(registry: ScenarioDef[]): string {
  const w = Math.max(...registry.map((s) => s.name.length));
  return registry.map((s) => `${s.name.padEnd(w)}  ${String(s.budgetSec).padStart(3)} s  ${s.description}${s.explicit ? ' (not in all)' : ''}`).join('\n');
}

export function formatDryRun(picked: ScenarioDef[], o: Options, lock = LOCK): string {
  const total = picked.reduce((a, s) => a + s.budgetSec, 0);
  return [
    'Dry run: nothing is started, locked or written.',
    `Would take ${lock} for the whole run (refuse if it is held).`,
    o.allowOpenDocs ? 'Would run beside other open presentations (--allow-open-docs) and leave them alone.' : 'Would refuse if PowerPoint has presentations open (--allow-open-docs to run anyway).',
    o.install ? 'Would install the add-in first.' : 'Would use the installed add-in (refuse if it is not the current build; --install to update it).',
    `Would run, in this order (${o.keepGoing ? 'all of them' : 'stopping at the first failure'}):`,
    ...picked.map((s, i) => `  ${String(i + 1).padStart(2)}. ${s.name.padEnd(12)} ${String(s.budgetSec).padStart(3)} s  ${s.description}`),
    `Worst case ${Math.ceil(total / 60)} min. Report: out/live/report.json, out/live/report.md.`,
    'PowerPoint would be left as found: quit at the end only if the run started it.',
  ].join('\n');
}

// ---------- pre-flight ----------

/** Whether the run may start, given what PowerPoint has open. */
export function preflight(p: { running: boolean; openDocs: string[]; allowOpenDocs: boolean }): void {
  if (p.running && p.openDocs.length && !p.allowOpenDocs) {
    throw new Refused(`PowerPoint has ${p.openDocs.length} presentation(s) open: ${p.openDocs.join(', ')}. Close them, or run with --allow-open-docs (they are left alone).`);
  }
}

// ---------- lock ----------

const alive = (pid: number) => { try { process.kill(pid, 0); return true; } catch (e) { return (e as NodeJS.ErrnoException).code === 'EPERM'; } };

/**
 * Takes the PowerPoint lock (a directory, as the other live tools use it) for this process.
 * A lock whose owner process is gone is taken over. Returns the release function.
 */
export function acquireLock(path = LOCK, pid = process.pid, isAlive = alive): () => void {
  const owner = join(path, 'owner');
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      mkdirSync(path);
      writeFileSync(owner, `${pid} ${new Date().toISOString()} live runner\n`);
      return () => {
        try { if (readFileSync(owner, 'utf8').startsWith(`${pid} `)) rmSync(path, { recursive: true, force: true }); } catch { /* gone */ }
      };
    } catch (e) {
      if ((e as NodeJS.ErrnoException).code !== 'EEXIST') throw e;
      let info = '';
      try { info = readFileSync(owner, 'utf8').trim(); } catch { /* no owner file: someone else's lock */ }
      const holder = Number(info.split(' ')[0]);
      if (attempt === 0 && holder && !isAlive(holder)) { rmSync(path, { recursive: true, force: true }); continue; }
      throw new Refused(`${path} is held${info ? ` (${info})` : ' (by another live tool or the owner)'}: only one PowerPoint test at a time.`);
    }
  }
  throw new Refused(`could not take ${path}`);
}

// ---------- running ----------

export interface Recorder {
  check(ok: boolean, what: string): boolean;
  note(text: string): void;
  time(name: string, ms: number): void;
}

export interface RunHooks<C> {
  /** Builds the scenario's context (with the recorder) and sets the PowerPoint deadline. */
  context(def: ScenarioDef, rec: Recorder, deadline: number): C;
  load(def: ScenarioDef): Promise<(ctx: C) => unknown>;
  /** After each scenario, also after a failure (close its presentations). Errors become notes. */
  after?(def: ScenarioDef, ctx: C | undefined): void;
  log?(line: string): void;
  now?(): number;
}

export async function runScenarios<C>(picked: ScenarioDef[], keepGoing: boolean, hooks: RunHooks<C>): Promise<ScenarioResult[]> {
  const log = hooks.log ?? ((l: string) => console.log(l));
  const now = hooks.now ?? Date.now;
  const results: ScenarioResult[] = [];
  let stop = false;
  for (const def of picked) {
    const r: ScenarioResult = { name: def.name, status: 'not-run', ms: 0, checks: [], timings: {}, notes: [] };
    results.push(r);
    if (stop) continue;
    const rec: Recorder = {
      check: (ok, what) => { r.checks.push({ what, ok }); log(`  ${ok ? 'ok  ' : 'FAIL'} ${what}`); return ok; },
      note: (text) => { r.notes.push(text); log(`  ${text}`); },
      time: (name, ms) => { (r.timings[name] ??= []).push(ms); },
    };
    log(`\n> ${def.name} (budget ${def.budgetSec} s): ${def.description}`);
    const t0 = now();
    let ctx: C | undefined;
    try {
      const run = await hooks.load(def);
      ctx = hooks.context(def, rec, t0 + def.budgetSec * 1000);
      await run(ctx);
      r.status = r.checks.every((c) => c.ok) ? 'pass' : 'fail';
      if (!r.checks.length) { r.status = 'fail'; r.error = 'the scenario made no checks'; }
    } catch (e) {
      if (e instanceof Skip) { r.status = 'skipped'; r.notes.push(`skipped: ${e.message}`); log(`  skipped: ${e.message}`); }
      else {
        r.status = 'fail';
        r.error = (e as Error).message.split('\n').slice(0, 6).join('\n');
        if (e instanceof LiveTimeout && e.dialog) r.dialog = e.dialog;
        log(`  ERROR ${r.error}${r.dialog ? `\n  PowerPoint dialog (dismissed): ${r.dialog}` : ''}`);
      }
    }
    try { hooks.after?.(def, ctx); } catch (e) { r.notes.push(`clean-up: ${(e as Error).message.split('\n')[0]}`); }
    r.ms = now() - t0;
    log(`< ${def.name}: ${r.status.toUpperCase()} in ${(r.ms / 1000).toFixed(1)} s`);
    if (r.status === 'fail' && !keepGoing) stop = true;
  }
  return results;
}

// ---------- report ----------

export const reportOk = (scenarios: ScenarioResult[]) => scenarios.every((s) => s.status === 'pass' || s.status === 'skipped');

const median = (v: number[]) => [...v].sort((a, b) => a - b)[v.length >> 1];

export function formatReport(r: Report): string {
  const out = [
    `# Live PowerPoint run: ${r.ok ? 'PASS' : 'FAIL'}`,
    '',
    `${r.started} to ${r.finished} (${(r.ms / 1000).toFixed(1)} s). PowerPoint was ${r.powerpoint.wasRunning ? `already running${r.powerpoint.openDocs.length ? ` with ${r.powerpoint.openDocs.length} presentation(s) open` : ''}` : 'not running (started and quit by the run)'}.`,
    `Command: npm run live -- ${r.argv.join(' ')}`,
    '',
    '| scenario | result | time | checks |',
    '|---|---|---|---|',
    ...r.scenarios.map((s) => `| ${s.name} | ${s.status.toUpperCase()} | ${s.status === 'not-run' ? '' : `${(s.ms / 1000).toFixed(1)} s`} | ${s.checks.filter((c) => c.ok).length}/${s.checks.length} |`),
  ];
  for (const s of r.scenarios.filter((x) => x.status !== 'not-run')) {
    out.push('', `## ${s.name}: ${s.status.toUpperCase()}`);
    if (s.error) out.push('', `Error: ${s.error}`);
    if (s.dialog) out.push('', `PowerPoint dialog (dismissed): ${s.dialog}`);
    if (s.checks.length) out.push('', ...s.checks.map((c) => `- ${c.ok ? 'ok' : '**FAIL**'} ${c.what.replace(/\n/g, ' ')}`));
    const t = Object.entries(s.timings);
    if (t.length) out.push('', 'Timings:', ...t.map(([k, v]) => `- ${k}: ${v.length > 1 ? `median ${median(v)} ms, max ${Math.max(...v)} ms (n=${v.length})` : `${v[0]} ms`}`));
    if (s.notes.length) out.push('', 'Notes:', ...s.notes.map((n) => `- ${n.replace(/\n/g, '\n  ')}`));
  }
  return `${out.join('\n')}\n`;
}

export function writeReport(dir: string, r: Report): { json: string; md: string } {
  mkdirSync(dir, { recursive: true });
  const json = join(dir, 'report.json'), md = join(dir, 'report.md');
  writeFileSync(json, `${JSON.stringify(r, null, 2)}\n`);
  writeFileSync(md, formatReport(r));
  return { json, md };
}
