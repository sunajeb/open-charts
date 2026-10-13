// Live tests in real PowerPoint for Mac (macOS only, never in CI; the owner approves live windows):
//   npm run live -- --list                 the scenarios
//   npm run live -- --dry-run all          what would run (touches nothing)
//   npm run live -- insert-all resize      some scenarios, one PowerPoint session
//   npm run live -- --install all          everything (= npm run e2e:powerpoint)
// Short independent scenarios, each with a fresh presentation and a time budget; a PowerPoint call
// that does not answer in time reads and dismisses PowerPoint's dialog, reports it and ends the
// scenario; the run stops at the first failed scenario unless --keep-going. Holds
// /private/tmp/open-charts-ppt.lock for the whole run, refuses to start beside the owner's open
// presentations (--allow-open-docs), and leaves PowerPoint as found. Report: out/live/report.{json,md}.
import { mkdirSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  acquireLock, formatDryRun, formatList, parseArgs, preflight, Refused, reportOk, runScenarios, selectScenarios, USAGE, writeReport,
  type Options, type Report, type ScenarioDef,
} from './core.ts';
import { SCENARIOS, type Ctx } from './scenarios.ts';

const ROOT = fileURLToPath(new URL('../../', import.meta.url));
const OUT = join(ROOT, 'out', 'live');

async function main(argv: string[]): Promise<number> {
  let o: Options, picked: ScenarioDef[];
  try {
    o = parseArgs(argv);
    if (o.help) { console.log(`${USAGE}\n\nScenarios:\n${formatList(SCENARIOS)}`); return 0; }
    if (o.list) { console.log(formatList(SCENARIOS)); return 0; }
    picked = selectScenarios(SCENARIOS, o);
  } catch (e) {
    console.error((e as Error).message);
    return 2;
  }
  if (o.dryRun) { console.log(formatDryRun(picked, o)); return 0; }
  if (process.platform !== 'darwin') { console.log('The live scenarios run on macOS with PowerPoint installed; skipped.'); return 0; }

  // Loaded only for a real run: the PowerPoint side and the add-in build.
  const { PowerPoint, installedPrefix } = await import('./powerpoint.ts');
  const { buildAddin } = await import('../../src/addin/package.ts');
  const pp = new PowerPoint();
  let release = () => {};
  let restoreVerbose = () => {};
  const started = new Date();
  try {
    release = acquireLock();
    // The scenarios read the add-in log's step lines, which are written only with verbose logging on.
    restoreVerbose = (await import('../../src/addin/server.ts')).verboseUntilRestored();
    pp.wasRunning = pp.running();
    const openDocs = pp.openDocuments();
    preflight({ running: pp.wasRunning, openDocs, allowOpenDocs: o.allowOpenDocs });
    pp.prefix = buildAddin().prefix;
    if (!picked.some((s) => s.name === 'install') && installedPrefix() !== pp.prefix) {
      throw new Refused(`The installed add-in is ${installedPrefix() ? 'not this build' : 'missing'}: run with --install.`);
    }
    // Leave PowerPoint as found, also on Ctrl+C.
    const cleanUp = () => { pp.closeOwned(); pp.quit(); restoreVerbose(); release(); };
    process.once('SIGINT', () => { console.error('\nInterrupted: closing what the run opened.'); try { cleanUp(); } finally { process.exit(130); } });

    const shared = new Map<string, unknown>();
    const scenarios = await runScenarios<Ctx>(picked, o.keepGoing, {
      load: async (def) => (await import(def.module)).run,
      context: (def, rec, deadline) => {
        const out = join(OUT, def.name);
        rmSync(out, { recursive: true, force: true });
        mkdirSync(out, { recursive: true });
        pp.deadline = deadline;
        if (def.name !== 'install') pp.ensureLoaded();
        return { ...rec, pp, out, root: ROOT, opts: o, shared };
      },
      after: () => { pp.deadline = Infinity; pp.closeOwned(); },
    });
    pp.deadline = Infinity;
    pp.closeOwned();
    if (pp.quit()) console.log('\nPowerPoint quit (the run started it).');
    else if (pp.wasRunning) console.log('\nPowerPoint left running (it was running before the run).');
    const finished = new Date();
    const report: Report = {
      started: started.toISOString(), finished: finished.toISOString(), ms: finished.getTime() - started.getTime(),
      ok: reportOk(scenarios), argv, powerpoint: { wasRunning: pp.wasRunning, openDocs }, scenarios,
    };
    const files = writeReport(OUT, report);
    console.log(`\n${scenarios.map((s) => `${s.status.toUpperCase().padEnd(7)} ${s.name}${s.status === 'not-run' ? '' : ` (${(s.ms / 1000).toFixed(1)} s)`}`).join('\n')}`);
    console.log(`Report: ${files.md}\n${report.ok ? 'OK' : 'FAILED'}`);
    return report.ok ? 0 : 1;
  } catch (e) {
    if (e instanceof Refused) { console.error(`Not started: ${e.message}`); return 2; }
    throw e;
  } finally {
    pp.deadline = Infinity;
    pp.closeOwned();
    pp.quit(); // no-op unless the run started PowerPoint and it is still running
    restoreVerbose();
    release();
  }
}

process.exitCode = await main(process.argv.slice(2));
