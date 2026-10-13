// Smoke test of a built engine executable (CI's windows job; on a Mac, dist/host/open-charts):
//   node --experimental-strip-types --no-warnings scripts/smoke-exe.ts dist/win/open-charts.exe
// Runs it the ways the add-in does: lists the gallery, lays out the example deck, builds the add-in
// from its embedded VBA, then starts the add-in server (on a spare port, with its own home and
// resource folders), pings it, runs an engine request over HTTP like the VBA does, loads the editor
// page and stops it. Prints timings; exits 1 on the first failure.
import { execFileSync, spawn } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

const exe = resolve(process.argv[2] ?? 'dist/win/open-charts.exe');
const port = Number(process.env.OPEN_CHARTS_PORT) || 4199;
const dir = mkdtempSync(join(tmpdir(), 'oc-smoke-'));
const env = { ...process.env, LOCALAPPDATA: dir, OPEN_CHARTS_HOME: join(dir, 'home'), OPEN_CHARTS_CACHE: join(dir, 'cache'), OPEN_CHARTS_PORT: String(port) };
const base = `http://127.0.0.1:${port}`;

const step = (what: string, f: () => unknown) => {
  const t = Date.now();
  try {
    const r = f();
    console.log(`ok   ${what} (${Date.now() - t} ms)`);
    return r;
  } catch (err) {
    console.error(`FAIL ${what}: ${(err as Error).message}`);
    process.exit(1);
  }
};
const run = (...args: string[]) => execFileSync(exe, args, { env, encoding: 'utf8', timeout: 60_000, windowsHide: true });
const check = (ok: boolean, why: string) => { if (!ok) throw new Error(why); };

try {
  step('first start (unpacks resources): addin types', () => check(/stacked\s+Stacked\s+available/.test(run('addin', 'types')), 'no stacked in the gallery'));
  step('warm start: addin types', () => run('addin', 'types'));
  step('layout examples/deck.json', () => {
    run('examples/deck.json', '-o', join(dir, 'deck'));
    check(existsSync(join(dir, 'deck.pptx')), 'no deck.pptx');
  });
  step('addin build (embedded VBA source)', () => {
    run('addin', 'build', '-o', join(dir, 'open-charts.ppam'));
    check(readFileSync(join(dir, 'open-charts.ppam')).subarray(0, 2).toString() === 'PK', 'not a zip');
  });
  const server = spawn(exe, ['addin', 'serve'], { env, stdio: 'ignore', windowsHide: true });
  try {
    const t = Date.now();
    let up = false;
    for (let i = 0; i < 150 && !up; i++) {
      await new Promise((r) => setTimeout(r, 100));
      up = await fetch(`${base}/api/addin/ping`).then((r) => r.ok, () => false);
    }
    step(`addin serve answers /api/addin/ping after ${Date.now() - t} ms`, () => check(up, 'no answer within 15 s'));
    const insert = await fetch(`${base}/api/addin/run`, { method: 'POST', body: 'insert\nwaterfall\n48,96,864,408\n\n' }).then((r) => r.text());
    step('POST /api/addin/run insert waterfall', () => check(insert.startsWith('OC\t') && /\nS\t/.test(insert), insert.slice(0, 200)));
    const page = await fetch(`${base}/`).then((r) => r.text());
    step('editor page', () => check(page.includes('window.OC_SPEC'), 'no editor page'));
    await fetch(`${base}/api/addin/quit`, { method: 'POST' }).catch(() => undefined);
    const gone = await new Promise<boolean>((r) => { server.once('exit', () => r(true)); setTimeout(() => r(false), 5000); });
    step('POST /api/addin/quit stops it', () => check(gone, 'still running'));
  } finally {
    server.kill();
  }
  console.log('smoke test passed');
} finally {
  rmSync(dir, { recursive: true, force: true });
}
