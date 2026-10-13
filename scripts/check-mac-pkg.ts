// `npm run check:mac-pkg [dist/mac/open-charts-<version>.pkg]`: installs the Mac package into a
// temporary home folder the way Installer does, without Installer and without PowerPoint, and
// checks the result:
//   1. pkgutil --expand: per-user domain only, the payload's files, the scripts;
//   2. the payload unpacked into <tmp home>/Library/Application Support/open-charts and the
//      postinstall run as Installer runs it ($2 = that folder), with no Node on PATH;
//   3. every file where `open-charts addin install` puts it, the bridge script calling the engine
//      executable, the open-charts command and the engine itself working (scripts/smoke-exe.ts);
//   4. `open-charts addin uninstall` removing all of it.
// OPEN_CHARTS_NO_POWERPOINT=1 keeps PowerPoint and launchd out of it.
import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, readdirSync, readFileSync, realpathSync, rmSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = fileURLToPath(new URL('../', import.meta.url));
const version = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8')).version;
const pkg = resolve(process.argv[2] ?? join(ROOT, 'dist', 'mac', `open-charts-${version}.pkg`));
const tmp = realpathSync(mkdtempSync(join(tmpdir(), 'oc-pkg-'))); // /var is a link: the engine knows itself by its real path
const home = join(tmp, 'home');
const dest = join(home, 'Library', 'Application Support', 'open-charts');
const NO_NODE_PATH = '/usr/bin:/bin:/usr/sbin:/sbin';
const env = { PATH: NO_NODE_PATH, HOME: home, USER: process.env.USER ?? '', OPEN_CHARTS_NO_POWERPOINT: '1', LOCALAPPDATA: join(tmp, 'res'), OPEN_CHARTS_PORT: '4198' };

let failed = false;
function step(what: string, f: () => unknown) {
  const t = Date.now();
  try {
    f();
    console.log(`ok   ${what} (${Date.now() - t} ms)`);
  } catch (err) {
    console.error(`FAIL ${what}: ${(err as Error).message}`);
    failed = true;
  }
}
const check = (ok: unknown, why: string) => { if (!ok) throw new Error(why); };
const sh = (cmd: string, args: string[], e: NodeJS.ProcessEnv = env) => execFileSync(cmd, args, { encoding: 'utf8', env: e, timeout: 120_000 });

try {
  const x = join(tmp, 'expanded');
  step(`pkgutil --expand ${pkg}`, () => sh('pkgutil', ['--expand', pkg, x], process.env));
  const component = readdirSync(x).find((f) => f.endsWith('.pkg'))!;
  step('distribution: per-user only (no administrator password), arm64 + x86_64', () => {
    const d = readFileSync(join(x, 'Distribution'), 'utf8');
    check(/<domains enable_anywhere="false" enable_currentUserHome="true" enable_localSystem="false"\/>/.test(d), 'domains');
    check(/hostArchitectures="arm64,x86_64"/.test(d), 'architectures');
  });
  step('component: install location and scripts', () => {
    const info = readFileSync(join(x, component, 'PackageInfo'), 'utf8');
    check(info.includes('install-location="/Library/Application Support/open-charts"'), info.slice(0, 300));
    check(/<pkg-info [^>]*auth="none"/.test(info), 'asks for root authorization');
    check(/<relocate\/>/.test(info), 'a bundle is relocatable (Installer could update another copy of it)');
    for (const s of ['preinstall', 'postinstall']) check(statSync(join(x, component, 'Scripts', s)).mode & 0o100, `${s} not executable`);
  });
  step('payload into the temporary home', () => {
    sh('mkdir', ['-p', dest]);
    execFileSync('/bin/sh', ['-c', `gunzip -dc "$1" | (cd "$2" && cpio -i --quiet)`, 'sh', join(x, component, 'Payload'), dest]);
    for (const f of ['engine/open-charts', 'OpenChartsPopup.app/Contents/MacOS/open-charts-popup', 'package/open-charts.ppam', 'package/addin.json', 'package/Uninstall open-charts.app']) {
      check(existsSync(join(dest, f)), `payload has no ${f}`);
    }
  });
  const engine = join(dest, 'engine', 'open-charts');
  step('engine and helper are universal and signed', () => {
    for (const f of [engine, join(dest, 'OpenChartsPopup.app/Contents/MacOS/open-charts-popup')]) {
      const archs = sh('lipo', ['-archs', f]).trim().split(/\s+/).sort().join(' ');
      check(archs === 'arm64 x86_64', `${f}: ${archs}`);
    }
    for (const f of [engine, join(dest, 'OpenChartsPopup.app'), join(dest, 'package', 'Uninstall open-charts.app')]) sh('codesign', ['--verify', '--strict', f]);
  });
  step('no Node on PATH', () => check(!/\/node\n?$/.test(execFileSync('/bin/sh', ['-c', 'command -v node || true'], { env, encoding: 'utf8' })), 'node is on PATH'));
  step('preinstall + postinstall, as Installer runs them', () => {
    sh(join(x, component, 'Scripts', 'preinstall'), [pkg, dest, home, '/']);
    const out = sh(join(x, component, 'Scripts', 'postinstall'), [pkg, dest, home, '/']);
    check(/Add-in written to/.test(out) && /Bridge script compiled/.test(out), out);
  });

  const ppam = join(home, 'Library', 'Application Support', 'Microsoft', 'open-charts', 'open-charts.ppam');
  const scpt = join(home, 'Library', 'Application Scripts', 'com.microsoft.Powerpoint', 'open-charts.scpt');
  const launcher = join(dest, 'bin', 'open-charts');
  const uninstaller = join(home, 'Applications', 'Uninstall open-charts.app');
  step('files where `addin install` puts them', () => {
    for (const f of [ppam, scpt, launcher, uninstaller, join(dest, 'addin.json')]) check(existsSync(f), `missing ${f}`);
    const cfg = JSON.parse(readFileSync(join(dest, 'addin.json'), 'utf8'));
    const staged = JSON.parse(readFileSync(join(dest, 'package', 'addin.json'), 'utf8'));
    check(cfg.prefix === staged.prefix && /^OC[0-9A-F]{24}$/.test(cfg.prefix), `prefix ${cfg.prefix}`);
    check(cfg.engine === engine, `engine ${cfg.engine}`);
    check(readFileSync(ppam).equals(readFileSync(join(dest, 'package', 'open-charts.ppam'))), 'ppam differs from the package\'s');
  });
  step('the bridge script runs the engine executable (no Node)', () => {
    const cmd = sh('osascript', ['-e', `set s to load script (POSIX file ${JSON.stringify(scpt)})`, '-e', 'tell s to engineCommand()']).trim();
    check(cmd === `'${engine}'`, cmd);
  });
  step('the open-charts command', () => check(/stacked\s+Stacked Column\/Bar\s+available/.test(sh(launcher, ['addin', 'types'])), 'no gallery'));
  step('the engine as the add-in runs it (scripts/smoke-exe.ts)', () => {
    execFileSync(process.execPath, ['--experimental-strip-types', '--no-warnings', join(ROOT, 'scripts', 'smoke-exe.ts'), engine], { stdio: 'inherit', cwd: ROOT, env: { ...process.env, OPEN_CHARTS_PORT: '4198' } });
  });
  step('addin uninstall removes it all', () => {
    sh(launcher, ['addin', 'uninstall']);
    for (const f of [ppam, scpt, launcher, uninstaller, engine, join(dest, 'OpenChartsPopup.app'), join(dest, 'package')]) check(!existsSync(f), `still there: ${f}`);
  });
} finally {
  rmSync(tmp, { recursive: true, force: true });
}
if (failed) process.exit(1);
console.log('\nThe package installs and uninstalls cleanly (temporary home, no Node, no PowerPoint).');
