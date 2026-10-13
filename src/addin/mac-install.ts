// The Mac install from open-charts-<version>.pkg (scripts/build-mac.ts), and the pieces it shares
// with the install from source (`open-charts addin install` in a checkout, src/addin/cli.ts).
//
// The package's payload puts, under ~/Library/Application Support/open-charts (addinHome()):
//   engine/open-charts       the engine as one executable (Node single-executable application,
//                            arm64 + x86_64): no Node on the Mac
//   OpenChartsPopup.app      the Edit data pop-up helper
//   package/                 open-charts.ppam, addin.json (its callback prefix), Uninstall open-charts.app
// and its postinstall runs `engine/open-charts addin install`, which lands here (installPackaged):
// the files end up exactly where a source install puts them, the bridge runs the executable instead
// of Node, and PowerPoint is never started.
import { execFileSync } from 'node:child_process';
import { cpSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { addinHome, installLauncher } from './launcher.ts';
import { log } from './server.ts';

const ROOT = fileURLToPath(new URL('../../', import.meta.url));
const SCRIPT_TEMPLATE = join(ROOT, 'addin', 'applescript', 'open-charts.applescript');
export const PPAM_DIR = join(homedir(), 'Library', 'Application Support', 'Microsoft', 'open-charts');
export const SCRIPTS_DIR = join(homedir(), 'Library', 'Application Scripts', 'com.microsoft.Powerpoint');
/** The package identifier of open-charts-<version>.pkg. */
export const PKG_ID = 'io.github.open-charts.addin';
export const UNINSTALLER_APP = 'Uninstall open-charts.app';
export const REGISTER_AGENT = 'io.github.open-charts.register';
export const agentPlist = () => join(homedir(), 'Library', 'LaunchAgents', `${REGISTER_AGENT}.plist`);
export const uninstallerPath = () => join(homedir(), 'Applications', UNINSTALLER_APP);

/** Tests (and the package build's checks) set this: nothing talks to PowerPoint or launchd. */
export const handsOff = () => !!process.env.OPEN_CHARTS_NO_POWERPOINT;

export function registerScript(ppam: string): string {
  return `tell application "Microsoft PowerPoint"
  set theAddIn to register add in ${JSON.stringify(ppam)}
  set auto load of theAddIn to true
  set loaded of theAddIn to true
end tell`;
}

export function osascript(script: string, stderr: 'inherit' | 'pipe' = 'inherit'): string {
  return execFileSync('osascript', ['-e', script], { encoding: 'utf8', stdio: ['ignore', 'pipe', stderr], timeout: 120_000 }).trim();
}

/**
 * The AppleScript bridge, filled in and compiled into PowerPoint's Application Scripts folder.
 * `node` empty: `cli` is the engine executable (package install).
 */
export function compileBridge(home: string, node: string, cli: string): string {
  const source = readFileSync(SCRIPT_TEMPLATE, 'utf8')
    .replace('__NODE__', node)
    .replace('__CLI__', cli)
    .replace('__HOME__', home);
  const sourceFile = join(home, 'open-charts.applescript');
  mkdirSync(home, { recursive: true });
  writeFileSync(sourceFile, source);
  mkdirSync(SCRIPTS_DIR, { recursive: true });
  const out = join(SCRIPTS_DIR, 'open-charts.scpt');
  execFileSync('osacompile', ['-o', out, sourceFile]);
  return out;
}

/** The install the package's postinstall runs (this engine executable, `addin install`). */
export function installPackaged(register: boolean) {
  const home = addinHome();
  const exe = process.execPath;
  const pkg = join(home, 'package');
  if (!existsSync(join(pkg, 'addin.json'))) throw new Error(`${pkg} is missing: install open-charts again from its .pkg.`);
  const { prefix } = JSON.parse(readFileSync(join(pkg, 'addin.json'), 'utf8')) as { prefix: string };
  // A source install's engine copy is replaced by this executable.
  for (const p of ['src', 'vendor', 'package.json']) rmSync(join(home, 'engine', p), { recursive: true, force: true });
  writeFileSync(join(home, 'addin.json'), `${JSON.stringify({ prefix, engine: exe }, null, 2)}\n`);

  mkdirSync(PPAM_DIR, { recursive: true });
  const ppam = join(PPAM_DIR, 'open-charts.ppam');
  cpSync(join(pkg, 'open-charts.ppam'), ppam);
  console.log(`Add-in written to ${ppam}`);
  console.log(`Bridge script compiled to ${compileBridge(home, '', exe)}`);

  const uninstaller = join(pkg, UNINSTALLER_APP);
  if (existsSync(uninstaller)) {
    const dest = uninstallerPath();
    rmSync(dest, { recursive: true, force: true });
    mkdirSync(dirname(dest), { recursive: true });
    cpSync(uninstaller, dest, { recursive: true, verbatimSymlinks: true });
    console.log(`Uninstaller copied to ${dest}`);
  }
  for (const line of installLauncher('', exe, home)) console.log(line);
  if (register) for (const line of registerWhenRunning(ppam, exe)) console.log(line);
  else console.log(`Load it in PowerPoint: Tools > PowerPoint Add-ins… > + > ${ppam}`);
}

/** True when PowerPoint is running, asked without an Apple event (which would start it). */
export function powerPointRunning(): boolean {
  try { execFileSync('/usr/bin/pgrep', ['-xq', 'Microsoft PowerPoint'], { stdio: 'ignore' }); return true; } catch { return false; }
}

/**
 * The LaunchAgent that registers the add-in once PowerPoint runs: every 30 s a `pgrep` (the engine
 * is not started) and, while PowerPoint is up, `addin register`, which removes the agent when done.
 */
export function registerAgentPlist(exe: string): string {
  const x = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const q = (s: string) => `'${s.replace(/'/g, `'\\''`)}'`;
  return `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>${REGISTER_AGENT}</string>
  <key>ProgramArguments</key>
  <array>
    <string>/bin/sh</string>
    <string>-c</string>
    <string>${x(`/usr/bin/pgrep -xq 'Microsoft PowerPoint' || exit 0; sleep 5; exec ${q(exe)} addin register`)}</string>
  </array>
  <key>RunAtLoad</key><true/>
  <key>StartInterval</key><integer>30</integer>
  <key>ProcessType</key><string>Background</string>
</dict>
</plist>
`;
}

const launchctl = (...args: string[]) => {
  try { execFileSync('/bin/launchctl', args, { stdio: 'ignore', timeout: 10_000 }); return true; } catch { return false; }
};
const guiDomain = () => `gui/${process.getuid?.() ?? ''}`;
const attemptsFile = (home = addinHome()) => join(home, 'register-attempts');

/** Registers the add-in now if PowerPoint runs, else as soon as it runs. Returns the lines to print. */
export function registerWhenRunning(ppam: string, exe: string): string[] {
  if (handsOff()) return ['PowerPoint not touched (OPEN_CHARTS_NO_POWERPOINT).'];
  if (powerPointRunning()) {
    try {
      osascript(registerScript(ppam), 'pipe');
      return ['Registered and loaded in PowerPoint: look for the open-charts tab.'];
    } catch { /* no Automation permission from the installer: the agent asks from the user's session */ }
  }
  const plist = agentPlist();
  mkdirSync(dirname(plist), { recursive: true });
  writeFileSync(plist, registerAgentPlist(exe));
  writeFileSync(attemptsFile(), '0');
  launchctl('bootout', `${guiDomain()}/${REGISTER_AGENT}`);
  return [launchctl('bootstrap', guiDomain(), plist)
    ? 'open-charts loads as soon as PowerPoint runs (start it now, or it loads the next time).'
    : 'open-charts loads the next time you log in and start PowerPoint.'];
}

/** `addin register`: the agent's call while PowerPoint runs. Gives up after 10 failures (e.g. Automation denied). */
export function registerPending() {
  if (handsOff() || !powerPointRunning()) return;
  const home = addinHome(), ppam = join(PPAM_DIR, 'open-charts.ppam');
  let done = false;
  try {
    osascript(registerScript(ppam), 'pipe');
    log('add-in registered in PowerPoint', home);
    done = true;
  } catch (err) {
    const n = (Number(existsSync(attemptsFile(home)) ? readFileSync(attemptsFile(home), 'utf8') : 0) || 0) + 1;
    writeFileSync(attemptsFile(home), String(n));
    log(`add-in registration failed (${n}/10): ${(err as Error).message.split('\n')[0]}`, home);
    if ((done = n >= 10)) log(`giving up: load it in PowerPoint with Tools > PowerPoint Add-ins… > + > ${ppam}`, home);
  }
  if (done) removeRegisterAgent();
}

export function removeRegisterAgent() {
  rmSync(attemptsFile(), { force: true });
  if (!existsSync(agentPlist())) return;
  rmSync(agentPlist(), { force: true });
  if (!handsOff()) launchctl('bootout', `${guiDomain()}/${REGISTER_AGENT}`); // last: run by the agent, this ends the process
}

/** The package install's own files, removed by `addin uninstall` along with the shared ones. */
export function removePackaged() {
  for (const p of [join(addinHome(), 'package'), uninstallerPath()]) rmSync(p, { recursive: true, force: true });
  removeRegisterAgent();
  try { execFileSync('/usr/sbin/pkgutil', ['--volume', homedir(), '--forget', PKG_ID], { stdio: 'ignore' }); } catch { /* not installed from the package */ }
}
