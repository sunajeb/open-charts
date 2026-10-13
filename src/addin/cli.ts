// `open-charts addin …`: build and install the PowerPoint add-in, and the engine end of its bridge.
import { isSea } from 'node:sea';
import { cpSync, existsSync, mkdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createMeasurer, reportOnce } from '../fonts.ts';
import { handleRequest, parseRequest } from './engine.ts';
import { httpSessions, saveSettings, serveAddin, stopServer } from './server.ts';
import { ensureToken, tokenPath } from './token.ts';
import { addinHome, installLauncher, removeLauncher } from './launcher.ts';
import { buildAddin } from './package.ts';
import { galleryItems, sampleFor } from './registry.ts';
import { checkForUpdate, readSettings, writeSettings } from '../update.ts';
import { HELPER_APP } from './popup.ts';
import { buildPopup, canBuildPopup } from './build-popup.ts';
import { compileBridge, handsOff, installPackaged, osascript, PPAM_DIR, registerPending, registerScript, removePackaged, SCRIPTS_DIR } from './mac-install.ts';

export { ADDIN_USAGE } from './usage.ts';
import { ADDIN_USAGE } from './usage.ts';

const ROOT = fileURLToPath(new URL('../../', import.meta.url));
export { osascript, PPAM_DIR, registerScript, SCRIPTS_DIR };
export const ADDIN_NAME = 'open-charts';

export function addinMain(args: string[], opts: { out?: string; register?: boolean }): void | Promise<void> {
  const [sub, arg] = args;
  switch (sub) {
    case 'run': {
      // Always answers in the line format and exits 0: the VBA side shows E records to the user.
      let request = '';
      try { request = readFileSync(!arg || arg === '-' ? 0 : arg, 'utf8'); } catch { request = ''; }
      return handleRequest(parseRequest(request), createMeasurer(), httpSessions(), reportOnce)
        .catch((err: Error) => `OC\t1\nE\t${err.message.replace(/[\t\n\r]/g, ' ')}\n`)
        .then((out) => { process.stdout.write(out); });
    }
    case 'serve':
      serveAddin(Number(process.env.OPEN_CHARTS_PORT) || undefined); // another port: tests only (the add-in uses 4174)
      return;
    case 'types':
      for (const g of galleryItems()) {
        const s = sampleFor(g.id);
        console.log(`${g.id.padEnd(12)} ${g.label.replace(/\s*\n\s*/g, ' ').padEnd(22)} ${'spec' in s ? 'available' : 'not yet'}`);
      }
      return;
    case 'build': {
      const out = opts.out ?? join(ROOT, 'dist', 'open-charts.ppam');
      const built = buildAddin();
      mkdirSync(dirname(out), { recursive: true });
      writeFileSync(out, built.ppam);
      console.log(`Wrote ${out} (${built.ppam.length} bytes, callback prefix ${built.prefix})`);
      return;
    }
    case 'verbose': {
      if (arg !== 'on' && arg !== 'off') { console.error(ADDIN_USAGE); process.exit(1); }
      saveSettings({ verbose: arg === 'on' });
      console.log(`Verbose logging ${arg} (${join(addinHome(), 'addin.log')}). The add-in picks it up when PowerPoint starts next.`);
      return;
    }
    case 'install':
      return stopServer().then(() => install(opts.register !== false));
    case 'update':
      return update(arg);
    case 'register':
      // Run by the LaunchAgent a package install leaves when PowerPoint was not running.
      registerPending();
      return;
    case 'uninstall':
      return stopServer().then(uninstall);
    default:
      console.error(ADDIN_USAGE);
      process.exit(1);
  }
}

/** `addin update [on|off]`: turn the daily "update available" check on or off, or check now. */
async function update(arg?: string) {
  const home = addinHome();
  if (arg === 'on' || arg === 'off') {
    writeSettings(home, { updateCheck: arg === 'on' });
    console.log(`Update check ${arg === 'on' ? 'on: the add-in asks GitHub Releases once a day and tells you about a new version' : 'off'}.`);
    return;
  }
  if (arg) throw new Error(ADDIN_USAGE);
  const u = await checkForUpdate();
  console.log(u.available
    ? `open-charts ${u.latest} is available (you have ${u.current}): ${u.url}${u.asset ? ` (${u.asset})` : ''}`
    : `open-charts ${u.current} is up to date (latest release ${u.latest}).`);
  console.log(`Daily check: ${readSettings(home).updateCheck ? 'on' : 'off'} (open-charts addin update on|off).`);
}

function install(register: boolean) {
  if (process.platform === 'win32') throw new Error('On Windows, run install.ps1 from the open-charts-win folder (README, Install on Windows).');
  if (process.platform !== 'darwin') throw new Error('addin install is for PowerPoint for Mac.');
  if (isSea()) return installPackaged(register); // the engine executable, installed from the .pkg
  const home = addinHome();
  // 1. Engine copy outside ~/Downloads & co., which macOS privacy protection guards from the bridge.
  const engine = join(home, 'engine');
  if (resolve(ROOT) === resolve(engine)) {
    throw new Error('This is the installed engine copy. Run addin install from the open-charts folder you downloaded.');
  }
  rmSync(engine, { recursive: true, force: true });
  mkdirSync(engine, { recursive: true });
  for (const p of ['src', 'vendor', 'package.json']) cpSync(join(ROOT, p), join(engine, p), { recursive: true });
  console.log(`Engine copied to ${engine}`);

  // 1b. The local server's secret (src/addin/token.ts): kept across reinstalls, user-only permissions.
  ensureToken(home);
  console.log(`Server token in ${tokenPath(home)}`);

  // 2. The add-in, where PowerPoint's sandbox may read it (~/Library/Application Support/Microsoft).
  const built = buildAddin();
  writeFileSync(join(home, 'addin.json'), `${JSON.stringify({ prefix: built.prefix, engine: engine }, null, 2)}\n`);
  mkdirSync(PPAM_DIR, { recursive: true });
  const ppam = join(PPAM_DIR, 'open-charts.ppam');
  writeFileSync(ppam, built.ppam);
  console.log(`Add-in written to ${ppam}`);

  // 3. The AppleScript bridge, compiled into PowerPoint's Application Scripts folder.
  console.log(`Bridge script compiled to ${compileBridge(home, process.execPath, join(engine, 'src', 'cli.ts'))}`);

  // 3b. The Edit data pop-up helper (Mac app, built here with swiftc; without it Edit data opens
  // the browser as before).
  for (const line of installPopup(home)) console.log(line);

  // 4. The open-charts command (runs this folder, so a later `open-charts addin install` updates).
  for (const line of installLauncher(process.execPath, join(ROOT, 'src', 'cli.ts'), home)) console.log(line);

  // 5. Register and load it in PowerPoint (starts PowerPoint if needed).
  if (register) {
    osascript(registerScript(ppam));
    console.log('Registered and loaded in PowerPoint: look for the open-charts tab.');
  } else {
    console.log(`Load it in PowerPoint: Tools > PowerPoint Add-ins… > + > ${ppam}`);
  }
}

/** Copies (or builds) OpenChartsPopup.app into the add-in's home. Returns the lines to print. */
export function installPopup(home = addinHome()): string[] {
  const dest = join(home, HELPER_APP), dist = join(ROOT, 'dist', HELPER_APP);
  const newest = (...p: string[]) => Math.max(...p.map((f) => statSync(f).mtimeMs));
  const src = join(ROOT, 'addin', 'popup-mac');
  try {
    if (existsSync(join(dist, 'Contents', 'MacOS')) && newest(join(dist, 'Contents', 'MacOS', 'open-charts-popup')) >= newest(join(src, 'main.swift'), join(src, 'Info.plist'))) {
      rmSync(dest, { recursive: true, force: true });
      cpSync(dist, dest, { recursive: true });
      return [`Pop-up helper copied to ${dest}`];
    }
    if (!canBuildPopup()) return ['Pop-up helper not built (no swiftc: run `xcode-select --install`, then install again). Edit data opens in the browser.'];
    console.log('Building the pop-up helper (swiftc, about 30 s)…');
    buildPopup(dest);
    return [`Pop-up helper built into ${dest}`];
  } catch (err) {
    return [`Pop-up helper not installed (${(err as Error).message.split('\n')[0]}). Edit data opens in the browser.`];
  }
}

// PowerPoint's `add ins` collection cannot be walked by reference (the loop never ends), so go by name.
export function unregisterScript(): string {
  return `tell application "Microsoft PowerPoint"
  repeat with n in (get name of every add in)
    if (n as text) starts with "${ADDIN_NAME}" then
      set loaded of add in (n as text) to false
      set auto load of add in (n as text) to false
    end if
  end repeat
end tell`;
}

function uninstall() {
  try {
    if (!handsOff()) osascript(unregisterScript(), 'pipe');
  } catch { /* PowerPoint not available: files are still removed */ }
  for (const p of [join(SCRIPTS_DIR, 'open-charts.scpt'), PPAM_DIR, join(addinHome(), 'engine'), join(addinHome(), HELPER_APP), tokenPath()]) {
    if (existsSync(p)) rmSync(p, { recursive: true, force: true });
  }
  removeLauncher();
  removePackaged();
  console.log('open-charts add-in removed.');
}
