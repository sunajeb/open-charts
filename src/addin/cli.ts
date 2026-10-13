// `open-charts addin …`: build and install the PowerPoint add-in, and the engine end of its bridge.
import { execFileSync } from 'node:child_process';
import { cpSync, existsSync, mkdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createMeasurer } from '../fonts.ts';
import { handleRequest, parseRequest } from './engine.ts';
import { httpSessions, serveAddin, stopServer } from './server.ts';
import { addinHome, installLauncher, removeLauncher } from './launcher.ts';
import { buildAddin } from './package.ts';
import { galleryItems, sampleFor } from './registry.ts';
import { HELPER_APP } from './popup.ts';
import { buildPopup, canBuildPopup } from './build-popup.ts';

export { ADDIN_USAGE } from './usage.ts';
import { ADDIN_USAGE } from './usage.ts';

const ROOT = fileURLToPath(new URL('../../', import.meta.url));
const SCRIPT_TEMPLATE = join(ROOT, 'addin', 'applescript', 'open-charts.applescript');
export const PPAM_DIR = join(homedir(), 'Library', 'Application Support', 'Microsoft', 'open-charts');
export const SCRIPTS_DIR = join(homedir(), 'Library', 'Application Scripts', 'com.microsoft.Powerpoint');
export const ADDIN_NAME = 'open-charts';

export function addinMain(args: string[], opts: { out?: string; register?: boolean }): void | Promise<void> {
  const [sub, arg] = args;
  switch (sub) {
    case 'run': {
      // Always answers in the line format and exits 0: the VBA side shows E records to the user.
      let request = '';
      try { request = readFileSync(!arg || arg === '-' ? 0 : arg, 'utf8'); } catch { request = ''; }
      return handleRequest(parseRequest(request), createMeasurer(), httpSessions())
        .catch((err: Error) => `OC\t1\nE\t${err.message.replace(/[\t\n\r]/g, ' ')}\n`)
        .then((out) => { process.stdout.write(out); });
    }
    case 'serve':
      serveAddin();
      return;
    case 'types':
      for (const g of galleryItems()) {
        const s = sampleFor(g.id);
        console.log(`${g.id.padEnd(12)} ${g.label.padEnd(18)} ${'spec' in s ? 'available' : 'not yet'}`);
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
    case 'install':
      return stopServer().then(() => install(opts.register !== false));
    case 'uninstall':
      return stopServer().then(uninstall);
    default:
      console.error(ADDIN_USAGE);
      process.exit(1);
  }
}

function install(register: boolean) {
  if (process.platform !== 'darwin') throw new Error('addin install is for PowerPoint for Mac.');
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

  // 2. The add-in, where PowerPoint's sandbox may read it (~/Library/Application Support/Microsoft).
  const built = buildAddin();
  writeFileSync(join(home, 'addin.json'), `${JSON.stringify({ prefix: built.prefix, engine: engine }, null, 2)}\n`);
  mkdirSync(PPAM_DIR, { recursive: true });
  const ppam = join(PPAM_DIR, 'open-charts.ppam');
  writeFileSync(ppam, built.ppam);
  console.log(`Add-in written to ${ppam}`);

  // 3. The AppleScript bridge, compiled into PowerPoint's Application Scripts folder.
  const source = readFileSync(SCRIPT_TEMPLATE, 'utf8')
    .replace('__NODE__', process.execPath)
    .replace('__CLI__', join(engine, 'src', 'cli.ts'))
    .replace('__HOME__', home);
  const sourceFile = join(home, 'open-charts.applescript');
  writeFileSync(sourceFile, source);
  mkdirSync(SCRIPTS_DIR, { recursive: true });
  execFileSync('osacompile', ['-o', join(SCRIPTS_DIR, 'open-charts.scpt'), sourceFile]);
  console.log(`Bridge script compiled to ${join(SCRIPTS_DIR, 'open-charts.scpt')}`);

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

export function registerScript(ppam: string): string {
  return `tell application "Microsoft PowerPoint"
  set theAddIn to register add in ${JSON.stringify(ppam)}
  set auto load of theAddIn to true
  set loaded of theAddIn to true
end tell`;
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
    osascript(unregisterScript(), 'pipe');
  } catch { /* PowerPoint not available: files are still removed */ }
  for (const p of [join(SCRIPTS_DIR, 'open-charts.scpt'), PPAM_DIR, join(addinHome(), 'engine'), join(addinHome(), HELPER_APP)]) {
    if (existsSync(p)) rmSync(p, { recursive: true, force: true });
  }
  removeLauncher();
  console.log('open-charts add-in removed.');
}

export function osascript(script: string, stderr: 'inherit' | 'pipe' = 'inherit'): string {
  return execFileSync('osascript', ['-e', script], { encoding: 'utf8', stdio: ['ignore', 'pipe', stderr], timeout: 120_000 }).trim();
}
