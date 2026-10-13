// `npm run build:mac`: the Mac installer, dist/mac/open-charts-<version>.pkg (docs/release.md).
//
// One per-user package ("Install for me only": no administrator password) that needs no Node and
// no Terminal. Its payload goes to ~/Library/Application Support/open-charts:
//   engine/open-charts        the engine as a Node single-executable application (SEA), universal
//                             (arm64 + x86_64): the official Node binaries of this Node version,
//                             checked against nodejs.org's SHASUMS256.txt, with the same blob as
//                             open-charts.exe injected (scripts/build-win.ts)
//   OpenChartsPopup.app       the Edit data pop-up helper, universal
//   package/open-charts.ppam  the add-in, its addin.json (callback prefix) and Uninstall open-charts.app
// The postinstall (addin/mac/postinstall) runs `engine/open-charts addin install`
// (src/addin/mac-install.ts): add-in, bridge script, open-charts command, uninstaller in
// ~/Applications, and registration in PowerPoint without starting it.
//
// Signing (env; all optional, unsigned otherwise and the build says so):
//   APPLE_DEVELOPER_ID   "Developer ID Application: Name (TEAMID)"  signs the engine, helper, uninstaller
//   APPLE_INSTALLER_ID   "Developer ID Installer: Name (TEAMID)"    signs the .pkg
//   APPLE_NOTARY_PROFILE  a `xcrun notarytool store-credentials` profile, or
//   APPLE_ID + APPLE_TEAM_ID + APPLE_APP_PASSWORD (app-specific password): notarize and staple.
import { execFileSync } from 'node:child_process';
import { chmodSync, copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildAddin } from '../src/addin/package.ts';
import { buildPopup } from '../src/addin/build-popup.ts';
import { HELPER_APP } from '../src/addin/popup.ts';
import { PKG_ID, UNINSTALLER_APP } from '../src/addin/mac-install.ts';
import { bundle, engineAssets, inject, officialNodeFile, seaBlob } from './build-win.ts';

const ROOT = fileURLToPath(new URL('../', import.meta.url));
const OUT = join(ROOT, 'dist', 'mac');
const WORK = join(OUT, 'work');
const STAGE = join(WORK, 'payload');
const VERSION = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8')).version as string;
export const INSTALL_LOCATION = '/Library/Application Support/open-charts'; // under the user's home
export const ARCHS = ['arm64', 'x64'] as const;
const env = process.env;

const run = (cmd: string, args: string[], opts: { quiet?: boolean } = {}) =>
  execFileSync(cmd, args, { encoding: 'utf8', stdio: ['ignore', opts.quiet ? 'pipe' : 'inherit', 'inherit'], maxBuffer: 1 << 26 });
const rel = (p: string) => relative(ROOT, p);

/** Developer ID signature (hardened runtime, timestamp) when APPLE_DEVELOPER_ID is set, else ad hoc. */
function sign(path: string, entitlements?: string) {
  const id = env.APPLE_DEVELOPER_ID;
  const args = id
    ? ['--force', '--sign', id, '--options', 'runtime', '--timestamp', ...(entitlements ? ['--entitlements', entitlements] : [])]
    : ['--force', '--sign', '-', '--timestamp=none'];
  run('codesign', [...args, path], { quiet: true });
}

/** Lowest macOS the official Node binaries of this version run on (Node 23+: 13.5). */
export const minMacOS = (nodeMajor = Number(process.versions.node.split('.')[0])) => (nodeMajor >= 23 ? '13.5' : '11.0');

export function distributionXml(version: string, componentPkg: string, minOS = minMacOS()): string {
  return `<?xml version="1.0" encoding="utf-8"?>
<installer-gui-script minSpecVersion="2">
  <title>open-charts ${version}</title>
  <welcome file="welcome.html" mime-type="text/html"/>
  <conclusion file="conclusion.html" mime-type="text/html"/>
  <options customize="never" require-scripts="false" hostArchitectures="arm64,x86_64"/>
  <!-- Per user only: installs into the home folder, no administrator password. -->
  <domains enable_anywhere="false" enable_currentUserHome="true" enable_localSystem="false"/>
  <volume-check>
    <allowed-os-versions><os-version min="${minOS}"/></allowed-os-versions>
  </volume-check>
  <choices-outline>
    <line choice="default"><line choice="${PKG_ID}"/></line>
  </choices-outline>
  <choice id="default"/>
  <choice id="${PKG_ID}" visible="false"><pkg-ref id="${PKG_ID}"/></choice>
  <pkg-ref id="${PKG_ID}" version="${version}" onConclusion="none">${componentPkg}</pkg-ref>
</installer-gui-script>
`;
}

/** A component's PackageInfo without the root authorization pkgbuild always asks for. */
export function noRootAuth(packageInfo: string): string {
  const out = packageInfo.replace(/(<pkg-info\b[^>]*?)\sauth="root"/, '$1 auth="none"');
  if (!/<pkg-info\b[^>]*\sauth="none"/.test(out)) throw new Error('PackageInfo: no auth attribute to change');
  return out;
}

const page = (body: string) => `<!doctype html><html><head><meta charset="utf-8"><style>body{font:13px -apple-system,sans-serif;line-height:1.5}</style></head><body>${body}</body></html>\n`;
const WELCOME = page(`<p>This installs the <b>open-charts</b> add-in for PowerPoint for Mac, for you only (no administrator password).</p>
<p>PowerPoint can stay open. If it is running, the open-charts tab appears in it a few seconds after the installation; otherwise the next time you start it.</p>`);
const CONCLUSION = page(`<p><b>open-charts is installed.</b> Look for the open-charts tab in PowerPoint.</p>
<p>macOS may ask whether open-charts may control PowerPoint: choose <b>OK</b>, it needs this to load the add-in and to apply edits.</p>
<p>To remove it, open <b>Uninstall open-charts</b> in the Applications folder of your home folder.</p>`);

/** The official `node` binary for darwin-<arch> of this Node version (verified, cached), copied to `dest`. */
async function officialMacNode(arch: string, dest: string) {
  const name = `node-${process.version}-darwin-${arch}`;
  const tgz = await officialNodeFile(`${name}.tar.gz`);
  const tmp = mkdtempSync(join(tmpdir(), 'oc-node-'));
  try {
    run('tar', ['-xzf', tgz, '-C', tmp, `${name}/bin/node`]);
    copyFileSync(join(tmp, name, 'bin', 'node'), dest);
    chmodSync(dest, 0o755);
  } finally {
    rmSync(tmp, { recursive: true, force: true });
  }
}

/**
 * Runs the engine the way the installed add-in does, with no Node anywhere on PATH and its own
 * resource folder: `addin types` must list the gallery. `prefix` e.g. ['arch', '-x86_64'].
 */
export function checkEngine(exe: string, prefix: string[] = []): string {
  const dir = mkdtempSync(join(tmpdir(), 'oc-engine-'));
  try {
    const [cmd, ...args] = [...prefix, exe, 'addin', 'types'];
    const out = execFileSync(cmd, args, {
      encoding: 'utf8', timeout: 60_000,
      env: { PATH: '/usr/bin:/bin:/usr/sbin:/sbin', HOME: dir, LOCALAPPDATA: join(dir, 'res'), OPEN_CHARTS_HOME: join(dir, 'home') },
    });
    if (!/stacked\s+Stacked Column\/Bar\s+available/.test(out)) throw new Error(`${exe} answered:\n${out}`);
    return out;
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

const rosetta = () => { try { execFileSync('arch', ['-x86_64', '/usr/bin/true'], { stdio: 'ignore' }); return true; } catch { return false; } };

function notarize(pkg: string): boolean {
  const auth = env.APPLE_NOTARY_PROFILE
    ? ['--keychain-profile', env.APPLE_NOTARY_PROFILE]
    : env.APPLE_ID && env.APPLE_TEAM_ID && env.APPLE_APP_PASSWORD
      ? ['--apple-id', env.APPLE_ID, '--team-id', env.APPLE_TEAM_ID, '--password', env.APPLE_APP_PASSWORD]
      : undefined;
  if (!auth) return false;
  console.log('Notarizing (xcrun notarytool, a few minutes)…');
  run('xcrun', ['notarytool', 'submit', pkg, ...auth, '--wait']);
  run('xcrun', ['stapler', 'staple', pkg]);
  return true;
}

async function main() {
  if (process.platform !== 'darwin') throw new Error('build:mac runs on a Mac (pkgbuild, productbuild, codesign, swiftc).');
  rmSync(OUT, { recursive: true, force: true });
  mkdirSync(join(STAGE, 'engine'), { recursive: true });
  mkdirSync(join(STAGE, 'package'), { recursive: true });

  // 1. The engine: one bundle, one SEA blob (no code cache or snapshot: the same for every
  //    architecture), injected into each official node binary, then joined into one universal file.
  const { assets, assetHash } = engineAssets();
  const cjs = join(OUT, 'open-charts.cjs'), blob = join(WORK, 'sea-prep.blob');
  await bundle(cjs, assets, assetHash);
  await seaBlob(cjs, blob, assets);
  const slices: string[] = [];
  for (const arch of ARCHS) {
    const slice = join(WORK, `open-charts-${arch}`);
    await officialMacNode(arch, slice);
    run('codesign', ['--remove-signature', slice]);
    await inject(slice, blob, true);
    slices.push(slice);
  }
  const engine = join(STAGE, 'engine', 'open-charts');
  run('lipo', ['-create', '-output', engine, ...slices]);
  sign(engine, join(ROOT, 'addin', 'mac', 'engine.entitlements'));
  console.log(`Engine: ${rel(engine)} (${(statSync(engine).size / 1e6).toFixed(0)} MB, ${run('lipo', ['-archs', engine], { quiet: true }).trim()}, Node ${process.version})`);
  checkEngine(engine);
  console.log(`  runs with no Node on PATH (${process.arch})`);
  if (process.arch === 'arm64' && rosetta()) {
    checkEngine(engine, ['arch', '-x86_64']);
    console.log('  runs as x86_64 too (Rosetta)');
  }

  // 2. Pop-up helper (universal), the add-in, the uninstaller.
  const popup = buildPopup(join(STAGE, HELPER_APP), [...ARCHS], false);
  sign(popup);
  const built = buildAddin();
  writeFileSync(join(STAGE, 'package', 'open-charts.ppam'), built.ppam);
  writeFileSync(join(STAGE, 'package', 'addin.json'), `${JSON.stringify({ prefix: built.prefix }, null, 2)}\n`);
  const uninstaller = join(STAGE, 'package', UNINSTALLER_APP);
  run('osacompile', ['-o', uninstaller, join(ROOT, 'addin', 'mac', 'uninstall.applescript')]);
  sign(uninstaller);
  console.log(`Pop-up helper, open-charts.ppam (callback prefix ${built.prefix}), ${UNINSTALLER_APP}`);

  // 3. The component package: never relocated to another copy of an app with the same bundle id.
  const scripts = join(WORK, 'scripts');
  mkdirSync(scripts, { recursive: true });
  for (const s of ['preinstall', 'postinstall']) {
    copyFileSync(join(ROOT, 'addin', 'mac', s), join(scripts, s));
    chmodSync(join(scripts, s), 0o755);
  }
  const plist = join(WORK, 'component.plist');
  run('pkgbuild', ['--analyze', '--root', STAGE, plist], { quiet: true });
  const bundles = (JSON.parse(run('plutil', ['-convert', 'json', '-o', '-', plist], { quiet: true })) as unknown[]).length;
  for (let i = 0; i < bundles; i++) run('plutil', ['-replace', `${i}.BundleIsRelocatable`, '-bool', 'NO', plist]);
  const component = 'open-charts-component.pkg';
  const raw = join(WORK, 'raw.pkg'), expanded = join(WORK, 'raw');
  run('pkgbuild', ['--root', STAGE, '--component-plist', plist, '--identifier', PKG_ID, '--version', VERSION, '--ownership', 'preserve',
    '--install-location', INSTALL_LOCATION, '--scripts', scripts, raw], { quiet: true });
  // pkgbuild always asks for root authorization (auth="root"), which makes Installer ask for an
  // administrator password even for "Install for me only". Nothing here needs root: auth="none".
  run('pkgutil', ['--expand', raw, expanded]);
  const info = join(expanded, 'PackageInfo');
  writeFileSync(info, noRootAuth(readFileSync(info, 'utf8')));
  run('pkgutil', ['--flatten', expanded, join(WORK, component)]);
  rmSync(raw);
  rmSync(expanded, { recursive: true });

  // 4. The product package (per-user domain, welcome and conclusion pages), signed if it can be.
  const resources = join(WORK, 'resources');
  mkdirSync(resources, { recursive: true });
  writeFileSync(join(resources, 'welcome.html'), WELCOME);
  writeFileSync(join(resources, 'conclusion.html'), CONCLUSION);
  const distribution = join(WORK, 'distribution.xml');
  writeFileSync(distribution, distributionXml(VERSION, component));
  const pkg = join(OUT, `open-charts-${VERSION}.pkg`);
  run('productbuild', ['--distribution', distribution, '--resources', resources, '--package-path', WORK,
    ...(env.APPLE_INSTALLER_ID ? ['--sign', env.APPLE_INSTALLER_ID] : []), pkg], { quiet: true });
  const notarized = env.APPLE_INSTALLER_ID ? notarize(pkg) : false;
  rmSync(cjs, { force: true });
  if (!process.argv.includes('--keep-work')) rmSync(WORK, { recursive: true, force: true });

  console.log(`\nWrote ${rel(pkg)} (${(statSync(pkg).size / 1e6).toFixed(0)} MB)`);
  if (!env.APPLE_DEVELOPER_ID || !env.APPLE_INSTALLER_ID) {
    console.log('UNSIGNED: no APPLE_DEVELOPER_ID / APPLE_INSTALLER_ID. macOS blocks it when downloaded:');
    console.log('  right-click the .pkg > Open (macOS 15+: System Settings > Privacy & Security > Open Anyway). See docs/release.md.');
  } else {
    console.log(`Signed (${env.APPLE_DEVELOPER_ID}; ${env.APPLE_INSTALLER_ID})${notarized ? ', notarized and stapled' : ', NOT notarized (no APPLE_NOTARY_PROFILE or APPLE_ID/APPLE_TEAM_ID/APPLE_APP_PASSWORD)'}`);
  }
  if (existsSync(join(OUT, 'work'))) console.log(`Work files kept in ${rel(WORK)}`);
}

if (/build-mac\.ts$/.test(process.argv[1] ?? '')) await main();
