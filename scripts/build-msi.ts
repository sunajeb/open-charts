// `npm run build:msi` (Windows, after `npm run build:win`): the Windows Installer packages from
// dist/win, with the WiX Toolset (addin/windows/open-charts.wxs):
//   dist/win/open-charts-<version>.msi          per user, no admin rights: double-click, or
//                                               msiexec /i open-charts-<version>.msi /quiet
//   dist/win/open-charts-<version>-machine.msi  per machine for IT (Intune, as SYSTEM): files in
//                                               %ProgramFiles%\open-charts, set up for each user at sign-in
// Needs the wix command (WiX v5, a .NET tool) and its Util and UI extensions:
//   dotnet tool install --global wix --version 5.0.2
//   wix extension add --global WixToolset.Util.wixext/5.0.2 WixToolset.UI.wixext/5.0.2
// Signing is a separate step (Azure Trusted Signing in CI, docs/release.md): sign open-charts.exe
// before this runs, the .msi files after.
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = fileURLToPath(new URL('../', import.meta.url));
const DIST = join(ROOT, 'dist', 'win');
export const WIX_VERSION = '5.0.2';

/** Windows Installer versions are numeric a.b.c (no pre-release tag). */
export function msiVersion(version: string): string {
  const m = /^(\d+)\.(\d+)\.(\d+)/.exec(version);
  if (!m) throw new Error(`version ${version}: expected a.b.c`);
  return `${m[1]}.${m[2]}.${m[3]}`;
}

export const msiName = (version: string, scope: 'perUser' | 'perMachine') => `open-charts-${version}${scope === 'perMachine' ? '-machine' : ''}.msi`;

/** The license as minimal RTF, for the installer's license page. */
export function rtf(text: string): string {
  const esc = (s: string) => s.replace(/[\\{}]/g, (c) => `\\${c}`).replace(/[^\x00-\x7f]/g, (c) => `\\u${c.charCodeAt(0)}?`);
  return `{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Segoe UI;}}\\fs16 ${text.split(/\r?\n/).map(esc).join('\\par\n')}}\n`;
}

/** The `wix build` arguments for one package. */
export function wixArgs(version: string, scope: 'perUser' | 'perMachine', dist = DIST): string[] {
  const v = msiVersion(version);
  return ['build', join(ROOT, 'addin', 'windows', 'open-charts.wxs'), '-arch', 'x64',
    '-ext', 'WixToolset.Util.wixext', '-ext', 'WixToolset.UI.wixext', // added globally at WIX_VERSION
    '-d', `Version=${v}`, '-d', `Scope=${scope}`, '-d', `Src=${dist}`, '-d', `LicenseRtf=${join(dist, 'License.rtf')}`,
    '-d', `ActiveSetupVersion=${v.replace(/\./g, ',')}`,
    '-o', join(dist, msiName(version, scope))];
}

function main() {
  const version = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8')).version as string;
  for (const f of ['open-charts.exe', 'open-charts.ppam', 'addin.json', 'install.ps1', 'uninstall.ps1']) {
    if (!existsSync(join(DIST, f))) throw new Error(`dist/win/${f} is missing: run npm run build:win first.`);
  }
  writeFileSync(join(DIST, 'License.rtf'), rtf(readFileSync(join(ROOT, 'LICENSE'), 'utf8')));
  for (const scope of ['perUser', 'perMachine'] as const) {
    execFileSync('wix', wixArgs(version, scope), { stdio: 'inherit' });
    console.log(`Wrote dist/win/${msiName(version, scope)}`);
  }
}

if (/build-msi\.ts$/.test(process.argv[1] ?? '')) main();
