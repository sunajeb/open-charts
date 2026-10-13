// The installers (docs/release.md): the Mac .pkg (scripts/build-mac.ts, src/addin/mac-install.ts),
// the Windows MSIs (scripts/build-msi.ts, addin/windows/open-charts.wxs) and the update check
// (src/update.ts), as far as they can be checked without building them. The builds themselves:
// `npm run build:mac && npm run check:mac-pkg` on a Mac, the windows job in CI.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { distributionXml, INSTALL_LOCATION, minMacOS, noRootAuth } from '../scripts/build-mac.ts';
import { msiName, msiVersion, rtf, wixArgs } from '../scripts/build-msi.ts';
import { assetFiles } from '../scripts/build-win.ts';
import { PKG_ID, registerAgentPlist, REGISTER_AGENT } from '../src/addin/mac-install.ts';
import { launcherScript } from '../src/addin/launcher.ts';
import { checkForUpdate, compareVersions, installerAsset, lastUpdateCheck, maybeCheckForUpdate, writeSettings } from '../src/update.ts';

const read = (f: string) => readFileSync(new URL(`../${f}`, import.meta.url), 'utf8');
const mac = process.platform === 'darwin';

test('Mac package: per-user domain only, no root authorization, installs where addin install does', () => {
  const d = distributionXml('1.2.3', 'c.pkg', '13.5');
  assert.match(d, /<domains enable_anywhere="false" enable_currentUserHome="true" enable_localSystem="false"\/>/);
  assert.match(d, new RegExp(`<pkg-ref id="${PKG_ID}" version="1.2.3" onConclusion="none">c.pkg</pkg-ref>`));
  assert.match(d, /<os-version min="13.5"\/>/);
  assert.equal(minMacOS(24), '13.5');
  assert.equal(minMacOS(22), '11.0');
  assert.equal(INSTALL_LOCATION, '/Library/Application Support/open-charts', 'under the home: addinHome()');
  assert.equal(noRootAuth('<pkg-info format-version="2" auth="root" relocatable="false">'), '<pkg-info format-version="2" auth="none" relocatable="false">');
  assert.throws(() => noRootAuth('<pkg-info format-version="2">'), /no auth attribute/);
  // The engine executable unpacks the bridge template too.
  assert.ok(assetFiles().includes('addin/applescript/open-charts.applescript'));
  const post = read('addin/mac/postinstall');
  assert.match(post, /"\$exe" addin install/);
  assert.doesNotMatch(post, /sudo|\bnode\b/);
});

test('Mac package: launcher and bridge run the engine executable, not Node', { skip: !mac && 'Mac only' }, () => {
  const dir = mkdtempSync(join(tmpdir(), 'oc-pkgtest-'));
  try {
    const exe = join(dir, "open-charts's engine");
    writeFileSync(exe, '#!/bin/sh\necho "engine $*"\n', { mode: 0o755 });
    const launcher = join(dir, 'launcher');
    writeFileSync(launcher, launcherScript('', exe), { mode: 0o755 });
    assert.equal(execFileSync(launcher, ['addin', 'types'], { encoding: 'utf8', env: { PATH: '/usr/bin:/bin' } }).trim(), 'engine addin types');
    // The bridge's engineCommand(): the quoted executable when installed from the package, else node + CLI.
    const bridge = read('addin/applescript/open-charts.applescript');
    const call = (node: string, cli: string) => {
      const src = join(dir, 'b.applescript'), scpt = join(dir, 'b.scpt');
      writeFileSync(src, bridge.replace('__NODE__', node).replace('__CLI__', cli).replace('__HOME__', dir));
      execFileSync('osacompile', ['-o', scpt, src]);
      return execFileSync('osascript', ['-e', `set s to load script (POSIX file ${JSON.stringify(scpt)})`, '-e', 'tell s to engineCommand()'], { encoding: 'utf8' }).trim();
    };
    assert.equal(call('', '/Users/a b/engine/open-charts'), "'/Users/a b/engine/open-charts'");
    assert.equal(call('/bin/sh', '/x/cli.ts'), "'/bin/sh' --experimental-strip-types --no-warnings '/x/cli.ts'");
    // The registration agent: a valid plist that only starts the engine while PowerPoint runs.
    const plist = join(dir, 'agent.plist');
    writeFileSync(plist, registerAgentPlist(exe));
    execFileSync('plutil', ['-lint', plist], { stdio: 'ignore' });
    const json = JSON.parse(execFileSync('plutil', ['-convert', 'json', '-o', '-', plist], { encoding: 'utf8' }));
    assert.equal(json.Label, REGISTER_AGENT);
    assert.equal(json.StartInterval, 30);
    const sh = json.ProgramArguments[2];
    assert.match(sh, /^\/usr\/bin\/pgrep -xq 'Microsoft PowerPoint' \|\| exit 0;/);
    assert.ok(sh.endsWith(`exec '${exe.replace(/'/g, `'\\''`)}' addin register`), sh);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('Windows MSI: per-user package matches install.ps1; machine package runs it per user', () => {
  assert.equal(msiVersion('0.2.0'), '0.2.0');
  assert.equal(msiVersion('1.4.10-beta.1'), '1.4.10');
  assert.throws(() => msiVersion('next'));
  assert.equal(msiName('0.2.0', 'perUser'), 'open-charts-0.2.0.msi');
  assert.equal(msiName('0.2.0', 'perMachine'), 'open-charts-0.2.0-machine.msi');
  assert.equal(rtf('a {b} \\c'), '{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Segoe UI;}}\\fs16 a \\{b\\} \\\\c}\n');
  const args = wixArgs('0.3.1', 'perMachine', 'D:\\dist');
  for (const d of ['Version=0.3.1', 'Scope=perMachine', 'ActiveSetupVersion=0,3,1']) assert.ok(args.includes(d), d);
  assert.ok(args.at(-1)!.endsWith('open-charts-0.3.1-machine.msi'));

  const wxs = read('addin/windows/open-charts.wxs');
  const [perUser, perMachine] = wxs.split('<?else?>\n    <StandardDirectory Id="ProgramFiles64Folder">');
  assert.ok(perMachine, 'two branches');
  const user = perUser.slice(perUser.indexOf('<?if $(var.Scope) = "perUser" ?>\n    <StandardDirectory'));
  // Same places as install.ps1 (test/windows.test.ts pins those).
  assert.match(user, /<StandardDirectory Id="LocalAppDataFolder">\s*<Directory Id="ENGINEDIR" Name="open-charts" \/>/);
  assert.match(user, /<StandardDirectory Id="AppDataFolder">\s*<Directory Id="MICROSOFTDIR" Name="Microsoft">\s*<Directory Id="ADDINSDIR" Name="AddIns" \/>/);
  assert.match(user, /Key="Software\\Microsoft\\Office\\16\.0\\PowerPoint\\AddIns\\open-charts" Name="Path" Type="string" Value="\[ADDINSDIR\]open-charts\.ppam"/);
  assert.match(user, /Key="Software\\Microsoft\\Office\\16\.0\\PowerPoint\\AddIns\\open-charts" Name="AutoLoad" Type="integer" Value="1"/);
  assert.match(user, /Trusted Locations\\open-charts" Name="Path" Type="string" Value="\[ADDINSDIR\]"/);
  assert.doesNotMatch(user, /HKLM|ProgramFiles/, 'the per-user package needs no admin rights');
  const install = read('addin/windows/install.ps1');
  for (const k of ['AddIns\\open-charts', 'Security\\Trusted Locations\\open-charts']) assert.ok(install.includes(`'${k}'`), k);
  for (const f of ['open-charts.exe', 'addin.json', 'uninstall.ps1', 'open-charts.ppam']) assert.ok(user.includes(`$(var.Src)\\${f}`), f);
  assert.match(wxs, /<util:CloseApplication Id="StopEngine" Target="open-charts\.exe"[^>]*TerminateProcess="1"/, 'a running engine is stopped');
  // Machine: install.ps1 for each user at sign-in (Active Setup).
  assert.match(perMachine, /Active Setup\\Installed Components\\open-charts/);
  assert.match(perMachine, /-File &quot;\[INSTALLDIR\]install\.ps1&quot; -SkipEngineCheck/);
  assert.match(install, /\[switch\]\$SkipEngineCheck/);
  for (const f of ['open-charts.exe', 'open-charts.ppam', 'addin.json', 'install.ps1', 'uninstall.ps1']) assert.ok(perMachine.includes(`$(var.Src)\\${f}`), f);
});

test('CI: the windows job builds the MSIs and installs and removes them silently', () => {
  const ci = read('.github/workflows/test.yml');
  assert.match(ci, /dotnet tool install --global wix --version 5\.0\.2/);
  assert.match(ci, /npm run build:msi/);
  assert.match(ci, /'\/i', "`"\$msi`"", '\/quiet'/);
  assert.match(ci, /'\/x', "`"\$msi`"", '\/quiet'/);
  assert.match(ci, /azure\/trusted-signing-action/);
});

test('update check: versions, installer per platform, off by default, once a day', async () => {
  assert.equal(compareVersions('0.3.0', '0.2.9'), 1);
  assert.equal(compareVersions('v0.10.0', '0.9.0'), 1);
  assert.equal(compareVersions('1.0.0-rc.1', '1.0.0'), -1);
  assert.equal(compareVersions('1.0.0', '1.0.0'), 0);
  const names = ['open-charts-0.3.0.pkg', 'open-charts-0.3.0.msi', 'open-charts-0.3.0-machine.msi', 'SHA256SUMS.txt'];
  assert.equal(installerAsset(names, 'darwin'), 'open-charts-0.3.0.pkg');
  assert.equal(installerAsset(names, 'win32'), 'open-charts-0.3.0.msi');
  assert.equal(installerAsset(names, 'linux'), undefined);

  let calls = 0;
  const release = { tag_name: 'v0.3.0', html_url: 'https://github.com/sunajeb/open-charts/releases/tag/v0.3.0', assets: names.map((name) => ({ name, browser_download_url: `https://x/${name}` })) };
  const fake = (async () => { calls++; return new Response(JSON.stringify(release), { status: 200 }); }) as typeof fetch;
  const u = await checkForUpdate({ current: '0.2.0', fetch: fake, platform: 'darwin', now: new Date('2026-10-13T00:00:00Z') });
  assert.deepEqual(u, { current: '0.2.0', latest: '0.3.0', available: true, url: release.html_url, asset: 'open-charts-0.3.0.pkg', checkedAt: '2026-10-13T00:00:00.000Z' });
  assert.equal((await checkForUpdate({ current: '0.3.0', fetch: fake })).available, false);
  assert.equal((await checkForUpdate({ current: '0.2.0', fetch: (async () => new Response(JSON.stringify({ ...release, prerelease: true }))) as typeof fetch })).available, false);
  await assert.rejects(checkForUpdate({ fetch: (async () => new Response('', { status: 404 })) as typeof fetch }), /no published release/);

  const home = mkdtempSync(join(tmpdir(), 'oc-update-'));
  try {
    calls = 0;
    const opts = { current: '0.2.0', fetch: fake, platform: 'win32' };
    assert.equal(await maybeCheckForUpdate(home, opts), undefined, 'off by default');
    assert.equal(calls, 0, 'nothing asked');
    writeSettings(home, { updateCheck: true });
    const first = await maybeCheckForUpdate(home, { ...opts, now: new Date('2026-10-13T08:00:00Z') });
    assert.equal(first?.asset, 'open-charts-0.3.0.msi');
    await maybeCheckForUpdate(home, { ...opts, now: new Date('2026-10-13T20:00:00Z') });
    assert.equal(calls, 1, 'once a day');
    await maybeCheckForUpdate(home, { ...opts, now: new Date('2026-10-14T09:00:00Z') });
    assert.equal(calls, 2);
    assert.equal(lastUpdateCheck(home)?.checkedAt, '2026-10-14T09:00:00.000Z');
    const offline = (async () => { throw new Error('offline'); }) as typeof fetch;
    assert.equal((await maybeCheckForUpdate(home, { ...opts, fetch: offline, now: new Date('2026-10-16T09:00:00Z') }))?.latest, '0.3.0', 'offline: the last result, no error');
  } finally {
    rmSync(home, { recursive: true, force: true });
  }
});
