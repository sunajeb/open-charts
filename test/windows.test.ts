// The Windows end of the add-in (src/addin/win.ts, addin/windows/*.ps1, the #Else branches of the
// VBA, scripts/build-win.ts), as far as it can be tested without PowerPoint. On Windows (CI) the
// PowerShell parts also run for real.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { macroRunner } from '../src/addin/runner.ts';
import { applyErrorMessage } from '../src/addin/server.ts';
import { encodeCommand, HELPER_PS, openCommand, POWERSHELL_ARGS, powershellRunner, selfCommand } from '../src/addin/win.ts';
import { assetFiles, banner, replaceImportMeta, shasumFor, stripSignature, wrapTopLevelAwait } from '../scripts/build-win.ts';

const read = (p: string) => readFileSync(new URL(`../${p}`, import.meta.url), 'utf8');
const win = process.platform === 'win32';

/** The VBA as compiled with the given constants: #If/#ElseIf/#Else/#End If on Mac, VBA7, Win64. */
export function preprocess(src: string, consts: Record<string, boolean>): string {
  const out: string[] = [];
  const stack: { active: boolean; taken: boolean; parent: boolean }[] = [];
  const cond = (expr: string) => {
    const e = expr.trim();
    const not = /^Not\s+/i.test(e);
    const v = !!consts[e.replace(/^Not\s+/i, '')];
    return not ? !v : v;
  };
  const active = () => stack.every((s) => s.active);
  for (const line of src.split(/\r?\n/)) {
    let m;
    if ((m = /^\s*#If\s+(.+?)\s+Then\s*$/i.exec(line))) {
      const parent = active(), c = cond(m[1]);
      stack.push({ active: c, taken: c, parent });
    } else if ((m = /^\s*#ElseIf\s+(.+?)\s+Then\s*$/i.exec(line))) {
      const s = stack.at(-1)!; const c = !s.taken && cond(m[1]); s.active = c; s.taken ||= c;
    } else if (/^\s*#Else\s*$/i.test(line)) {
      const s = stack.at(-1)!; s.active = !s.taken; s.taken = true;
    } else if (/^\s*#End If\s*$/i.test(line)) {
      assert.ok(stack.pop(), 'unbalanced #End If');
    } else if (active()) out.push(line);
  }
  assert.equal(stack.length, 0, 'unclosed #If');
  return out.join('\n');
}

const bas = read('addin/vba/OpenCharts.bas');
const mac = preprocess(bas, { Mac: true, VBA7: true });
const win64 = preprocess(bas, { Mac: false, VBA7: true, Win64: true });
const win32 = preprocess(bas, { Mac: false, VBA7: true, Win64: false });

test('VBA: the Mac compiles the AppleScriptTask bridge only, Windows the HTTP bridge only', () => {
  for (const id of ['WinEngine', 'WinHttp', 'WinLog', 'WinStartServer', 'ServerXMLHTTP', 'Declare', 'Shell ', 'LOCALAPPDATA']) {
    assert.ok(!mac.includes(id), `Mac source has ${id}`);
  }
  const code = (v: string) => v.split('\n').filter((l) => !/^\s*'/.test(l)).join('\n');
  assert.ok((code(mac).match(/AppleScriptTask\b/g) ?? []).length >= 5, 'startServer, copyOut, engine, logLine, elementSelection');
  assert.doesNotMatch(code(win64), /AppleScriptTask\b/, 'no AppleScriptTask on Windows: it would not compile');
  assert.ok(!win64.includes('TMPDIR'));
  // Every Windows bridge procedure that is called is defined, and they are closed.
  for (const fn of ['WinEngine', 'WinHttp', 'WinLog', 'WinStartServer']) {
    assert.match(win64, new RegExp(`Private (Function|Sub) ${fn}\\(`));
    assert.ok((win64.match(new RegExp(`\\b${fn}\\b`, 'g')) ?? []).length > 1, `${fn} is used`);
  }
  for (const v of [mac, win64, win32]) {
    assert.equal((v.match(/^\s*(Public |Private )?(Sub|Function) /gm) ?? []).length, (v.match(/^\s*End (Sub|Function)\b/gm) ?? []).length);
  }
});

test('VBA: Windows declares compile on 32- and 64-bit Office, 127.0.0.1 only, no proxy, engine hidden', () => {
  assert.match(win64, /^Private Declare PtrSafe Sub Sleep Lib "kernel32" \(ByVal dwMilliseconds As Long\)$/m);
  assert.equal(preprocess(bas, { Mac: false, VBA7: false }).match(/Declare (PtrSafe )?Sub Sleep/g)?.join(), 'Declare Sub Sleep', 'pre-VBA7 Office: no PtrSafe');
  // Declarations come before the first procedure (VBA requires it).
  assert.ok(win64.indexOf('Declare PtrSafe') < win64.search(/^(Public|Private) (Sub|Function) /m));
  assert.match(win64, /Private Const ENGINE_URL As String = "http:\/\/127\.0\.0\.1:4174\/api\/addin\/"/);
  assert.match(win64, /http\.setProxy 1/);
  assert.match(win64, /Shell """" & exe & """ addin serve", vbHide/);
  assert.match(win64, /Private Const ENGINE_EXE As String = "\\open-charts\\open-charts\.exe"/);
  // The same messages and silent Apply: only the reinstall hint differs.
  assert.match(win64, /"open-charts is not responding \(" & Err\.Description & "\)\." & vbLf & vbLf & WIN_REINSTALL/);
  assert.match(win64, /Report "EXPORT OK " & dest/);
  // ApplyEdit (run by the server through COM) is the same code on both.
  const body = (src: string) => src.slice(src.indexOf('Public Function OCPREFIX_ApplyEdit'), src.indexOf('Private Function Fmt3'));
  assert.equal(body(win64), body(mac));
});

test('VBA: removing the Windows branches gives back exactly the Mac code', () => {
  // Lines only Windows compiles, and the directives, are all that was added.
  const macLines = mac.split('\n');
  const plain = bas.split(/\r?\n/).filter((l) => !/^\s*#/.test(l));
  let i = 0;
  for (const l of plain) if (i < macLines.length && l === macLines[i]) i++;
  assert.equal(i, macLines.length, 'every Mac line, in order');
});

test('PowerShell helper: never starts PowerPoint, same answers as the Mac helper, fits a command line', () => {
  assert.match(HELPER_PS, /\[Runtime\.InteropServices\.Marshal\]::GetActiveObject\('PowerPoint\.Application'\)/);
  assert.doesNotMatch(HELPER_PS, /New-Object\s+-ComObject|CreateObject|Start-Process/i);
  assert.match(HELPER_PS, /'execution error: PowerPoint is not running\. \(-600\)'/);
  assert.match(HELPER_PS, /InvokeMember\('Run'/);
  assert.match(HELPER_PS, /-2147418111/, 'RPC_E_CALL_REJECTED is retried');
  assert.ok(/^[\x09\x0a\x0d\x20-\x7e]*$/.test(HELPER_PS), 'ASCII');
  const encoded = encodeCommand(HELPER_PS);
  assert.equal(Buffer.from(encoded, 'base64').toString('utf16le'), HELPER_PS);
  assert.ok(['powershell.exe', ...POWERSHELL_ARGS, encoded].join(' ').length < 32_000, 'under the Windows command-line limit');
  assert.deepEqual(POWERSHELL_ARGS.slice(-1), ['-EncodedCommand']);
  // Apply errors from the helper read as on the Mac; a timeout names what blocks PowerPoint on Windows.
  assert.match(applyErrorMessage({ message: 'execution error: PowerPoint is not running. (-600)' }, 'win32'), /PowerPoint isn't running/);
  assert.match(applyErrorMessage({ message: 'x', killed: true }, 'win32'), /Close any open dialog in PowerPoint/);
  assert.match(applyErrorMessage({ message: 'x', killed: true }, 'darwin'), /click Allow/);
  assert.match(applyErrorMessage({ message: 'execution error: Call was rejected by callee. (-2147418111)' }, 'win32'), /could not update the chart \(Call was rejected/);
});

test('macro runner protocol (fake helper): one request per line, timeouts restart the helper', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'oc-win-runner-'));
  try {
    const n = join(dir, 'n');
    writeFileSync(n, '0');
    // A stand-in for the PowerShell helper, in node so it runs everywhere: the first instance hangs.
    const fake = `const fs = require('fs'); const c = +fs.readFileSync(${JSON.stringify(n)}, 'utf8'); fs.writeFileSync(${JSON.stringify(n)}, String(c + 1));
if (c === 0) setInterval(() => {}, 1000);
else require('readline').createInterface({ input: process.stdin }).on('line', (l) => {
  const r = JSON.parse(l);
  const reply = r.ping ? { out: 'pong' } : r.connect ? { out: 'not running' } : r.macro === 'boom' ? { error: 'execution error: Call was rejected by callee. (-2147418111)' } : { out: r.macro + '|' + r.args.join('|') };
  process.stdout.write(JSON.stringify(reply) + '\\n');
});`;
    const runner = macroRunner(process.execPath, ['-e', fake]);
    // Long enough for node to start the first helper even on a loaded machine (it must record its run).
    const hung = await runner.run('x', [], 3000);
    assert.ok('error' in hung && hung.error.killed, 'a helper that never answers times out');
    assert.equal(await runner.ping(10_000), true, 'and a fresh helper takes the next request');
    assert.deepEqual(await runner.connect(5000), { out: 'not running' });
    assert.deepEqual(await runner.run('OC1_ApplyEdit', ['id', 'a\tb\nc', 'é ✓'], 5000), { out: 'OC1_ApplyEdit|id|a\tb\nc|é ✓' }, 'tabs, newlines and Unicode survive');
    assert.deepEqual(await runner.run('boom', [], 5000), { error: { message: 'execution error: Call was rejected by callee. (-2147418111)' } });
    runner.stop();
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('the real PowerShell helper answers without PowerPoint (Windows)', { skip: !win }, async () => {
  const runner = powershellRunner();
  try {
    assert.equal(await runner.ping(60_000), true);
    const running = (() => { try { return /POWERPNT/i.test(execFileSync('tasklist', ['/FI', 'IMAGENAME eq POWERPNT.EXE'], { encoding: 'utf8' })); } catch { return false; } })();
    if (!running) {
      assert.deepEqual(await runner.connect(30_000), { out: 'not running' });
      assert.deepEqual(await runner.run('OC_ApplyEdit', ['a', 'b', 'c'], 30_000), { error: { message: 'execution error: PowerPoint is not running. (-600)' } });
    }
  } finally {
    runner.stop();
  }
});

test('the engine starts itself as the exe, the bundle or the TypeScript CLI; the browser opens without a console', () => {
  const cli = new URL('../src/cli.ts', import.meta.url);
  const [cmd, args] = selfCommand(cli, ['addin', 'serve']);
  assert.equal(cmd, process.execPath);
  assert.deepEqual(args, ['--experimental-strip-types', '--no-warnings', fileURLToPath(cli), 'addin', 'serve']);
  assert.deepEqual(openCommand('http://localhost:4174/?session=a', 'win32'), ['rundll32.exe', ['url.dll,FileProtocolHandler', 'http://localhost:4174/?session=a']]);
  assert.deepEqual(openCommand('http://x', 'darwin'), ['open', ['http://x']]);
});

test('build-win: top-level await wrapped, import.meta.url per source file, assets, checksums, signature stripped', () => {
  const cli = wrapTopLevelAwait(read('src/cli.ts'));
  assert.ok(!cli.startsWith('#!'));
  const body = cli.slice(cli.indexOf('(async () => {'));
  assert.ok(cli.indexOf('import { parseArgs }') < cli.indexOf('(async () => {'), 'imports stay at the top');
  assert.doesNotMatch(body, /^import .* from '/m);
  assert.match(cli, /\}\)\(\)\.catch\(\(err\) => \{ console\.error\(err\); process\.exit\(1\); \}\);\n$/);
  assert.equal(replaceImportMeta("new URL('../web/editor.html', import.meta.url)", 'src/addin/server.ts'), 'new URL(\'../web/editor.html\', __ocUrl("src/addin/server.ts"))');
  const assets = assetFiles();
  for (const a of ['src/web/editor.html', 'addin/vba/OpenCharts.bas', 'addin/vba/AppEvents.cls']) assert.ok(assets.includes(a), a);
  assert.ok(assets.every((a) => !a.includes('\\')), 'asset keys use /');
  const b = banner(assets, 'abc');
  assert.match(b, /process\.env\.LOCALAPPDATA/);
  assert.match(b, /getRawAsset\(key\)/);
  new Function('require', '__filename', '__dirname', b); // valid JavaScript

  const sums = 'aa'.repeat(32) + '  win-x64/node.exe\n' + 'bb'.repeat(32) + '  win-x64/node.lib\n';
  assert.equal(shasumFor(sums, 'win-x64/node.exe'), 'aa'.repeat(32));
  assert.equal(shasumFor(sums, 'win-x86/node.exe'), undefined);

  // A minimal PE32+ with a 16-byte certificate table at its end.
  const pe = Buffer.alloc(0x200 + 16);
  pe.writeUInt32LE(0x80, 0x3c);
  pe.writeUInt32LE(0x4550, 0x80);
  pe.writeUInt16LE(0x20b, 0x80 + 24);
  const entry = 0x80 + 24 + 112 + 32;
  pe.writeUInt32LE(0x200, entry); pe.writeUInt32LE(16, entry + 4);
  const out = stripSignature(Buffer.from(pe));
  assert.equal(out.length, 0x200);
  assert.equal(out.readUInt32LE(entry), 0);
  assert.equal(out.readUInt32LE(entry + 4), 0);
  assert.throws(() => stripSignature(Buffer.alloc(0x100)), /not a PE/);
});

const ps1 = ['addin/windows/install.ps1', 'addin/windows/uninstall.ps1', 'tools/e2e-powerpoint-win.ps1'];

test('installer scripts: per-user only, autoload registration, unblocked files, ASCII for Windows PowerShell 5.1', () => {
  for (const f of ps1) {
    const s = read(f);
    assert.ok(/^[\x09\x0a\x0d\x20-\x7e]*$/.test(s), `${f}: ASCII (PowerShell 5.1 reads BOM-less scripts as ANSI)`);
    assert.doesNotMatch(s, /HKLM:\\[^']*'\s*-(Value|Force)|New-Item\s+-Path\s+'?HKLM|RunAs/i, `${f}: no admin`);
    assert.doesNotMatch(s, /\$\w+:[^\\\w{]/, `${f}: no "$name:" in strings (PowerShell reads a drive)`);
  }
  const install = read('addin/windows/install.ps1');
  assert.match(install, /\$OfficeKey = 'HKCU:\\Software\\Microsoft\\Office\\16\.0\\PowerPoint'/);
  assert.match(install, /Join-Path \$OfficeKey 'AddIns\\open-charts'/);
  assert.match(install, /-Name 'Path' -Value \$ppam -PropertyType String/);
  assert.match(install, /-Name 'AutoLoad' -Value 1 -PropertyType DWord/);
  assert.match(install, /Join-Path \$env:LOCALAPPDATA 'open-charts'/);
  assert.match(install, /Join-Path \$env:APPDATA 'Microsoft\\AddIns'/);
  assert.match(install, /Unblock-File -LiteralPath \$ppam/);
  assert.match(install, /api\/addin\/quit/, 'stops the running engine before replacing it');
  assert.match(install, /Trusted Locations\\open-charts/);
  const uninstall = read('addin/windows/uninstall.ps1');
  for (const k of ["'AddIns\\open-charts'", "'Security\\Trusted Locations\\open-charts'"]) assert.ok(uninstall.includes(k), k);
});

test('PowerShell scripts parse in Windows PowerShell (Windows)', { skip: !win }, () => {
  for (const f of ps1) {
    const file = fileURLToPath(new URL(`../${f}`, import.meta.url));
    const errors = execFileSync('powershell.exe', ['-NoProfile', '-Command',
      `$e = $null; [void][Management.Automation.Language.Parser]::ParseFile('${file}', [ref]$null, [ref]$e); $e | ForEach-Object { $_.Message }`], { encoding: 'utf8' }).trim();
    assert.equal(errors, '', `${f}: ${errors}`);
  }
});
