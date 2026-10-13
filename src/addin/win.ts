// The Windows end of the add-in. Same add-in, ribbon, VBA and tags as on the Mac; only the bridge
// differs (docs/windows.md):
//   - VBA talks to the add-in server (127.0.0.1:4174) over HTTP and starts open-charts.exe itself
//     (the engine as one Node single-executable, %LOCALAPPDATA%\open-charts\open-charts.exe);
//   - Apply: the server runs the add-in's ApplyEdit macro through COM from one long-lived Windows
//     PowerShell helper, the twin of the Mac's JavaScript-for-Automation helper (runner.ts), with
//     the same one-JSON-request-per-line protocol:
//       {"macro": "<name>", "args": ["…"]}  ->  {"out": "<the macro's return value>"} | {"error": "…"}
//       {"ping": true}                      ->  {"out": "pong"}        (no COM call)
//       {"connect": true}                   ->  {"out": "connected"}   | {"out": "not running"}
//     It attaches to a running PowerPoint (GetActiveObject) and never starts one.
import { spawn } from 'node:child_process';
import { isSea } from 'node:sea';
import { fileURLToPath } from 'node:url';
import { macroRunner } from './runner.ts';

/**
 * The PowerShell helper. Windows PowerShell 5.1 (powershell.exe, on every Windows 10/11), not
 * PowerShell 7: Marshal.GetActiveObject only exists on .NET Framework. PowerPoint answers "busy"
 * (RPC_E_CALL_REJECTED / RPC_E_SERVERCALL_RETRYLATER) while it is in the middle of something, e.g.
 * the add-in waiting for the engine: the call is retried until the server's timeout stops the helper.
 * A macro of a loaded add-in is found by its plain name; the add-in-qualified name is the fallback.
 */
export const HELPER_PS = `$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$stdin = New-Object System.IO.StreamReader([Console]::OpenStandardInput(), $utf8)
$stdout = New-Object System.IO.StreamWriter([Console]::OpenStandardOutput(), $utf8)
$stdout.AutoFlush = $true
function Reply($o) { $stdout.WriteLine((ConvertTo-Json -InputObject $o -Compress)) }
function PowerPoint { try { [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch { $null } }
function Busy($e) { $e -is [Runtime.InteropServices.COMException] -and (($e.HResult -eq -2147418111) -or ($e.HResult -eq -2147417846)) }
function RunMacro($ppt, [object[]]$params) {
  for (;;) {
    try { return [System.__ComObject].InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $ppt, $params) }
    catch {
      $e = $_.Exception; while ($e.InnerException) { $e = $e.InnerException }
      if (Busy $e) { Start-Sleep -Milliseconds 100 } else { throw $e }
    }
  }
}
while ($null -ne ($line = $stdin.ReadLine())) {
  try {
    $req = ConvertFrom-Json -InputObject $line
    if ($req.ping) { Reply @{ out = 'pong' }; continue }
    $ppt = PowerPoint
    if ($req.connect) { if ($ppt) { [void]$ppt.Name; Reply @{ out = 'connected' } } else { Reply @{ out = 'not running' } }; continue }
    if (-not $ppt) { Reply @{ error = 'execution error: PowerPoint is not running. (-600)' }; continue }
    $macroArgs = @($req.args | ForEach-Object { [string]$_ })
    try { $r = RunMacro $ppt ([object[]](@([string]$req.macro) + $macroArgs)) }
    catch {
      $first = $_
      try { $r = RunMacro $ppt ([object[]](@('open-charts.ppam!OpenCharts.' + $req.macro) + $macroArgs)) } catch { throw $first }
    }
    Reply @{ out = [string]$r }
  } catch {
    $e = $_.Exception; while ($e.InnerException) { $e = $e.InnerException }
    Reply @{ error = 'execution error: ' + $e.Message.Trim() + ' (' + $e.HResult + ')' }
  }
}
`;

/** A script for powershell.exe -EncodedCommand (UTF-16LE, base64): stdin stays free for requests. */
export const encodeCommand = (script: string) => Buffer.from(script, 'utf16le').toString('base64');

export const POWERSHELL_ARGS = ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand'];

/** The Apply helper on Windows. */
export const powershellRunner = (command = 'powershell.exe') => macroRunner(command, [...POWERSHELL_ARGS, encodeCommand(HELPER_PS)]);

/** True in open-charts.exe (and in the bundle, run by plain node for tests). */
export const bundled = () => isSea() || typeof (globalThis as { __OC_BUNDLE__?: string }).__OC_BUNDLE__ === 'string';

/** The command that runs `open-charts <args>` with this engine: the exe itself, the bundle, or the TypeScript CLI. */
export function selfCommand(cli: URL, args: string[]): [string, string[]] {
  if (isSea()) return [process.execPath, args];
  const bundle = (globalThis as { __OC_BUNDLE__?: string }).__OC_BUNDLE__;
  if (bundle) return [process.execPath, [bundle, ...args]];
  return [process.execPath, ['--experimental-strip-types', '--no-warnings', fileURLToPath(cli), ...args]];
}

/** Opens a URL in the default browser (the Edit data window). rundll32 shows no console window. */
export const openCommand = (url: string, platform = process.platform): [string, string[]] =>
  platform === 'win32' ? ['rundll32.exe', ['url.dll,FileProtocolHandler', url]] : ['open', [url]];

export function openUrl(url: string) {
  const [cmd, args] = openCommand(url);
  spawn(cmd, args, { stdio: 'ignore', detached: true, windowsHide: true }).unref();
}
