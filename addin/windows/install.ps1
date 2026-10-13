<#
.SYNOPSIS
  Installs the open-charts PowerPoint add-in for the current Windows user (no admin rights needed).

.DESCRIPTION
  Run from the open-charts-win folder (open-charts.exe, open-charts.ppam, addin.json next to this
  script), with PowerPoint closed:

    powershell -ExecutionPolicy Bypass -File .\install.ps1

  It
    1. stops a running open-charts engine (so its file can be replaced),
    2. copies open-charts.exe (the engine, no Node needed) to %LOCALAPPDATA%\open-charts,
    3. copies open-charts.ppam to %APPDATA%\Microsoft\AddIns (PowerPoint's add-ins folder),
    4. removes the "downloaded from the internet" mark from both (Office blocks macros in marked files),
       and creates the engine's secret token (%LOCALAPPDATA%\open-charts\token, kept across installs:
       the local server answers only callers that can read it),
    5. registers the add-in to load with PowerPoint:
         HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\open-charts  Path, AutoLoad=1
    6. makes sure the add-ins folder is a trusted location for PowerPoint (HKCU, unless your
       organisation's policy manages trusted locations),
    7. checks that the engine starts on this PC.
  Start PowerPoint afterwards: the open-charts tab appears next to Home.

  -DryRun prints what would be done and changes nothing. The other parameters exist for tests.
#>
[CmdletBinding()]
param(
  [string]$Source = $PSScriptRoot,
  [string]$EngineDir = (Join-Path $env:LOCALAPPDATA 'open-charts'),
  [string]$AddInsDir = (Join-Path $env:APPDATA 'Microsoft\AddIns'),
  [string]$OfficeKey = 'HKCU:\Software\Microsoft\Office\16.0\PowerPoint',
  [switch]$DryRun,
  [switch]$SkipEngineCheck
)
$ErrorActionPreference = 'Stop'
# Windows PowerShell 5.1 can leave $PSScriptRoot empty in a param default (seen with -File).
if (-not $Source) { $Source = Split-Path -Parent $MyInvocation.MyCommand.Path }
Set-StrictMode -Version 2

$AddInKey = Join-Path $OfficeKey 'AddIns\open-charts'
$TrustKey = Join-Path $OfficeKey 'Security\Trusted Locations\open-charts'
$exe = Join-Path $EngineDir 'open-charts.exe'
$ppam = Join-Path $AddInsDir 'open-charts.ppam'
$tokenFile = Join-Path $EngineDir 'token'

function Step([string]$what) { if ($DryRun) { Write-Host "[dry run] $what" } else { Write-Host $what } }
function Fail([string]$why) { Write-Host ''; Write-Host "open-charts was not installed: $why" -ForegroundColor Red; exit 1 }

foreach ($f in 'open-charts.exe', 'open-charts.ppam', 'addin.json') {
  if (-not (Test-Path -LiteralPath (Join-Path $Source $f))) { Fail "$f is missing in $Source. Run install.ps1 from the open-charts-win folder (unzip it first)." }
}
$prefix = (Get-Content -LiteralPath (Join-Path $Source 'addin.json') -Raw | ConvertFrom-Json).prefix
if (-not $prefix) { Fail "addin.json in $Source has no callback prefix." }

# 1. Stop a running engine: Windows does not let a running exe be overwritten.
Step 'Stopping a running open-charts engine (if any)'
if (-not $DryRun) {
  $headers = @{}
  if (Test-Path -LiteralPath $tokenFile) {
    $line = (Get-Content -LiteralPath $tokenFile -TotalCount 1)
    if ($line -match '^X-OC-Token: ([0-9a-f]{64})') { $headers['X-OC-Token'] = $Matches[1] }
  }
  try { Invoke-WebRequest -UseBasicParsing -Method Post -Uri 'http://127.0.0.1:4174/api/addin/quit' -Headers $headers -TimeoutSec 2 | Out-Null } catch { }
  for ($i = 0; $i -lt 25; $i++) {
    $running = @(Get-Process -Name 'open-charts' -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq $exe })
    if ($running.Count -eq 0) { break }
    if ($i -ge 10) { $running | Stop-Process -Force -ErrorAction SilentlyContinue } # not gone after 2 s
    Start-Sleep -Milliseconds 200
  }
}

# 2.-4. Files.
Step "Copying the engine to $exe"
Step "Copying the add-in to $ppam"
if (-not $DryRun) {
  New-Item -ItemType Directory -Force -Path $EngineDir, $AddInsDir | Out-Null
  try {
    Copy-Item -LiteralPath (Join-Path $Source 'open-charts.exe') -Destination $exe -Force
    Copy-Item -LiteralPath (Join-Path $Source 'open-charts.ppam') -Destination $ppam -Force
  } catch {
    Fail "a file could not be replaced ($($_.Exception.Message)). Close PowerPoint and run install.ps1 again."
  }
  foreach ($f in 'uninstall.ps1', 'e2e-powerpoint-win.ps1') {
    $p = Join-Path $Source $f
    if (Test-Path -LiteralPath $p) { Copy-Item -LiteralPath $p -Destination (Join-Path $EngineDir $f) -Force }
  }
  Get-ChildItem -LiteralPath $EngineDir -File | Unblock-File
  Unblock-File -LiteralPath $ppam
  # The engine's token (src/addin/token.ts): 32 random bytes as hex, in the header line VBA sends.
  # %LOCALAPPDATA% is readable by this user only; an existing token is kept.
  $valid = (Test-Path -LiteralPath $tokenFile) -and ((Get-Content -LiteralPath $tokenFile -TotalCount 1) -match '^X-OC-Token: [0-9a-f]{64}$')
  if (-not $valid) {
    $bytes = New-Object byte[] 32
    $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
    $rng.GetBytes($bytes)
    $rng.Dispose()
    $hex = -join ($bytes | ForEach-Object { $_.ToString('x2') })
    [IO.File]::WriteAllText($tokenFile, "X-OC-Token: $hex`n", (New-Object System.Text.UTF8Encoding($false)))
  }
  $config = [ordered]@{ prefix = $prefix; engine = $exe; ppam = $ppam }
  [IO.File]::WriteAllText((Join-Path $EngineDir 'addin.json'), ((ConvertTo-Json -InputObject $config) + "`n"))
}

# 5. Load with PowerPoint.
Step "Registering the add-in: $AddInKey (Path, AutoLoad=1)"
if (-not $DryRun) {
  New-Item -Path $AddInKey -Force | Out-Null
  New-ItemProperty -Path $AddInKey -Name 'Path' -Value $ppam -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $AddInKey -Name 'AutoLoad' -Value 1 -PropertyType DWord -Force | Out-Null
}

# 6. Trusted location for the add-ins folder (PowerPoint usually trusts it already).
$trusted = $false
$locations = Join-Path $OfficeKey 'Security\Trusted Locations'
if (Test-Path $locations) {
  foreach ($k in Get-ChildItem $locations) {
    $p = (Get-ItemProperty -Path $k.PSPath -Name 'Path' -ErrorAction SilentlyContinue)
    if ($p -and ([Environment]::ExpandEnvironmentVariables($p.Path).TrimEnd('\') -eq $AddInsDir.TrimEnd('\'))) { $trusted = $true }
  }
}
if ($trusted) {
  Step "$AddInsDir is already a trusted location for PowerPoint"
} else {
  Step "Adding $AddInsDir as a trusted location for PowerPoint ($TrustKey)"
  if (-not $DryRun) {
    New-Item -Path $TrustKey -Force | Out-Null
    New-ItemProperty -Path $TrustKey -Name 'Path' -Value ($AddInsDir.TrimEnd('\') + '\') -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $TrustKey -Name 'Description' -Value 'open-charts add-in' -PropertyType String -Force | Out-Null
  }
}
$policy = 'HKCU:\Software\Policies\Microsoft\Office\16.0\PowerPoint\Security'
if (Test-Path $policy) {
  $p = Get-ItemProperty -Path $policy -ErrorAction SilentlyContinue
  if ($p -and ($p.PSObject.Properties.Name -contains 'VBAWarnings') -and $p.VBAWarnings -eq 4) {
    Write-Host 'Note: your organisation disables all macros in PowerPoint (policy VBAWarnings=4), so the add-in cannot run. Ask IT to allow macros from trusted locations.' -ForegroundColor Yellow
  }
}

# 7. The engine starts on this PC (application control can block unsigned programs).
if (-not $DryRun -and -not $SkipEngineCheck) {
  Step 'Checking that the engine starts'
  try {
    $out = & $exe addin types 2>&1 | Out-String
  } catch {
    Fail "open-charts.exe could not be started ($($_.Exception.Message)). If your PC only allows approved programs (AppLocker, Defender Application Control), ask IT to allow $exe."
  }
  if ($out -notmatch 'available') { Fail "open-charts.exe did not answer as expected: $out" }
}

Write-Host ''
Write-Host 'open-charts is installed. Start PowerPoint (restart it if it is open): the open-charts tab is next to Home.' -ForegroundColor Green
if (Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue) { Write-Host 'PowerPoint is running: quit and start it again to load the add-in.' -ForegroundColor Yellow }
