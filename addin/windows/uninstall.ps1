<#
.SYNOPSIS
  Removes the open-charts PowerPoint add-in for the current Windows user (no admin rights needed).

.DESCRIPTION
    powershell -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\open-charts\uninstall.ps1"

  Stops the engine, removes the add-in registration and the trusted location install.ps1 added,
  open-charts.ppam from the add-ins folder and %LOCALAPPDATA%\open-charts. Charts already on slides
  stay as they are (plain PowerPoint shapes). -DryRun prints what would be done.
#>
[CmdletBinding()]
param(
  [string]$EngineDir = (Join-Path $env:LOCALAPPDATA 'open-charts'),
  [string]$AddInsDir = (Join-Path $env:APPDATA 'Microsoft\AddIns'),
  [string]$OfficeKey = 'HKCU:\Software\Microsoft\Office\16.0\PowerPoint',
  [switch]$DryRun
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2

$exe = Join-Path $EngineDir 'open-charts.exe'
function Step([string]$what) { if ($DryRun) { Write-Host "[dry run] $what" } else { Write-Host $what } }

Step 'Stopping the open-charts engine (if running)'
if (-not $DryRun) {
  $headers = @{}
  $tokenFile = Join-Path $EngineDir 'token'
  if ((Test-Path -LiteralPath $tokenFile) -and ((Get-Content -LiteralPath $tokenFile -TotalCount 1) -match '^X-OC-Token: ([0-9a-f]{64})')) { $headers['X-OC-Token'] = $Matches[1] }
  try { Invoke-WebRequest -UseBasicParsing -Method Post -Uri 'http://127.0.0.1:4174/api/addin/quit' -Headers $headers -TimeoutSec 2 | Out-Null } catch { }
  Start-Sleep -Milliseconds 500
  Get-Process -Name 'open-charts' -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq $exe } | Stop-Process -Force -ErrorAction SilentlyContinue
}

foreach ($key in (Join-Path $OfficeKey 'AddIns\open-charts'), (Join-Path $OfficeKey 'Security\Trusted Locations\open-charts')) {
  if (Test-Path $key) {
    Step "Removing $key"
    if (-not $DryRun) { Remove-Item -Path $key -Recurse -Force }
  }
}

$ppam = Join-Path $AddInsDir 'open-charts.ppam'
foreach ($p in $ppam, $EngineDir) {
  if (Test-Path -LiteralPath $p) {
    Step "Removing $p"
    if (-not $DryRun) {
      try { Remove-Item -LiteralPath $p -Recurse -Force }
      catch { Write-Host "Could not remove $p ($($_.Exception.Message)). Close PowerPoint and run uninstall.ps1 again." -ForegroundColor Yellow; exit 1 }
    }
  }
}
Write-Host 'open-charts add-in removed. Restart PowerPoint if it is open.' -ForegroundColor Green
