<#
.SYNOPSIS
  open-charts acceptance run in real PowerPoint for Windows. Writes a log and slide PNGs into a zip
  to send back.

.DESCRIPTION
  Close PowerPoint (save your work first), then in the open-charts-win folder:

    powershell -ExecutionPolicy Bypass -File .\e2e-powerpoint-win.ps1

  It installs the add-in from this folder (install.ps1), starts PowerPoint through COM and, through
  the add-in's own VBA entry points (the same Dispatch the ribbon buttons use):
    - inserts every available gallery item, one per slide, and one into a content placeholder,
    - updates a chart (Update chart) after moving and shrinking it,
    - resizes charts and checks they are laid out again at the new size,
    - opens Edit data (an editor window opens in your browser; leave it), then applies edited data
      the way the editor's Apply button does, and checks the chart on the slide changed,
    - exports the deck and one PNG per slide,
  then quits PowerPoint and writes open-charts-e2e-<time>.zip (log, addin.log, PNGs, deck) next to
  this script. The run takes a few minutes; don't use the PC meanwhile.
#>
[CmdletBinding()]
param(
  [string]$Source = $PSScriptRoot,
  [string]$OutDir = '',
  [switch]$SkipInstall
)
$ErrorActionPreference = 'Stop'
# Windows PowerShell 5.1 can leave $PSScriptRoot empty in a param default (seen with -File).
if (-not $Source) { $Source = Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $OutDir) { $OutDir = Join-Path $Source ('open-charts-e2e-' + (Get-Date -Format 'yyyyMMdd-HHmmss')) }
$home_ = Join-Path $env:LOCALAPPDATA 'open-charts'
$addinLog = Join-Path $home_ 'addin.log'
$base = 'http://127.0.0.1:4174'
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$logFile = Join-Path $OutDir 'e2e.log'
$failures = New-Object System.Collections.Generic.List[string]

function Log([string]$line) { $line | Tee-Object -FilePath $logFile -Append | Write-Host }
function Failed([string]$what) { $failures.Add($what); Log "FAIL $what" }
function LogLines { if (Test-Path -LiteralPath $addinLog) { @(Get-Content -LiteralPath $addinLog) } else { @() } }

Log "open-charts Windows acceptance run, $(Get-Date -Format s)"
Log "Windows $([Environment]::OSVersion.VersionString), 64-bit OS: $([Environment]::Is64BitOperatingSystem), PowerShell $($PSVersionTable.PSVersion)"
Log "think-cell installed: $([bool](Get-ChildItem 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins', 'HKLM:\Software\Microsoft\Office\PowerPoint\Addins', 'HKLM:\Software\WOW6432Node\Microsoft\Office\PowerPoint\Addins' -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -like '*think*' }))"

if (Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue) {
  Log 'PowerPoint is running. Save your work, quit PowerPoint and run this script again.'
  exit 1
}

# 1. Install from this folder.
if (-not $SkipInstall) {
  Log '--- install.ps1'
  & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $Source 'install.ps1') -Source $Source *>&1 | ForEach-Object { Log "  $_" }
  if ($LASTEXITCODE -ne 0) { Log 'install.ps1 failed; stopping.'; exit 1 }
}
$prefix = (Get-Content -LiteralPath (Join-Path $home_ 'addin.json') -Raw | ConvertFrom-Json).prefix
$exe = Join-Path $home_ 'open-charts.exe'
Log "callback prefix $prefix"
# The engine's server answers only callers with the install's token (src/addin/token.ts).
$auth = @{ 'X-OC-Token' = ((Get-Content -LiteralPath (Join-Path $home_ 'token') -TotalCount 1) -replace '^X-OC-Token: ', '') }

# 2. Gallery items the engine can insert.
$items = @(& $exe addin types | Where-Object { $_ -match '\savailable$' } | ForEach-Object { ($_ -split '\s+')[0] })
Log "insertable items: $($items -join ', ')"
# The first-run welcome dialog would block the scripted session: count it as shown.
& $exe addin welcome | Out-Null

# 3. PowerPoint.
Log '--- PowerPoint'
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
Log "PowerPoint $($ppt.Version) build $($ppt.Build), $(if ([Environment]::Is64BitProcess) { '64' } else { '32' })-bit PowerShell"
$addin = $null
foreach ($a in $ppt.AddIns) { Log "  add-in: $($a.Name) loaded=$($a.Loaded) autoload=$($a.AutoLoad) path=$($a.FullName)"; if ($a.Name -like 'open-charts*') { $addin = $a } }
if (-not $addin) { Failed 'open-charts is not in PowerPoint''s add-ins list (registry AutoLoad did not work)' }
elseif ($addin.Loaded -ne -1) { Log 'open-charts registered but not loaded: loading it'; $addin.Loaded = -1 }
$pres = $ppt.Presentations.Add(-1)

$runMember = [Reflection.BindingFlags]::InvokeMethod
$ready = $false
$qualified = $false
function Macro([string]$proc, [string[]]$macroArgs = @()) {
  $before = (LogLines).Count
  for ($attempt = 1; ; $attempt++) {
    try {
      $r = [System.__ComObject].InvokeMember('Run', $runMember, $null, $ppt, [object[]](@("${prefix}_$proc") + $macroArgs))
      break
    } catch {
      $first = $_
      # The add-in-qualified name, in case this PowerPoint does not find add-in macros by plain name.
      try {
        $r = [System.__ComObject].InvokeMember('Run', $runMember, $null, $ppt, [object[]](@("open-charts.ppam!OpenCharts.${prefix}_$proc") + $macroArgs))
        if (-not $script:qualified) { Log '  (macros run by their add-in-qualified name)' }
        $script:qualified = $true
        break
      } catch { }
      # Right after loading, PowerPoint may still be compiling the add-in's VBA.
      if ($script:ready -or $attempt -ge 8) { throw $first }
      Start-Sleep -Seconds 3
    }
  }
  $script:ready = $true
  $new = @(LogLines | Select-Object -Skip $before)
  if (($new -join "`n") -match 'still updating a chart') { Start-Sleep -Milliseconds 500; return Macro $proc $macroArgs }
  return (($new + @([string]$r)) | Where-Object { $_ }) -join "`n"
}
function WaitLog([string]$pattern, [int]$from, [int]$ms = 20000) {
  $end = (Get-Date).AddMilliseconds($ms)
  while ((Get-Date) -lt $end) {
    $hit = @(LogLines | Select-Object -Skip $from | Where-Object { $_ -match $pattern })
    if ($hit.Count) { return $hit[0] }
    Start-Sleep -Milliseconds 100
  }
  return ''
}
function NewSlide([int]$layout = 12) { # ppLayoutBlank; 2 = ppLayoutText (title + content)
  $n = $pres.Slides.Count + 1
  [void]$pres.Slides.Add($n, $layout)
  $ppt.ActiveWindow.View.GotoSlide($n)
  return $n
}
function ChartOn([int]$n) { foreach ($s in $pres.Slides.Item($n).Shapes) { if ($s.Tags.Item('OC_V') -ne '') { return $s } }; return $null }
function SpecOf($shape) {
  $s = $shape.Tags.Item('OC_SPEC')
  if ($s -eq '') { $n = [int]('0' + $shape.Tags.Item('OC_SPEC_N')); for ($i = 0; $i -lt $n; $i++) { $s += $shape.Tags.Item("OC_SPEC_$i") } }
  return $s
}
function SpecKey([string]$spec) {
  $sha = [Security.Cryptography.SHA1]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes($spec))
  return (-join ($sha | ForEach-Object { $_.ToString('x2') })).Substring(0, 16)
}
function Ms([datetime]$t) { [int]((Get-Date) - $t).TotalMilliseconds }

try {
  # 4. Insert every item.
  Log '--- insert'
  $slideOf = @{}
  foreach ($id in $items) {
    $slideOf[$id] = NewSlide
    $t = Get-Date
    $out = Macro 'Run' @("chart_$id")
    Log "  $id ($(Ms $t) ms): $out"
    if ($out -notmatch 'RESULT \S+ OK') { Failed "insert $id" }
    elseif (-not (ChartOn $slideOf[$id])) { Failed "insert ${id}: no chart on slide $($slideOf[$id])" }
  }
  $ph = NewSlide 2
  $box = $pres.Slides.Item($ph).Shapes.Item(2)
  $want = @($box.Left, $box.Top, $box.Width, $box.Height)
  Log "  placeholder slide: $(Macro 'Run' @('chart_waterfall'))"
  $c = ChartOn $ph
  if (-not $c) { Failed 'placeholder: no chart' }
  else {
    $frame = @($c.Tags.Item('OC_FRAME') -split ',' | ForEach-Object { [double]$_ })
    for ($i = 0; $i -lt 4; $i++) { if ([math]::Abs($frame[$i] - $want[$i]) -gt 1) { Failed "placeholder: chart frame $($c.Tags.Item('OC_FRAME')) is not the placeholder $($want -join ',')"; break } }
  }

  # 5. Move + shrink, then Update chart.
  if ($slideOf['stacked']) {
    Log '--- update'
    $c = ChartOn $slideOf['stacked']
    $c.Left = 300; $c.Top = 150; $c.Width = 400; $c.Height = 300
    $ppt.ActiveWindow.View.GotoSlide($slideOf['stacked'])
    [void](Macro 'SelectChart' @('1'))
    [void](Macro 'SelectChart' @('1'))
    $t = Get-Date
    $out = Macro 'Run' @('btnUpdate')
    Log "  update ($(Ms $t) ms): $out"
    if ($out -notmatch 'RESULT btnUpdate OK') { Failed 'update' }
    $c = ChartOn $slideOf['stacked']
    if ($c -and ([math]::Abs($c.Left - 300) -gt 40 -or $c.Width -gt 420)) { Failed "update: not laid out in the moved frame (left $($c.Left), width $($c.Width))" }
  }

  # 6. Resize: shrink to half, click elsewhere -> laid out again at the new size.
  Log '--- resize'
  foreach ($id in 'line', 'clustered') {
    if (-not $slideOf[$id]) { continue }
    $ppt.ActiveWindow.View.GotoSlide($slideOf[$id])
    [void](Macro 'SelectChart' @('1'))
    [void](Macro 'ResizeChart' @('1', '0.5'))
    $before = (LogLines).Count; $t = Get-Date
    [void](Macro 'SelectChart' @('0'))
    $done = WaitLog 'RELAYOUT' $before
    if (-not $done) { [void](Macro 'SelectChart' @('1')); $done = WaitLog 'RELAYOUT' $before 5000 }
    Log "  $id x0.5 ($(Ms $t) ms): $(if ($done) { $done } else { 'no automatic re-layout' })"
    if ($done -notmatch 'RELAYOUT OK') { Failed "resize $id (selection-change event not delivered?)" }
    $c = ChartOn $slideOf[$id]
    $w = [double](($c.Tags.Item('OC_FRAME') -split ',')[2])
    if ([math]::Abs($c.Width - $w) -gt 1) { Failed "resize ${id}: group width $($c.Width) differs from OC_FRAME width $w" }
  }

  # 7. Edit data -> Apply (as the editor does) -> the chart on the slide changes.
  if ($slideOf['waterfall']) {
    Log '--- edit data + apply'
    $ppt.ActiveWindow.View.GotoSlide($slideOf['waterfall'])
    [void](Macro 'SelectChart' @('1'))
    $spec = SpecOf (ChartOn $slideOf['waterfall'])
    $id = SpecKey $spec
    $t = Get-Date
    $out = Macro 'Run' @('btnEditData')
    Log "  edit data ($(Ms $t) ms): $out"
    try { $session = Invoke-RestMethod -Uri "$base/api/session/$id" -Headers $auth -TimeoutSec 5; Log "  session $id state $($session.state)" }
    catch { Failed "edit data: no session $id on the server ($($_.Exception.Message))" }
    $chart = $spec | ConvertFrom-Json
    $chart.categories[1] = 'Edited on Windows'
    $body = ConvertTo-Json -InputObject @{ chart = $chart; state = 'applied' } -Depth 50 -Compress
    $t = Get-Date
    $rev = (Invoke-RestMethod -Method Post -Uri "$base/api/session/$id" -Headers $auth -Body ([Text.Encoding]::UTF8.GetBytes($body)) -ContentType 'application/json' -TimeoutSec 10).rev
    $s = $null
    $end = (Get-Date).AddSeconds(30)
    while ((Get-Date) -lt $end) {
      $s = Invoke-RestMethod -Uri "$base/api/session/$id" -Headers $auth -TimeoutSec 5
      if ($s.rev -gt $rev -and ($s.state -eq 'drawn' -or $s.state -eq 'error')) { break }
      Start-Sleep -Milliseconds 100
    }
    Log "  apply ($(Ms $t) ms): session $($s.state) $($s.message)"
    $applied = @(LogLines | Where-Object { $_ -match "APPLY|APPLIED" } | Select-Object -Last 3)
    $applied | ForEach-Object { Log "    $_" }
    if ($s.state -ne 'drawn') { Failed "apply: session $($s.state) $($s.message)" }
    if ((SpecOf (ChartOn $slideOf['waterfall'])) -notmatch 'Edited on Windows') { Failed 'apply: the chart on the slide does not have the edited data' }
  }

  # 8. Export.
  Log '--- export'
  Log "  $(Macro 'ExportCopy' @((Join-Path $OutDir 'deck.pptx'), 'pptx'))"
  Log "  $(Macro 'ExportCopy' @((Join-Path $OutDir 'png'), 'png'))"
  $pngs = @(Get-ChildItem -Path (Join-Path $OutDir 'png') -Filter '*.png' -ErrorAction SilentlyContinue)
  if ($pngs.Count -lt $pres.Slides.Count) { Failed "export: $($pngs.Count) PNGs for $($pres.Slides.Count) slides" }
} catch {
  Failed "stopped: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)"
} finally {
  try { $pres.Close() } catch { }
  try { $ppt.Quit() } catch { }
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ppt)
}

# 9. Result + zip.
if (Test-Path -LiteralPath $addinLog) { Copy-Item -LiteralPath $addinLog -Destination (Join-Path $OutDir 'addin.log') }
Log ''
if ($failures.Count) { Log "RESULT: $($failures.Count) failure(s):"; $failures | ForEach-Object { Log "  - $_" } } else { Log 'RESULT: all checks passed' }
$zip = "$OutDir.zip"
Compress-Archive -Path (Join-Path $OutDir '*') -DestinationPath $zip -Force
Write-Host ''
Write-Host "Send this file back: $zip" -ForegroundColor Green
exit [int]($failures.Count -gt 0)
