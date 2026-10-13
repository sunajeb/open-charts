# open-charts on Windows

The PowerPoint add-in works on Windows PCs as well as on the Mac. It is **one add-in**: the same
`open-charts.ppam`, ribbon, VBA module and chart tags. A chart made on a Mac updates and edits on a
PC, and the other way round. Only the bridge between the VBA and the chart engine is different.

Install steps are in the [README](../README.md#install-on-windows). Troubleshooting is in the
[user guide](USER-GUIDE.md#windows).

## How it works

| | Mac | Windows |
|---|---|---|
| Engine | Node + the TypeScript sources (`~/Library/Application Support/open-charts/engine`) | `open-charts.exe`: Node and the engine bundled into one file, so the PC needs no Node (`%LOCALAPPDATA%\open-charts`) |
| VBA → engine | `AppleScriptTask` → bridge script → `curl` to the local server (or a one-off `node`) | `MSXML2.ServerXMLHTTP` POST to `http://127.0.0.1:4174/api/addin/run`, set to bypass any proxy |
| Server token | `~/Library/Application Support/open-charts/token` (mode 600), sent by `curl -H @file` | `%LOCALAPPDATA%\open-charts\token`, created by `install.ps1`, read by the VBA and sent as `X-OC-Token` |
| Starting the engine | the bridge runs `node … addin serve` in the background | VBA runs `open-charts.exe addin serve` with `Shell …, vbHide` (no console window), then waits up to 20 s for it to answer |
| Apply from the editor | the server runs `ApplyEdit` through one long-lived `osascript` (JavaScript for Automation) | the server runs `ApplyEdit` through one long-lived Windows PowerShell process that calls `Application.Run` over COM (`GetActiveObject`: it never starts PowerPoint) |
| Edit data window | the resident pop-up helper (`OpenChartsPopup.app`), next to the chart | Microsoft Edge in app mode (a window without tabs or address bar) next to the chart; the default browser if Edge is missing |
| Element mini toolbar | a panel of the pop-up helper | not yet: on Windows, selecting a chart element makes no call, so it never slows PowerPoint down |
| Add-in log | `~/Library/Application Support/open-charts/addin.log` | `%LOCALAPPDATA%\open-charts\addin.log` |
| Install | `open-charts-<version>.pkg` (per user), or `open-charts addin install` from source | `open-charts-<version>.msi` (per user, no admin rights; `-machine.msi` for IT), or `install.ps1` |

The VBA picks its path with `#If Mac Then … #Else … #End If`. The Mac compiles exactly the code it
compiled before Windows support was added (`test/windows.test.ts` checks this). The Windows branch
declares `Sleep` with `PtrSafe` under `#If VBA7`, so it compiles in 32-bit and 64-bit Office.

Apply on Windows has the same outcomes as on the Mac. The session becomes `drawn`, or `error` with a
message for the user. The 20-second timeout is the same. When PowerPoint is busy, for example while
the add-in waits for the engine, it rejects the COM call (`RPC_E_CALL_REJECTED`). The helper retries
until the timeout. If PowerPoint shows a dialog, the editor says: *Close any open dialog in PowerPoint
(and leave text editing), then Apply again.*

Fonts: the engine measures text with Arial from `%WINDIR%\Fonts` (`arial.ttf`, `arialbd.ttf`).

## Building `dist/win`

```sh
npm install
npm run build:win
```

This runs on macOS, Linux or Windows, and writes `dist/win/`:

| File | What it is |
|---|---|
| `open-charts.exe` | The engine as a Node [single executable application](https://nodejs.org/api/single-executable-applications.html) (about 80 MB, x64). |
| `open-charts.ppam` | The add-in, the same one the Mac installs. |
| `addin.json` | Its callback prefix, which the server needs to run `ApplyEdit`. |
| `install.ps1`, `uninstall.ps1` | Per-user install and removal. |
| `e2e-powerpoint-win.ps1` | The acceptance run for a real PC (below). |

On Windows, `npm run build:msi` then wraps this folder into the MSIs ([release.md](release.md#windows-the-msis-scriptsbuild-msits-addinwindowsopen-chartswxs)).

How `open-charts.exe` is made (`scripts/build-win.ts`):

1. **Bundle.** esbuild bundles `src/cli.ts` into one CommonJS file (a SEA can only run CommonJS).
   Two changes are needed for that:
   - Each `import.meta.url` becomes the URL of the source file it was in, under a resource root.
   - The CLI's top-level `await` runs inside an async function.

   The files the engine reads while it runs (`src/web/*`, `addin/vba/*`) are embedded in the exe as
   SEA assets. On first start they are unpacked to `%LOCALAPPDATA%\open-charts\res\<hash>\`.
2. **Blob.** `node --experimental-sea-config` creates the blob, without a code cache or snapshot so
   the blob does not depend on the build machine's platform.
3. **Node runtime.** The script downloads the official `win-x64/node.exe` for the Node version that
   is running the build (a blob only runs in the Node version that made it). It checks the file's
   SHA-256 against that release's `SHASUMS256.txt` on nodejs.org and caches it in
   `~/.cache/open-charts/node/`. It falls back to `curl` when Node's own `fetch` cannot connect.
4. **Inject.** postject injects the blob. Injecting invalidates Microsoft's Authenticode signature,
   so the script removes it: the exe is unsigned.

`--host` also builds `dist/host/open-charts` for the build machine, so the exe's code paths (unpacking
resources, starting itself) can be tried without a PC. `scripts/smoke-exe.ts <exe>` runs a built exe
the way the add-in does: gallery, layout, `addin build`, then `addin serve` with ping, an engine
request, the editor page and quit.

## Living with think-cell and other add-ins

- **Ribbon.** open-charts has its own tab (`tabOpenCharts`, before Home). Ribbon ids are private to
  each add-in unless it uses `idQ`, so they cannot clash with think-cell's. Every callback carries a
  per-build prefix (`OC` + 24 hex digits), so no other add-in has the same macro names.
- **Charts.** open-charts charts are ordinary grouped shapes with `OC_*` tags. think-cell ignores
  them. open-charts acts only on shapes tagged `OC_V`, so it never changes think-cell elements.
- **Events.** The add-in listens to `WindowSelectionChange` to re-lay out resized charts. It does
  nothing for shapes that are not open-charts charts.

## Corporate laptops

- **No admin rights needed.** Everything goes under the user's own folders and `HKCU`.
- **Macro security.** `install.ps1` removes the "downloaded from the internet" mark (Office blocks
  macros in marked files). It also makes sure PowerPoint trusts `%APPDATA%\Microsoft\AddIns`. That
  folder is usually trusted already; if it is not, the script adds an `HKCU` trusted location. If your
  organisation sets *Disable all macros without notification* (policy `VBAWarnings=4`), the add-in
  cannot run until IT allows macros from trusted locations. `install.ps1` warns when it finds that
  policy.
- **Application control (AppLocker, WDAC).** `open-charts.exe` is unsigned. When such a policy only
  allows signed or approved programs, `install.ps1` stops with a message that names the file to ask IT
  to allow.
- **Proxies.** All traffic stays on `127.0.0.1`. The VBA tells `ServerXMLHTTP` to connect directly,
  so a WinHTTP proxy setting does not get in the way.
- **Local server security.** The engine's server answers only callers that send the per-install
  token from `%LOCALAPPDATA%\open-charts\token` (32 random bytes; `install.ps1` creates it and keeps
  it on reinstall; `uninstall.ps1` removes the folder with it), and only requests addressed to
  `127.0.0.1`/`localhost` that come from no other web site. A web page in the user's browser cannot
  drive PowerPoint or read the charts being edited. Details: [ux-editor.md](ux-editor.md), Security.
- **Log.** `%LOCALAPPDATA%\open-charts\addin.log` gets results and errors. For timings too, run
  `open-charts.exe addin verbose on` (and `off` afterwards).
- **Antivirus.** The first start of a new `open-charts.exe` can take a few seconds while it is
  scanned. The add-in waits up to 20 s. After that the server stays up and each call is fast.

## What is verified where

- **CI (`windows-latest`, no PowerPoint).** The whole unit suite on Windows, including the real
  PowerShell helper (`ping`, and *not running* without PowerPoint). `npm run build:win`, then the
  smoke test of the built exe. The VBA's HTTP calls made the way VBA makes them (`ServerXMLHTTP` from
  COM, Unicode both ways). `install.ps1` (with `-DryRun` first) and `uninstall.ps1` against a
  temporary profile, including removing the download mark. The job uploads the `open-charts-win`
  artifact.
- **The owner's PC.** Everything that needs PowerPoint for Windows: the VBA compiling in Office, the
  ribbon next to think-cell, insert, update, resize, Edit data, Apply over COM, and export. Run
  `e2e-powerpoint-win.ps1`, below.

## Acceptance run on a PC

1. Get `dist/win`: the `open-charts-win` artifact of the latest CI run (Actions → *test* → artifact),
   or `npm run build:win` on any machine. Copy the folder to the PC and unzip it.
2. Save your work and quit PowerPoint.
3. Right-click the folder's empty space → *Open in Terminal* (or open PowerShell and `cd` there), then
   run:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\e2e-powerpoint-win.ps1
   ```

4. Don't use the PC for a few minutes. PowerPoint opens, charts are inserted, and an Edge window with
   the datasheet opens. Leave it all running.
5. Send back `open-charts-e2e-<time>.zip`, which is written next to the script. It contains the log,
   `addin.log`, the exported deck and one PNG per slide.
