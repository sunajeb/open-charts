# open-charts 0.2.0 — acceptance pass

**Date:** 2026-10-13 · **Tester:** final acceptance agent, run as a new user from a clean install.

**Environment:** macOS (Darwin 27.0), PowerPoint for Mac 16.105 (think-cell add-in also loaded),
Node 22.12.0 (nvm), Playwright Chromium 1243 (headless) for the editor, default browser Arc.
Accessibility was available to the terminal; Screen Recording was not used for evidence (all images
below come from PowerPoint's own slide export, or from Playwright for the editor).

Evidence from the run is in `out/e2e-powerpoint/` and `out/acceptance/` (not committed: `results.log`,
exported decks and PNGs). The images below are copied into `docs/acceptance/`.

## Summary

| # | Check | Result |
|---|---|---|
| 1 | Clean install, launcher, auto-load, ribbon | **Pass** after a fix (`addin uninstall` did not unload the add-in) |
| 2 | `npm run e2e:powerpoint`, every gallery item, PNGs judged by eye | **Pass** (19 types + placeholder slide; 3 cosmetic notes) |
| 3 | Edit data round trip with the real editor; Cancel | **Pass** (bug found and fixed: an applied edit could later land on another chart) |
| 4 | Resize: small, wide, tall, below minimum, grow back | **Pass** |
| 5 | Persistence: save, quit, reopen, Edit data, Update; ungroup | **Pass** (ungroup: no longer an open-charts chart, as expected) |
| 6 | Interop: CLI decks open in PowerPoint; add-in deck imports | **Pass** |
| 7 | Timings vs targets | Insert **pass** (cold 1.8 s → 1.3 s after O5); update **pass** in VBA, about 1 s by wall clock including AppleScript |
| 8 | Every ribbon control incl. nothing selected | **Pass** |
| 9 | Windows: same add-in on PowerPoint for Windows | Built and tested without PowerPoint (see [Windows](#9-windows)); **live run pending on the owner's PC** |

Bugs fixed on this branch: 3 (see [Bugs](#bugs)). Open items: see [Open](#open-items).

## 1. Clean install

- `node … src/cli.ts addin uninstall` (old build installed): printed
  `96:109: execution error: No result was returned from some part of this expression. (-2763)` and
  then "open-charts add-in removed.", **but PowerPoint still had open-charts loaded and auto-loading**
  (`get {name, loaded, auto load} of every add in` → `open-charts … true … true`). Cause:
  `repeat with a in add ins` never ends in PowerPoint's AppleScript (a test loop logged item 1, 2, 3, …
  without stopping). **Fixed** (d1a95de): unload by name from `name of every add in`; osascript errors are
  no longer echoed. Re-test: `loaded false, auto load false`, no stray output.
- README steps, exactly: `npm install` (3 packages, 7 s), then
  `node --experimental-strip-types --no-warnings src/cli.ts addin install` (4.4 s, PowerPoint was
  closed and was started by the install). Output named the engine copy, `.ppam`, bridge script,
  launcher, and `Linked ~/.local/bin/open-charts`.
- Launcher: `open-charts addin types` and `open-charts --help` work from a new shell (all 19 types
  `available`).
- PowerPoint: `open-charts` loaded + auto load true; the **open-charts** tab sits before Home (after
  thinkcell). Its controls in the accessibility tree: *Insert chart* (menu button, enabled), *Edit data*
  and *Update chart* (disabled with nothing selected, enabled once a chart is selected). No macro prompt
  appeared.

## 2. `npm run e2e:powerpoint`

`OK: 19 chart types inserted and saved, updates applied.` All 20 slides inspected
(`out/e2e-powerpoint/png/slide1…20.png`); add-in rendering matches the CLI writer on every slide (worst:
19 pixels beyond a 3-px shift on slide 7, under the threshold).

By eye, as a consultant: all 19 types are clean, labelled, no overlaps, no stretched text; numbers
check out (stacked CAGR 91→207 over 4 years = +22.8 %; waterfall 100+18+9−4−15+6−7 = 107).
Cosmetic notes (not fixed, engine layout):

- **Combination** (`docs/acceptance/e2e-combo.png`): the *Plan* line's value labels (108m, 118m, 126m)
  sit inside the *Services* segments in white, where they read like segment values.
- **Pie / doughnut:** the boxed "3%" label for the thin *Other* slice is wider than the slice.
- **Stacked at small sizes:** a boxed segment label wider than its column (e.g. "118m") shows its dark
  box sticking out of the column (`resize-small-260x170.png`, `resize-tall-280x460.png`).

## 3. Edit data round trip (real editor)

Driven by `out/acceptance/accept.ts edit`: a stacked chart is selected and *Edit data* is run through the
add-in's dispatcher. The add-in opened `http://localhost:4174/?session=650a9a4e065a9347` in the default
browser (Arc). The same URL was driven in Playwright Chromium:

- Editor loaded in 0.6 s: "Editing **Stacked chart** from PowerPoint", first cell focused.
- Changed Enterprise 2021 42 → 77 and renamed category 2021 → FY21; clicked **Apply to slide**: log
  `APPLIED 650a9a4e065a9347 OK (845 ms)`, editor status "Applied to slide"
  (`docs/acceptance/editor-applied.png`).
- Second edit (Mid-market 2022 → 40) applied with ⌘↵: `APPLIED … OK (588 ms)`.
- Then typed 999 without applying and clicked **Cancel**: page shows "Editing cancelled — The slide
  keeps the last applied chart"; nothing was logged and the slide kept the applied data.
- Saved copy, read back (`readCharts`, same as `open-charts inspect`): categories
  `["FY21","2022","2023","2024","2025"]`, series `[[77,58,…],[31,40,…],[18,…]]`. Slide PNG
  (`docs/acceptance/slide-after-apply.png`): FY21, 77m, 40m, totals 126m/119m and CAGR recomputed to
  +13.2 %.

**Bug found and fixed (9928eaf):** edit sessions are keyed by the chart's `OC_SPEC`, and a session
stayed `applied` after it had been drawn. When the user closed the editor without Cancel (the
normal way to finish), any other chart that still had the same data — a freshly inserted sample of the
same type, or a copy of the chart — silently took that edit on its next *Update chart* or resize. The engine now sets
the session back to `open` once it has drawn it (the next Apply re-arms it); unit test updated, contract
noted in `docs/ux-editor.md`. Reproduced first in PowerPoint (pie on two slides; Apply on slide 1, close
the window, Update chart on slide 2); see the Apply permission note below for the live re-test.

**macOS permission prompt (expected one-time step).** On a later Apply, macOS showed
*"Microsoft PowerPoint" wants access to control "Microsoft PowerPoint". Allowing control will provide
access to documents and data in "Microsoft PowerPoint", and to perform actions within that app.*
(Allow / Don't Allow). While it is unanswered, Apply does nothing to the slide: the server's
`run VB macro` call times out after 60 s (`APPLY … failed: … AppleEvent timed out`) and the editor
still says "Applied to slide". README and user guide now quote the real dialog and say **Allow**
(7a006fa). The editor now says so instead (O1, fixed).

## 4. Resize

`accept.ts resize`: each chart's group is set to a new box with AppleScript, then selected and
deselected (the selection change re-lays it out). Every case logged `RELAYOUT OK`, and group bounds equal
`OC_FRAME` exactly.

| Case | Box (pt) | Result | Image |
|---|---|---|---|
| small stacked | 260 × 170 | Re-laid out; text keeps its size; thin 2021 SMB label moved outside with a leader | `docs/acceptance/resize-small-260x170.png` |
| wide stacked | 880 × 170 | Re-laid out, CAGR arrow kept | `docs/acceptance/resize-wide-880x170.png` |
| tall stacked | 280 × 460 | Re-laid out | `docs/acceptance/resize-tall-280x460.png` |
| below minimum | 140 × 80 | "Chart too small: make it at least 200 × 120 pt" | `docs/acceptance/resize-below-min-140x80.png` |
| grow back | 140 × 80 → 864 × 408 | Full chart restored (49 shapes) | `docs/acceptance/resize-grown-back.png` |
| small line / mekko | 260 × 170 / 300 × 180 | Clean; mekko wraps "Mid-/market" | `out/acceptance/resize-png/slide6-7.png` |
| gantt | 300 × 180 | Placeholder "make it at least 320 × 180 pt" (its minimum) | `out/acceptance/resize-png/slide8.png` |
| pie below minimum | 90 × 60 | Placeholder, text fits the box | `out/acceptance/resize-png/slide9.png` |
| tall table | 280 × 460 | Columns compressed, numbers right-aligned | `out/acceptance/resize-png/slide10.png` |

## 5. Persistence

A deck with waterfall, mekko and clustered charts was saved (PowerPoint `SaveCopyAs`, copied out of
the sandbox), PowerPoint quit, then restarted and the file reopened (6.3 s, no prompts; windows: just the
deck).

- Slide 1 selected → *Edit data*: `RESULT btnEditData OK`; the session holds the deck's data (categories
  `EBITDA 2024 … EBITDA 2025`, identical to the saved `OC_SPEC`).
- Slide 2 → *Update chart*: `RESULT btnUpdate OK (1529 ms)` (first engine call after reopen).
- **Ungroup** (slide 3 → 26 loose shapes): the shapes keep their look; *Edit data* and *Update chart*
  on a piece both say "Select an open-charts chart first (click it once), then choose Edit data or
  Update chart." A selection change does nothing; `inspect` no longer lists a chart on that slide.
  Regrouping cannot bring the tags back. As expected for this release. The message could say "this is no longer an
  open-charts chart" (O3).

## 6. Interop

- `npm run examples` → `out/deck.pptx` (4 slides) opened in PowerPoint with no repair prompt (only the
  document window); exported PNGs look right. The e2e run also opened a 20-slide CLI-written deck of
  every type (`compare.pptx`) without a prompt.
- The add-in-made deck (`out/e2e-powerpoint/addin.pptx`, 20 charts) reopened without a prompt;
  `open-charts import` recovered all 20 charts (`stacked … waterfall`), and `open-charts imported.json
  --strict` rebuilt it with exit code 0.
- Content note: the example deck's title says "Revenue grew **2.4x** in four years" but its data goes
  97m → 216m (2.2x, +22.2 % CAGR on the slide). See O4.

## 7. Timings

Wall clock includes AppleScript `run VB macro` (about 0.3–0.4 s that a real ribbon click does not pay);
"VBA" is the add-in's own measurement (engine + draw).

| Operation | Wall clock | VBA (engine + draw) | Target | Verdict |
|---|---|---|---|---|
| PowerPoint start + new presentation | 5.2 s | — | — | — |
| Cold first insert (server stopped, first macro after start) | 1.81 s (e2e 1.81 s); after O5 1.23–1.36 s (e2e 1.31 s) | 1.33 s (engine 0.83 s); after O5 0.71–0.96 s (engine 0.22–0.37 s) | < 1.5 s | Pass after O5 (see below) |
| Warm insert | median 0.88 s, max 1.16 s (e2e, n=18) | median 0.49 s | < 1.5 s | Pass |
| Update chart | 0.77–1.06 s (n=4) | 0.41–0.63 s | < 1 s | Pass in VBA; ~1 s wall |
| Edit data → session open | 0.93–1.11 s | 0.36 s | ~1 s | Pass |
| Editor loads session | 0.62 s | — | — | — |
| Apply in editor → chart updated | 1.45 s first, 0.96 s second (click to log); e2e 2.25 s; after O7 0.21–0.28 s steady (pair 0.64–0.67 s), first right after Edit data 1.4–4.4 s (see O7) | 0.59–0.85 s; after O7 0.19–0.25 s steady, 0.52–0.97 s first | ≤ 2 s, VBA ≤ 0.8 s | Pass in VBA; wall clock see O7 |
| Resize → re-layout on selection change | 0.90–1.64 s | 0.24–0.76 s | — | Pass |

## 8. Ribbon controls

- Real clicks (System Events, accessibility): *Insert chart* opens a gallery with group headers
  (Column and bar, Line and area, Mekko, XY, Pie, Timeline and table) and all 19 types; clicking
  *Waterfall* drew "Waterfall chart" (864 pt wide) and left it selected, which enabled *Edit data* and
  *Update chart*. *Update chart* ran with no dialog; *Edit data* opened the session (rev 2) in the
  browser.
- Nothing selected: the two buttons are **disabled** once PowerPoint has reported a selection
  change. Before that, they are enabled and show a MsgBox. That MsgBox path was tested through the
  dispatcher in silent mode (it would block automation):
  - Edit data / Update chart → "Select an open-charts chart first (click it once), then choose Edit data
    or Update chart."
  - Unknown control id → "This button is not known to this version of the add-in (btnNope)."
  - Insert with no slide → "Click a slide (Normal view) first, then insert the chart."
- Unsupported-type path: every gallery type is available in 0.2.0, so the e2e "not available" branch
  had nothing to test.

## Bugs

Fixed on this branch:

1. **`addin uninstall` left the add-in loaded and auto-loading** and printed a raw osascript error
   (d1a95de).
2. **An applied edit could later be applied to another chart with the same data** (fresh sample of the
   same type, or a copy) once the editor was closed without Cancel (9928eaf).
3. **Docs named the wrong permission dialog and button** ("click OK"; macOS says *Allow*), and did not
   say how to recover from *Don't Allow* (7a006fa).

## Open items

- **O1 — Apply while macOS permission is unanswered or denied — fixed.** Was: the editor said "Applied
  to slide" although nothing happened (the server's AppleScript call timed out after 60 s). Now the
  session records the outcome: `applied` → `drawn` once `run VB macro` has returned with the chart
  redrawn, or `error` with a message. The server's call has a 20 s timeout and never starts
  PowerPoint. The add-in's server answers sessions with `confirms: true`; the editor then shows
  "Updating the slide…", after 2.5 s "Waiting for PowerPoint… If macOS asks whether PowerPoint may
  control PowerPoint, click Allow", and finally **Updated on slide** or, in a bar under the toolbar,
  e.g. "PowerPoint did not respond. If macOS is asking whether PowerPoint may control PowerPoint, click
  Allow in that dialog, then Apply again." / "macOS does not allow PowerPoint to be controlled (Don't
  Allow was chosen). Turn it on in System Settings › Privacy & Security › Automation › Microsoft
  PowerPoint, then Apply again." / "PowerPoint isn't running. Open the presentation, click the chart
  once, then Apply again." / "PowerPoint could not find the chart…". Plain hosts (`open-charts serve`)
  keep "Applied to slide". ApplyEdit no longer shows a dialog (it would block the server's call).
  Contract in `docs/ux-editor.md`. Tests: unit (each outcome and message, rev races), editor E2E
  against a host that confirms through `onSession` (drawn, permission timeout, not running), and live
  `e2e:powerpoint`: Apply → session `drawn` (`APPLIED … OK (629 ms)`, 2.25 s from Apply to `drawn`
  including osascript); Apply after quitting PowerPoint → `error: PowerPoint isn't running…`, and
  PowerPoint stays closed. The permission prompt was already allowed on this Mac, so the timeout path
  is covered by the simulated host, not live.
- **O6 — Apply after PowerPoint was restarted froze PowerPoint — fixed.** Was (since O5, pre-existing at
  6eb802e): quit PowerPoint, start it again, Edit data → Apply: PowerPoint stopped answering Apple events
  for 20 s+, ApplyEdit never ran, the editor said "PowerPoint did not respond…", and later scripting
  calls also timed out (-1712). Cause: the warm server outlives PowerPoint, and macOS attributes it (and
  every osascript it runs) to the PowerPoint that spawned it (responsible pid: that PowerPoint while it
  runs; after it quits, the server's osascript still carries PowerPoint's identity under its own pid).
  An Apple event from such a process to the new PowerPoint goes through PowerPoint's own permission
  check (`TCCAccessRequestIndirect` in the TCC log; PowerPoint's Apple-event thread waits on a semaphore
  in `designatedCodeRequirementsCacheQueue`, main thread idle) and never gets through. A server spawned
  by the running PowerPoint sends to itself and needs no check. Fix: the server records the PowerPoint
  it started with (pgrep); once that process is gone and another PowerPoint runs, `/api/addin/ping`
  (the new PowerPoint's startServer) and `/api/addin/run` answer 503 and the server exits, so the
  bridge starts a fresh server from the new PowerPoint; an Apply that still reaches the old server
  fails at once with "PowerPoint was restarted after this editor window opened…" instead of sending the
  event. Live, minimal repro (insert, Edit data, Apply; quit PowerPoint; repeat): before,
  session 2 `APPLY … failed after 20035 ms: timed out` every time (also at 6eb802e); after, `SERVER
  retired: PowerPoint 45249 has quit` → `SERVER started (… PowerPoint 45480)` → session 2 `drawn` in
  4.6 s. `e2e:powerpoint` now edits the same-scale pair's data and applies it in the third PowerPoint
  session (drawn, partner `SYNC revenue 1 of 1`): 2.7 s and 3.8 s in two consecutive passing runs.
- **O7 — Apply timing — fixed.** Was ~3.3 s from Apply to drawn (VBA ~1.7 s) after the theme and
  same-scale work, against ~1.9 s / 0.6 s before. Profile (new phase timing in `addin.log`: VBA
  `[find, read, theme, scale, shapes, group, tags, names, zorder, delete, sync, context]`, server
  `ENGINE <op> … (sessions, layout)` and `APPLY … (layout, to macro, macro, back)`): the theme read
  (0–2 ms) and same-scale scan (0 ms without `OC_SGROUP`) were never the cost. Every AppleScriptTask
  from VBA costs a fixed ~217 ms (a no-op handler as much as the engine call, whose server work is
  ~10 ms), and Apply made two (engine, then the `APPLIED` log line), four for a same-scale pair; each
  Apply also started osascript (0.12–0.25 s, 1.5–2.8 s while the editor's browser launches). Fix:
  Edit data sends the chart's context (frame, bounds, theme, scale line); on Apply the server runs the
  same `update` request in process (`layoutForApply`) and passes the response to `ApplyEdit`, which
  only draws and answers its log lines and the new context through `run VB macro`'s return value (a
  moved/resized/re-themed chart answers its context and is laid out again); the Apple event goes
  through one long-lived JXA helper (`src/addin/runner.ts`), whose connection Edit data opens with one
  read-only event. Live A/B, same machine, back to back (repeated Applies, steady state): single chart
  0.81–1.15 s → 0.21–0.28 s (VBA 0.43–0.56 s → 0.19–0.25 s); same-scale pair 1.48–1.97 s →
  0.64–0.67 s (the partner still goes through the engine call). The first Apply posted within a second
  of Edit data (what `e2e:powerpoint` does) is dominated by the editor's browser starting: on this
  8 GB Mac with 6.6 GB of swap in use, PowerPoint did not take the Apple event for up to 4.7 s and drew
  3–6× slower in that window, before and after alike; e2e single-chart Apply 3.6 s before, 1.4–4.4 s
  after, VBA 2.0 s before, 0.52–0.97 s after.
- **O2 — Cosmetic layout:** combo line labels inside bars; pie "3%" box wider than its slice; boxed
  segment labels wider than narrow columns (section 2).
- **O3 — Ungrouped chart message — fixed.** A selected shape that carries `OC_ROLE` with no group
  carrying `OC_SPEC` (ungrouped, or regrouped) now gets "This chart was ungrouped, so it can't be
  edited as a chart. Undo (⌘Z) the ungroup or insert it again." from Edit data and Update chart; the
  buttons stay enabled for such a selection so the click can explain. Live `e2e:powerpoint` ungroups a
  chart, selects a piece and checks the message (`RESULT btnEditData This chart was ungrouped…`).
- **O4 — Example content:** `examples/deck.json` (and the editor's sample deck) titles "Revenue grew
  2.4x in four years" over data that grows 2.2x. Changing the text needs a golden re-baseline.
- **O5 — Cold insert — fixed.** Was 1.8 s wall clock (1.3 s in VBA), because the first call ran a
  one-off node while starting the server. Profile of that call: engine 0.73–1.16 s in VBA, of which
  node + layout is only ~0.25 s (`addin run` from a shell; a fresh server answers its first request in
  14 ms and is up 0.27 s after spawn); the rest is the first AppleScriptTask (loading the bridge
  script, `findNode`, the parallel server spawn). Font metrics were already cached and node's compile
  cache gave nothing measurable. Fix: the ribbon's onLoad calls a new bridge handler `startServer`
  that spawns the warm server in the background if it is not up (returns in ~0.1 s), so the first
  insert takes the curl fast path and the bridge script is already loaded. Not from Auto_Open: an
  AppleScriptTask while PowerPoint is launching made `register add in` fail (-43) during
  `addin install`. Live, first insert right after starting PowerPoint (n=3 each, same machine):
  before 1.72–2.41 s wall, VBA 1.24–1.80 s (engine 0.73–1.16 s); after 1.23–1.36 s wall (one
  2.27 s outlier inside AppleScript's `run VB macro`, which a ribbon click does not pay), VBA
  0.71–0.96 s (engine 0.22–0.37 s). `e2e:powerpoint`: first insert 1.31 s (was 1.81 s), warm median
  0.89 s.

## 9. Windows

Windows support uses the same `.ppam`, ribbon, VBA and tags as the Mac, with an HTTP bridge, a
single-executable engine and a COM Apply helper ([docs/windows.md](windows.md)). Where each part
was verified:

| Check | Where | Result |
|---|---|---|
| VBA: the Mac compiles exactly the code it did before; Windows compiles no `AppleScriptTask`; `PtrSafe` declares for 32- and 64-bit Office; same `ApplyEdit` on both | unit (`test/windows.test.ts`, VBA run through a `#If` preprocessor) | **Pass** |
| PowerShell Apply helper protocol: one request per line, timeout restarts it, tabs, newlines and Unicode survive | unit, with a stand-in helper (macOS) | **Pass** |
| Real PowerShell helper: ping, and *not running* without PowerPoint | unit on Windows (CI) | runs in the `windows` job |
| `npm run build:win` on macOS: bundle, node.exe v22.12.0 checked against SHASUMS256.txt, blob injected, SEA fuse set, signature removed | Mac | **Pass** (83.8 MB PE32+ x64) |
| The same bundle as a macOS SEA (`--host`): resources unpacked, gallery, example deck, `addin build`, `addin serve`, ping, engine request, editor page, quit | Mac, `scripts/smoke-exe.ts dist/host/open-charts` | **Pass** (first start 3.1 s, warm 0.36 s, server up in 1.1 s) |
| `open-charts.exe` runs: the same smoke test | CI `windows` job | runs in the `windows` job |
| The VBA's HTTP call made as VBA makes it (`MSXML2.ServerXMLHTTP`, direct connection), Unicode label both ways | CI `windows` job | runs in the `windows` job |
| `install.ps1 -DryRun` changes nothing; install into a temporary profile (files, HKCU `Path` and `AutoLoad=1`, trusted location, download mark removed, engine starts); `uninstall.ps1` removes it all and stops the engine | CI `windows` job | runs in the `windows` job |
| Unit suite on Windows (paths, line endings, no `sh`) | CI `windows` job | runs in the `windows` job |
| Live PowerPoint for Windows with think-cell: VBA compiles, ribbon, every gallery item, update, resize, Edit data + Apply over COM, PNG export | owner's PC: `e2e-powerpoint-win.ps1` ([how](windows.md#acceptance-run-on-a-pc)) | **pending** |

Not verified yet, and to look for in the PC run's zip:
- whether `Application.Run` from COM finds the add-in macro by its plain name (the helper falls back to
  `open-charts.ppam!OpenCharts.<name>`; the log says when it had to),
- whether `WindowSelectionChange` fires for selection changes made through COM (the resize check),
- first start time of the unsigned exe under the PC's antivirus.

