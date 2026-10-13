# open-charts — live checklist

Plan of record: docs/MASTER-PLAN.md · Repo: https://github.com/sunajeb/open-charts

## Goal (2026-10-13) — lean first build
Add every chart type into PowerPoint, tested end to end: every editor button in a real browser, every chart type rendered through LibreOffice and PowerPoint for Mac.
**Out of the first build (owner, 2026-10-13):** AI features, other apps (Excel, Keynote, Sheets), CSV/XLSX inputs, template fill, HTTP API, SDKs.

### Not buildable on this machine (stays planned, not faked)
- P4 native Windows add-in (needs Windows + Office; gated at M7, skipped by default when solo)
- Windows PowerPoint golden runner, blind consultant panel, AppSource submission
- P5 Gantt / agenda / Excel live links (52 ew; after P3)

## Shared contracts (all agents code against these)
- `Scene.roles?: string[]`, parallel to `Scene.shapes`, assigned by the engine only. Format: `bar:s<i>:c<j>`, `label:value:s<i>:c<j>`, `label:total:c<j>`, `label:category:c<j>`, `label:series:s<i>`, `axis`, `axis:tick:<k>`, `gridline:<k>`, `arrow:cagr`, `arrow:diff`, `connector:c<j>`, `legend:s<i>`, `marker:s<i>:c<j>`, `line:s<i>`; wave 2: `area:s<i>`, `wedge:s0:c<j>`, `leader:c<j>`, `legend:c<j>`, `ball:s<i>:c<j>`, `label:total` (doughnut grand total), `shade:<k>` (Gantt shaded period)
- Tag contract v1 (MASTER-PLAN §3): group gets `OC_V`="1", `OC_ENGINE`, `OC_SPEC` (chunked above 30 KB), `OC_FRAME`="x,y,w,h", `OC_FONT`; each child gets `OC_ROLE`
- `specVersion: 1` at chart and deck root; unversioned input = v0, migrated by `migrate()`

## Wave 1 (three agents in parallel, git worktrees)
- [x] **A — engine (P1):** B3 date CAGR (30/360), B5, B7 hhea/OS2 line height, B8 long-word overflow, B9 legend wrap, B10, B11, B12 gate v2, B13, B14 numbers/locale service, B15 path-addressed validation; spec v1 + `migrate` + hand-written JSON Schema; `Scene.roles`; horizontal bar orientation
- [x] **B — I/O (P1 tags):** writer emits tag contract v1; `inspect` and `import`; `--strict` / `--report`
- [x] **C — QA + editor (P1 editor polish):** Playwright E2E for every editor button; LibreOffice golden pipeline (soffice → pdftoppm → pixel diff); editor transpose / reverse / all chart types; fix whatever the E2E finds

## Wave 2
- [x] Merge wave 1 (68 unit, 30 editor E2E, 13 goldens green; goldens rebaselined after review)
- [x] Scene `path` shape + preset outlines (diamond, roundRect, triangle)
- [ ] Chart catalogue (owner's list, 2026-10-13). ✅ = exists today
  - Column/bar: ✅ stacked, ✅ clustered, ✅ 100%, ✅ percent-axis variant (`percent` + `percentData`), ✅ waterfall (subtotals, connectors, below-zero), ✅ combination (stacked/clustered + line); [ ] bar orientation for all (agent A)
  - Line/area: ✅ line, ✅ ortho/step line, ✅ stacked area, ✅ area 100%
  - Marimekko: ✅ Mekko (% and units), ✅ Mekko XY (variable width and height)
  - XY: ✅ scatter, ✅ bubble
  - Circular: ✅ pie, ✅ doughnut
  - Special: ✅ Gantt (timeline: tasks, milestones, scales), ✅ Harvey ball, ✅ table
  - Needs new scene shapes first: `path` (polygon/step/area fills, wedges via arcs) for writer, SVG and add-in
  - `.ppam` = ribbon (`customUI14.xml`: own "open-charts" tab, Charts gallery with every type, Edit data, Update) + one thin VBA module that only forwards the control id
  - VBA → engine through `AppleScriptTask` (the sandbox-approved Mac bridge; the reference uses a `Declare` into its native dylib) → `node` engine lays out the chart → VBA draws native shapes via the object model (AddShape, AddLine + arrowheads, AddTextbox, BuildFreeform), groups and names them, writes the `OC_*` tags
  - Chart identity = tags + shape names (reference: GUID + shape name + persisted state); `OC_SPEC` is the stored datasheet
  - Edit data opens the datasheet editor; Update re-lays-out the selected chart in place
  - `install` copies the add-in into Office's add-in folder and the bridge script into `~/Library/Application Scripts/com.microsoft.Powerpoint`
  - Tests: unit (scene serialiser, vbaProject writer, package), and an end-to-end run in real PowerPoint via AppleScript `run VB macro` for every chart type → `inspect` the saved deck → export PNGs and check by eye

## Owner acceptance criteria (2026-10-13)
- [x] Easy data editing (live: session opens ~0.7 s; Apply → chart updates in place in 1.4–2.3 s, no extra click): Edit data on a selected chart opens its datasheet in ~1 s; Apply updates the chart on the slide in place, no extra clicks
- [x] Resize without breaking (3,172-layout matrix clean; live add-in re-layout on resize; min-size placeholder): resizing a chart re-lays it out at the new size (no stretched text); every type is clean from small to large frames; below a minimum it shows a clear "too small" message
- [x] Fast and seamless (live: insert ~530 ms in VBA / ~0.9 s wall, cold 1.8 s; update 675–930 ms; warm layout p95 ≤ 0.7 ms; 100-slide deck 81 ms): insert < 1.5 s, update < 1 s in real PowerPoint; layout p95 < 20 ms per chart; 100-slide deck < 2 s

## Follow-ups queued (after polish-b + resize-perf merge, to avoid golden churn)
- [x] `textOn` contrast: switch to dark text near luminance 0.21 (white on #7FA3CC / #A6A6A6 is ~2.5:1)
- [x] `niceTicks`: 620m max picks a 0–800m axis (22% empty) — prefer tighter steps
- [x] MASTER-PLAN §7.1: kerning decision changed — writer and add-in now set kern so PowerPoint never kerns (it kerns when the attribute is missing)

- [x] Editor: "label hidden (no room)" notices dominate the issue list on the default chart — collapse into one low-severity line ("4 labels too small to show"); keep real problems prominent
- [x] Thin segments: place labels outside with a leader (reference behaviour) instead of hiding, where there is room

- [x] Minor: mekko-3 Asia Pacific "3%" gets a long, nearly flat leader (valid, not pretty) — leaders leave the mark below their label; flat leaders cost extra
- [ ] Later: user drag pins a label (needs a fixed-position field in the placer + add-in support)

## Wave 3 — verification
- [x] Full suite + E2E + LibreOffice goldens; open generated decks in PowerPoint for Mac (no repair prompt), export PNGs, check by eye
- [x] CI runs unit + editor E2E on push/PR; goldens nightly (Linux fonts differ from the Mac baselines)
- [x] Review section below

## Open decisions (owner)
- [x] Provenance: "clean-room" claim dropped (2026-10-13)
- [ ] O1 staffing · O2 native Windows · O3 spec freeze — defaults from MASTER-PLAN §7.2 assumed

## Review

## Questions for the owner (collected overnight, 2026-10-13)
- **Scope beyond P1:** the goal says "complete the master plan" (P1–P5), but you cut the first build to "every chart type into PowerPoint, no AI or other apps". P1 (chart quality parity) is being finished now. Should P2 (AI, data inputs, template fill), P3 (Office.js for Windows/web), P4 (native Windows) and P5 (Excel links, agenda) be built next, or stay out?
- Optional: grant Accessibility permission to the terminal app (System Settings → Privacy & Security → Accessibility) so live PowerPoint tests can click their own "Enable Macros"/"OK" dialogs.
- Old test entries vA–vD may still be listed in PowerPoint → Tools → PowerPoint Add-ins (files deleted); remove them with "−".

### Final review (2026-10-13, open-charts 0.2.0)
- **Scope delivered (lean first build):** all 17 chart types from the owner's list (+ step line, % axis variant) inserted into PowerPoint for Mac through a .ppam add-in built like the reference (ribbon + one thin VBA module + AppleScriptTask bridge to the engine, native shapes, OC_* tags); Edit data → datasheet editor → Apply updates the slide in place; resize re-lays out; CLI + web editor.
- **Tests on master:** 253 unit, 67 editor E2E (every button), 35 LibreOffice goldens (each reviewed by eye), 3,172-layout resize matrix, perf budgets, docs examples validated; live PowerPoint E2E passed (19 gallery items, update, resize, edit→apply, export; add-in output matches the CLI writer within a few px). Acceptance report: docs/ACCEPTANCE.md.
- **Timings (live):** insert median 0.86 s (VBA 0.50 s), first insert 1.3 s, update ~1 s, resize ~1 s, edit-data open 0.95 s, apply 1.9 s end to end.
- **Not built (planned, by owner decision for the first build):** AI (P2 AI), data inputs / template fill / SDKs (P2), Office.js and native Windows add-ins (P3/P4), Excel links, agenda (P5), user label dragging. Gantt and tables are built at first-build depth.
- **Known minor:** a few labels render 1–3 px differently between add-in and CLI decks in PowerPoint; at very small frames some labels are hidden (and reported) rather than overflowing.

## P1 completion — unmerged branches (state at usage limit, 2026-10-13)
- `worktree-agent-a7009d645a2127149` (bench): 28 bench fixtures + snapshots, gallery 63 pages, `npm run quality` (gate passes), 3 engine fixes. Ready to merge after full `npm test`/`e2e`/goldens. Behaviour-spec and optional-rows sub-work may be missing (check for WIP commits).
- `worktree-agent-aaf1bf041b0c27de5` (theme, WIP 8556bfd): theme slots + schemeClr, style.theme/numbers, scaleGroup, multi-chart slides, add-in format v2. Left: docs (2 failing doc tests), editor swatches for slots, examples/goldens, e2e-powerpoint theme comparison, live run.
- `worktree-agent-a4e3bbc3ce7f9e418` (decorations, WIP cbb3f94, NOT merge-ready): level/difference arrows (3 modes), value lines, 100% indicator + percentBase, category gaps/breaks, value axis break. Left: waterfall/combo gaps, combo secondary axis drawing, schema/docs, engine.test roles/options, resolve/quality roles, unit tests, editor controls, example deck, full suites + goldens.
- Merge order suggestion: bench → theme → decorations; reinstall the add-in from the main repo after theme.
