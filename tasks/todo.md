# open-charts — live checklist

Plan of record: docs/MASTER-PLAN.md · Repo: https://github.com/sunajeb/open-charts

## Goal (2026-10-13)
A working, end-to-end tested open-charts covering everything in the master plan that can be built and proven on this Mac.
Every editor button and every chart type must be exercised in a real browser and rendered through LibreOffice and PowerPoint for Mac.

### Not buildable on this machine (stays planned, not faked)
- P4 native Windows add-in (needs Windows + Office; gated at M7, skipped by default when solo)
- Windows PowerPoint golden runner, blind consultant panel, AppSource submission
- P5 Gantt / agenda / Excel live links (52 ew; after P3)

## Shared contracts (all agents code against these)
- `Scene.roles?: string[]`, parallel to `Scene.shapes`, assigned by the engine only. Format: `bar:s<i>:c<j>`, `label:value:s<i>:c<j>`, `label:total:c<j>`, `label:category:c<j>`, `label:series:s<i>`, `axis`, `axis:tick:<k>`, `gridline:<k>`, `arrow:cagr`, `arrow:diff`, `connector:c<j>`, `legend:s<i>`, `marker:s<i>:c<j>`, `line:s<i>`
- Tag contract v1 (MASTER-PLAN §3): group gets `OC_V`="1", `OC_ENGINE`, `OC_SPEC` (chunked above 30 KB), `OC_FRAME`="x,y,w,h", `OC_FONT`; each child gets `OC_ROLE`
- `specVersion: 1` at chart and deck root; unversioned input = v0, migrated by `migrate()`

## Wave 1 (three agents in parallel, git worktrees)
- [ ] **A — engine (P1):** B3 date CAGR (30/360), B5, B7 hhea/OS2 line height, B8 long-word overflow, B9 legend wrap, B10, B11, B12 gate v2, B13, B14 numbers/locale service, B15 path-addressed validation; spec v1 + `migrate` + hand-written JSON Schema; `Scene.roles`; horizontal bar orientation
- [ ] **B — I/O (P1 tags + P2 data-to-deck):** writer emits tag contract v1; `inspect` and `import`; CSV/XLSX inputs; `fill` template; `--strict` / `--report`; HTTP `POST /api/render`
- [ ] **C — QA + editor (P1 editor polish):** Playwright E2E for every editor button; LibreOffice golden pipeline (soffice → pdftoppm → pixel diff); editor transpose / reverse / all chart types; fix whatever the E2E finds

## Wave 2
- [ ] Merge wave 1, resolve conflicts, full test suite once
- [ ] More chart types (area, pie/doughnut, scatter, combo) — P2
- [ ] Office.js add-in (P3 path A): task pane, sideload on PowerPoint for Mac
- [ ] `@open-charts/ai` with injectable client + verifier (no live key needed for tests)

## Wave 3 — verification
- [ ] Full suite + E2E + LibreOffice goldens; open generated decks in PowerPoint for Mac (no repair prompt), export PNGs, check by eye
- [ ] CI runs unit + E2E + goldens on PRs
- [ ] Review section below

## Open decisions (owner)
- [x] Provenance: "clean-room" claim dropped (2026-10-13)
- [ ] O1 staffing · O2 native Windows · O3 spec freeze — defaults from MASTER-PLAN §7.2 assumed

## Review
