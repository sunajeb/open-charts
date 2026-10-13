# Editor UX plan (`src/web/editor.html`)

The editor is two products in one file: the standalone web editor (a deck of chart slides, download
`.pptx`) and the **"Edit data" window** the PowerPoint add-in opens for one chart (`?session=<id>`).
The owner's bar for the add-in case: editing a chart's data must be easy and fast — open straight into
the datasheet, the common edits one action away, Apply in one click.

Screenshots: `docs/ux-shots/before-*.png` (as found) and `docs/ux-shots/after-*.png` — 1280×800,
1024×700, 900×700, 2000×1100, gallery, context menu, colour picker, invalid cell, issue markers, dark,
add-in session (editing and applied).

## Problems found → change → why

| # | Problem (as found, 1280×800 and 1024×700) | Change | Why / reference |
|---|---|---|---|
| 1 | Datasheet sat in a 460 px column under the form: at 1024 px it was below the fold and the 2024/2025E columns were scrolled off to the right. | App-shell layout: **preview on top, datasheet below it in the same wide column, options in a right sidebar**; each region scrolls on its own and a draggable splitter sets the preview/datasheet split. Works from 900 px to ultra-wide (below 900 px it stacks). | The reference's datasheet is a wide Excel-like sheet under the chart ("Element datasheets"); data is the thing users edit most. |
| 2 | Chart type was a plain `<select>` of 17 names. | **Visual gallery** with an inline-SVG icon per type, grouped *Column & bar · Line & area · Mekko · XY · Circular · Special*; keyboard navigable; current type shown on the trigger button. | Mirrors the reference's insert gallery ("Insert elements") and its chart-type switcher on the mini toolbar ("Change the chart type"). |
| 3 | Datasheet had no clear header/series areas, no range selection, no copy, no undo, no Tab, delete only via "selected cell" buttons. | Excel-like sheet: sticky category header row and series-name column, **range selection** (drag, Shift+click, Shift+arrows), **copy/cut/paste of ranges** (TSV, Excel compatible; a single value pasted into a range fills it), Delete clears a range, **Tab/Shift+Tab, Enter/Shift+Enter, arrows**, **undo/redo** (⌘Z / ⇧⌘Z, Ctrl+Y; typing in one cell coalesces into one step), "+" ghost column/row to add a category/series in one click, and a **context menu** (right-click, Shift+F10) to insert/delete series and categories at the clicked position, clear and copy. | "Element datasheets": insert/delete rows and columns from the toolbar, undo/redo, cut/copy/paste; "Context menu": options specific to what you right-clicked, Esc closes. |
| 4 | Per-type data layouts (points, dates, width row) were explained in one generic paragraph. | A one-line **per-type hint** above the sheet for every type, and the general paste/keyboard help moved into a `?` tooltip. | "Edit datasheet layout": the layout depends on the chart type. |
| 5 | Series colours could not be changed. | **Colour swatch** at the start of every series row → palette popover (theme palette + accents + "Automatic" + hex). Hidden where series colours do not apply (pie/doughnut slices, Harvey, table, Gantt rows). | "Chart fill scheme": fills follow the scheme unless a series overrides it. |
| 6 | Only unit suffix and decimals; no magnitude, locale, prefix units or Excel format. | **Numbers** section: unit + prefix toggle (`$12m`), decimals, magnitude (none/auto/k/m/bn), locale (decimal/grouping), custom Excel format. | "Format chart labels → Number format / Magnitude". |
| 7 | All options in one 3-column grid; unrelated controls side by side; no explanation. | Sidebar **sections**: *Chart* (type, title, orientation), *Numbers*, *Axes & layout*, *Labels & legend*, *Decorations* (totals, CAGR, difference arrows, waterfall totals). Only applicable controls are shown; a section with nothing applicable disappears. Every control has a **one-line tooltip** (hover or keyboard focus). | Reference chapters: Chart labels, Legends, Axes, Chart decorations. |
| 8 | Bar orientation existed in the spec but not in the editor. | A **Direction** radio group (▲ ▼ ◀ ▶, only the type's directions enabled) in the Chart section, and ▲ ▼ ◀ ▶ arrows on the gallery's tiles with a live preview of the hovered choice. | "Rotate and flip elements"; the reference's chart gallery. |
| 9 | Preview was small at 1024 px, fixed to the column width, no zoom; issues were a flat grey list. | Preview **fits the available box** (width and height), with **zoom** (−, Fit, +, 100 %). The quality-gate issues get **numbered markers on the slide** and a compact, clickable list: clicking an issue pulses the shapes it names and, for a hidden label, selects the datasheet cell it belongs to. | The reference flags problem labels in place ("exclamation mark next to the label"). |
| 10 | Errors from the engine were text only. | Validation errors land **on the offending cell** (red cell, message shown under the cell while focused) and in the list (click → go to cell); the preview is marked stale instead of blanking. | "No flicker" and keep the last good chart visible. |
| 11 | No loading/empty states; no add-in mode. | Loading skeleton, a clear "session not found" state, a stale-preview banner. **Add-in mode** (`?session=<id>`): no slide tabs or import/export; opens with the first data cell focused; primary **Apply to slide** (also ⌘↵ / Ctrl+Enter) with a visible confirmation once the slide is updated ("Updated on slide", or why not); **Cancel**; unsaved-changes indicator. | Owner requirement: fast round trip. |
| 12 | Focus styling inconsistent; no landmark/ARIA labels on the sheet. | Visible focus rings everywhere, `role=grid`-style labelling (`aria-label` "Enterprise, 2022"), menus/popovers with Esc and arrow keys, `aria-live` status. Neutral light theme with a dark-mode variant for the chrome (the slide stays white). | Accessibility baseline. |

Performance budget: edit → new SVG < 200 ms (E2E asserts it on every edit, including the largest
example); the SVG is swapped in one assignment and only when it changed, so there is no flicker.

## Session API (add-in ⇄ editor contract)

All JSON. The server keeps sessions in memory (last 200), `cache-control: no-store`, no CORS headers
(only pages served by the server itself may call it: see Security below).
Ids: `^[A-Za-z0-9_-]{1,64}$` (use a random UUID).

| Call | Body | Response |
|---|---|---|
| `POST /api/session/<id>` | `{ chart: ChartSpec, title?: string, state?: "open" }` | `200 { id, rev, state, chart, title, updatedAt }` |
| `GET /api/session/<id>` | — | `200` same shape, or `404 { error }` |
| `POST /api/session/<id>` (editor, Apply) | `{ chart, state: "applied" }` | `200` session; `rev` increments on every POST |
| `POST /api/session/<id>` (editor, Cancel) | `{ state: "cancelled" }` (chart kept) | `200` session |
| `POST /api/session/<id>` (host, outcome of an Apply) | `{ state: "drawn" }` or `{ state: "error", message }` (chart kept) | `200` session |

States: `open` → `applied` (editor) → `drawn` (the chart on the slide was updated) or `error`
(`message`: why not and what to do, written for the user); `cancelled` from any state. A host that
confirms Apply (`serve(…, { confirmsApply: true })`, the add-in's server) answers every session with
`confirms: true`; the editor then waits for `drawn`, `error` or `cancelled` at a higher `rev`
and shows **Updated on slide** or the
message in a bar under the toolbar. After 2.5 s it shows "Waiting for PowerPoint… If macOS asks whether
PowerPoint may control PowerPoint, click Allow"; after 30 s without an outcome it says the update was
not confirmed. Without `confirms` (plain `open-charts serve`, other hosts) it says **Applied to slide**
as soon as the POST succeeds, as before.

Errors: `400 { error }` for a bad id, an unknown state, a non-string `message`, or a new session
without a chart; `403 { error }` for a request the server refuses (Security).

Server hooks for the add-in (`src/server.ts`):
`serve(port, fonts, opts?: { routes?; onSession?; confirmsApply?; token?; isPublic? })`
— `routes` runs before the built-in routes (return `true` when it handled the request, e.g. the add-in's
`/api/addin/run` and `/api/addin/ping`); `onSession` is called after every successful
`POST /api/session/<id>`; `confirmsApply` is described above; `token` and `isPublic` in Security
below. `Session`
(`{ id, rev, state, chart, title?, updatedAt, message?, confirms? }`), `setSession(id, body)` (the
POST body, without calling `onSession`) and `getSession(id)` are exported for a host in the same
process. Covered by `test/server.test.ts`; the editor's side by `test/e2e/editor.spec.ts` against
`test/e2e/host.ts`, a host that confirms through `onSession` (drawn, permission timeout, PowerPoint
not running).

Flow:
1. Add-in: `POST /api/session/<id>` with `{ chart, title }` (id = 16 hex chars derived from the chart's
   `OC_SPEC`; state `open`) on its own server (port 4174), then open
   `http://127.0.0.1:4174/?session=<id>#oc-token=<token>` (the pop-up, Edge or the browser).
2. Editor: `GET` the session, edits the chart (type, data, options), **Apply** → `POST { chart, state: "applied" }`.
   The window stays open; the user may apply again (each Apply bumps `rev`).
3. Add-in: reacts in `onSession` (PowerPoint for Mac VBA cannot poll): on `state: "applied"` it runs
   the add-in's `ApplyEdit` macro (`run VB macro`, 20 s timeout; it never starts PowerPoint) through
   one long-lived AppleScript helper (`src/addin/runner.ts`: no osascript start-up per Apply; Edit data
   has it open its connection to PowerPoint with one read-only Apple event). The server lays the chart out itself, for the chart's context (frame, bounds, slide
   theme, same-scale line) that Edit data sent and each Apply answers, and passes the response to
   `ApplyEdit`, which only draws it (no engine round trip from PowerPoint: each AppleScriptTask costs
   about 0.2 s). If the chart was moved, resized or re-themed meanwhile, `ApplyEdit` answers its
   context and the server lays it out again for that. As it hands the response over, the server sets
   the session back to `open` (chart kept), so the edit is never applied to another chart that has the
   same data; once the macro has returned (the shapes are drawn) the server sets it to `drawn`. The
   next Apply sets it to `applied` again. If the call fails — macOS permission prompt
   unanswered (timeout), permission denied (-1743), PowerPoint not running, macro error — or the macro
   found no chart, the server sets `error` with a message for the user. An `error` session is still
   picked up by an explicit **Update chart** on that chart (never by the automatic re-layout). On
   `"cancelled"` it stops. The server only sends Apple events to the PowerPoint that started it (macOS
   attributes the server to that process; an event to a later PowerPoint froze it): once that
   PowerPoint has quit and another runs, an Apply fails at once with "PowerPoint was restarted after
   this editor window opened…", and the new PowerPoint's first ping or engine call gets 503, which
   makes the bridge start a fresh server. Other hosts may instead poll `GET`
   and act when `rev` grows. When the editor runs inside an Office dialog it
   also calls `Office.context.ui.messageParent(JSON.stringify({ type: "open-charts:apply" | "open-charts:cancel", id, rev }))`,
   and posts the same object to `window.opener` / `window.parent` when present, so the add-in can
   react without polling.

## Security (local server)

The add-in's server on `127.0.0.1:4174` runs macros in PowerPoint (Apply) and holds the charts being
edited. Listening on 127.0.0.1 keeps other computers out, but any web page open in the user's browser
can still send requests there. So every request is checked (`src/guard.ts`, `src/addin/token.ts`):

- **Host** must be `127.0.0.1:<port>` or `localhost:<port>`. A site that points its own name at
  127.0.0.1 (DNS rebinding) sends that name and is refused.
- **Origin**, when the browser sends one, must be the server itself (`http://127.0.0.1:<port>` or
  `http://localhost:<port>`). Another site's `fetch` or form post is refused, preflights included.
  There are no CORS headers, so no other site can read an answer either.
- **Token.** Every request except the static pages (`/`, `/index.html`, `/toolbar`, with any query)
  and `GET /api/addin/ping` must carry `X-OC-Token: <token>`. Without it, or with a wrong one, the
  answer is `403` and nothing changes. That covers the engine (`/api/addin/run`, `/selection`,
  `/format`, `/properties`, `/handles`, `/quit`), the sessions and `/api/render`, `/api/pptx`.

The token is 32 random bytes (hex), one per install. `open-charts addin install` (Mac) and
`install.ps1` (Windows) create it; a reinstall keeps it; the server creates it if it is missing. It
lives in the add-in's home, readable by the user alone (mode 600 on the Mac; `%LOCALAPPDATA%` is
per user on Windows): `~/Library/Application Support/open-charts/token` or
`%LOCALAPPDATA%\open-charts\token`. The file holds the header line itself, `X-OC-Token: <hex>`.

How each caller gets it:

| Caller | How |
|---|---|
| Mac VBA → AppleScript bridge → `curl` | `curl -H @<home>/token` reads the header line from the file, so the token is never on a command line (`ps` shows those) |
| Windows VBA (`MSXML2.ServerXMLHTTP`) | reads the file, `setRequestHeader "X-OC-Token"` |
| Node (the server's own calls, `addin run`, `stopServer`, the live test tools) | `tokenHeaders()` |
| Editor pop-up, toolbar panel, Edge pop-up, browser fallback | the URL the server opens ends in `#oc-token=<token>`. A fragment is never sent to a server, logged or put in a `Referer`. The page's first script moves it into `sessionStorage` (this tab only, so a reload works), removes it from the address bar and history (`history.replaceState`), and adds the header to the page's own same-origin `fetch` calls. Never in the query string. |
| `install.ps1` / `uninstall.ps1` (stop the engine) | read the file, `-Headers` |

The plain editor (`open-charts serve`) has no token (it does not drive PowerPoint) but makes the same
Host and Origin checks. So does the Office.js development host (`npm run officejs:serve`); for Office
on the web through a tunnel, name the tunnel's host with `--allow-host <name>`.

Limits: a program running as the same user can read the token file, as it can read everything else
of that user. On a Mac the Edit data pop-up gets its URL over a pipe; only the browser fallback runs
`open <url>`, whose command line is visible to `ps` for the moment it runs. Tests:
`test/security.test.ts` (refusal without/with a wrong token, foreign Host and Origin, the preflight,
`curl -H @file`, fragment-only page URLs, the Office.js host) and `test/e2e/popup.spec.ts` (the page
reads the token from its fragment and drops it from the address; another origin in the same browser
can neither change nor read a session).

Logging: `addin.log` gets results and errors (RESULT, APPLY, RELAYOUT, SYNC, failures).
`open-charts addin verbose on` (setting `verbose` in `settings.json` next to it) adds timings and the
pop-up, toolbar and drag-handle steps; `off` turns them off again. The server picks it up within
2 seconds, the VBA when PowerPoint starts next.

## Pop-up mode (`?session=<id>&popup=1`, or `?popup=1` preloaded)

The add-in's Edit data window over PowerPoint (src/addin/popup.ts: Mac helper panel, Edge app mode on
Windows). Same session API; differences from the session page above:

- Layout: datasheet only (no preview: the slide is the preview), compact toolbar with status,
  **Options** (type and option sections as an overlay), **Cancel**, **Done ⌘↵**.
- Live apply: each edit schedules `POST { chart, state: "applied" }` 400 ms after the last change;
  while one is in flight (until `drawn`/`error` at a higher `rev`) later edits set a flag and the
  latest chart is sent once it returns. Nothing is sent while the engine reports an error, or when the
  chart equals the last one sent.
- **Esc / Cancel**: if anything was sent, `POST { chart: <the chart the pop-up opened with>, state:
  "applied" }`, then `{ state: "cancelled" }` (the host draws applies in order, so the revert lands
  last). **⌘↵ / Done / close button**: sends the last edit if needed, keeps the slide.
- Host hooks: `window.ocPopupOpen(id)` loads a session into the preloaded page (state reset);
  `window.ocPopupClose("done" | "cancel")`. The page posts `{type: "ready", title}` once the session is
  drawn and `{type: "close", reason}` to `window.webkit.messageHandlers.openCharts` (the Mac helper);
  without that handler (Edge) it calls `window.close()`.
- `body.dataset.live` (`idle`/`busy`), `body.dataset.applies` and `body.dataset.closed` expose the
  state to tests (`test/e2e/popup.spec.ts`).
