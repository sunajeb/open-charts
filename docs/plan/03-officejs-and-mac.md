# 03 — Cross-platform PowerPoint add-in (Office.js): Windows, Mac, web, iPad

Status: draft plan. Companion to `docs/native-windows-design.md` (native COM/VSTO add-in for Windows). This document covers the one add-in that can run everywhere Office.js runs; where the native add-in does something better on Windows, that is noted and not re-designed here.

## 1. Goal and shape of the thing

One web add-in, hosted as static files, that lets a user:

1. **Insert** a chart on the current slide from the datasheet UI we already have (`src/web/editor.html`).
2. **Edit** an existing open-charts chart: select it, the task pane loads its spec, change data/options, press Update.
3. **Refresh / re-layout** after the user resized the group, so labels are re-planned instead of scaled.

Everything runs in the browser: layout engine, quality check, text measurement, and (for one insertion path) the OOXML writer. No backend, no account, no data leaves the machine. The engine stays a single codebase shared with the CLI and the Node editor; the add-in is a thin host layer in `addin/`.

```
addin/
  manifest.xml            # XML manifest (all four platforms)
  manifest.json           # unified manifest (added once Mac/iPad support is confirmed, §2)
  src/taskpane.ts         # datasheet UI (ported from src/web/editor.html) + Office host glue
  src/commands.ts         # ribbon command handlers (Insert, Update, Re-layout)
  src/host/powerpoint.ts  # Scene → PowerPoint.Shape* calls; tags; selection
  src/host/measure.ts     # browser Measurer (bundled Liberation Sans, canvas fallback)
  build.mjs               # esbuild
  dist/                   # published static site
```

## 2. Manifest, ribbon, task pane, dialogs

**XML vs unified JSON manifest.** The unified (JSON) manifest is Microsoft's direction and is required for a Teams-style combined app, but platform coverage for PowerPoint on **Mac and iPad** has lagged Windows/web. Decision: ship the **XML manifest** first (works on all four platforms, accepted by AppSource and centralized deployment), keep the manifest surface small enough that converting later is mechanical (`office-addin-manifest convert`). Spike S1 (§8) checks the current unified-manifest platform matrix; if Mac + iPad are covered we ship JSON from the start.

Manifest contents (XML):
- `<Hosts><Host Name="Presentation"/>` only.
- `<Requirements>`: declare **PowerPointApi 1.2** as the hard floor (enough for the slide-insert path, §4.1) and feature-detect everything else with `Office.context.requirements.isSetSupported('PowerPointApi', '1.x')` at runtime. Declaring a high floor would hide the add-in on older Mac builds instead of degrading gracefully.
- `<VersionOverrides>`: one ribbon group "open-charts" on the **Home** tab (or Insert; decide in UX review): **Insert chart** (ShowTaskpane), **Update** (ExecuteFunction, acts on selection), **Re-layout** (ExecuteFunction). Function-file commands run in a hidden runtime; they must load `office.js`, do the work with `PowerPoint.run`, and call `event.completed()`.
- Shared runtime (`<Runtimes>`/`Requirements SharedRuntime 1.1`) so task pane and commands share state and the font bundle is loaded once. SharedRuntime is Windows/Mac/web; on iPad fall back to the task pane only.
- Icons at 16/32/80 px, `SupportUrl`, `AppDomains` for any dialog host.

Skeleton of the parts that matter (abridged):

```xml
<OfficeApp xsi:type="TaskPaneApp"> <Id>…</Id> <Version>0.1.0.0</Version>
  <Hosts><Host Name="Presentation"/></Hosts>
  <Requirements><Sets DefaultMinVersion="1.1"><Set Name="PowerPointApi" MinVersion="1.2"/></Sets></Requirements>
  <DefaultSettings><SourceLocation DefaultValue="https://addin.open-charts.org/v0.1/taskpane.html"/></DefaultSettings>
  <VersionOverrides xsi:type="VersionOverridesV1_0">
    <Hosts><Host xsi:type="Presentation">
      <Runtimes><Runtime resid="Taskpane.Url" lifetime="long"/></Runtimes>   <!-- shared runtime -->
      <DesktopFormFactor>
        <FunctionFile resid="Taskpane.Url"/>
        <ExtensionPoint xsi:type="PrimaryCommandSurface"><OfficeTab id="TabHome">
          <Group id="oc.group"> <!-- Insert chart (ShowTaskpane), Update (ExecuteFunction), Re-layout (ExecuteFunction) --> </Group>
        </OfficeTab></ExtensionPoint>
      </DesktopFormFactor>
    </Host></Hosts>
  </VersionOverrides>
</OfficeApp>
```

**Task pane.** Port of `editor.html`: datasheet (paste from Excel), options, live SVG preview, issues list. The preview is produced by `sceneToSvg` in-browser, so it is pixel-identical to the Node editor. Width target 320–420 px; the current editor's two-column layout becomes a vertical stack.

**Dialog API** (`Office.context.ui.displayDialogAsync`, DialogApi 1.1): used only for things that don't fit a 350 px pane — a wide datasheet for many series, and a "paste JSON spec" view. Dialogs run in a separate page, same origin, message back with `messageParent`. Works on Windows, Mac, web; iPad supports it with size limits. Not used for anything essential.

**iPad** is tier 2: task pane works, ribbon add-in commands are not available on iPad in general (the add-in is launched from Insert ▸ Add-ins). Everything must be reachable from inside the pane; ribbon buttons are accelerators only.

### 2.1 PowerPoint API requirement sets we depend on

Confidence column is honest: ✅ sure, ⚠️ believed but verify in S1 (print `isSetSupported` for 1.1…1.10 on each platform and exercise the call), ❌ believed absent.

| Capability | API | Set (believed) | Confidence | Needed for |
|---|---|---|---|---|
| Insert slide(s) from a .pptx base64 | `presentation.insertSlidesFromBase64` | 1.2 | ✅ | §4.1 "insert as new slide" path, exact fidelity |
| Add rect / ellipse / rounded rect | `slide.shapes.addGeometricShape(GeometricShapeType.*, {left,top,width,height})` | 1.3 | ✅ | bars, markers, CAGR bubble |
| Add straight line | `shapes.addLine(ConnectorType.straight, …)` | 1.3 | ✅ | axes, connectors, arrows |
| Add text box | `shapes.addTextBox(text, …)` | 1.3 | ✅ | all labels |
| Fill / line colour, weight, dash, visibility | `shape.fill.setSolidColor`, `shape.lineFormat.{color,weight,dashStyle,visible}` | 1.3 | ✅ | styling |
| Font name/size/bold/colour | `textFrame.textRange.font.*` | 1.3 | ✅ | labels |
| Paragraph alignment | `textRange.paragraphFormat.horizontalAlignment` | 1.4 | ⚠️ | left/center/right labels |
| Vertical anchor, margins, no-wrap, no-autofit | `textFrame.verticalAlignment`, `.leftMargin…`, `.wordWrap`, `.autoSizeSetting` | 1.3/1.4 | ⚠️ | match writer's `lIns=0 wrap=none noAutofit` |
| Tags on shapes/slides/presentation | `shape.tags.add(key, value)`, `.getItemOrNullObject` | 1.3 | ✅ | storing the spec on the chart (§4.3) |
| Group / ungroup, group bounds | `shapes.addGroup([...shapes])`, `shape.group.shapes`, `ShapeGroup.ungroup()` | 1.4 | ⚠️ | one object per chart, as the CLI writes |
| Selected shapes / slides / text | `presentation.getSelectedShapes()`, `getSelectedSlides()`, `slide.setSelectedShapes` | 1.5 | ✅ | edit flow |
| Selection-changed event | Common API `Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged)` | Common | ✅ | detect "user clicked a chart" |
| Shape z-order | `shape.setZOrder` | 1.9? | ⚠️ | only if group order ≠ creation order |
| Bindings to shapes | `presentation.bindings.add` | 1.6 | ⚠️ | optional: stable handle to a chart across sessions |
| Custom XML parts (presentation-level) | `presentation.customXmlParts` | 1.7 | ⚠️ | optional: deck-level style/theme, large specs |
| Arrowheads on lines | `lineFormat.beginArrowheadStyle/endArrowheadStyle` | — | ❌ (not in PowerPoint JS as far as we know) | CAGR / difference arrows |
| Shape rotation | `shape.rotation` | — | ⚠️ | arrowhead workaround |
| Preset-geometry adjustments (corner radius) | — | — | ❌ | rounded CAGR bubble radius |
| Line flipH/flipV | — | — | ❌ | direction only matters for arrowheads |
| Insert OOXML into current slide | `setSelectedDataAsync(…, {coercionType:'ooxml'})` | Word only | ❌ | would have been the easy path; it isn't available |
| Shape resize/move event | — | — | ❌ | re-layout must be triggered by selection or button |

Consequences: the **shape-by-shape path can express ~95 % of the scene** natively; the known gaps are arrowheads, rounded-rect radius, and line flips. Fallbacks in §4.2. The **slide-insert path has 100 % fidelity** because it reuses `src/render/pptx.ts` unchanged.

## 3. Running the engine in the browser

### 3.1 Build
- The repo keeps its "no build step" promise for CLI and Node editor. The add-in is the one thing that needs bundling: `addin/build.mjs` runs **esbuild** (`--bundle --format=esm --target=es2020 --platform=browser`), emitting `dist/taskpane.js`, `dist/commands.js`, static HTML, icons, fonts. esbuild resolves our `./x.ts` extension imports as-is, needs no config, and builds in <100 ms. Vite is heavier than we need (no HMR benefit inside the Office web view; use `office-addin-debugging start` for live reload instead). `esbuild` is a devDependency of `addin/` only.
- `office.js` is **not** bundled: loaded from `https://appsforoffice.microsoft.com/lib/1/hosted/office.js` (AppSource requirement). Types from `@types/office-js` as devDependency.
- `src/core/text.ts` currently imports `node:fs`. Split: `core/text.ts` exports `createMeasurerFromFonts(regular: ArrayBuffer, bold?: ArrayBuffer)` plus `wrap`; `node/fonts.ts` keeps the path search and file reads and calls the former. Nothing else in `src/core` or `src/render` touches Node APIs (`zip.ts` uses `TextEncoder`/`Uint8Array`; fine in browsers). Add a CI check: `esbuild --bundle --platform=browser src/core/index.ts src/render/pptx.ts` must succeed with no Node built-ins.

### 3.2 Font metrics in the browser
The engine's promise is "what it plans matches what PowerPoint draws", which rests on real advance widths of the deck font. Two options:

| | Bundled metric-compatible font (Liberation Sans / Arimo via opentype.js) | `canvas.measureText` with installed font |
|---|---|---|
| Parity with Node/CLI | Exact: same code, same font file, same numbers → shared golden tests pass in both | Differs per OS/renderer by ~1–3 %, hinting and rounding |
| Fonts covered | Arial only (plus whatever we bundle) | Any installed font (Calibri, theme fonts, corporate fonts) |
| Payload | Regular + Bold TTF ≈ 700 KB; Latin subset ≈ 120 KB; cached after first load | 0 |
| Works in Office web / sandboxed iframe | Yes | Only if the font is installed and the iframe may use it; Office web users on Linux/ChromeOS lack Arial |
| Offline | Yes, once cached | Yes |

Decision: **bundled metrics are the primary measurer** (default style font Arial → Liberation Sans Regular/Bold, subsetted to Latin-1 + common symbols; opentype.js parses TTF, not WOFF2, so ship TTF or decompress). **Canvas is the fallback** when the deck uses another font: `measure.ts` tries the bundle, else `document.fonts.check()` + canvas, and surfaces a pane notice "measuring with installed font X; layout may differ slightly". Later: let users drop a `.ttf` into the pane (stored in IndexedDB) for corporate fonts, mirroring `--font` on the CLI.

### 3.3 Parity with Node
- `test/core.test.ts` golden scenes run in Node and, via Playwright or `node --test` + jsdom-free harness, in Chromium/WebKit against the same bundled font. Assertion: scene JSON byte-identical.
- One `Measurer` interface, no browser-only branches in `src/core`.
- The add-in pins the engine version it was built with and writes it into the chart's tags, so a later engine can migrate specs.

## 4. Insert / edit / refresh flows

### 4.1 Two insertion paths
**Path A — "Insert as new slide"** (milestone 1): run `layoutChart` + `writePptx` in the browser for a one-slide deck, base64 it, call `insertSlidesFromBase64(b64, { formatting: 'UseDestinationTheme', targetSlideId })`. Pros: exact output, 1.2-level API, three lines of host code, works on every platform the day it ships. Cons: it is a new slide, not a shape on the user's slide; with `UseDestinationTheme` our "Blank" layout maps to the destination's Blank layout, with `KeepSourceFormatting` our stub master comes along (avoid). Our colours and fonts are explicit, so theme mapping doesn't restyle the chart. Keep this path forever as the fidelity fallback and for "insert all slides of a deck.json".

**Path B — "Insert on this slide"** (milestone 2): walk the `Scene` and create shapes:

```
PowerPoint.run(async ctx => {
  const slide = ctx.presentation.getSelectedSlides().getItemAt(0);
  const created = scene.shapes.map(s => addShape(slide.shapes, s));  // queued, no sync
  created.forEach(sh => sh.load('id'));
  await ctx.sync();                                                  // sync #1
  const group = slide.shapes.addGroup(created);
  group.name = name; group.tags.add('OC_SPEC', json); group.tags.add('OC_V', version);
  group.tags.add('OC_FRAME', JSON.stringify(frame));
  await ctx.sync();                                                  // sync #2
});
```

`addShape` maps `RectShape → addGeometricShape(rectangle|ellipse|roundRectangle)`, `LineShape → addLine(straight)` (normalise so left/top is the min corner; direction is irrelevant without arrowheads), `TextShape → addTextBox` with margins 0, wordWrap off, autosize none, anchor/alignment/font from the shape. Z-order follows creation order, matching the writer. Every child also gets a tag `OC_ROLE` (`bar:s1:c2`, `label:value:s1:c2`, `axis`, …) so later diffing can tell what the user changed.

Scene → Office.js mapping, field by field (mirrors `shapeXml`/`textXml` in `src/render/pptx.ts` so the two hosts stay in step):

| Scene field (`src/core/types.ts`) | OOXML writer | Office.js |
|---|---|---|
| `RectShape.x/y/w/h` (pt) | `a:off`/`a:ext` in EMU | `{left, top, width, height}` in pt (no conversion) |
| `RectShape.ellipse` | `prstGeom ellipse` | `GeometricShapeType.ellipse` |
| `RectShape.fill` / `stroke` / `strokeWidth` | `solidFill` / `a:ln w` | `fill.setSolidColor('#…')`, `lineFormat.color/weight`, `lineFormat.visible=false` when no stroke |
| `LineShape.x1,y1→x2,y2` | `cxnSp` + `flipH/flipV` | `addLine(straight, {left:min x, top:min y, width:abs(dx), height:abs(dy)})` — flips dropped |
| `LineShape.dash` | `prstDash dash` | `lineFormat.dashStyle = 'Dash'` |
| `LineShape.arrowEnd` | `tailEnd triangle` | **gap** → head drawn as separate triangle (§4.1 fallbacks) |
| `TextShape.text` (multi-line) | one `a:p` per line | `addTextBox(text)` with `\n`; verify paragraph split in S5 |
| `TextShape.size/bold/color` | `a:rPr sz b solidFill latin` | `textRange.font.{size,bold,color,name}` |
| `TextShape.align/valign` | `algn`, `anchor` | `paragraphFormat.horizontalAlignment`, `textFrame.verticalAlignment` |
| writer defaults `wrap=none lIns=0… noAutofit` | `a:bodyPr` | `textFrame.wordWrap=false`, margins `=0`, `autoSizeSetting='AutoSizeNone'` |
| `TextShape.fill/stroke/rounded` (CAGR bubble) | `roundRect adj 50000` | `GeometricShapeType.roundRectangle` + fill/line; radius not adjustable |
| group `name`, `descr` | `cNvPr name descr` | `group.name`; alt text via `shape.altTextDescription` (⚠️ verify in S5) |

The host advertises what it could not express through a small capability record, which the engine already has a place to honour (the quality gate and label planner read it like any other constraint):

```ts
interface HostCaps { arrowheads: boolean; lineFlip: boolean; roundedRadius: boolean; altText: boolean }
```

Fallbacks for the §2.1 gaps: arrowheads → draw the head as a filled `triangle`/`isoscelesTriangle` rotated to the line angle if `shape.rotation` exists, else axis-aligned `downArrow`-style preset for vertical difference arrows and a chevron text glyph for diagonal CAGR arrows; rounded bubble → `roundRectangle` default radius (visually acceptable; writer uses adj 50 000). Each fallback is a `hostCaps` flag so the engine can be told "no arrowheads" and plan a slightly longer line.

### 4.2 Performance and batching
A typical chart is 60–250 shapes. Rules: never `sync` inside a loop; set all properties on the proxy before the first `sync`; load only `id`. Two syncs for insert, three for update (read tags + bounds → delete → create/group). Office web and Mac are the slow hosts; budget ≤ 1.5 s for 250 shapes, measured in S4. If creation proves slow on web, Path A remains an "exact but new slide" option and we can offer "insert as picture (SVG)" for drafts.

### 4.3 Storing the spec on the chart
- `OC_SPEC` tag on the group: compact JSON of `ChartSpec` (+ `style` overrides). Tags persist in the .pptx (`ppt/tags/tagN.xml` via `p:custDataLst`), survive copy/paste between decks, and are invisible to users. Alt text (`descr`) stays human-readable — never JSON; screen readers announce it.
- `OC_V` engine version, `OC_FRAME` the frame the layout was planned for, `OC_FONT` the measured font.
- Tag value size limits are undocumented; S2 measures them. If a spec exceeds the limit, fall back to a presentation-level custom XML part keyed by group id (1.7) or chunked tags `OC_SPEC_0..n`.
- **Parity with the CLI:** teach `src/render/pptx.ts` to emit the same tags part, so a deck generated by the CLI is editable in the add-in, and the native Windows add-in (which reads the same tags, see `docs/native-windows-design.md`) interoperates. Tag schema is shared and versioned in `docs/plan/` as the single contract.

Tag contract v1 (keys are upper-cased by PowerPoint; values are strings):

| Key | On | Value | Notes |
|---|---|---|---|
| `OC_V` | group | `"1"` | contract version, not engine version; bump only on incompatible change |
| `OC_ENGINE` | group | `"0.1.0"` | engine version that produced the layout; used for migrations |
| `OC_SPEC` | group | JSON `ChartSpec` + `{style?: Partial<Style>}` | whitespace-free; chunked to `OC_SPEC_0…n` + `OC_SPEC_N` if over the S2 limit |
| `OC_FRAME` | group | `"x,y,w,h"` in pt | frame the layout was planned for; compared on select |
| `OC_FONT` | group | `"Arial"` | font the text was measured with |
| `OC_ROLE` | each child | `bar:s1:c2`, `marker:s0:c3`, `label:value:s1:c2`, `label:total:c2`, `label:category:c2`, `label:series:s1`, `axis`, `tick:3`, `connector:c2`, `arrow:cagr`, `bubble:cagr`, `legend:s1` | stable across rebuilds so overrides can be matched |
| `OC_HIDDEN` | group | JSON string[] of roles the user deleted | written by the override capture (§4.6) |

### 4.4 Edit flow
1. `DocumentSelectionChanged` → `getSelectedShapes()` → load `name`, `tags`, `left/top/width/height` of the first shape (one sync). If a child is selected, walk `shape.parentGroup` (if available; else match by `OC_ROLE` + slide scan).
2. If `OC_SPEC` present: pane switches to edit mode, datasheet populated, "Update" enabled. Unknown shape: insert mode.
3. **Update**: re-run layout with the group's *current* bounds as the frame (keeps the user's position and size), delete the group, recreate, re-tag. Keep `name` and alt text updated.

Task-pane states, so the UI is specified before it is built:

| State | Trigger | Pane shows | Ribbon |
|---|---|---|---|
| Insert | nothing / foreign shape selected | datasheet with sample data, chart type, options, preview, **Insert on slide** / **Insert as new slide** | Update, Re-layout disabled |
| Edit | open-charts group (or a child) selected | datasheet filled from `OC_SPEC`, preview of *current* spec at *current* bounds, **Update**, **Revert** | Update, Re-layout enabled |
| Resized | Edit + bounds ≠ `OC_FRAME` | banner "Chart was resized — re-layout to new size" | Re-layout highlighted |
| Stale engine | `OC_ENGINE` older than bundle | notice "Rebuilt with a newer engine; labels may move" on Update | — |
| Unsupported | API floor missing for Path B | only **Insert as new slide**; explanation with platform/build | Insert only |

Multiple charts on a slide need no special handling: each group carries its own tags, and selection decides which one the pane is bound to. Multi-selection (several groups) disables Update.

### 4.5 Re-layout on resize
PowerPoint scales a group's children on resize, including text box frames but not font sizes — exactly what breaks label fit. There is no resize event, so: on every selection of one of our groups compare bounds with `OC_FRAME`; if they differ by > 0.5 pt show a one-click "Re-layout to new size" banner (and a setting "always re-layout automatically on select"). The ribbon **Re-layout** button does the same explicitly.

### 4.6 User manual formatting
v1 policy: Update/Re-layout **rebuilds the chart** and tells the user beforehand that manual formatting on the chart will be replaced. Mitigations, in order of delivery:
1. Per-shape `OC_ROLE` tags make it possible to read back common manual edits before deletion — series fill colours, label font size/bold, deleted labels — and write them into the spec as overrides (`series[i].color`, `options.decimals`, a `hidden` list). Milestone 4.
2. Users can always ungroup and own the shapes; the group then has no tags and the add-in treats it as plain shapes.
3. Never silently reformat: the diff is shown ("kept: 2 colour changes; discarded: 1 moved label") before committing.

### 4.7 Undo
PowerPoint pushes each Office.js mutation onto the normal undo stack; a batch within one `sync` is believed to appear as one or a few entries, but this is not documented and differs on web. S3 measures how many Ctrl+Z presses undo an insert on each platform. Mitigations: keep sync count minimal; keep the previous spec+frame in pane memory and offer "Revert last update" in the pane, which rebuilds from it (one more rebuild, but deterministic). Document the behaviour honestly in the pane help.

## 5. Mac specifically
- **Why native is impractical for third parties.** Mac Office has no COM/VSTO; the only in-process extension points are VBA (sandboxed, no ribbon customization, no distribution story, disabled in many enterprises) and AppleScript (external, no UI). Microsoft's supported model for third-party Mac add-ins is exclusively Office.js. Hooking into the Mac app's private frameworks would be unsupported and fragile. So on Mac this add-in *is* the product; there is no Mac counterpart to `docs/native-windows-design.md`.
- **What Office.js gives on Mac.** Task pane and commands in a WKWebView (Safari/WebKit engine — test in Safari, not only Chromium); the same PowerPointApi sets, usually shipping a little later than Windows (Mac Insider builds first); Dialog API; `insertSlidesFromBase64`; SharedRuntime. Sideload by dropping the manifest into `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`. Debug with Safari Web Inspector after `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`.
- **Gaps.** API-set lag (feature-detect, never assume); unified manifest support status (S1); WebKit-specific issues (IndexedDB/service-worker behaviour in WKWebView, `document.fonts` availability); no file-system or clipboard beyond standard web APIs (download of a .pptx from the pane opens a save sheet, fine); no process-level font enumeration — hence the bundled-metrics decision. Performance on Mac is between Windows and web.

## 6. Distribution, hosting, offline, security

- **Sideload** (dev and early users): `npx office-addin-debugging start addin/manifest.xml` on Windows/Mac; "Upload My Add-in" in PowerPoint for the web; wef folder on Mac; iPad via Settings ▸ Add-ins is impractical, use centralized deployment.
- **Microsoft 365 centralized deployment** (Integrated apps in the admin centre): the route for firms deploying to their consultants; takes our manifest URL or file; honours the `Requirements` floor. Windows-only network share catalogs are covered by the native doc.
- **AppSource**: publish once stable. Requirements that shape the code: HTTPS everywhere, office.js from CDN, no `eval`, accessibility (keyboard-navigable datasheet, contrast), privacy policy + support URL, an `AppDomains` list, test notes for validators (sample data!). Expect 1–2 review rounds; budget 2–3 weeks calendar time.
- **Hosting**: a static site is all we need — GitHub Pages or Cloudflare Pages from `addin/dist/` with immutable, hashed asset names; manifest `SourceLocation` points at a versioned path so old manifests keep working. No server, no database, no telemetry by default (opt-in anonymous usage counter at most, off by default, disclosed).
- **Offline**: Office loads the pane from the URL every time; it does not cache add-ins for you. A service worker precaching `dist/` makes the add-in work offline after first use on Windows (WebView2) and in Office web (browser rules); WKWebView service-worker support must be verified (S5). Fallback: an "offline bundle" — the same static site run from `open-charts serve` on localhost with a sideloaded manifest pointing at `https://localhost:3000` (the standard dev-cert flow), documented for air-gapped users.
- **Security / privacy**: all computation in the pane; the only data crossing a network is the add-in's own static assets. Strict CSP (`default-src 'self'; script-src 'self' https://appsforoffice.microsoft.com`), SRI on our bundles, no third-party scripts, no fonts from Google. Specs stored in tags live in the user's file only. No SSO or Entra registration needed because we never call Graph.

### 6.1 Testing strategy
- **Engine**: existing golden tests plus the browser run (§3.3). No Office needed.
- **Host layer**: `host/powerpoint.ts` is written against a tiny interface (`addRect/addLine/addText/group/tag`), with a fake implementation that records calls; unit tests assert the call sequence and sync count for each example chart. The real implementation is a thin adapter over `PowerPoint.*`.
- **Integration**: a manual checklist per platform (Windows current + semi-annual channel, Mac current, web, iPad) run before each release, scripted where Office supports it (Office web via Playwright on a test tenant). Each example in `examples/` is inserted both ways and the saved .pptx is diffed against the CLI output at the XML level (ignoring ids and `p:custDataLst` ordering).
- **Round-trip**: CLI → .pptx → open in add-in → Update → save → CLI reads tags back (a small `open-charts inspect deck.pptx` command) → spec equal.

## 7. Milestones and estimates

Estimates are engineer-weeks for one engineer familiar with the engine, excluding AppSource review wait time.

| # | Milestone | Deliverable | Est. |
|---|---|---|---|
| 0 | Spikes (§8) | Verified API matrix, numbers for tag size, undo, perf; go/no-go on Path B per platform | 2 w |
| 1 | Browser engine + Path A | `core/text.ts` split, esbuild bundle, bundled Liberation metrics, browser golden tests; task pane = ported editor; "Insert as new slide" on Win/Mac/web; sideload docs | 2 w |
| 2 | Path B insert | Scene → shapes on current slide, grouped, tagged; fallbacks for arrowheads/rounded; perf budget met | 2 w |
| 3 | Edit + re-layout | Selection handling, spec read-back, Update in place, resize banner, Revert; CLI writer emits tags | 2 w |
| 4 | Manual-formatting preservation | `OC_ROLE` diffing, override capture, diff preview | 2 w |
| 5 | Polish + distribution | Accessibility pass, iPad pass, service worker, CSP/SRI, hosting, centralized-deployment guide, AppSource submission | 2 w |
| | **Total** | | **~12 w** |

Milestone 1 alone is useful (every platform, exact output) and should ship as 0.x to early users; milestones 2–3 are the product.

## 8. First two weeks: spikes
Each spike is a small throwaway page under `addin/spikes/` sideloaded on Windows, Mac, and web (iPad where possible), with results recorded in `docs/plan/spike-results.md`.

| ID | Question | Method | Exit criterion |
|---|---|---|---|
| S1 | Which PowerPointApi sets (1.1–1.10) and which manifest format are live on each platform/channel? | Print `isSetSupported` matrix; try unified manifest sideload on Mac and iPad | Table in §2.1 corrected; manifest decision made |
| S2 | Tag value size limit and persistence | Write 1 KB, 32 KB, 256 KB tag values; save; reopen; copy-paste the group to another deck; open in Keynote/LibreOffice and re-save | Max size known; spec storage strategy fixed (tag vs chunked vs custom XML) |
| S3 | Undo granularity | Insert a 150-shape chart via 2 syncs; count Ctrl+Z steps per platform; same for delete+recreate | Behaviour documented; decide whether in-pane Revert is required |
| S4 | Shape-creation performance | Create 50/150/300 shapes + group, time per platform; vary sync strategy | Budget met or Path B gated per platform |
| S5 | Arrowheads, rotation, roundRectangle, paragraph alignment, textFrame margins/wrap/autosize actually settable? | Call each API, inspect saved XML | Fallback list in §4.1 confirmed or shortened |
| S6 | Selection event reliability and group handling | Register `DocumentSelectionChanged`; select child inside group, group, nothing; on all platforms | Edit-flow entry point proven |
| S7 | Browser measurer parity | Run golden tests with bundled Liberation in Chromium and WebKit vs Node | Byte-identical scenes, or documented deltas |
| S8 | Offline / service worker in WebView2 and WKWebView | Precache dist, go offline, open pane | Offline story in §6 confirmed |
| S9 | `insertSlidesFromBase64` theme mapping | Insert our one-slide deck with both `formatting` options into a corporate template | Correct layout mapping; no stray master |

## 9. Risks
- **API gaps are wider than believed** (groups, alignment, margins missing on some platform): Path A still ships; Path B degrades per platform via `hostCaps`. Mitigated by S1/S5 before committing design.
- **Performance on Office web/Mac** makes 200-shape charts feel slow: batching discipline; Path A fallback; "insert as picture" for drafts.
- **Undo surprises** erode trust: measured in S3; in-pane Revert; honest documentation.
- **Rebuild destroys manual edits**: explicit warning in v1, override capture in milestone 4.
- **Font parity**: Arial is covered exactly; other deck fonts go through canvas and may differ by a few percent — the quality gate still prevents overlaps because measurement and layout use the same numbers, but PowerPoint's rendering may differ. Document; allow custom `.ttf` upload.
- **Manifest churn** (XML deprecation timeline, unified-manifest platform matrix): keep manifest minimal; conversion is mechanical.
- **Spec-on-shape contract drift** between this add-in, the CLI and the native Windows add-in: one versioned tag schema document, one shared test fixture.
- **AppSource validation** delays: not on the critical path for sideload/centralized users; submit early with milestone 3.
