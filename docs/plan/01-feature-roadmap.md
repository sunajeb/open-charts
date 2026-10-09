# 01 — Feature inventory and phased roadmap

This section lists every user-facing feature of the reference product (a best-in-class consulting PowerPoint chart add-in, studied through its public manual and ribbon layout), scores each one, and assigns it to a phase of open-charts. It is the master checklist the other plan sections build on.

## How to read the tables

| Column | Meaning |
|---|---|
| **Value** | What the feature is worth to a consulting / corporate-finance user. H = used on most decks; M = used weekly; L = niche or cosmetic. |
| **Diff.** | Build difficulty for us. H = new subsystem or hard algorithm; M = a few days of focused work on top of the engine; L = hours. |
| **Surface** | Where the feature has to live. **E** = layout engine + OOXML writer (headless, spec in → shapes out). **ED** = web editor (datasheet, preview, download). **OJ** = Office.js task-pane add-in (insert/replace shapes, read selection, no sub-shape handles). **NW** = native Windows add-in (COM/VSTO: selection of sub-shapes, drag handles, context menus, command overrides). **XL** = Excel side (add-in or script). |
| **Phase** | P1–P5 as defined in §4. **—** = deliberately not built (reason in §5/§6). |

Status marks in the Notes column: ✅ exists in open-charts today, 🔶 partial, blank = not started.

## 1. The reference product in one paragraph

A ~75 MB native core runs inside PowerPoint and Excel. PowerPoint gets its own ribbon tab (Structure · Elements · Data · Text · Layout · Color · Quality · Collaboration · More) plus groups on Insert, Design, Review and Shape Format; it also overrides built-in commands (Group, Ungroup, Regroup, Print, Mark as Final, Redo) so its elements survive them. Excel gets a tab with Elements, Formulas (rounding, CAGR), Data (chart scanner) and Styles groups. Every element (chart, table, text box, Harvey ball, checkbox, agenda, Gantt) has an embedded Excel datasheet and can instead be linked to an external workbook, a per-slide workbook, or a Tableau view. Direct manipulation is the core UX: selection frames, drag handles, a context menu and a mini toolbar per element and per chart feature. Around that sit a content library, a brand/style file system, clean-up tools, a COM API plus JSON (.ppttc) batch automation with a server mode, and an AI side pane. Three licence tiers are visible: Essentials (layout tools, text boxes, tables, agenda, Harvey balls), Suite (all charts, datasheets, data links, automation) and an Assist add-on (AI).

## 2. Feature inventory

### 2.1 Charts by type

| Feature | Value | Diff. | Surface | Phase | Notes |
|---|---|---|---|---|---|
| Stacked column (simple = one series) | H | M | E | P1 | ✅ totals, CAGR, difference arrows, series labels |
| Stacked bar (rotated column) | H | L | E | P1 | Rotation of the column layout, labels re-flowed horizontally |
| 100% stacked column / bar | H | M | E | P1 | ✅ 🔶 needs largest-remainder rounding so labels sum to 100; columns may total ≠100% |
| Clustered column / bar | H | M | E | P1 | ✅ legend, value axis |
| Stacked-clustered (0% gap + category gaps) | M | M | E | P1 | Compose from gap controls, no new type |
| Waterfall / bridge, build-up and build-down | H | H | E | P1 | 🔶 "e" (equals) subtotal columns, connectors, mixed-sign columns, start-from-right calculation |
| Mekko with % axis (marimekko) | H | H | E | P1 | 🔶 widths from column totals, 100% indicator, X value axis |
| Mekko with units (submarine) | M | M | E | P1 | Explicit X-extent row |
| Line chart (category or value X axis, dates) | H | M | E | P1 | 🔶 markers, end labels, interpolation of gaps, smoothing |
| Profile chart (rotated line) | L | L | E | P1 | Free once line + rotation exist |
| Combination (columns + lines, per-series type switch) | H | M | E | P2 | Per-series `kind` in spec |
| Secondary Y axis for lines | M | M | E | P2 | Two independent nice-tick scales |
| Error bars; football-field (rotated line + thick error bars) | M | M | E | P2 | Valuation decks need football fields |
| Stacked area / 100% area | M | M | E | P2 | Same data model as line |
| Pie / doughnut (rotation, explode, hole size, order) | M | M | E | P2 | Inside/outside labels with leader lines |
| Scatter / bubble (groups, marker scheme, area-or-diameter sizing, size legend) | M | H | E | P2 | Needs a real point-label placer (simulated annealing or greedy with leader lines) |
| Trendlines (linear, power, exp, log, poly 2–4) and partition lines with region fills | L | M | E | P2 | Weighted regression for bubbles |
| Butterfly (two back-to-back bars, shared scale) | M | L | E | P2 | Composition of two charts + `sameScale` group |
| Net-line series and step lines | L | M | E | P2 | Series rendered as horizontal ticks; optional connectors |
| "Other" series (fold small segments) | M | L | E | P2 | Spec-level transform |
| Segment / area / slice ordering (sheet, reverse, asc, desc per category) | M | L | E | P1 | Option on every stacked type |
| Category ordering (sheet, reverse, by total) + ordinal labels | M | L | E | P2 | |
| Gantt / timeline | H | H | E · ED · NW | P5 | See §2.10 |
| Charts with attached data table | M | M | E | P5 | Depends on tables (§2.11) |

### 2.2 Chart decorations / annotations

| Feature | Value | Diff. | Surface | Phase | Notes |
|---|---|---|---|---|---|
| Total difference arrow (between column totals) | H | M | E | P1 | ✅ 🔶 three modes: relative →, absolute ↔, relative ← |
| Level difference arrow (segment↔segment, point↔point, to value line) | H | M | E | P1 | Label bubble moves beside arrow when arrow is short |
| CAGR arrow (from category dates, 30/360 day count for partial years) | H | M | E | P1 | ✅ 🔶 multiple arrows for successive selected categories |
| Series CAGR in series label | M | L | E | P1 | Label-content variant |
| Value line (snaps to mean / weighted mean, label left or right) | M | L | E | P1 | |
| 100% indicator (100% charts, Mekko) | M | L | E | P1 | Left or right |
| Series connectors (stacked, Mekko) | M | L | E | P1 | Per chart or per segment |
| Waterfall connectors as part of the calculation | H | M | E | P1 | Removing a connector starts a new summation |
| Universal connector between any two anchor points, incl. across charts | L | M | E · NW | P4 | Needs anchor registry in scene |
| Axis break (straight / wiggle), extent adjustable | M | H | E | P2 | Compressed value range applied to all segments sharing it |
| Category axis break and category gaps (width-preserving or chart-width-preserving) | M | L | E | P1 | |
| Legend (box line style, font, custom order, resizable → horizontal) | H | L | E | P1 | ✅ 🔶 |
| Bubble size legend | L | L | E | P2 | |
| Doughnut title in hole | L | L | E | P2 | |
| Ridge outline on Mekko | L | L | E | P1 | |

### 2.3 Axes and scales

| Feature | Value | Diff. | Surface | Phase | Notes |
|---|---|---|---|---|---|
| Nice tick spacing, auto extent covering all data | H | M | E | P1 | ✅ Heckbert / Talbot |
| Manual extent and tick spacing | H | L | E | P1 | ✅ options; drag handles come in P4 |
| Axis line vs gridlines; tick marks; tick labels left/right; axis title | H | L | E | P1 | 🔶 |
| Absolute ↔ % axis type (converts stacked→100%, clustered→stacked, line→100% area) | M | L | E | P1 | Spec transform |
| Reverse axis (scatter, bubble, line) | L | L | E | P2 | |
| Y crosses at / between categories (line) | M | L | E | P1 | |
| Same scale across several charts on a slide | H | M | E | P1 | Deck-level `scaleGroup`; the one feature consultants ask for first |
| Logarithmic scale (powers of 10 only; negatives flagged) | L | M | E | P2 | |
| Date axis (yy/yyyy mixing, week-start, tick units in years/months/weeks) | M | M | E | P2 | |
| Line X axis auto-switches to value axis for strictly increasing numbers/dates | M | M | E | P2 | |

### 2.4 Text, labels and number formats

| Feature | Value | Diff. | Surface | Phase | Notes |
|---|---|---|---|---|---|
| Label types: category, series, segment, point, total, grand total, inside/outside (pie), percent indicator, tick, axis title, value-line | H | M | E | P1 | ✅ 🔶 |
| Automatic placement: centre; boxed background when label > segment; left/right stagger; outside with leader line | H | H | E | P1 | ✅ 🔶 the quality gate (no overlap, nothing outside frame) is already the differentiator |
| Manual placement overrides that survive relayout ("drag here for automatic") | H | M | E · NW | P4 | Engine: `pin` per label; UI: drag in P4 |
| Rotate labels ±90° | M | L | E | P1 | |
| Label content: absolute, %, category, series, combos, CAGR, ordinal, x/y/size | H | L | E | P1 | 🔶 |
| Number format by example ("1.000,00", "$1.2M", "+1,234", "(1.2M)", en-dash minus, prime/double-prime magnitudes) | H | M | E | P1 | 🔶 implement the by-example parser; it is small and high-leverage |
| Magnitude ×10³ / ×10⁶ without changing data | H | L | E | P1 | ✅ `unit` |
| Text fields: numbers stay live inside free text; "!" marker when overwritten | M | M | E · NW | P4 | Engine keeps `template` + `fields`; UI edits in P4 |
| Font / size / colour per label or per chart, auto-contrast text colour | H | L | E | P1 | ✅ contrast |
| Date format codes (d, ddd, w, www, M, MMM, Q, QQQ, yy, ¶ line break …) | M | L | E | P2 | Shared with Gantt |
| Decimal-symbol switch (point ↔ comma) per slide / deck | M | L | E · OJ | P3 | Trivial once formats are data |
| Proofing language drives placeholder text and RTL layout | L | M | E | P5 | RTL mirror of layouts |
| Replace & resize fonts across deck | L | L | OJ | P3 | Not chart-specific |
| Insert symbol dialog | L | L | OJ | — | PowerPoint has one |
| Paste text into multiple objects | M | L | OJ | P3 | Nice editor/add-in utility |

### 2.5 Chart area, plot area, auto-layout

| Feature | Value | Diff. | Surface | Phase | Notes |
|---|---|---|---|---|---|
| Chart area vs plot area, both selectable | H | M | E | P1 | ✅ 🔶 engine already distinguishes; expose both frames in scene |
| Lock / unlock edges; area grows in the unlocked direction to fit content | H | H | E · NW | P4 | Engine: `fit` rules per edge; UI: handles |
| Dynamic alignment between elements (edge snaps that persist on relayout) | H | H | E · NW | P4 | Constraint solver; the heart of "automatic layout" |
| Match size (same width/height/size) with reference object; size-constraint arrows; units pt/mm/cm/in/slide-fraction | M | M | NW | P4 | |
| Align to reference object; red snap guides while dragging | M | M | NW | P4 | |
| Swap two objects | L | L | OJ | P3 | |
| Scale a selection proportionally incl. text and line weights | M | M | OJ | P3 | |
| Preset chart placement layouts (from style file) | M | L | E | P2 | Slide templates with named chart slots |

### 2.6 Datasheet and data entry

| Feature | Value | Diff. | Surface | Phase | Notes |
|---|---|---|---|---|---|
| Per-element datasheet (grid, formulas, paste from Excel) | H | M | ED | P1 | ✅ 🔶 editor grid; formulas later |
| Absolute vs relative values; `100%=` row to override the base | H | L | E · ED | P1 | Needed for 100% and Mekko-% |
| Transpose; reverse row/column order | M | L | ED | P1 | |
| Edit data layout: optional rows/columns (series, category, 100%=, X-extent, group, size) | M | L | E · ED | P1 | Spec columns are optional already |
| Hidden rows/columns excluded from chart | L | L | E | P2 | |
| Chinese conversion, F7 spell-check in datasheet | L | — | — | — | Host features |
| Per-slide workbook (one Excel for all elements on a slide) | M | H | OJ · XL | P5 | Replace with "one data file per deck" (§5) |
| Chart scanner (image → data via OCR) | M | H | ED | P2 | Vision-model call instead of Tesseract; also fed by AI phase |
| Import Mekko Graphics charts | L | H | NW | — | Legacy-competitor migration; skip |

### 2.7 Excel, Tableau and external data links

| Feature | Value | Diff. | Surface | Phase | Notes |
|---|---|---|---|---|---|
| Create chart / table / Harvey ball / checkbox / image-from-URL from a selected Excel range | H | H | XL · OJ | P5 | Excel Office.js add-in sends range → engine → inserts into PowerPoint |
| Link existing element to a range; link slide text to a cell | H | H | XL · OJ | P5 | |
| Colour-coded link frames in Excel (green/red/blue) | M | M | XL | P5 | |
| Adjust linked range: move frame, add/remove rows, transpose | M | M | XL | P5 | |
| Update once / revert / update automatically; pending-update indicator | H | M | OJ | P5 | |
| Data Links dialog: all links, thumbnails, filters, alternate sources, bulk update | M | M | OJ | P5 | |
| Restore disconnected links; switch to alternate workbook | M | M | OJ | P5 | |
| Match formatting to Excel cells (number format, bold, colour, fill, borders, conditional formatting) | M | M | E · XL | P5 | |
| Table from Excel with formatting; table as image | M | M | XL · OJ | P5 | |
| Tableau view → element via browser extension | L | H | browser | — | Replace with CSV/API connectors (§5) |
| Delete disconnected links (Excel side) | L | L | XL | P5 | |

### 2.8 Layout and alignment tools (non-chart PowerPoint utilities)

| Feature | Value | Diff. | Surface | Phase | Notes |
|---|---|---|---|---|---|
| Duplicate in a direction with auto-alignment (Ctrl+Alt+arrows) | M | M | NW | P4 | |
| Rotate / flip elements to snap targets | L | L | NW | P4 | |
| Group / ungroup / regroup that preserve element identity (command override) | H | M | NW | P4 | Needed so our grouped charts survive user grouping |
| Keyboard selection (Alt+arrows), Ctrl+A widening selection, outer selection frame | M | M | NW | P4 | |
| Rounded rectangles, pentagons/chevrons, process flows (auto-fit, striping) | M | M | E · NW | P5 | |
| Text boxes with auto-fit edges | M | M | E · NW | P5 | |

### 2.9 Agenda

| Feature | Value | Diff. | Surface | Phase | Notes |
|---|---|---|---|---|---|
| Chapter slides auto-inserted; agenda replicated and highlighted on every chapter slide | H | H | OJ · E | P5 | Deck-level model: chapters ↔ slide ranges |
| Reorder chapters moves their slides; multilevel hierarchy via text levels | M | H | OJ | P5 | |
| Hide empty parent chapter slides automatically | L | M | OJ | P5 | |
| Split agenda (backup section); table-of-contents slide | M | M | OJ | P5 | |
| 12 style states (current/top-level/in-subtree/…) | M | M | E | P5 | Encode as style matrix |
| Default agenda layout in slide master | L | M | E | P5 | |

### 2.10 Gantt / timeline

| Feature | Value | Diff. | Surface | Phase | Notes |
|---|---|---|---|---|---|
| Calendar scales (year, quarter, month, week, day; separators; label formats; fiscal calendar) | H | H | E | P5 | |
| Rows with activity / responsible / remark label columns, shading, separators | H | M | E | P5 | |
| Items: bars, process arrows, milestones (+ lines), brackets, shades | H | M | E | P5 | |
| Anchors (item ↔ item, bracket ↔ shade), item labels | M | M | E | P5 | |
| Datasheet with dates; Excel link | M | M | ED · XL | P5 | |
| Drag items / rows / date range | M | H | NW | P5 | |
| RTL timelines | L | L | E | P5 | |

### 2.11 Tables, Harvey balls, checkboxes

| Feature | Value | Diff. | Surface | Phase | Notes |
|---|---|---|---|---|---|
| Table = grid of auto-fit text boxes; add/delete rows/cols; striping | H | M | E | P5 | |
| Cell content types: text, Harvey ball, checkbox, chevron, rounded rect, image | M | M | E | P5 | |
| Table with datasheet / linked to Excel; "duplicate chart data as table" | M | M | E · ED | P5 | |
| Harvey ball: n segments (default 4), completion 0–n, fills, outline, size units | H | L | E | P2 | Cheap, popular; ship early as a standalone element |
| Checkbox: ✓ ✗ ☐ + custom symbols from style | M | L | E | P2 | |
| Harvey ball / checkbox from Excel cell value (0–4, %, formula) | M | L | XL | P5 | |

### 2.12 Library, templates, images

| Feature | Value | Diff. | Surface | Phase | Notes |
|---|---|---|---|---|---|
| Slide-asset library over local/network/OneDrive with slide-level text search, filters (date, author, master, layout, aspect), duplicate hiding | M | H | OJ | — | Out of charting scope; see §6 |
| 450+ slide templates adapting to the master | M | M | E | P2 | We ship a *chart-slide* template set as JSON specs instead |
| Insert slide objects onto current slide (destination/source formatting) | L | M | OJ | — | |
| Stock images / icons (Pexels, Unsplash, Magnific) and DAM connectors (Brandfolder, Canto, Getty) | L | M | OJ | — | |
| Image element: fill-frame / specify-height / specify-width modes, crop, alignment | L | M | NW | — | |
| SVG → shapes → image round-trip | L | L | OJ | — | |
| Browser extension: send web image to slide | L | M | browser | — | |

### 2.13 Styles and brand

| Feature | Value | Diff. | Surface | Phase | Notes |
|---|---|---|---|---|---|
| Theme-driven colours and fonts (body font, accent colours) | H | L | E | P1 | ✅ 🔶 read theme from a .potx/.pptx master |
| Style file: fill / line / marker lists; fill, line, marker *schemes*; default scheme per chart type; overflow rule when series > fills | H | M | E | P2 | Our style = JSON/YAML, diffable |
| Default label, annotation, legend, gridline, text styles; master text-level overrides | M | M | E | P2 | |
| Harvey ball, checkbox, Gantt, agenda defaults; arrow label shapes | M | L | E | P2/P5 | |
| Load style per slide master / layout / *region of a layout*; API to load/get/remove | M | M | E · OJ | P3 | Region-scoped styles map to named chart slots |
| Deploy default style org-wide (admin) | M | L | — | P3 | Config file + env var |
| File name / path text box in master | L | L | OJ | — | |

### 2.14 Collaboration and quality tools

| Feature | Value | Diff. | Surface | Phase | Notes |
|---|---|---|---|---|---|
| Clean Up: alt text, shape locks, animations, ink, data links, transitions, comments, notes, invisible/off-slide, workbooks, hidden slides, doc properties, unused masters/layouts, custom XML | M | M | E (pptx rewrite) · OJ | P3 | Headless `open-charts clean deck.pptx` is a differentiator |
| Sanitize numbers (randomise values, keep layout) | M | L | E | P2 | Trivially a spec transform; great for sharing examples |
| Send / save selected slides as a new deck (strip workbooks) | L | L | OJ | P3 | |
| Compatibility: files are 100% native; re-import elements later | H | M | E | P1 | ✅ native shapes; re-import = spec embedded as alt-text/custom XML |
| Mark-as-final / print overrides | L | — | NW | — | |

### 2.15 Automation and API

| Feature | Value | Diff. | Surface | Phase | Notes |
|---|---|---|---|---|---|
| JSON → deck from templates with named elements (`.ppttc`): templates array, datasheet rows, typed cells (string, number, percentage, date), empty rows, text fields, Harvey/checkbox values | H | M | E | P1/P2 | ✅ `deck.json` is already this; add template filling in P2 |
| Remote templates / remote data; server mode with auto-start service | M | M | E | P2 | HTTP service around the CLI |
| COM API (Windows): PresentationFromTemplate, UpdateBatch, UpdateChart | M | H | NW | P4 | Expose same operations over Office.js first |
| Style API: LoadStyle, LoadStyleForRegion, GetStyleName, RemoveStyles | L | L | E | P2 | |
| Element naming (AddRangeData / AddRangeImage) in templates | H | L | E | P2 | Named slots in .pptx via alt text |
| Excel-side formulas: TCROUND family (consistent rounding), TCCAGR | M | M | XL · E | P2 (engine) / P5 (Excel) | Engine gets consistent-rounding solver for totals |
| Excel table/cell style favourites | L | L | XL | — | |

### 2.16 AI features (the "Assist" add-on)

| Feature | Value | Diff. | Surface | Phase | Notes |
|---|---|---|---|---|---|
| Chat about the current slide; role prompts; web search | M | M | ED · OJ | P2 | |
| Rewrite / shorten / bullets / parallel phrasing; custom prompts | M | L | OJ | P3 | Text, not charts |
| Translate slides preserving formatting | M | M | OJ | P3 | |
| Action titles; executive-summary slide | H | M | ED · OJ | P2 | Title from chart data is the killer feature: we have the numbers |
| Generate tables from prompt + sources | M | M | ED | P2 | |
| Upload .pptx/.xlsx/.docx/.pdf/images as context | M | M | ED | P2 | |
| Research data: Morningstar financials, Statista → chart | M | H | ED | — | Licensed data; use open sources / user-supplied connectors |
| Excel: analyse / transform workbook, then "create chart in PowerPoint" with suggested chart types | H | M | ED · XL | P2 (editor) / P5 (Excel) | Chart-type recommendation from data shape |
| Chart-type auto-selection from data | H | M | E | P2 | Deterministic rules first, model second |

## 3. Phase overview

| Phase | Theme | Primary surface | Rough span | Headline deliverable |
|---|---|---|---|---|
| **P1** | Chart quality parity | E · ED | now → +3 months | Six core types indistinguishable from the reference at 100% zoom; the full decoration/label/axis set for them |
| **P2** | Data-to-deck + AI | E · ED | +3 → +6 months | Template filling, style files, remaining chart types, chart-type recommendation, action titles, chart scanner |
| **P3** | Office.js add-in | OJ | +6 → +9 months | Insert, re-open and edit any open-charts element inside PowerPoint (Windows, Mac, web) |
| **P4** | Native Windows direct manipulation | NW | +9 → +15 months | Handles, context menus, auto-layout constraints, command overrides |
| **P5** | Breadth | E · OJ · XL | +12 → +18 months | Gantt, agenda, tables, Excel links |

Phases overlap: P2 engine work starts while P1 polishes; P5 engine-only parts (Gantt layout, table layout) can start in P2 when contributors appear.

## 4. Phase definitions and exit criteria

### P1 — Chart quality parity (stacked, 100%, clustered, line, waterfall, Mekko)

Scope: everything marked P1 in §2 — the six types, all decorations that apply to them, axes, labels, number formats, same-scale groups, theme colours, datasheet basics in the editor.

Exit criteria (all must hold):
1. **Golden gallery.** ≥ 40 reference slides (≥ 6 per type) rebuilt from their data; a side-by-side review by two consultants finds no visible difference in label placement, number formatting, arrow geometry or spacing at 100% zoom. Differences are logged as issues until the count is zero.
2. **Quality gate green on every example**: zero label overlaps, zero out-of-frame shapes, 100% columns label-sum to exactly 100, waterfall `e` columns balance, Mekko widths sum to plot width within 0.5 pt.
3. **Native round-trip**: every output opens in PowerPoint (Win, Mac), Keynote and LibreOffice without repair prompts; ungroup → regroup keeps the chart intact; `open-charts import deck.pptx` recovers the spec from the embedded metadata.
4. **Text metrics**: measured label widths within ±2% of PowerPoint's for Arial/Liberation Sans, Calibri and one brand font loaded from a .potx.
5. **Editor**: paste a range from Excel, see the preview update < 200 ms, download .pptx; transpose, reverse, `100%=` and optional rows work.
6. **Spec is stable**: `ChartSpec` v1 frozen with a JSON Schema; all breaking changes thereafter are migrations.

### P2 — Data-to-deck + AI

Scope: template filling with named slots, style files, remaining chart types (combo, area, pie, scatter/bubble, error bars, secondary axis, axis breaks, date axes), Harvey balls and checkboxes as elements, chart-type recommendation, action titles, chart scanner, sanitize, server mode.

Exit criteria:
1. `open-charts fill template.pptx data.json` replaces every named chart, text field, Harvey ball and checkbox in a 30-slide template with new data, preserving positions, styles and annotations; difference/CAGR arrows recompute.
2. A `style.json` reproduces a client's existing deck look (fills, line styles, label fonts, gridlines, arrow styles) for all P1+P2 types; switching style files changes nothing but appearance.
3. Chart-type recommendation picks the type a consultant would choose on a 100-case benchmark ≥ 85% of the time; the model-assisted path explains its choice.
4. Action-title generation produces a title judged "usable with light edits" on ≥ 80% of the gallery slides.
5. Chart scanner recovers values from 50 published bar/column chart images within 3% of the true values for ≥ 90% of bars.
6. Server mode: POST spec → .pptx in < 1 s per slide at p95; runs as a container.

### P3 — Office.js add-in

Scope: task-pane add-in for PowerPoint (Windows, Mac, web) that embeds the editor; insert elements at the cursor; detect selection of an open-charts group and re-open it in the pane; apply decimal-symbol switch, replace fonts, clean-up, swap, scale; style loading per master.

Exit criteria:
1. Insert and re-edit any P1/P2 chart from the pane on PowerPoint for Windows, Mac and web; the inserted group is pixel-identical to the CLI output.
2. Edit → re-insert replaces the shapes in place (same z-order, same position, same name), preserving the user's manual moves of the whole group.
3. Add-in passes Microsoft's AppSource validation checklist (even if we only side-load).
4. Clean-up on a 200-slide deck finishes in < 10 s and removes every category listed in §2.14.
5. Documented limits: no sub-shape drag handles, no context menu on segments — these are explicitly deferred to P4.

### P4 — Native Windows direct manipulation

Scope: COM/VSTO add-in on Windows reusing the engine (via embedded Node or a compiled TS→.NET bridge): selection of chart features, drag handles for axis extent/tick spacing/column width/category gaps/label position/value lines/difference-arrow ends, context menu and mini toolbar per feature, lock/unlock edges, dynamic alignment, match size, duplicate-with-alignment, Group/Ungroup/Regroup overrides, COM automation API.

Exit criteria:
1. Every mouse interaction in the reference manual's chart chapters has an equivalent: a checklist of ≥ 60 interactions, each demonstrated in a recorded test.
2. Drag feedback renders at ≥ 30 fps on a 2-year-old laptop; relayout of a 12-series chart < 50 ms.
3. Manual label placements, locked edges and dynamic alignments persist through data edits, save/reload and the Office.js path (shared metadata format).
4. Group/Ungroup/Regroup/Undo never corrupt an element; fuzz test of 1,000 random command sequences.
5. The COM API offers `PresentationFromTemplate`-equivalent and `UpdateBatch`-equivalent calls with VBA and C# samples.

### P5 — Breadth: Gantt, agenda, tables, Excel links

Scope: Gantt engine + editor + handles; agenda deck model; tables with auto-fit cells and content types; text boxes, chevrons, rounded rectangles; Excel add-in with linked ranges, link frames, Data Links dialog, update/revert/auto; charts with data tables; slide-workbook replacement (deck data file); TCROUND/TCCAGR equivalents in Excel.

Exit criteria:
1. Gantt: rebuild 15 reference project timelines (quarter/month/week scales, milestones, brackets, shades, three label columns) indistinguishably; dates round-trip through the datasheet.
2. Agenda: insert / reorder / nest / split chapters on a 60-slide deck; every agenda slide updates; hidden-parent logic matches the reference.
3. Tables: 20 reference tables (striping, Harvey/checkbox cells, auto-fit rows) rebuilt; a table linked to a chart's data stays in sync.
4. Excel links: create, link, update, revert, auto-update, alternate source and bulk Data Links dialog work across a 50-element deck linked to three workbooks; link frames show correct colours.
5. Consistent rounding: totals of rounded addends equal rounded totals on the reference examples; the solver reports unsolvable cases instead of guessing.

## 5. Where open-charts deliberately differs and leapfrogs

| Reference approach | open-charts approach | Why it is better for us |
|---|---|---|
| Chart state lives in opaque binary data inside the .pptx; only the add-in can read it | **Spec-first.** The JSON spec is the source of truth, embedded in the .pptx (alt text / custom XML) and recoverable with `import`; shapes are a rendering | Diffable in git, reviewable in PRs, reproducible in CI, editable by scripts and by models |
| Direct manipulation is the primary (only) authoring path | **Code-first, UI-second.** CLI and HTTP first, editor second, add-ins third; every UI action is a spec edit | One engine, three surfaces, no drift; PowerPoint-free pipelines (Keynote, Google Slides export, PDF) |
| Style files are large XML schemas edited by consultants | **Style = small JSON/YAML** with the same concepts (fills, lines, markers, schemes, defaults) plus `extends` | Versionable per client; a style can be derived from an existing deck automatically |
| AI is a side pane bolted on; chart data is only visible to it via the datasheet | **AI-native.** The spec is the prompt: action titles, chart-type recommendation, anomaly flags and sanity checks run on structured data, not pixels | Higher accuracy, explainable, cheap; works headless in batch |
| Layout quality depends on the user noticing problems | **Layout that checks itself** — a quality gate per chart (no overlap, nothing out of frame, sums correct) with a machine-readable report | Safe for unattended generation of hundreds of slides |
| Per-slide Excel workbooks hidden inside the .pptx | **One data file per deck** (`deck.json` or a CSV folder) referenced from the spec; optional live link to Excel | Transparent, greppable, no hidden data leaking with copied slides |
| Tableau via browser extension; Morningstar/Statista via licence | **Open connectors**: CSV, Parquet, SQL, HTTP JSON, Google Sheets; user brings their own licensed sources | No vendor lock, no hidden data costs |
| Windows-only API, COM late binding | **TypeScript API + HTTP + CLI on every OS**; COM shim only in P4 | Usable from Python, Node, CI, serverless |
| Clean-up is a dialog | **Headless clean-up and sanitize** (`open-charts clean`, `open-charts sanitize`) runnable in CI before a deck leaves the firm | Governance by pipeline, not by discipline |
| Features gated by Essentials / Suite / Assist tiers | **Apache-2.0, everything included**; optional paid hosting of the HTTP service and model calls | Adoption inside firms with strict procurement; community contributions |
| Same-scale set by multi-selecting segments | **Declarative `scaleGroup`** at deck level, plus the gesture later | Works in batch; impossible to forget when data changes |
| Chart scanner = bundled OCR models (≈ 60 MB) | **Vision-model call** with a verification step against axis ticks | Smaller install, handles more chart styles, improves without shipping binaries |

Deliberate *non*-differences: we copy the reference's visual conventions exactly (label stagger rules, boxed labels, arrow geometry, waterfall `e` semantics, number-format-by-example, 30/360 CAGR) because they are the industry standard consultants are trained on, and "indistinguishable" is the P1 bar.

## 6. Explicitly out of scope (for now)

| Area | Reason |
|---|---|
| Slide/image asset library, stock-image and DAM connectors, browser extension for images | Not charting; PowerPoint and DAM vendors cover it. Revisit only if a contributor owns it. |
| Image element (fill-frame/crop modes) | Needs native handles; low value for the chart story. |
| Mekko Graphics import | Migration from a competitor we do not target. |
| Licensed data providers (Morningstar, Statista) | Commercial licensing; connectors are user-supplied. |
| Insert-symbol dialog, file-name text box, Excel table-style favourites, Mark-as-Final/Print overrides | Host already does it or value is L; native command overrides limited to Group/Ungroup/Regroup/Undo in P4. |
| Native **Mac** add-in | Office.js (P3) covers Mac; a native Mac core is a P6 question driven by demand. |

## 7. Inventory totals

| Area | Features | P1 | P2 | P3 | P4 | P5 | — |
|---|---|---|---|---|---|---|---|
| Charts by type | 24 | 11 | 11 | 0 | 0 | 2 | 0 |
| Decorations | 15 | 11 | 3 | 0 | 1 | 0 | 0 |
| Axes | 10 | 6 | 4 | 0 | 0 | 0 | 0 |
| Text / labels / formats | 15 | 7 | 1 | 3 | 2 | 1 | 1 |
| Areas / auto-layout | 8 | 1 | 1 | 2 | 4 | 0 | 0 |
| Datasheet | 9 | 4 | 2 | 0 | 0 | 1 | 2 |
| Excel / external links | 11 | 0 | 0 | 0 | 0 | 10 | 1 |
| Layout tools | 6 | 0 | 0 | 0 | 4 | 2 | 0 |
| Agenda | 6 | 0 | 0 | 0 | 0 | 6 | 0 |
| Gantt | 7 | 0 | 0 | 0 | 0 | 7 | 0 |
| Tables / Harvey / checkbox | 6 | 0 | 2 | 0 | 0 | 4 | 0 |
| Library / images | 7 | 0 | 1 | 0 | 0 | 0 | 6 |
| Styles / brand | 7 | 1 | 3 | 2 | 0 | 0 | 1 |
| Collaboration / quality | 5 | 1 | 1 | 2 | 0 | 0 | 1 |
| Automation / API | 7 | 1 | 4 | 0 | 1 | 0 | 1 |
| AI | 9 | 0 | 6 | 2 | 0 | 0 | 1 |
| **Total** | **152** | **43** | **39** | **11** | **12** | **33** | **14** |

Roughly: P1 delivers the 43 features that make a chart *correct*; P2 adds the 39 that make a *deck* from data; P3–P4 are about where the user is standing when they do it; P5 is the long tail the reference product took two decades to accrete. Features assigned to two phases are counted in the earlier one.
