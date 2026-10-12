# 02 — Engine architecture and quality

Scope: the layout engine (`src/core`), the renderers it feeds (`src/render`), and the quality system that proves the output is consulting-grade. Everything here is measured against one acceptance bar, defined in §5: a reviewer who knows the reference add-in cannot tell our slide from its slide.

Status today (commit `950b05c`): ~1,040 lines of TypeScript, 7 unit tests, three working chart families (stacked / 100% / clustered column, line). Waterfall and Mekko are stubs that only push `"not implemented yet"` into `scene.hidden` (`src/core/charts/waterfall.ts:4`, `src/core/charts/mekko.ts:4`) although the README lists both as supported. The code is small, readable and the right shape to grow; it is not yet an engine.

---

## 0. Honest critique of the current code

What is good: scene-first design (layout never touches OOXML), real font metrics, one scene feeding both renderers, a quality gate that runs on every build, largest-remainder rounding, Liang–Barsky for label/line hits, zero dependencies. Keep all of that.

### 0.1 Bugs

| # | Where | Problem |
|---|---|---|
| B1 | `package.json:12`, README "Tests" | `node --test test/` fails on Node 22.12 with `Cannot find module …/test` (directory args are not resolved). `test/core.test.ts` passes (7/7) when named explicitly. CI as documented is red. |
| B2 | `package.json:8` | `bin` points to `./bin/open-charts.js`, which does not exist. `npm i -g` yields a broken link. |
| B3 | `src/core/charts/column.ts:132` | CAGR uses `years = b - a`, i.e. category *index* distance. Quarterly or monthly categories (as in the stress-test slide in `examples/deck.json`) give a CAGR that is 4× / 12× too high. The reference tool reads dates from the category cells and uses a 30/360 day count for fractional years. |
| B4 | `src/core/charts/column.ts:93,104` | A series whose value in the last category is `0` or `null` is skipped (`if (!v) continue`) before `lastSegs.push`, so it gets **no right-hand series label at all**, silently. |
| B5 | `src/core/charts/column.ts:73-82` | Clustered: `null` is coerced to `0` by `raw()` (line 15), drawn as a 0.01 pt bar and labelled "0". Missing data should be an empty slot, not a zero. |
| B6 | `src/core/text.ts:43` | `getAdvanceWidth` applies GPOS/kern pair kerning by default. PowerPoint does **not** kern unless `<a:rPr kern="…">` is set (the writer never sets it, `src/render/pptx.ts:42`). Measured widths are therefore ~1–3 % narrower than drawn widths for strings like "AV", "Ta", "11". Either measure with `{kerning:false}` or emit `kern="1200"` — pick one and test both faces. |
| B7 | `src/core/text.ts:45` | Line height is a hard-coded `1.2 × size`. PowerPoint single spacing for Arial is `(ascender+descender+lineGap)/unitsPerEm = (1854+434+67)/2048 ≈ 1.149`. Multi-line category labels are planned 4 % taller than drawn; with other fonts the error is unbounded. Read `hhea`/`OS/2` from the loaded font. |
| B8 | `src/core/common.ts:51,59` + `src/render/pptx.ts:45` | `wrap()` splits on whitespace only (`text.ts:52`). A single word wider than the slot (or any CJK string) cannot wrap, the text box is emitted `wrap="none"` with `w = slot`, so the glyphs overflow the box. `checkScene` compares boxes, not glyph extents, so this overlap is **undetected**. |
| B9 | `src/core/common.ts:87-105` | Legend is one row, left-aligned, never wraps. Six series with long names walk out of the frame; the gate reports `out-of-frame` but nothing fixes it. |
| B10 | `src/core/charts/line.ts:79` | Lines are found by `sh.width === 2`. Any future 2 pt line (value line, axis) becomes a "data line" for collision purposes. Use node roles/ids, not stroke width. |
| B11 | `src/core/charts/line.ts:115` vs `line.ts:95-103` | Right-hand series names ignore value labels already placed to the *side* of the last point. Overlap is possible and is only *reported*. |
| B12 | `src/core/quality.ts:10-22` | Gate checks text×text and text×frame only. Text over bars, arrows crossing labels, marker over axis label, out-of-frame *shapes* (negative bars below `frame`), and the two bracket arrows' verticals passing through each other's bubbles are all invisible to it. |
| B13 | `test/core.test.ts:69` | `hidden:` issues are filtered out of the pass/fail. Hiding a label is a silent degradation path and never fails CI; `examples/line.json` is not covered at all (line 65). |
| B14 | `src/core/numbers.ts:37-39` | Locale is hard-coded `en-US`; unit is always a suffix with no space (`42m`, but never `$42`, `42 %`, `€ 42`, `(42)`); no magnitude scaling; `inferDecimals` (line 29-34) returns 1 for `0.123`. |
| B15 | `src/core/index.ts:45-53` | Validation does not catch `NaN`, `Infinity`, strings in `values`, duplicate series names, bad hex colours, unknown option keys, or `cagr.from === cagr.to`. Errors are not path-addressed, so the editor cannot point at the cell. |

### 0.2 Structural gaps

- **Fixed-point loops by hand.** `column.ts:48-53` and `line.ts:38-42` each iterate three times to settle slot width vs. right-label width vs. category height. This pattern will recur for every edge (axis title, legend wrap, data table, arrows) and needs a shared *space reservation* service with convergence detection (§1.4).
- **No outside-label fallback.** `column.ts:101-103` hides any segment label that does not fit. The reference behaviour (manual, "Automatic label placement") is: centred → offset left/right when neighbours clash → outside with a leader line → coloured backing rectangle when wider than the segment. Hiding is the last resort, not the first.
- **Decorations are special cases.** `cagr`/`difference` are two optional fields on `ChartOptions` (`types.ts:15-17`), both arrows are forced to be total-to-total and above the chart, and clustered charts get none (`column.ts:40,128`). Level arrows, series CAGR, value lines, 100 % indicators, connectors, category gaps all need a *decoration list* with a common anchor model.
- **Flat scene, no identity.** `Shape` has an optional `role` but no `id`, no owner (series/category), no z-layer, no group. The PPT writer therefore emits one flat group; the editor cannot select "all totals"; golden tests cannot diff by node.
- **Colours are baked hex.** `pptx.ts:18` always writes `srgbClr`. A deck can never be re-themed in PowerPoint; brand themes (§1.6) need `schemeClr` references.
- **Fixed frame.** `deck.ts:15` hard-codes one 16:9 frame per slide; no two-charts-per-slide, no 4:3, no explicit `frame` in the spec.
- **Heuristic thresholds without provenance.** `labelAll = nCat*nSer <= 24` (`line.ts:28`), narrow-band `< 0.4` (`line.ts:17`), `pad = 0.4·fs` (`common.ts:36`), `tol = 1` (`quality.ts:7`). Fine as defaults, but they must move into the theme file so a benchmark can tune them.

---

## 1. Engine architecture evolution

### 1.1 Spec schema with versioning

```jsonc
{ "$schema": "https://open-charts.dev/schema/chart/2.json", "version": 2,
  "type": "stacked", "orientation": "vertical",          // or "horizontal" → bar chart
  "categories": [{ "label": "Q1 FY24", "date": "2024-01-01" }, "Q2 FY24"],  // string or object
  "series": [{ "name": "Licences", "values": [120, null, {"kind":"subtotal"}], "role": "column" }],
  "labels":  { "content": "abs|pct|both", "format": "#,##0.0", "unit": {"text":"m","position":"suffix","space":false} },
  "axis":    { "value": { "show": true, "min": null, "max": null, "break": "auto", "gridlines": false, "title": "€ m" } },
  "decorations": [ { "kind": "cagr", "from": 0, "to": -1 },
                   { "kind": "difference", "mode": "level", "from": {"category":1,"series":0}, "to": {"category":3,"series":0} },
                   { "kind": "valueLine", "value": "mean", "label": "Ø {value}" } ],
  "theme": "default", "frame": { "x": 48, "y": 96, "w": 864, "h": 408 } }
```

- `version` is mandatory from v2; `migrate(spec)` upgrades v1 (today's `options.cagr` etc.) in place and is covered by fixture tests. We never break a stored JSON.
- JSON Schema is generated from the TypeScript types (one `schema/*.json` per version, checked in) and used by the editor, the CLI (`--strict` rejects unknown keys) and the docs.
- Validation returns `{ path: "series[1].values[3]", message }[]`, never throws on the first error.
- Decorations are a typed union; adding one touches the union, one plugin file and the schema — nothing in the chart layouts.

### 1.2 Scene graph

Replace `Shape[]` with nodes that carry identity and intent:

```ts
interface Node { id: string; kind: 'rect'|'line'|'path'|'text'|'group'; role: Role;
  owner?: { series?: number; category?: number; decoration?: string };
  layer: 'background'|'plot'|'connector'|'label'|'decoration'; z: number;
  bounds(): Box; children?: Node[] }
```

- `path` (pie wedges, area polygons, elbow connectors, axis-break glyph), `text` with **runs** (`{text, bold, size, color}[]`) so "**118** m" and footnote superscripts are one label, and a `measured: {w, h}` cache written by the measurer so later passes never re-measure.
- Groups nest: slide → chart → {plot, axes, labels, decorations, legend}. The PPT writer emits nested `grpSp` with names ("Chart 1 / Totals"), which is what makes the output feel native when a user ungroups it.
- The scene is plain JSON. Golden tests diff scenes first (exact, fast, explains *what* moved) and pixels second.

### 1.3 Plugin per chart type

```ts
interface ChartPlugin<O> { type: ChartType; optionsSchema: JSONSchema; defaults(spec): O;
  model(spec, o): Model;                 // pure data: stacks, totals, domains, widths, subtotals
  reserve(model, ledger, ctx): void;     // claims margins: axis width, legend rows, arrow rows…
  plot(model, ctx): Node[];              // marks only
  labels(model, ctx): LabelRequest[];    // what wants to be labelled, with candidates (§3)
  decorate(model, ctx): Node[] }         // anchors for decorations, resolved by shared code
```

The pipeline is shared and fixed; each stage has a declared input and output so a plugin cannot reach around it:

1. `migrate` — any spec version → current. Pure, fixture-tested.
2. `validate` — schema + semantic checks, returns path-addressed errors; the CLI exits here, the editor highlights cells.
3. `model` — plugin turns the spec into numbers only: stacks, nets, running sums, domains, widths, nice ticks, and the *final formatted strings* (via the number service), so every later stage measures the text that will actually be drawn.
4. `reserve` — plugin and shared decorations file claims on the ledger; the ledger iterates to a fixed point and yields the plot rectangle.
5. `plot` — marks only (rects, paths, lines, markers), each with owner and layer.
6. `decorate` — shared code resolves decoration anchors against model + plot (arrow endpoints, value-line y, connector corners) and emits their marks; their labels become `LabelRequest`s like any other.
7. `label` — the solver (§3) places every request against the obstacle index.
8. `verify` — the quality gate (§4.2) runs on the scene; failures feed the fallback.
9. `fallback` — shrink one role by one step (never below `theme.fonts.min`) and re-run 7–8; then hide by priority with a reason; then report.

A chart file contains only what is genuinely specific to that chart (today `column.ts` mixes all of these stages in 139 lines). Stages 1–2 and 6–9 are shared; a new chart type is `model + reserve + plot + labels`, typically 150–300 lines.

### 1.4 Shared layout services

1. **Space reservation ledger.** `ledger.claim('top', 'arrows', h)`, `ledger.claim('right', 'seriesLabels', w)`. Claims may depend on the plot size (category height depends on slot width); the ledger iterates to a fixed point (≤ 6 rounds, Δ < 0.1 pt) and emits a diagnostic when it oscillates. Replaces both hand-rolled loops.
2. **Label placement solver** — §3.
3. **Numbers & locale.** Excel-style format strings (`#,##0.0;(#,##0.0)`, `0%`, `+0.0%`), locale table (decimal/grouping chars, true minus U+2212 vs hyphen, percent spacing per locale — `42 %` in de-DE), magnitude auto-scaling (`k/m/bn` with a shared magnitude per chart), and **consistent rounding**: displayed segments must sum to the displayed total (largest remainder across each column *and* across the total row), the TCROUND idea applied to labels.
4. **Text measurement.** Font stack per script (Latin → Arial/Liberation; CJK → the deck's `ea` font, measured with Noto Sans CJK metrics on CI; symbols → Segoe UI Symbol/Arial Unicode), real line height from `hhea`, kerning parity with the writer (B6), a UAX #14 subset for break opportunities (space, hyphen, slash, en/em dash, CJK between any two ideographs, no break before closing punctuation), ellipsis/shrink policies, bold measured from the bold face, LRU cache keyed by `(font, size, bold, text)`.
5. **Style / theme system.** One JSON file per brand:
   `fonts {body, size, min}`, `palette {categorical[], sequential[], positive, negative, neutral, emphasis}`, `lines {axis, connector, arrow}`, `gaps {column: 0.5, cluster: 0.1, categoryGap: 1}`, `labels {contrastThreshold, backingBox, leaderStyle}`, `arrows {bubble, head}`, `locale`. The writer maps `palette.categorical[i]` to `accent1..6` and emits `schemeClr` so PowerPoint's Design › Variants recolours the deck. Themes are validated by a schema and previewed by the benchmark (§4) — every reference slide renders under every theme.

### 1.5 Renderers

PPT writer gains nested groups, `kern`, theme colours, `custGeom` paths, number-format-free text (we pre-format), and `descr` alt text per chart. SVG preview stays a 1:1 mirror and gains `path` and runs. A third renderer, **native `c:chart` export**, is explicitly *not* planned: it would lose per-label control, which is the product.

---

## 2. Chart catalogue

Each entry: layout notes, then edge cases the benchmark must cover. Orientation (`vertical`/`horizontal`) is a transform of the plot coordinates plus swapped label rules (horizontal bars: value labels end-aligned, category labels right-aligned on the left, growth arrows on the right).

**Stacked / clustered / 100 % column & bar.** Column width from gap ratio (gap = `g·colW`, default 0.5); cluster inner gap 0.1·barW; 100 % axis fixed 0–100 with the *100 % indicator* on the left by default; label content abs/pct/both; segment order sheet/reverse/ascending/descending; "Other" series collapses segments below a threshold; stacked-clustered via gap 0 + category gaps. *Edge cases:* negatives below baseline with a separate negative stack, totals = net; mixed-sign column in 100 % (reference: positive-only 100 %, we flag); zero and null segments (no shape, no label, still counted for series labels — B4); one category; 40 categories (labels rotate 90° or thin to every nth); single series (no legend); 12 series (legend wraps to rows, max 3, then "legend overflow" diagnostic); all-equal values; values spanning 1e-3..1e6 (magnitude scaling kicks in).

**Line / area.** Points at category centres, axis crossing at or between categories (theme default: between); markers optional; broken lines at nulls, optional interpolation (dashed); smoothed lines via Catmull-Rom → cubic Bézier path; stacked and 100 % area (percent axis converts line → stacked area); secondary value axis on the right with independent nice ticks and *matched zero* when both include zero; date axis when categories parse as dates (uneven spacing). *Edge cases:* narrow band far from zero → axis break glyph instead of silent non-zero baseline (today `line.ts:17`); two series crossing repeatedly (labels alternate sides by local slope); first/last point labels fighting the axis; 50-point lines (label only extrema + last).

**Combo.** Series carry `role: column|line|area`; column series lay out as stacked/clustered, line series overlay on the same or secondary axis; label solver treats line points as obstacles for column labels and vice versa. *Edge:* line on secondary axis with opposite sign.

**Waterfall (vertical & horizontal, multi-series).** Model: running sum; a value `{kind:"subtotal"}` ("e") is stretched to the baseline and computed; an `e` in the *first* column is solved backwards from the last; multiple segments per column allowed, at most one `e`; mixed signs in one column: connector gap = signed sum, each segment drawn with true extent. Connectors: one between consecutive columns by default, removable (`connect: false` starts a new summation), skew connectors flagged. Labels show extent (positive) by default, optional signed. Level-difference arrows may show "% of 100 %=" (category 100 % value). Build-up vs build-down is data only. Horizontal = same model, swapped axes. *Edge cases:* a step crossing zero (split into two rects, label the net); tiny steps (label outside with leader); last `e` equals 0; negative start; 25 steps (connector weight thins, labels rotate).

**Mekko (percent & unit).** Percent: widths ∝ column totals, heights 100 %, no gaps, baseline is a value axis with abs or % tick labels, optional ridge outline, "Other" series common. Unit ("submarine"): widths from an explicit `xExtent` row, heights absolute. Column labels are centred under variable widths with the solver handling thin columns (rotate → leader → hide). *Edge cases:* a column with total 0 (zero width, label kept in the gap row); 1 % column next to 60 % column; both-percent-and-absolute labels in a 20 pt wide segment.

**Pie / doughnut.** Treated as a 100 % stack of one category bent around 360°. Slice order sheet/ascending/descending, start angle snapped to 0/90/180/270, explode per slice, hole size (0 = pie). Inside labels pushed toward the rim; when a label is wider than its slice → backing box → outside label with elbow leader, outside labels spread with the 1-D `spread()` on each half, never crossing the vertical centre. *Edge cases:* 1 % slices in a run; doughnut centre title; 2 slices at 50/50 (labels dead-centre).

**Scatter / bubble.** Two value axes always shown, groups → marker scheme, bubble area ∝ size (option: diameter), smaller bubble drawn in front, bubble-size legend, optional trendline (OLS) and partition lines. Labels: the full 2-D solver (§3) with leader lines — this is the chart that justifies it. *Edge cases:* 200 points; coincident points; log axes.

**Butterfly.** Two horizontal bar charts back-to-back sharing a category column in the middle, *same scale* enforced, one set of category labels. Implemented as a composite of two plugins plus a shared ledger — the first test of multi-chart layout.

**Gantt (later).** Calendar scales (year/quarter/month/week/day) with separators, activity rows, bars/milestones/brackets, date format codes, locale. Different enough (date axis, table-like rows) to be its own milestone after the quantitative set.

**Decorations (shared, anchored to model points, resolved after plotting):**
totals and grand totals (bold, above/beside, net value); CAGR arrows (dates from categories, 30/360 fractional years, one row per arrow, stacking rows top-down with bubbles never crossing another arrow's vertical); series CAGR labels (per-series rate in the right-hand series label); difference arrows — *total* (column tops) and *level* (segment/point to segment/point, three modes: →relative, ↔absolute, ←relative), bubble moved beside the arrow when the arrow is shorter than the bubble; value lines (mean / weighted mean / fixed, label left or right, snapping); axis breaks (zig-zag glyph on axis and on every bar that crosses it, scale split into two linear pieces, labels keep true values); category gaps (extra slot width between selected categories; arrows prefer to land in gaps); series labels vs legend (right-hand labels aligned to last segments with leaders, legend 1..3 rows or vertical, order follows visual stacking order top-down); data tables (a row grid under the plot aligned to category slots, values formatted by the same number service, grouped with the chart).

---

## 3. Label placement: from heuristics to a scored solver

Today: column segments are centred-or-hidden (`column.ts:101`), line labels try six fixed spots in order (`line.ts:97-100`), series names use 1-D `spread()` (`place.ts:5`). Each is local, greedy, and unaware of the others.

**Target.** One solver receives every `LabelRequest {id, text runs, anchor, candidates[], priority}` of a chart after the marks exist, plus obstacle sets (marks, axes, frame, already-fixed text), and returns a position or `hidden(reason)` per label.

- **Candidates** are generated by label type from the theme, in preference order:

| Label type | Candidates (rank 0 → n) |
|---|---|
| segment (stack, waterfall, Mekko) | centre · centre with backing box · offset left / right (neighbour clash) · outside right + leader · outside above the column (top segment only) · rotated 90° (thin Mekko columns) |
| bar end (clustered, simple) | beyond the end · inside the end (contrast colour) · beside the bar (horizontal only) |
| point (line, area) | above · below · right · left · above-right / above-left / below-right / below-left, each at 1·pad and 2·pad |
| total | above column (vertical) / right of bar (horizontal) · inside top segment when arrows need the space |
| pie inside / outside | rim at 0.85·r, 0.7·r, 0.55·r · outside at 1.15·r with elbow leader, spread per half |
| scatter / bubble | 8 compass points × 2 distances · nearest free cell of a surrounding ring with a leader |
| series name (right) | 1-D spread aligned to last segment / point, leaders when displaced · legend fallback |
| category | one line · wrapped (≤ 3 lines) · rotated 90° · every nth shown |

- **Cost.** For label *i* at candidate *c*:
  `cost(i,c) = Σ_text ∞·overlap + Σ_mark w_mark·area(overlap)/area(label) + w_own·overlapsOwnMark + ∞·outsideFrame + w_rank·rank(c) + w_dist·|c − anchor| + w_leader·leaderLen + w_cross·(leader crosses mark) + w_side·(side ≠ sibling side) + w_axis·(crosses axis)`
  with all `w_*` in the theme (`labels.weights`) and the hard terms immune to tuning. Total cost is the sum over labels plus `w_shrink·shrinkSteps + w_hide·Σ priority(hidden)`.
- **Optimisation.** Deterministic greedy seeding (hardest labels first = fewest free candidates), then iterative improvement: repeatedly move the single label whose move lowers total cost most, allowing 2-label swaps; stop at a fixed-point or `maxIterations`. Then the fallback ladder: shrink to `theme.fonts.min` (one size step at a time, *all labels of that role together* so a chart never mixes sizes), then hide lowest-priority labels one by one. Seeded, so identical input ⇒ identical bytes.
- **Obstacle index:** uniform grid (cell ≈ 2·lh) so a candidate test is O(1) amortised; 2-D only where needed (scatter, pie), 1-D fast path (`spread`, rewritten as pool-adjacent-violators for minimum total displacement) for series names and pie outside labels.
- **Performance targets** (Apple M1, Node 22): ≤ 60 labels → < 5 ms; 200 labels → < 50 ms p95; a 50-slide deck end-to-end < 2 s; editor live preview keystroke-to-SVG < 30 ms for charts ≤ 60 labels. Budget enforced by a benchmark test that fails on 2× regression.
- **Explainability.** Every placed label records `{candidate, cost, rejectedBecause[]}`; the editor shows it on hover and the benchmark report aggregates it ("38 % of hidden labels: segment < 0.9·lh").

---

## 4. Quality benchmark

### 4.1 Reference slides (`bench/slides/*.json`, ~20)

| # | Slide | What it stresses |
|---|---|---|
| 1 | Stacked revenue, 4 series, CAGR (the README chart) | baseline; arrow row; right-hand series labels |
| 2 | Stacked, 8 quarters × 5 series, thin segments, difference arrow | segment label ladder; quarterly dates (B3) |
| 3 | Stacked with negative segments and net totals | negative stack; totals placement |
| 4 | 100 % mix with rounding traps (33.33 ×3; 0.4 % slice) | largest remainder; tiny segments; 100 % indicator |
| 5 | Clustered with value axis and negatives | bar-end labels both sides of zero; axis width |
| 6 | Horizontal bar, 12 categories, 40-character names | orientation; wrapping; category column width |
| 7 | Line NRR 95–120 % | narrow band → axis break |
| 8 | Line 12 months × 3 series with nulls | broken lines; dense point labels; legend |
| 9 | Line single series crossing zero | dashed zero line; labels alternate sides |
| 10 | Combo: columns + margin line on secondary axis | two axes, matched zero; line obstacles for column labels |
| 11 | EBITDA waterfall with two subtotals | running sum; `subtotal`; connectors |
| 12 | Horizontal waterfall build-down | orientation of waterfall; backwards `e` |
| 13 | Waterfall with step crossing zero and 0.3 steps | split step; tiny-step labels with leaders |
| 14 | Mekko % market map 6 × 5 | variable widths; ridge; Other series |
| 15 | Mekko unit (submarine) | independent width/height; baseline value axis |
| 16 | Pie, 7 slices incl. three < 3 % | inside → outside ladder; outside spread |
| 17 | Doughnut with centre title | hole; title obstacle |
| 18 | Scatter, 40 points, 2 groups, trendline | 2-D solver with leaders; OLS |
| 19 | Bubble with size legend | area scaling; z-order; legend |
| 20 | Butterfly age pyramid | composite; shared scale; shared category column |
| 21 | Two charts on one slide | multi-frame; consistent magnitude |
| 22 | CJK categories, 10-series legend, de-DE numbers | font stack; legend wrap; locale |

Each slide ships with the expected `scene.json` and, once approved, golden PNGs from both renderers. A slide is added to this list whenever a user-reported layout bug is fixed, with the bug's data.

### 4.2 Automated metrics (run on every slide, every theme, in `test/bench.test.ts`)

| Metric | Definition | Pass |
|---|---|---|
| overlaps | pairs of text nodes, or text×mark of another owner, intersecting by > 0.25 pt² after glyph-extent (not box) measurement | 0 |
| out-of-frame | any node bounds outside `frame` by > 0.5 pt (shapes included) | 0 |
| hidden labels | count and reasons | 0 on slides 1–21 unless the fixture declares `allowHidden` |
| rounding consistency | Σ displayed segments = displayed total per column; Σ displayed shares = 100.0 per 100 % column; CAGR label = recomputed from displayed endpoints ±0.05 pp | all |
| alignment | labels of one role share a baseline / centre line within 0.25 pt where the layout intends it (totals row, axis ticks, legend row, category row) | all |
| whitespace balance | plot area ≥ 55 % of frame height (column) / ≥ 60 % (line); left/right margins beyond reserved claims differ by < 4 pt | all |
| min font size | no run below `theme.fonts.min` (default 8 pt); no chart mixes sizes within one role | all |
| determinism | layout twice ⇒ identical scene JSON; scene hash recorded | all |
| performance | §3 budgets | all |
| OOXML validity | Open XML SDK validator (Windows job) = 0 errors; `python-pptx` loads (Linux job); LibreOffice converts without warnings | all |

### 4.3 Golden images

- **Linux CI, every PR:** `soffice --headless --convert-to pdf` → `pdftoppm -r 144` → compare with stored PNGs. Metric: SSIM ≥ 0.995 *and* no 8×8 block with mean abs diff > 24/255 (catches a moved label that SSIM averages away). Fonts pinned (Liberation Sans, Noto Sans CJK) in the container image.
- **Windows CI, nightly + on `release/*`:** self-hosted runner with PowerPoint 365 (hosted runners have no Office). PowerShell COM: `Presentation.Slides(i).Export(png, "PNG", 1920, 1080)`. Same diff. This is the one that matters — LibreOffice differs from PowerPoint in text box insets and line-end geometry.
- Goldens update only via `npm run bench:accept` which writes a markdown diff report (before/after thumbnails) into the PR; a human approves it.

### 4.4 Human review rubric

Two reviewers per release, blind, each slide scored 1–5 on: (a) label legibility and placement, (b) proportion and whitespace, (c) typographic consistency, (d) annotation clarity (arrows, totals, lines), (e) "would I put this in front of a client unchanged". Any 1–2 is a blocking defect with an issue link.

### 4.5 Definition of "indistinguishable" (release gate)

1. §4.2 metrics all pass for all reference slides under all shipped themes.
2. PowerPoint goldens pass; LibreOffice goldens pass or carry an annotated, approved difference.
3. Blind A/B panel: ≥ 6 people who use the reference add-in weekly see 20 pairs (ours vs. a slide rebuilt by hand in the add-in from the same JSON) and guess which is ours. Accept when the pooled hit rate is ≤ 60 % (chance = 50 %) **and** no slide is identified by ≥ 5/6 panellists.
4. Rubric: mean ≥ 4.3, no score below 3.
5. Editing test: a non-engineer opens the .pptx in PowerPoint, changes one value label and one colour, regroups, saves — no repair prompt, no layout damage.

---

## 5. Backlog (one engineer, ordered; days are focused days)

| # | Item | Days | Done when |
|---|---|---|---|
| 1 | Fix B1, B2, B13: test glob, bin entry, line.json in tests, `hidden` fails CI unless allowed | 0.5 | `npm test` green from a clean clone |
| 2 | Text parity: kerning decision (B6), line height from font metrics (B7), bold face check, measurement cache | 1.5 | 30 strings measured vs. PowerPoint screenshots within 0.5 pt |
| 3 | Numbers & locale service: format strings, locale table, units prefix/suffix/space, magnitude, parentheses; consistent rounding across segments+total | 3 | B14 closed; rounding metric green on slide 4 |
| 4 | Spec v2 + migration + JSON Schema + path-addressed validation (B15); decorations as a list | 3 | v1 fixtures migrate byte-identically; schema published |
| 5 | Scene graph with ids/owners/layers/groups, runs, `path`; renderers updated (nested groups, kern, `custGeom`) | 3 | scene JSON goldens for existing slides |
| 6 | Space reservation ledger; port column + line to it (removes both 3-pass loops) | 2 | identical output on existing fixtures |
| 7 | Plugin interface + pipeline; split `column.ts` into model/reserve/plot/labels | 2 | all charts go through one pipeline |
| 8 | Label solver v1 (candidates, penalties, greedy + improvement, fallback ladder, obstacle grid, explain data); replace line/column label code; fix B4, B5, B8, B10, B11 | 6 | slides 1–9 pass §4.2; perf budget test |
| 9 | Quality gate v2: glyph-extent overlap, shape out-of-frame, text×mark, alignment, whitespace, min-font, determinism (B12) | 2 | metrics table implemented |
| 10 | Bench harness: 20 slide fixtures, LibreOffice golden pipeline in GitHub Actions, accept script with diff report | 3 | PR check runs in < 4 min |
| 11 | Theme system: JSON themes, schema, `schemeClr` mapping, thresholds moved out of code; 3 themes (default, mono, dark) | 2.5 | all slides × 3 themes in bench |
| 12 | Horizontal orientation for column family (bar charts) + long-label rules + legend wrapping (B9) | 2.5 | slide 6 |
| 13 | Decorations: date-aware CAGR (B3), level/total difference arrows with 3 modes, series CAGR labels, value lines, 100 % indicator, category gaps, arrow-row packing | 4 | slides 1–3, 5 with all decorations |
| 14 | Waterfall: model with `subtotal`, backward solve, multi-segment columns, connectors, sign labels, horizontal | 5 | slides 11–13 |
| 15 | Mekko % and unit, ridge, Other series, segment ordering (shared with stacks) | 4 | slides 14–15 |
| 16 | Area / stacked area / 100 % area, secondary axis, date axis, smoothing, interpolation; combo charts | 4 | slides 7, 8, 10 |
| 17 | Axis breaks (scale split, glyphs on axis and bars) | 2 | slide 7 switches from silent non-zero baseline to a break |
| 18 | Pie / doughnut with inside→outside label ladder | 3.5 | slides 16–17 |
| 19 | Scatter / bubble: two axes, groups, size legend, trendline, 2-D solver with leaders | 5 | slides 18–19 |
| 20 | Composite layouts: multiple charts per slide, shared scales, butterfly | 3 | slides 20–21 |
| 21 | Data tables under charts | 2 | fixture with table |
| 22 | Text: UAX #14 breaks, CJK font stack + metrics on CI, locale de-DE/ja-JP fixtures | 2.5 | slide 22 |
| 23 | Windows PowerPoint golden runner (self-hosted), OOXML validator job, edit-roundtrip test | 3 | nightly green |
| 24 | Blind A/B panel tooling (pair sheets, scoring form, report) and first panel run | 1.5 | §4.5 numbers published |
| 25 | Gantt (calendar scales, rows, items, date formats) — after everything above | 8 | separate milestone |

Total without Gantt: **≈ 77 days**; with Gantt ≈ 85. Items 1–11 (~29 days) are the engine; nothing after them should require touching shared code again. Items 1–3 are worth doing this week regardless of the rest of the plan: they fix shipped bugs and the test command.

---

## 6. Risks and open questions

- **PowerPoint parity is only provable on Windows.** Until item 23 exists, every "matches PowerPoint" claim rests on LibreOffice plus manual screenshots. Budget the self-hosted runner early (a licensed VM is cheaper than a week of guessing text-box insets).
- **Kerning / line-height decisions (B6, B7) change every existing golden.** Do them before goldens are frozen (item 2 precedes item 10 for that reason).
- **Solver determinism vs. quality.** Greedy + local improvement can plateau on dense scatter plots. Keep a seeded annealing fallback behind a flag; never let timing decide output.
- **Theme colours via `schemeClr`** break the SVG preview's 1:1 promise unless the preview resolves the same theme. Resolve at scene level (store both `themeRef` and resolved hex).
- **Sourcing.** We take inspiration from the reference tool's manual, behaviour and internals, its output as a benchmark pair in §4.5, and other information available on the internet — whatever is the most efficient and effective route to parity. Record the source of every rule in `docs/behaviour-spec.md` (still to be written — see `tasks/todo.md`).
- **Scope creep from the editor.** The editor should consume `scene.json` and the solver's explain data; it must never grow its own layout logic.
