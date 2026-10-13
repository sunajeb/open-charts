# open-charts behaviour spec

Every rule the engine implements, one row each: **what it does → where (file: function) → why
(source) → which tests hold it**. Rules were found by reading `src/core`; when a rule had no test,
one was added to `test/behaviour.test.ts` (its test names start with the rule id).

**Sources.** *Manual* = the reference product's user manual (local copy under
standard. *House rule* = an open-charts decision where the manual is silent; the reason is given
(often a bug id `Bn` from [plan/02](plan/02-engine-and-quality.md) §1). Paths are relative to
`src/core/`; tests to `test/`.

**Count: 214 rules** in 27 areas. Change a rule → change its row and its test in the same commit.

- [1 Spec and validation](#1-spec-and-validation) · [2 Hidden rows](#2-hidden-rows)
- [3 Numbers and locales](#3-numbers-and-locales) · [4 Rounding](#4-rounding) · [5 Dates and 30/360](#5-dates-and-30360)
- [6 Value axes](#6-value-axes) · [7 Category labels](#7-category-labels) · [8 Series labels and legends](#8-series-labels-and-legends)
- [9 Label placement and the placer](#9-label-placement-and-the-placer) · [10 Totals, CAGR and difference arrows](#10-totals-cagr-and-difference-arrows) · [10b Decorations](#10b-decorations)
- [11 Stacked, clustered and 100% columns and bars](#11-stacked-clustered-and-100-columns-and-bars) · [12 Waterfall](#12-waterfall)
- [13 Mekko](#13-mekko) · [14 Mekko XY](#14-mekko-xy) · [15 Line](#15-line) · [16 Area](#16-area) · [17 Combination](#17-combination)
- [18 Pie and doughnut](#18-pie-and-doughnut) · [19 Scatter and bubble](#19-scatter-and-bubble) · [20 Harvey balls](#20-harvey-balls)
- [21 Table](#21-table) · [22 Gantt](#22-gantt)
- [23 Text, contrast and minimum sizes](#23-text-contrast-and-minimum-sizes) · [24 Resizing and collision resolution](#24-resizing-and-collision-resolution) · [25 Quality gate](#25-quality-gate)
- [26 Theme colours, style and same scale](#26-theme-colours-style-and-same-scale)

## 1 Spec and validation

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| SPEC-1 | Unversioned input is v0; `migrate` stamps `specVersion: 1` and removes options the type does not read, with one warning each; newer versions are rejected, not guessed | `spec.ts: migrate, migrateDeck, validateSpec` | House rule (MASTER-PLAN §4.2 "spec version field") | engine › "spec v1: migrate stamps v0…", "options: migrate() strips options…" |
| SPEC-2 | Every error is path-addressed (JSON Pointer) and all errors are reported at once; `layoutChart` throws one `InvalidSpecError` listing them | `spec.ts: validateSpec, InvalidSpecError, prepareSpec` | House rule (B15) | engine › "B15: validation is path-addressed…", "B15: layoutChart throws one error…" |
| SPEC-3 | Unknown keys anywhere (chart, series, options, sub-objects, deck, style) are errors | `spec.ts: Errors.unknown` | House rule (B15) | engine › "B15: validation…"; engine › "JSON Schema files stay in sync…" |
| SPEC-4 | Values must be finite numbers or `null`; `"e"` only in a waterfall; `[x,y]`/`[x,y,size]` (size ≥ 0) in scatter/bubble; dates in Gantt; text in tables; Mekko XY widths ≥ 0 | `spec.ts: validateSpec, checkPoint` | House rule (B15); DATA-FORMAT.md | engine › "B15…"; charts-b › "gantt: validation of dates, bars and shades" |
| SPEC-5 | Each series has as many values as there are categories; series names are unique; colours are 6-digit hex or theme slots (THM-1) | `spec.ts: validateSpec, checkColor` | House rule (B15) | engine › "B15: validation…" |
| SPEC-6 | An option on a chart type that does not read it is an error; `OPTIONS[k].types` is the one applicability table (validator, schema, editor, docs) | `spec.ts: OPTIONS, appliesTo, validateOptions` | House rule | engine › "options: every option on a type that does not read it…"; docs › "options table is the one generated…" |
| SPEC-7 | Numeric limits (`decimals` 0–10 whole, `gap` 0–10) are defined once | `spec.ts: LIMITS` | House rule | engine › "options: limits are defined once…" |
| SPEC-8 | `cagr`/`difference`: indices in range, `from` < `to`, `years` > 0 | `spec.ts: validateOptions` | Manual en-chartdecorations § CAGR arrows (start and end categories) | engine › "B15…" |
| SPEC-9 | Gantt needs Start and End series; each bar has both dates, end after start; shades/brackets `to` after `from` | `spec.ts: checkTypeSeries, validateOptions` | Manual en-gantt § Datasheets | charts-b › "gantt: validation of dates, bars and shades" |
| SPEC-10 | Mekko XY needs a width series and at least one height series | `spec.ts: checkTypeSeries` | Manual en-mekko (X extent row) | charts-b › "mekkoxy: width from the first series…" |
| SPEC-11 | `optionDefaults(type)` is what the engine draws when an option is left out; data-dependent defaults are described by `OPTIONS[k].auto` | `spec.ts: optionDefaults`; `common.ts: makeCtx` | House rule (a control always shows what is drawn) | engine › "options: optionDefaults() is what the engine draws…" |
| SPEC-12 | Decks validate their own keys, style (font, size 4–72, colours, palette) and each slide's chart with paths from the deck root | `spec.ts: validateDeck` | House rule | engine › "validateDeck addresses errors from the deck root" |

## 2 Hidden rows

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| HID-1 | `Series.hidden: true` leaves the row out of the layout; its data stays in the spec and is still validated | `visible.ts: visibleSeries`; `index.ts: layoutChart` | Manual en-axes § Adjust the scale of a value axis ("hide the respective datasheet row"); en-element-datasheets § Edit datasheet layout | hidden › "an extra hidden row changes nothing in any chart type"; e2e › "optional rows…" |
| HID-2 | Hidden rows count for nothing: totals, shares, value range, decimals, legend and alt text see visible rows only | `visible.ts: visibleSeries`; `deck.ts: altText` | Manual en-axes § Adjust the scale of a value axis | hidden › "a hidden middle row leaves the chart…", "table and Harvey rows drop out; alt text…"; e2e › "optional rows…" (totals 66m) |
| HID-3 | Visible rows keep the palette colour of their datasheet position (the palette is re-ordered per type offset: waterfall +1, Mekko XY −1) | `visible.ts: PALETTE_OFFSET, visibleSeries` | House rule (showing a row again changes nothing else) | hidden › "a hidden middle row…"; e2e › "optional rows…" |
| HID-4 | Roles keep the series' index in the spec (`bar:s2:c0` stays `s2` when `s1` hides) | `visible.ts: restoreRoles` | House rule (tags and overrides keep naming the same row) | hidden › "a hidden middle row…", "pie and doughnut show the first visible series" |
| HID-5 | At least one series stays visible; Gantt Start/End and the Mekko XY width row cannot hide; Mekko XY keeps one visible height row | `spec.ts: checkTypeSeries`; editor `cannotHide` | House rule (rows read by position) | hidden › "validated…"; e2e › "optional rows…" |
| HID-6 | Waterfall categories marked `"e"` only in a hidden row stay totals | `visible.ts: visibleSeries` | Manual en-waterfall § Create waterfall charts (e = computed total) | hidden › "waterfall categories marked "e"…" |
| HID-7 | Pie and doughnut draw the first visible series | `visible.ts`; `charts/pie.ts: layoutPie` | Manual en-pie § Introduction (one series) | hidden › "pie and doughnut show the first visible series" |
| HID-8 | Options that name series by index (`secondaryAxis`, level-arrow ends) follow the visible rows; a level arrow at a hidden row is left out and reported | `visible.ts: visibleSeries`; `index.ts: layoutChart` | House rule | hidden › "level arrows and the secondary axis follow the visible rows" |

## 3 Numbers and locales

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| NUM-1 | Decimals are inferred: the fewest (0–2) that show every value exactly, else two significant digits of the smallest non-zero value (max 3) | `numbers.ts: inferDecimals` | House rule (B14); manual en-textlabels § Number format | numbers › "B14: inferDecimals keeps two significant digits…"; core › "number formatting" |
| NUM-2 | Explicit `decimals` win; `format` overrides `decimals` | `common.ts: makeCtx (fmt)`; `numbers.ts: formatValue` | Manual en-textlabels § Number format | engine › "B14: locale, unit prefix, magnitude and format reach the labels and the axis" |
| NUM-3 | Locales en-US, en-GB, en-IN (3-then-2 grouping), de-DE, fr-FR (narrow no-break space grouping): decimal mark, grouping, true minus U+2212, percent spacing (NBSP in de, NNBSP in fr) | `numbers.ts: LOCALES, groupDigits, digits` | Convention: Unicode CLDR number symbols | numbers › "B14: locales" |
| NUM-4 | Locale lookup: exact id, then language ("de" → de-DE), then en-US | `numbers.ts: localeOf` | House rule | numbers › "B14: locales"; behaviour › "NUM-4" |
| NUM-5 | No sign on a value that rounds to zero (no "−0"); "+" only for signed positives (growth, deltas) | `numbers.ts: formatValue, applyFormat` | Convention (Excel shows 0.0, not −0.0) | numbers › "B14: Excel format strings"; behaviour › "NUM-5" |
| NUM-6 | Units: suffix (default) or prefix, optional NBSP; `%` takes the locale's spacing; sign goes before a prefix unit (−$42) | `numbers.ts: formatValue, normalizeUnit` | Manual en-textlabels § Number format | numbers › "B14: units as prefix or suffix…" |
| NUM-7 | Magnitude is one per chart: `auto` picks bn/m/k from the largest absolute value; suffix sits right after the number, before the unit | `numbers.ts: pickScale, formatValue`; `common.ts: makeCtx` | Manual en-textlabels § Magnitude | numbers › "B14: magnitude scaling is shared per chart" |
| NUM-8 | Excel format codes: ≤ 3 sections (pos;neg;zero), `0 # ?`, grouping comma, trailing comma ÷ 1000, `%` ×100, quoted text, `\x`, `_x`; `[colour]` and `*` ignored; a missing negative section adds the minus | `numbers.ts: applyFormat, splitSections` | Convention: Microsoft Excel number format codes | numbers › "B14: Excel format strings" |
| NUM-9 | Axis tick labels use the chart's unit, magnitude and locale, with decimals inferred from the ticks | `common.ts: axisFormatter` | Manual en-axes § Value axes | engine › "B14: locale, unit prefix, magnitude and format reach the labels and the axis" |

## 4 Rounding

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| RND-1 | Largest-remainder rounding: rounded parts add up to the rounded total (also with negatives) | `numbers.ts: roundToTotal` | Convention: largest remainder (Hamilton) method; manual en-excel-data-rounding § Introduction | core › "largest-remainder rounding…"; numbers › "B14: consistent rounding works with negatives" |
| RND-2 | 100% labels (columns, Mekko, area100, pie) are rounded together so each column/pie reads exactly 100% | `charts/column.ts: model`; `charts/mekko.ts`; `charts/area.ts`; `charts/pie.ts` | Manual en-excel-data-rounding; en-column-line-area-charts § 100% stacked | core › "100% chart labels always sum to 100"; mekko › "percent mekko…"; area › "100% area…"; pie › "shares add up to exactly 100…" |
| RND-3 | Stacked, unit-Mekko, Mekko XY, area and combo segment labels are rounded together so they add up to the total shown (not with a custom `format`) | `charts/column.ts: model (rounded)`; `mekko.ts`; `mekkoxy.ts: roundedOf`; `area.ts`; `combo.ts` | Manual en-excel-data-rounding § Introduction (consistent rounding) | engine › "B14: stacked segment labels add up to the total shown"; quality › "stacked segments must add up…" |

## 5 Dates and 30/360

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| DATE-1 | Category labels read as dates: years, FY/CY, estimate markers (E/A/F/P/B, est., act., fcst.), quarters (Q1 2023, 2023-Q1, 1Q23), halves, months (Jan 2023, Mar-24, 2023-01, 01/2023, 2024 Mar), ISO dates; each maps to the first day of its period | `dates.ts: parseCategoryDate` | Manual en-chartdecorations § CAGR arrows (dates from category cells) | engine › "B3: category labels parse as dates"; behaviour › "DATE-2" |
| DATE-2 | Two-digit years pivot at 70 (`'24` → 2024, `FY99` → 1999) | `dates.ts: year` | Convention (Excel two-digit year cutoff) | behaviour › "DATE-2" |
| DATE-3 | Year fractions use the 30/360 (US) day count | `dates.ts: years360` | Manual en-chartdecorations § CAGR arrows ("30/360 day count convention") | engine › "B3: 30/360 day count" |
| DATE-4 | CAGR years: explicit `years`, else 30/360 between the two labels' dates, else the index distance | `dates.ts: cagrYears` | Manual en-chartdecorations § CAGR arrows; house rule for the fallbacks (B3) | engine › "B3: quarterly CAGR uses dates…" |

## 6 Value axes

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| AX-1 | A value axis asks for 3–5 ticks by its length (fewer on a chart resized small) | `common.ts: valueTicks` | House rule (tick labels never crowd) | behaviour › "AX-1" |
| AX-2 | Nice ticks: steps 1, 2, 2.5, 5 × 10ⁿ, 3–9 ticks covering the data, scored by wasted range, distance from the asked count and a 2.5 penalty | `numbers.ts: niceTicks` | Convention: Talbot, Lin & Hanrahan 2010 (extended Wilkinson), simplified | core › "nice ticks use 1/2/5 steps…"; behaviour › "AX-2/3" |
| AX-3 | Empty axis range costs more than an extra tick (620 → 0–700, not 0–800) | `numbers.ts: niceTicks (score)` | House rule | behaviour › "AX-2/3" |
| AX-4 | The axis always includes the data; with a value axis the plot spans the outer ticks | every chart type (`ticks → dmin/dmax`) | Manual en-axes § Adjust the scale ("must always include the range…") | charts-b › "100% chart with a value axis…"; core › "nice ticks…" |
| AX-5 | Value axis off by default except for line charts: values are labelled on the chart | `spec.ts: optionDefaults` | Manual en-axes § Value axes | engine › "options: optionDefaults()…" |
| AX-6 | Left axis: line, 3 pt ticks, right-aligned muted labels; horizontal bars put it along the bottom | `common.ts: valueAxis`; `charts/column.ts: layoutBars` | Convention | engine › "horizontal stacked bars…"; charts-b › "100% chart with a value axis…" |

## 7 Category labels

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| CAT-1 | Each name wraps into its slot (≤ 3 lines); if any does not fit, the whole row shrinks one point at a time to the theme minimum; what still does not fit is hidden and reported | `common.ts: fitCategories, reportFit` | House rule (B8); manual en-chart-area-plot-area § Automatically resize | engine › "B8: a word wider than its slot shrinks…" |
| CAT-2 | Category labels sit centred under their slot, top-aligned below the baseline | `common.ts: categoryLabels` | Convention | engine › "every shape… has a role"; area › "every category label is centred under its point…" |
| CAT-3 | A crowded time axis (every label a date) shows every 2nd, 3rd, … label on one line instead of hiding at random; end labels kept inside the frame, thinning further if they then clash | `common.ts: fitCategories (thin), categoryLabels` | House rule (how a consultant thins a year axis) | behaviour › "CAT-3"; resize › matrix |
| CAT-4 | Names that cannot all show in uniform slots ask for a wider frame (TooSmall → placeholder) when the needed width is ≤ 864 pt, else stay hidden and reported | `common.ts: askForRoom`; `index.ts: layoutChart` | House rule (an axis with gaps reads as missing data) | behaviour › "CAT-4" |
| CAT-5 | Horizontal bars: category column as wide as the longest name, at most 30% of the frame, right-aligned | `charts/column.ts: layoutBars` | Convention | engine › "horizontal stacked bars: categories on the left…" |

## 8 Series labels and legends

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| SER-1 | Default identification: names right of the last column (stacked, 100%, Mekko, area, line, stacked combo); a legend for clustered, horizontal bars, multi-series waterfall, clustered combo; scatter: legend with several groups, none with one | `common.ts: makeCtx`; per type | Manual en-legends; en-textlabels § Label types (series labels) | behaviour › "SER-1", "SC-6/7"; engine › "multi-series waterfall: legend swatches…" |
| SER-2 | Right-hand names are aligned to their segment centres, spread apart without overlap (keeping order), with a thin leader when moved more than 1 pt | `common.ts: seriesLabelsRight`; `place.ts: spread` | Manual en-textlabels § Automatic label placement | core › "spread keeps order and removes overlaps"; behaviour › "SER-2" |
| SER-3 | Names beside the plot that would take more than a third of the frame width become a legend | `common.ts: legend` | House rule (a resized narrow chart keeps every name) | behaviour › "SER-3" |
| SER-4 | A series with an empty last value keeps its name anchored at the stack edge | `charts/column.ts: layoutVertical` | House rule (B4) | labels › "stacked: the sample deck labels every thin Services segment…" |
| SER-5 | Right names never collide with the labels at the last point (line, combo) | `charts/line.ts`, `charts/combo.ts` | House rule (B11) | engine › "B11: right-hand series names never collide…" |
| LEG-1 | Legend: swatch + name left to right from the top-left, wrapping onto as many rows as needed | `common.ts: legend` | Manual en-legends ("text automatically reformatted to fit") ; house rule (B9) | engine › "B9: the legend wraps onto more rows…" |
| LEG-2 | Swatch follows the mark: box (columns, areas), short line (lines), circle (scatter/bubble), diamond (Gantt milestones) | `common.ts: legend`; `charts/gantt.ts` | Convention | behaviour › "LEG-2"; charts-b-polish › "scatter and bubble: legend swatches are circles…" |

## 9 Label placement and the placer

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| LBL-1 | Labels are placed as constrained candidate selection: each label has allowed positions with a cost; the cheapest valid one wins (inside bounds, clear of marks, labels and leaders) | `labels.ts: placeLabels` | Reference engine internals (MASTER-PLAN §12.2, §12.5); manual en-textlabels § Automatic label placement | labels › "placer: cheapest valid candidate wins…" |
| LBL-2 | A segment label that fits sits centred in its segment (text width + 2 ≤ segment width, 0.9 line ≤ height) in the contrast colour | `charts/column.ts`, `mekko.ts`, `mekkoxy.ts`, `waterfall.ts`, `combo.ts` | Manual en-textlabels § Automatic label placement ("place all labels centered") | core › "stacked column chart: totals and segments line up" |
| LBL-3 | A label slightly larger than its segment is boxed on it (a rectangle in the segment colour) while it overflows by ≤ half its height and its text fits the column width | `labels.ts: segmentCandidates` | Manual en-textlabels § Automatic label placement ("colored rectangle underneath"); en-style-files (`avoidSegmentLabelBoxing`) | labels › "placer: a label overflowing its segment a little is boxed…", "a boxed label never pokes out of its column…" |
| LBL-4 | Otherwise it goes outside with a leader, on the allowed sides in order (column: right, left, then above/below the stack; never right of the last column when names are there) | `labels.ts: segmentCandidates, placePending`; `charts/column.ts` | Manual en-textlabels § Automatic label placement ("place the label outside… add a line") | labels › "100% chart: thin shares go outside…", "horizontal bars: thin segments…"; behaviour › "LBL-11" |
| LBL-5 | Outside labels may stagger 1–3 label sizes along their side (cost 10 + side rank + 4 per step); up/left before down/right | `labels.ts: segmentCandidates` | Manual en-textlabels ("offset one to the left and the other to the right") | labels › "placer: two labels too close together go one to the right…" |
| LBL-6 | A leader runs away from its mark: its sideways run beyond its outward run costs like a stagger step per label height | `labels.ts: lead` | House rule | labels › "placer: leaders run away from their mark…", "mekko-3: the thin… share gets a short leader" |
| LBL-7 | Leaders never cross each other or another label; labels keep 1 pt of air | `labels.ts: placeLabels (conflict, GAP)` | Manual en-textlabels § Automatic label placement | labels › "placer: leaders never cross each other…" |
| LBL-8 | Greedy in order of fewest valid positions, then up to 4 improvement rounds; a label left out may evict one blocker to its next valid spot | `labels.ts: placeLabels` | Reference engine internals (min-cost candidates, MASTER-PLAN §12.5) | labels › "placer: leaders never cross… a blocked label evicts one blocker" |
| LBL-9 | A boxed label never covers a neighbouring segment whose own label points at it from outside | `labels.ts: placePending (anchors)` | House rule | behaviour › "LBL-9" |
| LBL-10 | A label with no valid position is hidden and reported, never overlapped | `labels.ts: placePending` | House rule (B13) | labels › "stacked: labels too thin for any spot are hidden and reported…" |
| LBL-11 | Segment labels of the last column never go right when the series names are there | `charts/column.ts: layoutVertical (sides)` | House rule | behaviour › "LBL-11" |
| PLC-1 | Labels along one axis are spread apart keeping order: pushed down from the top, then up from the bottom, each as close to its centre as possible | `place.ts: spread` | House rule (two-pass 1-D packing) | core › "spread keeps order and removes overlaps" |
| PLC-2 | Two boxes collide when they overlap by more than a tolerance (0.5 pt by default, 1 pt in the gate) | `place.ts: overlaps` | House rule | engine › "B12: gate v2 catches…" |
| PLC-3 | A segment crosses a box when any part lies inside it (Liang–Barsky clipping); leaders and arrows are tested this way | `place.ts: segmentHitsBox` | Convention: Liang–Barsky line clipping | behaviour › "PLC-3" |

## 10 Totals, CAGR and difference arrows

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| TOT-1 | Totals: bold, above each column (stacked default on; 100% shows the absolute total; Mekko and Mekko XY on top; area above the top band; doughnut in the hole) | `charts/column.ts`, `mekko.ts`, `mekkoxy.ts`, `area.ts`, `pie.ts`, `combo.ts` | Manual en-textlabels § Label types (total labels) | core › "stacked column chart: totals…"; behaviour › "AR-4/5"; pie › "doughnut: … total in the middle" |
| TOT-2 | Totals wider than their slots are all left out (not every other one) and reported | `charts/column.ts: layoutVertical (totalsFit)` | House rule (gaps read as missing data) | behaviour › "TOT-2" |
| ARR-1 | CAGR and difference arrows use the column totals (100% charts: the absolute totals) | `charts/column.ts: arrowLabel, model.totalsFor` | Manual en-chartdecorations § Series CAGR labels ("always calculated based on the column totals") | behaviour › "ARR-1"; quality › "CAGR label is recomputed from the shown totals" |
| ARR-2 | CAGR = (b/a)^(1/n) − 1, one decimal, signed, in % | `common.ts: cagrLabel` | Manual en-chartdecorations § CAGR arrows | engine › "B3: quarterly CAGR uses dates…"; quality › "CAGR label…" |
| ARR-3 | Total difference arrow, three modes: `absolute` (default, ↔, signed b − a in the chart's format, no sign when zero), `relative` (→, change from a to b in %), `reverse` (←, change from b to a in %); heads follow the mode; a relative label needs a positive base | `charts/column.ts: arrowLabel, headsOf`; `decorations.ts: diffLabel` | Manual en-chartdecorations § Arrow modes | decorations › "total difference arrow: three modes…"; behaviour › "ARR-3" |
| ARR-5 | Arrows are brackets from both columns up to a shared line with a rounded, bold label bubble; one row per arrow above the totals; horizontal bars put them in columns to the right | `common.ts: bracketArrow, bracketArrowH, arrowRowHeight`; `charts/column.ts` | Manual en-chartdecorations § Difference arrows, CAGR arrows | e2e › "CAGR and difference arrows"; engine › "horizontal stacked bars…" |
| ARR-6 | Arrows are decorations: the collision pass drops them before any label | `resolve.ts: priority` | House rule | behaviour › "RES-1/2/3" |

## 10b Decorations

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| DEC-1 | Level difference arrows between two levels — a column total, a segment's level (`series`) or a value line — relative by default, in any of the three modes | `charts/column.ts: levelArrows`; `charts/line.ts`; `decorations.ts: diffLabel, drawLevelArrow` | Manual en-chartdecorations § Level difference arrows, § Arrow modes | decorations › "level arrows: between totals (relative by default)…" |
| DEC-2 | A level arrow's shaft stands beside the later column (below the later bar), or with `position: "between"` in the gap before it; dotted guides run from the other level; line charts connect points of two series | `decorations.ts: drawLevelArrow`; `charts/column.ts`, `charts/line.ts` | Manual en-chartdecorations § Level difference arrows | decorations › "level arrows: "between" sits in the gap…" |
| DEC-3 | The bubble sits on the shaft (centred, else slid along it), else beside it when the arrow is too short; the shaft crosses no label; with no room the arrow is left out and reported | `decorations.ts: drawLevelArrow` | Manual en-chartdecorations § Level difference arrows ("the bubble is automatically placed next to the arrow") | decorations › "level arrows…", "decorations degrade at small sizes…" |
| DEC-4 | Level arrows on clustered and line charts must name a series; both ends must differ; a value-line end must exist | `spec.ts: validateDecorations` | House rule | decorations › "level arrows: validation needs a series…" |
| DEC-5 | Value lines: dashed, parallel to the category axis; `mean` = arithmetic mean of the column totals (clustered: of the bars), labelled "Ø value"; a fixed value widens the axis; `{value}` in a custom label is filled in; label left or right of the chart with a small triangle | `charts/column.ts`; `decorations.ts: valueLineText, drawValueLine` | Manual en-chartdecorations § Value lines | decorations › "value lines: mean of the column totals with "Ø"…" |
| DEC-6 | Decoration lines step around every label they would cross (1.5 pt clear) | `decorations.ts: lineAround` | House rule | decorations › "value lines step around labels they would cross" |
| DEC-7 | 100% indicator on 100% charts: an arrow at the 100% level with "100%", right by default, left when series names take the right, none when a value axis already reads 100% on that side; `none` removes it | `charts/column.ts: indicatorSide`; `decorations.ts: drawIndicator, indicatorSize` | Manual en-chartdecorations § 100% indicators | decorations › "100% indicator: on by default for 100% charts…" |
| DEC-8 | `percentBase` (the datasheet's 100%= row): shares are of the base, not the column sum; with a base, percentage data show absolute totals | `charts/column.ts: model (base)` | Manual en-mekko § percent axis (100%= row); en-waterfall § Difference arrow label: Percent of datasheet 100%= | decorations › "100%= row: shares of the base…" |
| DEC-9 | A category gap is one more slot (a column plus a gap) after the category; the chart keeps its size, so columns get narrower; a gap after the last category is an error | `decorations.ts: slotLayout, slotUnits`; `spec.ts` | Manual en-axes § Category gaps ("as wide as one column width plus one gap width") | decorations › "category gaps widen the space after a category…" |
| DEC-10 | A category axis break draws a break symbol on the baseline after the category | `decorations.ts: categoryBreakSymbol` | Manual en-axes § Category axis breaks | decorations › "category gaps… category breaks draw a symbol on the baseline" |
| DEC-11 | Value axis break: the range from–to is compressed to a short gap; it must lie inside the data and no segment may end inside it (else ignored and reported); straight or wiggle symbols cut every bar that spans it and the axis; ticks are nice below and above, spaced clear of each other | `decorations.ts: resolveBreak, axisDistance, brokenTicks, spacedTicks, breakSymbol` | Manual en-axes § Value axis breaks | decorations › "value axis break: compressed range…" |
| DEC-12 | Combo `secondaryAxis`: listed line series are read against their own nice-tick axis on the right, without the primary unit; columns cannot use it | `charts/combo.ts`; `spec.ts` | Manual en-column-line-area-charts § Secondary axes | decorations › "combo secondary axis…" |
| DEC-13 | Every decoration is drawn after the labels it must not touch; what finds no room is left out and reported, never colliding, at every size | `decorations.ts` (all) | House rule | decorations › "decorations degrade at small sizes…"; resize › matrix |

## 11 Stacked, clustered and 100% columns and bars

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| COL-1 | Stacked: positives stack up from the baseline, negatives down | `charts/column.ts: model, layoutVertical` | Manual en-column-line-area-charts § Simple and stacked | behaviour › "COL-1" |
| COL-2 | Clustered: series side by side in each slot; `null` is an empty slot (no bar, no label) | `charts/column.ts: layoutVertical` | Manual en-column-line-area-charts § Clustered; house rule (B5) | engine › "B5: null in a clustered chart…" |
| COL-3 | Clustered value labels share one size, shrunk together to the theme minimum; a label may overhang its bar by 0.4 pad only where the neighbour does not reach it | `charts/column.ts: barLabelsFit` | House rule | bench-fixes › "clustered: value labels share one size…" |
| COL-4 | Clustered labels too wide even at the minimum: a value axis carries the values instead of a partial row | `charts/column.ts: layoutColumns` | House rule | bench-fixes › "clustered: labels too wide… give way to a value axis…" |
| COL-5 | 100%: each column scaled to 100%; negative values ignored and reported | `charts/column.ts: model` | Manual en-column-line-area-charts § 100% stacked | core › "100% chart labels always sum to 100"; behaviour › "COL-5" |
| COL-6 | `percentData`: values are shares already; a column keeps its own total on the % axis | `charts/column.ts: model (pdata)` | Manual en-mekko § Mekko charts with a percent axis (values entered as percentages) | charts-b › "percentData: values are shares already…" |
| COL-7 | Column width = slot / (1 + gap); default gap 0.5 | `charts/column.ts: layoutVertical`; `spec.ts: optionDefaults` | Manual en-column-line-area-charts (gap width) | behaviour › "COL-7"; e2e › "column gap changes bar width" |
| COL-8 | Horizontal orientation: bars from the left, values axis at the bottom, totals and arrows beyond the bar ends, series in a legend | `charts/column.ts: layoutBars` | Manual en-column-line-area-charts § Column and bar charts | engine › "horizontal stacked bars…", "horizontal clustered and percent bars pass the strict gate" |

## 12 Waterfall

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| WF-1 | Values are deltas floating from the running sum; totals stand on the baseline at the running sum | `charts/waterfall.ts: layoutWaterfall` | Manual en-waterfall § Create waterfall charts | waterfall › "bars float from the running sum…" |
| WF-2 | Totals: `totalsAt`, else the first category and every `"e"` category, else first and last | `charts/waterfall.ts` | Manual en-waterfall § Create waterfall charts ("e") | engine › "waterfall "e" cells mark computed totals…"; e2e › "waterfall: first and last are totals by default…" |
| WF-3 | A total at index 0 seeds the running sum; values of later totals are ignored | `charts/waterfall.ts` | Manual en-waterfall | waterfall › "labels are signed deltas and bold totals…"; behaviour › "WF-3" |
| WF-4 | Several series: positives stack up and negatives down from the entry level; the next column continues from entry + net | `charts/waterfall.ts` | Manual en-waterfall | waterfall › "multi-series stacks positives up…" |
| WF-5 | Colours: totals palette[0]; one series → style.positive / style.negative; several → palette[s + 1]; a series colour wins; legend swatches match | `charts/waterfall.ts: serColor, segColor` | House rule (consulting convention: green up, red down) | behaviour › "WF-5"; engine › "multi-series waterfall: legend swatches match…" |
| WF-6 | Thin connectors join each column's exit level to the next column | `charts/waterfall.ts` | Manual en-waterfall (connectors) | waterfall › "one connector between each pair…"; quality › "waterfall totals and connectors balance" |
| WF-7 | Labels: signed deltas inside when they fit, else above (increase) / below (decrease) for the outermost segment, else boxed or beside with a leader; bold unsigned totals above (below if negative) | `charts/waterfall.ts`; `labels.ts: placePending` | Manual en-textlabels § Automatic label placement | waterfall › "labels are signed deltas…"; labels › "waterfall: a thin segment inside a stack goes beside…" |
| WF-8 | A zero delta is labelled "0" and the chain continues | `charts/waterfall.ts` | House rule | waterfall › "a zero delta is labelled "0"…" |
| WF-9 | A bridge crossing below zero keeps the baseline and reserves a label row below | `charts/waterfall.ts` | House rule | waterfall › "a bridge crossing below zero…" |

## 13 Mekko

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| MK-1 | Column widths are proportional to the category totals and fill the plot; columns touch | `charts/mekko.ts: layoutMekko` | Manual en-mekko § Mekko charts with a percent axis | mekko › "mekko column widths are proportional…"; quality › "Mekko widths sum to the plot…" |
| MK-2 | `percent` (default): every column stacks to 100%, labels are shares rounded to 100 | `charts/mekko.ts` | Manual en-mekko § percent axis | mekko › "percent mekko: every column tops out at 100%…" |
| MK-3 | `unit`: absolute segment heights labelled with values | `charts/mekko.ts` | Manual en-mekko § Mekko charts with units | mekko › "unit mekko…" |
| MK-4 | Under each name, the column's share of the grand total, rounded to add up to 100% | `charts/mekko.ts (colShares)` | House rule (what the width encodes) | behaviour › "MK-4" |
| MK-5 | Negative values ignored and reported; a zero-total column is skipped and reported | `charts/mekko.ts` | House rule | mekko › "negative and zero data are tolerated" |
| MK-6 | A very narrow last column moves the series names to a legend unless `right` was asked for | `charts/mekko.ts (narrowLast)` | House rule | mekko › "a narrow last column moves the series names to a legend…" |
| MK-7 | Totals over narrow columns move apart symmetrically; still colliding, they stagger into a second row; beyond that hidden | `charts/mekko.ts: placeRow, cluster` | Manual en-textlabels ("offset one to the left and the other to the right") | mekko › "totals that collide… are staggered…", "a total over a narrow column is centred on it…" |
| MK-8 | No positive grand total: nothing drawn, reason reported | `charts/mekko.ts` | House rule | behaviour › "MK-8" |
| MK-9 | Thin segments of the first/last column go beside it (series names move right to make room) | `charts/mekko.ts (geometry outW)`; `labels.ts` | Manual en-textlabels § Automatic label placement | labels › "Mekko: thin segments of the last column go beside it…", "Mekko: the narrow Rest-of-world column…"; mekko › "a column too narrow for its labels…" |

## 14 Mekko XY

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| MXY-1 | First series = column widths (X extent), every further series a stacked height; columns touch | `charts/mekkoxy.ts: layoutMekkoXY` | Manual en-mekko § Mekko charts with units (X extent row) | charts-b › "mekkoxy: width from the first series…" |
| MXY-2 | Order: sheet, reverse, ascending or descending by column height | `charts/mekkoxy.ts (order)` | Manual en-mekko (Y extent ordering) | charts-b › "…sorted by height"; behaviour › "MXY-2/3/5" |
| MXY-3 | One height series: its value on top of each column, no series names, axis titles from the series names | `charts/mekkoxy.ts` | House rule (cost-curve convention) | behaviour › "MXY-2/3/5" |
| MXY-4 | Under each column its name and width, unless `valueAxis` turns the baseline into a cumulative-width axis | `charts/mekkoxy.ts` | Manual en-mekko | charts-b › "mekkoxy: value axes on both edges…" |
| MXY-5 | Heights get their own decimals and the unit; widths are shown without the unit | `charts/mekkoxy.ts: fmtH, fmtW` | House rule (a different quantity) | behaviour › "MXY-2/3/5" |
| MXY-6 | Thin segments of the last column go beside it with a leader; elsewhere hidden and reported | `charts/mekkoxy.ts (outs)` | Manual en-textlabels § Automatic label placement | charts-b-polish › "mekkoxy: a segment too thin…" |
| MXY-7 | Zero-width columns are skipped and reported | `charts/mekkoxy.ts` | House rule | charts-b › "mekkoxy: … zero width reported" |

## 15 Line

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| LN-1 | The axis includes zero unless the data sits in a narrow band far from it (all > 0, range < 40% of max); `zeroBased` forces either way | `charts/line.ts` | House rule (retention 95–120% would be a flat line) | behaviour › "LN-1"; e2e › "line: axis start auto / zero / fit data" |
| LN-2 | A floating axis keeps 10% air below the lowest point | `charts/line.ts` | House rule | bench-fixes › "line, floating axis: the lowest point keeps air…" |
| LN-3 | Every point is labelled up to 24 points, else only the last category | `charts/line.ts (labelAll)` | House rule (readability) | behaviour › "LN-3" |
| LN-4 | At a category the top series is labelled above, the bottom one below; then beside; step lines label above | `charts/line.ts` | Manual en-textlabels § Automatic label placement | behaviour › "LN-4" |
| LN-5 | Lines break at empty cells | `charts/line.ts` | Manual en-column-line-area-charts § Interpolation (off) | behaviour › "LN-5/6" |
| LN-6 | A dashed zero line when the data crosses zero | `charts/line.ts` | Convention | behaviour › "LN-5/6" |
| LN-7 | `step`: horizontal to the next category, then vertical | `charts/line.ts` | Manual en-column-line-area-charts § Line scheme (step lines) | step › "step line: every segment is horizontal or vertical…" |
| LN-8 | Value labels avoid data lines (found by role, not stroke width) and markers | `charts/line.ts` | House rule (B10) | engine › "B10: line labels avoid data lines…" |

## 16 Area

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| AR-1 | Bands are closed paths between cumulative sums; points span the plot edge to edge | `charts/area.ts: layoutArea` | Manual en-column-line-area-charts § Stacked area charts | area › "stacked area: bands are closed paths…" |
| AR-2 | Gaps are interpolated linearly; missing ends count as zero; interpolated points carry no label | `charts/area.ts: interpolate` | Manual en-column-line-area-charts § Interpolation | area › "stacked area: missing values are interpolated…" |
| AR-3 | 100% area: every category reaches 100%, labels add up to 100 | `charts/area.ts` | Manual en-column-line-area-charts § 100% stacked area | area › "100% area…" |
| AR-4 | Negative values ignored and reported | `charts/area.ts` | House rule | behaviour › "AR-4/5" |
| AR-5 | `totals` (stacked area only) sit above the top band, lifted clear of its edge across the label width | `charts/area.ts` | House rule | behaviour › "AR-4/5" |
| AR-6 | Labels sit inside their band where it is ≥ 0.9 line tall across the label's width | `charts/area.ts (span)` | Manual en-textlabels § Automatic label placement | area › "area samples and the legend variant pass the quality gate" |
| AR-7 | A single category is a flat band across the plot | `charts/area.ts` | House rule | behaviour › "AR-7" |
| AR-8 | First and last category labels may overhang the plot by half their width; the plot is inset just enough; with a value axis the label row drops clear of the bottom tick | `charts/area.ts` | House rule | area › "every category label is centred under its point…" |

## 17 Combination

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| CB-1 | Columns stacked (default) or clustered; series with `as: "line"` are lines with markers over the same value axis | `charts/combo.ts: layoutCombo` | Manual en-column-line-area-charts § Combination charts | charts-b › "combo: stacked columns plus a line series…", "combo: clustered columns with a line…" |
| CB-2 | Segment labels stay inside their segment, moved within it clear of the line, else hidden and reported | `charts/combo.ts` | House rule | charts-b-polish › "combo: line labels read as their point's, segment labels stay inside…" |
| CB-3 | Lines default to dark greys (text, muted, axis) over a 4 pt background halo; a series colour wins | `charts/combo.ts (lineColors)` | House rule (reads over any column colour) | behaviour › "CB-3" |
| CB-4 | Line labels above/below, then beside; a point inside its stack is labelled above the stack; on a segment only as a last resort | `charts/combo.ts: lineSpots` | House rule | charts-b-polish › "combo: a line point inside its stack…" |

## 18 Pie and doughnut

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| PIE-1 | One series; slices run clockwise from 12 o'clock in category order, sweeps proportional | `charts/pie.ts: layoutPie` | Manual en-pie § Introduction | pie › "slices start at 12 o'clock…" |
| PIE-2 | Shares rounded together to exactly 100 | `charts/pie.ts` | Manual en-excel-data-rounding | pie › "shares add up to exactly 100…" |
| PIE-3 | `labelContent`: percent (default), value, or "value (share)" | `charts/pie.ts` | Manual en-textlabels § Label content; en-pie § Introduction | pie › "label content shows values, or values and shares" |
| PIE-4 | Inside labels sit as close to the rim as they fit; two that would touch move towards the centre | `charts/pie.ts: plan (radial)` | Manual en-textlabels § Automatic label placement (pie inside labels) | pie › "inside labels sit within their slice" |
| PIE-5 | A label larger than its slice is boxed in the slice colour while the box is no wider than the slice; else it goes outside with the name | `charts/pie.ts: narrower` | Manual en-textlabels § Automatic label placement ("colored rectangle underneath") | pie › "a boxed share is never wider than its slice…" |
| PIE-6 | Category names outside, left or right by the slice's side; crowded ones stack in a column pushed away from the equator, each with a leader from the rim | `charts/pie.ts: outward` | House rule | pie › "small slices' shares stack in a column outside…" |
| PIE-7 | Leaders never enter the pie, cross each other or a label | `charts/pie.ts: plan (bad)` | Manual en-textlabels § Automatic label placement | pie › "small slices' shares… with clean leaders" |
| PIE-8 | Crowded labels shrink the pie (to 70% at most) and move it away from the crowd; the pie and its labels are centred together | `charts/pie.ts` | House rule | pie › "small slices' shares stack…" |
| PIE-9 | Doughnut hole 55% of the radius; `totals` puts the bold total in the hole | `charts/pie.ts (HOLE)` | Manual en-pie § Introduction | pie › "doughnut: ring slices with a hole…" |
| PIE-10 | Beyond the palette, colours repeat as lighter tints: no two slices match | `charts/pie.ts: tint` | House rule | behaviour › "PIE-10" |
| PIE-11 | On a narrow frame the names go to a legend on top | `charts/pie.ts` | House rule | behaviour › "PIE-11" |
| PIE-12 | A doughnut total too wide for the hole is left out and reported | `charts/pie.ts` | House rule | behaviour › "PIE-12" |
| PIE-13 | A single value is a full disc; zero and negative values get no slice (negatives reported) | `charts/pie.ts` | House rule | pie › "a single value is a full disc…" |

## 19 Scatter and bubble

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| SC-1 | Each category is a point labelled with its name; each series a group (colour, legend entry) | `charts/scatter.ts: layoutScatter` | Manual en-scatterbubble § Introduction, § Labels | charts-b › "scatter: points map linearly…" |
| SC-2 | Two value axes with nice ticks; each starts at zero unless the data sits in a narrow band (range < 60%); the domain grows so markers stay inside | `charts/scatter.ts: domain` | House rule | charts-b › "scatter: points map linearly onto two nice value axes…" |
| SC-3 | Bubble area is proportional to size; smaller bubbles in front | `charts/scatter.ts: radius` | Manual en-scatterbubble § Bubble charts | charts-b › "bubble: area is proportional to size…" |
| SC-4 | Labels try the spots around the marker (inside a big bubble first), crowded points first; never over a label or marker | `charts/scatter.ts` | Manual en-scatterbubble § Labels | charts-b › "scatter: labels that cannot be placed are hidden…" |
| SC-5 | A point with no free spot gets its label further out (rings in 16 directions) with a leader that crosses nothing | `charts/scatter.ts: leader` | Manual en-scatterbubble § Labels ("connecting lines if necessary") | charts-b-polish › "scatter: a crowded point gets its label further out…" |
| SC-6 | One group: no series names; several: a legend with circle swatches | `charts/scatter.ts` | House rule | behaviour › "SC-6/7"; charts-b-polish › "legend swatches are circles…" |
| SC-7 | The x label at the origin gives way to the y axis's lowest label | `charts/scatter.ts` | House rule | behaviour › "SC-6/7" |
| SC-8 | Bubble size legend: one bubble of a round size (1/2/5 × 10ⁿ ≤ half the max) drawn to scale, top right, with an optional caption | `charts/scatter.ts: roundDown` | Manual en-scatterbubble § Bubble charts (bubble size legend) | charts-b-polish › "bubble: the size legend bubble…"; e2e › "bubble: size legend…" |

## 20 Harvey balls

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| HB-1 | Values are quarters 0–4 (four segments), or a fraction 0–1 when every value is in 0–1 and one is fractional | `charts/harvey.ts: layoutHarvey` | Manual en-harvey-balls-and-checkboxes § Change the number of segments | harvey › "one ball per value, filled in quarters…", "fractions 0–1…" |
| HB-2 | Fill is a wedge from 12 o'clock clockwise; out-of-range values clamped and reported; empty cells stay empty | `charts/harvey.ts` | Manual en-harvey-balls-and-checkboxes § Change the completion state | harvey › "fractions 0–1 fill proportionally; out-of-range values are clamped…" |
| HB-3 | Grid: series are rows (names left, ≤ 35% of the frame), categories columns (headers on top) | `charts/harvey.ts` | House rule | harvey › "grid of rows (series) and columns (categories)…" |
| HB-4 | Rows keep a text-sized height (balls not stretched); the table is centred vertically | `charts/harvey.ts` | Manual en-harvey-balls-and-checkboxes § Resize | harvey › "the table is centred vertically…" |
| HB-5 | Balls use the series colour, else palette[0] | `charts/harvey.ts` | House rule | behaviour › "HB-5" |

## 21 Table

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| TBL-1 | Header of categories, one row per series (name first), numbers through the number service, text cells as typed | `charts/table.ts: layoutTable` | Manual en-tables-with-datasheets | charts-b › "table: header, row names and formatted cells…" |
| TBL-2 | Decimals per row when the headers are periods, per column otherwise | `charts/table.ts (byRow)` | House rule (financial summary vs comparison) | behaviour › "TBL-2/3" |
| TBL-3 | Text columns left-aligned; numbers (and text that reads as a number) right-aligned; headers follow their column | `charts/table.ts (align)` | Convention | behaviour › "TBL-2/3" |
| TBL-4 | Columns fit their content: headers and names wrap (≤ 3 lines), then the font shrinks; the table stretches to the frame width; columns that cannot fit at the minimum size make the chart a placeholder | `charts/table.ts: fit` | Manual en-table § Automatically resize table rows and columns | charts-b › "table: wide content wraps headers, then shrinks…"; resize › "a table whose columns cannot fit…" |
| TBL-5 | Rows that do not fit the frame height are dropped and reported | `charts/table.ts` | House rule | charts-b › "table: … reports rows that do not fit" |
| TBL-6 | Spare height is shared by the rows (≤ 2 lines more each), content centred; the top edge stays | `charts/table.ts (grow)` | Manual en-table § Resize tables | charts-b-polish › "table: rows share the spare frame height…" |

## 22 Gantt

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| GNT-1 | Categories are rows; series 1 starts, 2 ends, further series milestone kinds; a date is the beginning of its day | `charts/gantt.ts: layoutGantt` | Manual en-gantt § Datasheets, § Timeline items | charts-b › "gantt: bars span their dates…"; behaviour › "GNT-1" |
| GNT-2 | Calendar scale: the finest of week/month/quarter/year whose labels fit (≤ 60 units) is the lower tier, the next larger unit the upper tier | `charts/gantt.ts: pickScale` | Manual en-gantt § Calendar scales | charts-b › "gantt: the scale follows the range…" |
| GNT-3 | The range fits the data snapped to whole lower-tier units; weeks start on Monday | `charts/gantt.ts: floorTo, ceilTo` | Manual en-gantt § Modify the date range; convention ISO 8601 weeks | charts-b-polish › "gantt: a week scale spans whole weeks…" |
| GNT-4 | Milestones near either end widen the range for the diamond and its date label | `charts/gantt.ts` | House rule | charts-b-polish › "a week scale spans whole weeks even when a milestone label widens the range" |
| GNT-5 | Upper-tier separators solid through header and rows; lower-tier dashed | `charts/gantt.ts` | Manual en-gantt § Calendar scale separators | behaviour › "GNT-5/12/13" |
| GNT-6 | Scale labels: longest that fits (e.g. "Mar 2026" → "Mar" → "M"; "2026" → "'26") | `charts/gantt.ts: unitLabel` | Manual en-gantt § Calendar scale labels | charts-b › "gantt: the scale follows the range…" |
| GNT-7 | Row labels in a left column (≤ 30% of the frame, ≤ 2 lines), wrapped, shrunk, then hidden | `charts/gantt.ts`; `common.ts: fitCategories` | Manual en-gantt § Row labels | charts-b › "gantt: bars span their dates…" |
| GNT-8 | Rows share the frame height up to 4.5 lines; a short plan keeps its top edge | `charts/gantt.ts (rowMax)` | Manual en-gantt § Rows | charts-b-polish › "gantt: rows share the frame height…" |
| GNT-9 | Milestone date label: "Mar 15" within a year, "Mar 15, 2026" beyond, "Mar 2026" beyond three years; day first outside en-US; right of the diamond, else left, never over a bar | `charts/gantt.ts: dateLabel` | Manual en-gantt § Timeline item labels, § Date format control | behaviour › "GNT-9" |
| GNT-10 | `shades`: shaded periods behind every row; a labelled shade gets a bracket under the chart | `charts/gantt.ts` | Manual en-gantt § Types of timeline items | charts-b-polish › "gantt: brackets label periods…"; e2e › "gantt: today line, shaded periods…" |
| GNT-11 | `brackets` under the rows share lanes with shade labels; items that would touch take another lane | `charts/gantt.ts: addBand` | Manual en-gantt § Types of timeline items | charts-b-polish › "gantt: brackets label periods…" |
| GNT-12 | `today`: a red line through every row labelled "Today"; outside the range it is reported; never read from the clock | `charts/gantt.ts` | House rule (deterministic output) | charts-b-polish › "…the today line crosses every row"; behaviour › "GNT-5/12/13" |
| GNT-13 | A milestone key appears with two or more milestone kinds; `seriesLabels: none` hides it | `charts/gantt.ts (showKey)` | House rule | behaviour › "GNT-5/12/13" |

## 23 Text, contrast and minimum sizes

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| TXT-1 | Widths are the sum of glyph advances without kerning or ligatures, as PowerPoint draws by default | `text.ts: measurerFromMetrics, build` | Convention (PowerPoint `kern` off by default; B6) | text › "metrics measurer matches opentype.js to the bit…" |
| TXT-2 | Line height from the font: typo metrics when USE_TYPO_METRICS, else win ascent + descent plus extra hhea line gap | `text.ts: lineHeightFactor` | Convention: OpenType OS/2 & hhea, GDI/DirectWrite single spacing (B7) | text › "B7: line height comes from the font tables…" |
| TXT-3 | Line breaks (UAX #14 subset): after spaces, after hyphens/slashes/dashes (kept on the line), between two CJK ideographs/kana; never before closing punctuation; greedy, the last line keeps the overflow | `text.ts: breakPieces, wrap` | Convention: Unicode UAX #14 | text › "line breaking: spaces, hyphens, slashes and CJK"; behaviour › "TXT-3" |
| TXT-4 | East Asian characters missing from a Latin font are one em wide; shaped scripts (Arabic, Hebrew) go through the shaper | `text.ts: FULL_WIDTH, SHAPED` | Convention (PowerPoint East Asian font fallback) | bench-fixes › "text: CJK characters… one em wide"; text › "measurer… shapes Arabic and Hebrew…" |
| TXT-5 | Without a bold font, bold is measured 6% wider | `text.ts: build` | House rule (synthetic bold) | behaviour › "TXT-5" |
| CON-1 | Text on a fill is black (222222) or white, whichever has the higher WCAG contrast | `numbers.ts: textOn` | Convention: WCAG 2.x relative luminance and contrast ratio | core › "number formatting" (textOn) |
| MIN-1 | Theme minimum font: labels never shrink below min(theme size, 8 pt) | `common.ts: makeCtx (minFs)` | House rule (legibility) | resize › "every text keeps the theme minimum size…"; quality › "typography flags text below 8 pt…" |
| MIN-2 | Minimum frame per type (category charts grow by 14 pt per category, horizontal bars 12 pt per row, Mekko 24 pt, Harvey 24 pt per column and 16 per row, Gantt 14 pt per row) | `minsize.ts: minimumSize` | House rule (documented in the function) | behaviour › "MIN-2"; resize › matrix |
| MIN-3 | Below the minimum, or when content cannot fit (TooSmall), the scene is a placeholder: frame outline and the largest readable message saying how large to make it | `minsize.ts: placeholder`; `index.ts: layoutChart` | House rule | resize › "below the minimum the chart is an outlined placeholder…" |

## 24 Resizing and collision resolution

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| RES-1 | After a chart type lays out, text that still collides (frame, label, foreign mark, data line, arrow) is hidden — never shrunk — in priority order: category > legend > series name > axis title > total > tick > shade > value > other > arrow | `resolve.ts: resolveCollisions, priority` | House rule; manual en-resizing § Automatically resize elements to fit their contents | behaviour › "RES-1/2/3"; resize › matrix |
| RES-2 | Shapes of one decoration stand or fall together (arrow + bubble, legend swatch + name, series leader + name) | `resolve.ts: grouped` | House rule | behaviour › "RES-1/2/3" |
| RES-3 | A dropped label takes its pie leader or segment leader with it | `resolve.ts` | House rule | behaviour › "RES-1/2/3" |
| RES-4 | At the default frame the pass is a no-op (it judges like the gate) | `resolve.ts` + `quality.ts` (shared tests) | House rule | resize › every case; behaviour › "RES-4" |
| RES-5 | Layout is deterministic at every size; every type stays clean from 160×90 to 1400×800 pt or becomes a placeholder | `index.ts: layoutChart` | House rule (the add-in re-lays out on every resize) | resize › "layout is deterministic at every size", matrix |
| RES-6 | Every shape carries an engine-assigned role; roles are the contract for tags and selection names | `common.ts: Ctx.add`; `index.ts` | House rule (MASTER-PLAN §3) | engine › "every shape of every chart type has a role…"; pptx-io › "each child gets OC_ROLE…" |

## 25 Quality gate

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| QG-1 | Gate v2: text×text, text×frame, text over another owner's mark (except a boxed label on its own column and a line label inside a column), any shape out of frame, arrows through labels, glyph overflow with a measurer | `quality.ts: checkScene, coversMark` | House rule (B12) | engine › "B12: gate v2 catches…", "B8: the gate reports glyph overflow…"; labels › "quality gate: a boxed segment label may cover its own column only" |
| QG-2 | Strict mode reports every hidden label, with an allow-list for declared stress fixtures | `quality.ts: checkScene (strict)` | House rule (B13) | engine › "B13: hidden labels are issues in strict mode…" |
| QG-3 | Numeric checks: 100% columns sum to 100, segments add to totals, CAGR recomputed, waterfall balances, Mekko widths follow totals, typography ≥ 8 pt and one size per role | `scripts/quality.ts`, `src/report.ts` | MASTER-PLAN §4.3 P1 (2) | quality › all |

## 26 Theme colours, style and same scale

| ID | Rule | Where | Source | Tests |
|---|---|---|---|---|
| THM-2 | The default style follows the presentation: series `accent1`… `accent6`, text `tx1`, separators `bg1`; greys and the waterfall green/red stay fixed | `index.ts: defaultStyle` | Manual en-default-theme (theme colours for text, axes, annotations); en-format-and-style-charts § Chart fill scheme ("PowerPoint Theme" by default) | theme › "theme: the default style draws series in theme accents…" |
| THM-3 | The writer emits `schemeClr` for a slot and `srgbClr` for hex; a CLI deck's theme1.xml holds the style's theme colours; the SVG preview draws the resolved hex | `render/pptx.ts: clrXml, theme`; `render/svg.ts` | Manual en-style-files (`schemeClr` = PowerPoint theme colour) | theme › "theme: the writer emits schemeClr…" |
| THM-4 | The add-in gets slot names in the line format and sets `ObjectThemeColor`; slots resolve against the slide's own theme colours (request `theme=`), which win over `style.theme` | `addin/format.ts`; `addin/engine.ts: parseRequest, layout`; VBA `SetColor`, `ThemeLine` | House rule (label contrast matches what PowerPoint draws) | theme › "theme: the add-in line format carries…", "theme: the add-in resolves slots in the slide…" |
| THM-5 | A chart's own `style` wins over the deck's; `style.numbers` (locale, decimals, scale) applies where the type reads the option and a chart's options win; style errors are path-addressed (`/style/numbers/decimals`) | `index.ts: layoutChart`; `common.ts: makeCtx`; `spec.ts: validateStyle` | House rule | theme › "style.numbers: number defaults…", "style validation: colours are hex or theme slots…" |
| THM-6 | Same scale: charts on one slide with the same `scaleGroup` draw one value unit equally long — the smallest of their own scales; the others stretch their domain away from the baseline (ticks inside it, axis width reserved for them); a group of one is unchanged | `index.ts: layoutGroup`; `common.ts: sharedScale, axisWidth` | Manual en-axes § Match the scales of different axes ("resize smaller scales to match the largest one") | theme › "same scale: charts in one group draw one unit equally long…", "same scale: works with value axes on and both orientations…" |
| THM-7 | Same scale applies to stacked, clustered, waterfall, line, area and combination charts (primary axis); a chart with a value axis break has no linear scale and keeps its own | `spec.ts: OPTIONS.scaleGroup`; `charts/column.ts` | Manual en-axes § Match the scales of different axes | theme › "same scale: line, area and combo charts take part…"; behaviour › "THM-7" |
| THM-8 | A slide holds one `chart` or several `charts` (`{frame?, chart}`, never both); without frames they split the chart area into equal columns 24 pt apart; import gives them back with their frames | `deck.ts: slideCharts, buildDeck`; `spec.ts: validateDeck`; `ooxml/pptx-read.ts: importDeck` | House rule (MASTER-PLAN §4.1 P1 "two charts/slide") | theme › "deck: a slide holds several charts…", "deck: chart and charts together…" |
| THM-9 | Grouped charts carry `OC_SGROUP` and `OC_SCALE` (own pt per unit, plain digits); in PowerPoint, Update chart, Apply and resizing re-lay out the group's other charts with the smallest partner scale (`scale=`) | `deck.ts: chartContent, scaleText`; `addin/engine.ts`; VBA `ScaleLine`, `SyncGroup` | Manual en-axes § Match the scales ("the scales of all other charts set to the same scale are then adjusted as well") | theme › "deck: a slide holds several charts…"; live `npm run e2e:powerpoint` step 10 |
