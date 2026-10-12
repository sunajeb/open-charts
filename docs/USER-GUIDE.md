# open-charts user guide

This guide is for the people who build the slides: consultants, analysts, anyone who turns numbers
into a chart. It covers the PowerPoint add-in, the datasheet editor and each chart type. For the
exact JSON behind a chart, see [DATA-FORMAT.md](DATA-FORMAT.md).

- [The basics](#the-basics)
- [Working in PowerPoint](#working-in-powerpoint)
- [The datasheet](#the-datasheet)
- [Chart types](#chart-types)
- [Numbers: units, decimals, magnitudes, separators](#numbers-units-decimals-magnitudes-separators)
- [Decorations: totals, CAGR and difference arrows](#decorations-totals-cagr-and-difference-arrows)
- [Resizing](#resizing)
- [Hidden labels and why](#hidden-labels-and-why)
- [Keyboard shortcuts](#keyboard-shortcuts)
- [Troubleshooting](#troubleshooting)

## The basics

An open-charts chart is a group of ordinary PowerPoint shapes: rectangles, lines, text boxes. Anyone
can open the deck, with or without open-charts installed, and it looks the same. What makes it a
chart rather than a drawing is the data stored inside the group: you can select it at any time, edit
its data and have it laid out again.

open-charts makes the layout decisions for you: where each label goes, whether it fits inside its
segment, how the axis is scaled, how the totals and growth arrows sit. If a label cannot be placed
cleanly it is **hidden and reported**, never drawn on top of something else.

## Working in PowerPoint

Install the add-in first (see the [README](../README.md#install)). You then have an **open-charts**
tab on the ribbon, before Home, with one group, **Charts**:

| Control | What it does |
|---|---|
| **Insert chart** | A menu of chart types, grouped *Column and bar · Line and area · Mekko · XY · Pie · Timeline and table*. Inserts the chart with sample data on the current slide. |
| **Edit data** | Opens the selected chart's datasheet in your browser. Click **Apply to slide** there and the chart on the slide updates. |
| **Update chart** | Lays the selected chart out again at its current size and position. |

Edit data and Update chart are greyed out unless an open-charts chart is selected.

### Insert, edit, apply, resize

1. **Insert.** Click a slide in Normal view, then **Insert chart** and pick a type. If the slide has
   an empty content placeholder, the chart takes its place and size; otherwise it goes in the
   standard chart area. The chart arrives with sample data.
2. **Edit data.** With the chart selected, click **Edit data**. The datasheet opens in your browser
   with the first data cell ready for typing. Type or paste your numbers; the preview updates as you
   type.
3. **Apply.** Click **Apply to slide** (or press ⌘↵). The chart on the slide is replaced in place,
   keeping its position and stacking order. The window stays open, so you can keep editing and apply
   again. **Cancel** (or Esc twice) closes without changing the slide.
4. **Resize.** Drag the chart's corner handles like any shape. When you click elsewhere, open-charts
   lays the chart out again at the new size: text keeps its size, and the bars, labels and arrows are
   placed for the new space. Moving a chart only records its new position.

The first time Apply updates a slide, macOS may ask whether the editor may control Microsoft
PowerPoint; click **OK**.

### Editing an existing chart

Select the chart (click it once) and click **Edit data**. This works on any open-charts chart,
including ones written by the command line or the web editor, and on decks you received from
someone else: the data travels inside the file. Don't ungroup the chart or edit its shapes by hand
if you want to keep editing its data: the next Update chart or Apply redraws the chart from its data
and replaces those changes.

## The datasheet

The datasheet is the same in the add-in's **Edit data** window and in the standalone web editor.

- **Columns are categories** (years, regions, slices, activities). The header row holds their names.
- **Rows are series** (segments, lines, groups). The first column holds their names, with a colour
  swatch to change the series colour.
- A one-line hint above the sheet says what rows and columns mean for the current chart type.
- **Paste from Excel.** Copy a range in Excel and paste it into a cell: it fills the sheet from that
  cell, adding rows and columns as needed. Paste into the top-left corner cell to replace the whole
  table, with categories in the first row and series names in the first column. A single value
  pasted over a selected range fills every cell of it.
- **Numbers as you type them.** `1,234`, `(5)` (read as −5), `−3`, `$12` and `7%` are all
  understood; currency signs, `%` and spaces are dropped, so `7%` is stored as 7. Use the unit
  option to show `%` on the labels.
- **Empty cells** are allowed: nothing is drawn and no label is shown.
- **Text that isn't a number** (for the chart type) turns the cell red; the cell counts as empty and
  the message under it says what was expected. Validation errors from the engine also land on the
  cell they concern.
- **+ Series / + Category** add a row or column; **− Series / − Category** delete the one of the
  selected cell. Right-click a cell for insert left/right/above/below, delete, copy and clear.
- **Transpose** swaps rows and columns; **Reverse series** flips the row order.
- **Undo / redo** work for everything: typing, pasting, options, type changes.
- **Chart type** is a gallery at the top of the options panel. Switching type keeps your data where
  the shapes match (numbers stay numbers). Switching between numbers, points (scatter/bubble) and
  dates (Gantt) converts what it can or starts from sample data.

The options panel on the right only shows what applies to the current chart type, in sections:
*Chart*, *Numbers*, *Axes & layout*, *Labels & legend*, *Decorations*. Every control has a tooltip.

Under the preview, problems are listed and numbered on the slide: click one to highlight the shapes
it concerns, or, for a hidden label, to jump to its cell.

### The standalone web editor

Run `open-charts serve` (see the README) and open http://localhost:4173. It works on a deck of
slides: add, duplicate and delete slides, give each an action title, then **Download .pptx**.
**Export JSON** saves the deck as JSON; **Import JSON** opens a deck or single-chart JSON file. Your
work is kept in the browser between visits.

A few options exist only in the JSON, not in the editor's panel: `labelContent` (pie values or
shares), `sizeLegend` (bubble), `shades`, `brackets` and `today` (Gantt), and an explicit number of
years for CAGR. Set them in a JSON file and import it; the editor keeps them.

## Chart types

Each section says when to use the chart, how to fill the datasheet and which options apply. The
images are rendered from `examples/`.

### Stacked column (`stacked`)

![Stacked](../test/goldens/darwin/deck-1.png)

**Use it for** composition over time: revenue by segment, headcount by function.
**Datasheet:** each row is a segment stacked in every column; each column is a category.
Negative values stack downwards from the baseline.
**Options:** horizontal bars, value axis, column gap, series labels (right of the last column by
default, or a legend), totals (on by default), CAGR and difference arrows. Segment labels are
rounded so they add up exactly to the total shown on top.

### Clustered column (`clustered`)

![Clustered](../test/goldens/darwin/deck-3.png)

**Use it for** side-by-side comparison: win rate by region and segment.
**Datasheet:** each row is a series; its bars stand side by side in every category.
**Options:** horizontal bars, value axis, column gap, series labels (a legend by default). No totals
or arrows.

### 100% column (`percent`)

![100%](../test/goldens/darwin/percent-1.png)

**Use it for** mix: share of revenue, channel mix.
**Datasheet:** each row is a series; every column is scaled to 100%. Shares are rounded so every
column reads exactly 100%. Negative values are ignored (and reported).
If your data is already in percent and a column may add up to less than 100% (survey answers with
no-replies), tick **Values are %**: each column then keeps its own total.
**Options:** horizontal bars, value axis, gap, series labels, totals (off by default; the raw column
totals), CAGR and difference arrows (on the raw totals). The add-in also offers **100% with % axis**,
the same chart with the value axis switched on.

### Waterfall (`waterfall`)

![Waterfall](../test/goldens/darwin/waterfall-1.png)

**Use it for** bridges: EBITDA walk, cash bridge, price-volume-mix.
**Datasheet:** the first row drives the bridge. Positive values step up (green), negative values
step down (red). The first category is the start value. To make a category a computed total, type
**`e`** in its cell: the column stands on the baseline at the running sum. Without any `e`, the first
and last categories are totals.
**Options:** **Total columns** (pick exactly which categories are totals; overrides `e`), value axis,
gap. With several rows, increases stack upwards and decreases downwards, each row in its own colour,
named in a legend.

### Combination (`combo`)

![Combination](../test/goldens/darwin/combo-1.png)

**Use it for** actuals against a plan or an average: columns with a line on top.
**Datasheet:** each row is a series. Under *Chart*, **Series drawn as lines** picks which rows are
lines (the last row becomes a line when you switch to this type). Lines share the columns' value
axis.
**Options:** columns stacked or clustered, value axis, gap, series labels, totals (stacked columns).

### Line (`line`), including step lines

![Line](../test/goldens/darwin/line-1.png) ![Step line](../test/goldens/darwin/step-1.png)

**Use it for** trends over time.
**Datasheet:** each row is a line; each column is a point on the category axis. Empty cells leave a
gap.
**Options:** **Step line** (holds each value until the next category), axis start (zero, or fit to
the data; by default the axis starts at zero unless the values sit in a narrow band far from it),
value axis (**on** by default for lines), series labels. Every point is labelled while the chart has
at most 24 points in total; beyond that only the last point of each line is labelled.

### Stacked area and 100% area (`area`, `area100`)

![Area](../test/goldens/darwin/area-1.png) ![100% area](../test/goldens/darwin/area-2.png)

**Use it for** composition over time when the trend matters more than individual values.
**Datasheet:** each row is a band stacked on the ones above it; each column is a point. A missing
value is filled in from its neighbours (and carries no label); missing values at either end count as
zero. Negative values are ignored.
**Options:** value axis, series labels, totals above the top band (stacked area only). The 100% area
labels are rounded so each column reads exactly 100%.

### Mekko (`mekko`)

![Mekko](../test/goldens/darwin/mekko-1.png)

**Use it for** market maps: segment size as width, share within it as height.
**Datasheet:** columns are categories; their width is the column total. Rows are the segments.
Negative values are ignored.
**Options:** **Mekko heights**: *100% (share)* labels each segment with its share of the column;
*Absolute* keeps segment heights in units. Totals on top (on by default) and each column's share of
the grand total under its name. Value axis, series labels (switches to a legend when the last column
is too narrow).

### Mekko XY (`mekkoxy`)

![Mekko XY](../test/goldens/darwin/mekkoxy-1.png)

**Use it for** cost curves (capacity × cost) and spend per head (headcount × spend).
**Datasheet:** the first row is the column width (X extent); the rows below are stacked heights.
**Options:** **Order** (sheet order, reverse, ascending or descending height), value axis (also turns
the baseline into an axis of cumulative widths), totals, series labels.

### Scatter (`scatter`)

![Scatter](../test/goldens/darwin/scatter-1.png)

**Use it for** relationships between two measures: price vs. volume, accounts by size.
**Datasheet:** one point per column, labelled with the column name. Each cell is **`x; y`**, for
example `12.5; 42`. A row is a group with its own colour: put each point in the row of its group and
leave the other rows empty.
**Options:** X and Y axis titles, axis start. Groups are named in a legend when there are two or
more.

### Bubble (`bubble`)

![Bubble](../test/goldens/darwin/scatter-2.png)

**Use it for** portfolios: growth vs. share, sized by revenue.
**Datasheet:** like scatter, but each cell is **`x; y; size`**. Bubble *area* is proportional to
size; small bubbles are drawn in front of large ones.
**Options:** axis titles, axis start; a size legend via the JSON option `sizeLegend`.

### Pie and doughnut (`pie`, `doughnut`)

![Pie](../test/goldens/darwin/pie-1.png) ![Doughnut](../test/goldens/darwin/pie-2.png)

**Use them** sparingly, for a few shares of one whole.
**Datasheet:** only the first row is drawn: one slice per column. Negative values are ignored, zero
values get no slice.
**Options:** decimals (shares are rounded to add up to exactly 100%), series labels (category names
outside the slices, a legend, or none), totals (doughnut: the grand total in the hole). Labels show
shares by default; the JSON option `labelContent` shows values or both.

### Harvey balls (`harvey`)

![Harvey balls](../test/goldens/darwin/harvey-1.png)

**Use it for** qualitative assessments: vendors against criteria, workstream status.
**Datasheet:** each row is an item, each column a criterion. Scores are quarters from 0 to 4. If all
scores are between 0 and 1 and at least one is a fraction, they are read as completion (0.6 = 60%
filled). Out-of-range scores are clamped and reported. Empty cells stay empty.

### Table (`table`)

![Table](../test/goldens/darwin/table-1.png)

**Use it for** financial summaries and comparison tables built from native shapes.
**Datasheet:** rows become table rows, columns become table columns. Numbers are formatted with the
chart's number options; text cells (such as `+4.0 pp`) are shown as typed. Decimals are chosen per
row when the headers are periods (a financial summary), per column otherwise.

### Gantt (`gantt`)

![Gantt](../test/goldens/darwin/gantt-1.png)

**Use it for** roadmaps and project plans.
**Datasheet:** one activity per column. The first row is **Start**, the second **End**: dates like
`2026-03-15`, `Mar 2026` or `Q2 2026`. Every further row is a kind of milestone (a diamond with its
date), e.g. "Milestone" and "Decision". An activity needs both a start and an end, or neither (a row
with only milestones). The end must come after the start. A date means the beginning of that day, so
a bar ending `2026-03-16` covers the 15th.
**Options:** the calendar scale (weeks, months, quarters, years) is picked to fit. Shaded periods,
phase brackets and a "today" line are set in the JSON (`shades`, `brackets`, `today`). Separators
(locale) set the date style of milestone labels: US month-first, others day-first.

## Numbers: units, decimals, magnitudes, separators

The *Numbers* section applies to every label, total, arrow and axis tick of the chart:

| Control | Effect | Example |
|---|---|---|
| Unit | Text added to every value | `m` → `42m`; `%` → `42%` |
| Unit before the number | Puts the unit in front | `$` → `$42` |
| Decimals | Fixed digits after the decimal point; empty = automatic | `1` → `42.0` |
| Magnitude | Divides by a thousand, million or billion and adds k / m / bn; *Auto* picks one for the whole chart | 1,250,000 → `1.25m` |
| Separators | Decimal mark and grouping: US, UK, India (`12,34,567`), DE (`1.234,5`), FR (`1 234,5`) | |
| Number format | An Excel format code; overrides decimals; applied on Enter | `#,##0;(#,##0)` → `(480)` |

Automatic decimals use the fewest digits that show every value exactly (up to 2), so whole numbers
stay whole. Shares in 100%, Mekko and pie charts are rounded together so they always add up to
exactly 100%, and stacked segment labels add up exactly to their total. Negative numbers use a real
minus sign (−). See [DATA-FORMAT.md](DATA-FORMAT.md#numbers) for every detail of format codes.

## Decorations: totals, CAGR and difference arrows

- **Totals**: the sum above each column. On by default for stacked and Mekko charts; available on
  100%, combination (stacked), stacked area, Mekko XY and doughnut (in the hole).
- **CAGR arrow** (stacked and 100%): pick the two categories. The growth rate is computed from the
  column totals, over the number of years read from the category labels (`2021`, `FY24`, `Q1 2023`,
  `Jan 2024`, `2025E` all work), or, if the labels aren't dates, the number of columns between them.
  Both totals must be positive; otherwise the arrow is left out and reported.
- **Difference arrow** (stacked and 100%): the signed difference between two column totals.
- **Total columns** (waterfall): which categories stand on the baseline.

## Resizing

open-charts never stretches a chart. Text keeps its point size; the layout is recomputed for the new
box: columns get wider or narrower, labels move inside or outside segments, category names wrap or
shrink.

- **In PowerPoint**, resize the chart like any shape; it is laid out again when you click elsewhere.
  **Update chart** does the same on demand.
- **In the editor and the command line**, the chart fills the standard chart area of a 16:9 slide
  (below the title).

As the box gets smaller, open-charts adapts in this order: category names wrap and then shrink (never
below 8 pt); a crowded time axis shows every 2nd, 3rd, … label; series names move from the right of
the chart into a legend; labels that still do not fit are hidden and reported.

Below a minimum size the chart would be unreadable, so it is drawn as a grey outline with the message
**Chart too small: make it at least W × H pt** instead. Make the chart bigger and it comes back. The
minimums, in points (72 pt = 1 inch; a 16:9 slide is 960 × 540 pt):

| Chart type | Minimum width | Minimum height |
|---|---|---|
| Stacked, clustered, 100%, waterfall, combination, line, area (n = categories) | 200, or 40 + 14 per category | 120 |
| Horizontal bars | 200 | 120, or 30 + 12 per category |
| Mekko, Mekko XY | 240, or 24 per category | 140 |
| Pie, doughnut | 200 | 120 |
| Scatter, bubble | 200 | 140 |
| Harvey balls | 160, or 60 + 24 per column | 80, or 24 + 16 per row |
| Table | 120, more if the columns need it | 50 |
| Gantt | 320 | 120, or 40 + 14 per activity |

A chart with long category names may ask for more width than this, so the names fit.

## Hidden labels and why

A label is hidden when it would overlap another label, cross an edge or stick out of its segment.
The chart is still correct; only that label is missing. You see the list:

- in the editor, under the preview, with numbered markers on the slide (click one to jump to the
  cell);
- on the command line, as `hidden: …` lines (and `--strict` fails the run).

| Reason | What to do |
|---|---|
| `segment too small` | The value is small next to its column. Fine for slivers; otherwise give the chart more height, use fewer decimals or a magnitude, or merge small series into "Other". |
| `label too wide` / `bar too thin for label` | Make the chart wider, use fewer columns, fewer decimals or a magnitude (`1.2m` instead of `1,234,567`). |
| `category label too wide for its slot` / `category labels shrunk to … pt` | Shorten the category names or widen the chart. |
| `no room for label` (line, scatter, pie, Gantt) | Points or slices are too close together. Widen the chart or label fewer series. |
| `chart too small: … needs at least W × H pt` | The chart is shown as a placeholder. Make it at least that size. |
| `chart frame too small` | Make the chart bigger. |
| `needs positive totals` | A CAGR or difference arrow needs positive totals at both ends. |
| `negative values ignored in …` | 100%, Mekko, area and pie charts can't show negatives; use a stacked or waterfall chart. |
| `Harvey ball values outside 0–4 clamped` | Use scores 0–4 (or 0–1). |
| `today: outside the chart's date range` | The Gantt "today" date is before the first or after the last date. |

## Keyboard shortcuts

In the datasheet (⌘ is Ctrl on Windows and Linux browsers):

| Keys | Action |
|---|---|
| Tab / Shift+Tab | Next / previous cell (wraps to the next row) |
| Enter / Shift+Enter | Cell below / above |
| Arrow keys | Move (left/right once the cursor is at the edge of the text) |
| Shift+arrows, or drag | Select a range |
| ⌘C / ⌘X / ⌘V | Copy / cut / paste the range (Excel compatible) |
| Delete or Backspace | Clear the selected range |
| Esc | Collapse the selection to one cell; close a menu or popover |
| ⌘Z | Undo |
| ⇧⌘Z or Ctrl+Y | Redo |
| Right-click, Shift+F10 or the Menu key | Context menu: insert/delete series and categories, copy, clear |

In the **Edit data** window from PowerPoint:

| Keys | Action |
|---|---|
| ⌘↵ (Ctrl+Enter) | Apply to slide |
| Esc twice | Cancel and close |

## Troubleshooting

**The open-charts tab is not on the ribbon.** Run the install again (see the README). If you
installed with `--no-register`, load it yourself: in PowerPoint, *Tools › PowerPoint Add-ins…*,
click **+** and pick `~/Library/Application Support/Microsoft/open-charts/open-charts.ppam`. Make
sure it is ticked in that list.

**PowerPoint asks about macros.** The add-in is a small macro file whose only job is to pass your
click to the chart engine and draw the shapes it returns. Click **Enable Macros** for
`open-charts.ppam`. If macros were disabled and the buttons do nothing, open *PowerPoint ›
Preferences › Security & Privacy*, set *Macro Security* to warn you (or to enable macros), restart
PowerPoint and click **Enable Macros** when asked.

**"open-charts can't reach its chart engine".** The add-in calls Node through a small script that
the install placed in PowerPoint's scripts folder. Run the install again from Terminal and restart
PowerPoint. This also fixes **Node not found** after you upgraded or moved Node (for example with
nvm): the install records the path of the Node that ran it. Check `node --version` is 22.6 or newer
first.

**"This add-in and its engine do not match".** You updated open-charts but not the installed copy.
Run the install again.

**"Select an open-charts chart first".** Click the chart once (select the whole group, not a shape
inside it), then choose Edit data or Update chart.

**"Click a slide (Normal view) first".** Charts are inserted on the current slide in Normal view, not
in Slide Sorter or the outline.

**"The edited data is ready. Select the chart in PowerPoint and click Update chart."** The chart you
were editing was no longer found (for example, you deleted or moved it to another slide). Select it
and click **Update chart**.

**"open-charts is still updating a chart".** Wait a moment and click again.

**macOS asks whether something may control Microsoft PowerPoint.** Click **OK**: that is how Apply
in the browser updates the slide.

**The chart shows "Chart too small".** It is below its minimum size (see [Resizing](#resizing)). Drag
it bigger; it is laid out again when you click elsewhere.

**A label is missing.** See [Hidden labels and why](#hidden-labels-and-why).

**A cell is red.** The text is not what this chart type expects (a number, `x; y`, `x; y; size` or a
date). Hover or focus the cell to see the message.

**My manual formatting disappeared.** Update chart, Apply and resizing redraw the chart from its
data. Make changes through the datasheet and options instead of editing shapes inside the group.
