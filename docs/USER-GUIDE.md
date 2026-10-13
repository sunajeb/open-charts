# open-charts user guide

This guide is for the people who build the slides: consultants, analysts, anyone who turns numbers
into a chart. It covers the PowerPoint add-in, the datasheet editor and each chart type. For the
exact JSON behind a chart, see [DATA-FORMAT.md](DATA-FORMAT.md).

- [The basics](#the-basics)
- [Working in PowerPoint](#working-in-powerpoint)
- [The datasheet](#the-datasheet)
- [Chart types](#chart-types)
- [Numbers: units, decimals, magnitudes, separators](#numbers-units-decimals-magnitudes-separators)
- [Decorations: totals, arrows, value lines, gaps and breaks](#decorations-totals-arrows-value-lines-gaps-and-breaks)
- [Colours and your deck's theme](#colours-and-your-decks-theme)
- [Fonts](#fonts)
- [A client's look: style files](#a-clients-look-style-files)
- [Fill a template](#fill-a-template)
- [Two charts on a slide, same scale](#two-charts-on-a-slide-same-scale)
- [Agenda](#agenda)
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
tab on the ribbon, before Home, with the groups **Charts**, **Chart tools**, **Agenda** (see
[Agenda](#agenda)) and **Help**. The first time you click in a presentation after installing, a short
welcome says where to start and offers a sample deck: one chart of each type, with sample data, to
try Edit data on. **Help › Welcome** shows it again. The Charts group:

| Control | What it does |
|---|---|
| **Insert chart** | The chart gallery: one icon per chart, five per row (see [Choosing a chart and its direction](#choosing-a-chart-and-its-direction)). Inserts the chart with sample data on the current slide. |
| **Direction** ▲ ▼ ◀ ▶ | With a chart selected: draws it growing that way, in place. With nothing selected: the next chart you insert grows that way. The arrow of the selected chart's direction shows pressed; directions its type cannot take are greyed out. |
| **Edit data** | Opens the selected chart's datasheet in your browser. Click **Apply to slide** there and the chart on the slide updates. |
| **Update chart** | Lays the selected chart out again at its current size and position. |

Edit data and Update chart are greyed out unless an open-charts chart is selected.

Next to it, the **Chart tools** group holds the commands that change a chart without touching its
data. Each one is greyed out unless it applies to what is selected:

| Control | What it does |
|---|---|
| **Totals** | Shows or hides the totals above the columns (stacked, 100%, Mekko, Mekko XY, area, combination, doughnut). |
| **CAGR arrow** | Adds a compound annual growth arrow from the first to the last column, or between two columns you select first (Shift-click a segment of each). Click again to remove it; with two other columns selected, it moves there. Stacked and 100% charts. |
| **Difference arrow** | The same for the arrow showing the difference between two column totals. |
| **Same scale** | With two or more charts selected (Shift-click), puts them on one scale: equal values are drawn equally tall in all of them. See [Two charts on a slide, same scale](#two-charts-on-a-slide-same-scale). |
| **Remove from same scale** | The selected charts go back to their own scale. |
| **Convert to plain text** | Select one label of the chart (click the chart, then the label), then click: the label keeps the text it shows now, and data edits no longer change it. It stays in the chart and moves with it. |
| **Reset label position** | Puts the selected label back where open-charts places it after you moved it, or every moved label when the whole chart is selected. |

Formatting you changed with PowerPoint's own tools is kept by these commands, as by Update chart.

**Right-click.** The same commands, plus Edit data and Update chart, are added to PowerPoint's
right-click menu on an open-charts chart (and on its pieces, or on several charts selected together),
showing only the ones that apply. This uses the right-click menu extension of Office add-ins, which
PowerPoint for Windows shows; PowerPoint for Mac does not (see [Undo and the right-click menu on a
Mac](#undo-and-the-right-click-menu-on-a-mac)): there, use the ribbon.

### Insert, edit, apply, resize

1. **Insert.** Click a slide in Normal view, then **Insert chart** and pick a chart. If the slide has
   an empty content placeholder, the chart takes its place and size; otherwise it goes in the
   standard chart area. The chart arrives with sample data.
2. **Edit data.** With the chart selected, click **Edit data**. A small datasheet window pops up
   next to the chart, with the first data cell ready for typing (see [Editing data](#editing-data)).
   Type or paste your numbers: the chart on the slide follows as you type.
3. **Done.** Press ⌘↵ (or click **Done**, or close the window) to keep the changes; **Esc** (or
   **Cancel**) puts the chart back as it was. The chart is redrawn in place, keeping its position and
   stacking order. Without the pop-up helper (see below) the datasheet opens in your browser instead:
   there you click **Apply to slide** (⌘↵) and the editor says **Updated on slide** once PowerPoint has
   done it (or why not, see [Troubleshooting](#troubleshooting)); **Cancel** (or Esc twice) closes it.
4. **Resize.** Drag the chart's corner handles like any shape. When you click elsewhere, open-charts
   lays the chart out again at the new size: text keeps its size, and the bars, labels and arrows are
   placed for the new space. Moving a chart only records its new position.

The first time Apply updates a slide, macOS asks once: *"Microsoft PowerPoint" wants access to
control "Microsoft PowerPoint"*; click **Allow**.

### Editing an existing chart

Select the chart (click it once) and click **Edit data**. This works on any open-charts chart,
including ones written by the command line or the web editor, and on decks you received from
someone else: the data travels inside the file. Formatting you change with PowerPoint's own tools
(colours, fonts, sizes, line styles) is kept when the chart is updated. Two things are not: don't
ungroup the chart (it can't be edited as a chart afterwards), and keep your own notes and shapes
outside the chart's group, because Update chart and Apply redraw everything inside it from the data.

### Dragging handles

On the Mac, a selected chart shows small **yellow handles** next to PowerPoint's own selection
handles. Drag one to change the chart directly; the chart follows the mouse live, and when you let
go it is laid out once more in place (one update, like Apply). **Double-click a handle** to put it
back to automatic.

| Handle | Where | What dragging it does |
|---|---|---|
| Gap width | right edge of the first column (bottom edge of the first bar on a bar chart) | wider or thinner columns: `gap` |
| Axis top | top end of the value axis (right end on a bar chart; only while the value axis is shown) | more headroom above the data: `axisMax`. Dragging below the top of the data gives the automatic axis again |
| Level arrow | right edge of the arrow's bubble | moves the arrow beside the later column or into the gap before it |
| Label | the label itself, once you have selected it (click the chart, then the label) | puts the label where you drop it; it stays there through data edits until you double-click it |

The handles only take the mouse while the pointer is on one: every other click goes to PowerPoint
as usual. They hide while you scroll or zoom, while another app is in front and while you drag
something in PowerPoint, and come back once the chart has been measured again (a moment later).
Zooming with the keyboard (⌘+ / ⌘−) is not noticed: click the chart again to bring the handles back
to the right place.

**Permissions:** none. The handles are a transparent window of the open-charts pop-up helper; it
watches where the mouse is (macOS needs no permission for that) and only catches a press that
starts on a handle. It never asks for Accessibility or Input Monitoring and never sees your
keyboard. (The live test, `npm run live -- handles`, moves the mouse itself and needs
Accessibility for the terminal it runs in; users never do.)

Windows: not yet (the handles are a panel of the Mac helper).

## Choosing a chart and its direction

![The chart gallery](ux-shots/chart-gallery.png)

The gallery lists the charts in the same order everywhere — the ribbon's **Insert chart** and the
editor's chart-type button: Stacked, 100% and Clustered column/bar; Build-up and Build-down
waterfall; Mekko with a % axis; Combination; Line/Profile; Stacked and 100% area; Mekko with a unit
axis; Pie; Doughnut; Scatter; Bubble; Gantt/Timeline; then Mekko XY, Harvey balls and Table.

**Directions.** The small arrows around an icon are the directions that chart can grow in:

| Arrow | The chart | Charts |
|---|---|---|
| ▲ up | columns growing up from a baseline at the bottom (the default) | every chart with arrows |
| ▼ down | columns hanging from a baseline at the top; category names above them, totals below — for losses, reductions, or the lower half of a pair | stacked, 100%, clustered, both waterfalls, both Mekkos, combination |
| ◀ left | bars growing to the left, category names on the right — e.g. the left half of a butterfly chart | the ▼ charts, plus line and both areas |
| ▶ right | bars growing to the right, category names on the left | the same as ◀ |

A line chart turned sideways is a **profile chart**: the categories (criteria, say) run down the
side and each series is a line through its scores — the classic vendor comparison. A Mekko chart
turned sideways draws one row per category, as tall as its share of the total. Pie, doughnut,
scatter, bubble and Gantt charts have no arrows.

- **In the editor**, click an icon for its default (▲) or one of its arrows for that direction.
  With the keyboard: arrow keys move between icons, **Shift+arrow** picks that direction, **Enter**
  picks the icon. While you hover (or focus) an icon or an arrow, the strip under the gallery
  previews your data as that chart, that way, before you pick it. The **Direction** buttons in the
  Chart section turn the current chart.
- **In PowerPoint**, a ribbon gallery cannot have clickable arrows inside its tiles, so the four
  **Direction** buttons sit next to it. Select a chart and click an arrow to turn it in place (its
  data, size and position stay); or, with nothing selected, click an arrow and then insert — the new
  chart grows that way. Only the directions the selected chart can take are enabled.

**Build-up or build-down waterfall?** A *build-up* waterfall stacks components up to a total on the
right (cost elements adding up to total cost). A *build-down* waterfall starts with the total on the
left and subtracts its parts step by step down to the result (revenue down to net income). They are
the same chart with different data, exactly as in the reference manual: the two tiles just insert
different sample data. Either can then be turned with the arrows.

## Office.js add-in (Windows, web, Mac)

The Office.js add-in is a second way to use open-charts in PowerPoint: a task pane instead of a
ribbon tab and browser window, with the chart engine running inside the pane. It needs no Node, no
macros and no AppleScript permission at run time, and works on PowerPoint for Windows, PowerPoint on
the web and PowerPoint for Mac. Both add-ins read and write the same data inside the chart (tag
contract v1), so a chart inserted by one can be edited by the other, and by the command line.

### Set it up

From the repository folder (Node 22.6+):

```sh
npm install
npm run build:officejs      # builds dist/officejs/ (task pane, manifest.xml, fonts, icons)
npm run officejs:cert       # makes a certificate for https://localhost:4175
npm run officejs:serve      # serves the add-in; keep it running while you use PowerPoint
```

**Trust the certificate once.** Office only loads add-ins over trusted https. `officejs:cert` prints
the command; on a Mac it is

```sh
security add-trusted-cert -r trustRoot -k ~/Library/Keychains/login.keychain-db \
  "$HOME/Library/Application Support/open-charts/officejs-cert/localhost.crt"
```

(macOS asks for your password; or double-click the file and set it to *Always Trust* in Keychain
Access). On Windows: `certutil -user -addstore Root localhost.crt`.

**Sideload the manifest** (`dist/officejs/manifest.xml`):

| Where | How |
|---|---|
| PowerPoint for Mac | `npm run officejs:sideload` copies it to `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`. If macOS asks whether Terminal may access data from other apps, click **Allow** (or copy the file there in Finder). Restart PowerPoint. |
| PowerPoint for Windows | Put the manifest in a shared folder, add the folder's `\\machine\share` path under *File › Options › Trust Center › Trust Center Settings › Trusted Add-in Catalogs* (tick *Show in Menu*), restart PowerPoint, then *Insert › My Add-ins › Shared Folder*. |
| PowerPoint on the web | *Insert › Add-ins › My Add-ins › Upload My Add-in*, pick the manifest. The browser must trust the certificate. |
| A whole organisation | Microsoft 365 admin centre › *Integrated apps* › upload the manifest, after `npm run build:officejs -- --base https://your-host/path/` and hosting `dist/officejs/` there (any static https host; [OFFICEJS-HOSTING.md](OFFICEJS-HOSTING.md)). |

Then **Home › open-charts** opens the pane.

### Insert, edit, apply, revert

- **Insert.** With no chart selected the pane shows a sample chart and **Insert on slide**. Pick the
  type in the gallery, type or paste your data, and click **Insert on slide** (⌘↵): the chart is
  drawn on the current slide, in the standard chart area.
- **Edit.** Click any open-charts chart: the pane loads its data (the title reads *Editing …*).
  Change it and click **Apply to slide**: the chart is replaced in place — same position, size,
  name and stacking order — and stays selected.
- **Revert** appears after an Apply and puts back the chart as it was before it, one step per click,
  as long as the same chart stays selected. PowerPoint's own Undo also works, but may need several
  presses for one Apply.
- **Resize.** After you resize a chart, selecting it shows *This chart was resized* with
  **Fit to new size**, which lays it out again for the new size with its stored data (text
  keeps its size, labels are re-placed). Moving a chart needs nothing.

The preview, the issue list and every datasheet feature are the ones described below
([The datasheet](#the-datasheet)); in the pane they stack vertically.

### How the chart is drawn, by PowerPoint version

PowerPoint's add-in API grows with each version (`PowerPointApi` 1.4 … 1.10). The pane checks what
your PowerPoint offers and picks the best way to draw:

| Your PowerPoint offers | Columns, bars, waterfall, Mekko, scatter, table, Gantt | Lines, arrows, pie, doughnut, Harvey balls | Area charts |
|---|---|---|---|
| 1.10 (current Microsoft 365) | Shapes on the slide | Shapes on the slide (rotated lines and arrowheads, pie and arc presets) | Slide re-inserted with the chart |
| 1.8 – 1.9 | Shapes on the slide | Slide re-inserted with the chart | Slide re-inserted with the chart |
| 1.4 – 1.7 | New slide with the chart | New slide with the chart | New slide with the chart |

- **Shapes on the slide** (Path B): the chart is a group of native shapes, exactly as the `.ppam`
  add-in draws it.
- **Slide re-inserted** (Path A): area bands and other free-form shapes cannot be drawn through the
  add-in API, so the pane exports the slide, adds the chart to it with the same writer as the command
  line, inserts it after the original and removes the original. Everything on the slide is kept, but
  it becomes a new slide: links from other slides to it need re-pointing.
- **New slide** (Path A on older PowerPoints): Insert puts the chart on a new slide after the current
  one; Apply cannot replace a chart in place and says so.

The pane's first line in its log (`ready`) lists the API versions your PowerPoint has.

**Colours.** With 1.10 the chart is laid out in the slide's own theme colours; drawn as shapes they
are fixed colours, so after changing the deck's theme select the chart and Apply again. Charts drawn
through a re-inserted or new slide keep real theme colours.

**Your own formatting.** Apply redraws the chart from its data and its overrides (`overrides` in the
spec). Formatting you changed on the slide with PowerPoint's own tools (a bar's colour, a label's
size) is not read back by the pane yet and is replaced; the `.ppam` add-in keeps such edits.
Element overrides the pane cannot draw as shapes (pattern fills, theme fonts) send the chart through
the re-inserted slide.

**Fonts.** Arial is measured with the bundled Liberation Sans (same widths), so Arial charts fit
exactly. Other fonts: see [Fonts](#fonts) — exact when the pane is served by `npm run officejs:serve`,
else with the browser's copy of the font, else a wider stand-in, which the chart reports.

### Troubleshooting the pane

**The pane is blank or says it cannot be reached.** Is `npm run officejs:serve` running? Is the
certificate trusted (open `https://localhost:4175/taskpane.html` in Safari: no warning should
appear)?

**The pane says "Open this pane from PowerPoint".** The page was opened in a browser, not through
Home › open-charts; it previews but cannot insert.

**Debugging on a Mac.** `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`,
restart PowerPoint, right-click the pane › *Inspect Element*. The dev server also keeps the pane's log
at `https://localhost:4175/api/log` and in `~/Library/Application Support/open-charts/officejs.log`.

### Editing data

**Edit data** opens the datasheet in a small window that floats over PowerPoint, next to the chart
(to its right, or left, below or above when there is no room; inside PowerPoint's window when
PowerPoint cannot say where the chart is on screen). It opens in a fraction of a second: a small
helper app (*open-charts pop-up*, no Dock icon) keeps the datasheet loaded in the background while
PowerPoint runs.

- **Live.** Every change reaches the slide on its own, about 0.4 s after you stop typing. The status
  at the top says *Updating the slide…*, then *Updated on slide*. If you keep typing while PowerPoint
  draws, your later edits go in one more update with the latest numbers: PowerPoint is never asked to
  draw twice at once. A cell the chart cannot use (text in a number cell) is marked red and nothing
  is sent until it is fixed.
- **⌘↵ / Done / the window's close button** keep what is on the slide (sending your last edit first).
- **Esc / Cancel** put the chart back exactly as it was when the window opened, and close it.
- **Options** opens chart type and options over the datasheet. The slide itself is the preview, so
  the window has none. Copy, paste, undo, Tab and arrow keys work as in the full editor.
- The window stays with PowerPoint: it hides while another app is in front and comes back with
  PowerPoint. Drag it by its title bar; resize it from its edges.

**Selecting part of a chart.** Click a chart once to select it, then click a label or a bar inside
it: a small toolbar appears above the selection with the formatting that piece takes (today fill
colour, text size, bold, text colour and text direction, as each piece allows; more to come). A click there restyles
just those elements and redraws the chart; the change is stored in the chart's data, so it survives
Update chart, resizing and Edit data. The toolbar steps aside while you drag and hides when nothing
in a chart is selected. Telling the toolbar about a selection costs PowerPoint one short call (about
0.1 s) only when you select a different piece of an open-charts chart; other clicks cost nothing.

**Where it comes from.** `addin install` builds the helper on your Mac (with `swiftc` from Apple's
Command Line Tools: `xcode-select --install` if the install says it is missing) into
`~/Library/Application Support/open-charts/OpenChartsPopup.app`, ad-hoc signed (`codesign -s -`).
Because it is built on your Mac and never downloaded, macOS does not quarantine it and Gatekeeper does
not ask about it (`spctl --assess` says *rejected*, which only matters for downloaded apps). If you copy
a built helper to another Mac, the copy is quarantined and blocked: build it there instead
(`npm run build:popup`, then `addin install`), or clear the flag with
`xattr -dr com.apple.quarantine OpenChartsPopup.app`. If the helper is missing or does not start,
Edit data opens the datasheet in your browser as before, and nothing else changes.

**Windows.** Edit data opens the same pop-up as a borderless Microsoft Edge window (Edge ships with
Windows 10 and 11) next to the chart; without Edge it opens your default browser.

## The datasheet

The datasheet is the same in the add-in's **Edit data** window and in the standalone web editor.

- **Columns are categories** (years, regions, slices, activities). The header row holds their names.
- **Rows are series** (segments, lines, groups). The first column holds their names, with a colour
  swatch to change the series colour: a **theme colour** (Accent 1–6, which follows the deck's
  theme), one of the fixed accents, any hex colour, or **Automatic** (the next palette colour).
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
- **Hide a row** (optional rows): click the eye left of a series name, or right-click a cell and
  pick *Hide series*. The row turns grey and striped, and the chart is drawn as if it were not
  there — totals, shares, the axis and the legend only count visible rows — but its numbers stay in
  the sheet and in the saved chart. Click the eye again to bring it back. The other rows keep their
  colours. At least one row stays visible, and rows a chart reads by position (a Gantt chart's Start
  and End, a Mekko XY chart's widths) always show. Pie and doughnut charts draw the first visible row.
- **Transpose** swaps rows and columns (rows it creates start visible); **Reverse series** flips the
  row order.
- **Undo / redo** work for everything: typing, pasting, options, type changes.
- **Chart type** is a gallery at the top of the options panel. Switching type keeps your data where
  the shapes match (numbers stay numbers). Switching between numbers, points (scatter/bubble) and
  dates (Gantt) converts what it can or starts from sample data.

The options panel on the right only shows what applies to the current chart type, in sections:
*Chart*, *Numbers*, *Axes & layout*, *Labels & legend*, *Decorations*. Every control has a tooltip.
A control you have not touched shows the chart type's default, which is what the preview draws
(totals are ticked on a new Mekko, the value axis on a new line chart). Switching type drops the
options the new type does not use.

Under the preview, problems are listed and numbered on the slide: click one to highlight the shapes
it concerns, or, for a hidden label, to jump to its cell.

### The standalone web editor

Run `open-charts serve` (see the README) and open http://localhost:4173. It works on a deck of
slides: add, duplicate and delete slides, give each an action title, then **Download .pptx**.
**Export JSON** saves the deck as JSON; **Import JSON** opens a deck or single-chart JSON file. Your
work is kept in the browser between visits. Every option of the
[data format](DATA-FORMAT.md#options) has a control in the panel.

## Chart types

Each section says when to use the chart, how to fill the datasheet and which options apply. The
images are rendered from `examples/`.

### Stacked column (`stacked`)

![Stacked](../test/goldens/darwin/deck-1.png)

**Use it for** composition over time: revenue by segment, headcount by function.
**Datasheet:** each row is a segment stacked in every column; each column is a category.
Negative values stack downwards from the baseline.
**Options:** direction, value axis, column gap, series labels (right of the last column by
default, or a legend), totals (on by default), CAGR and difference arrows. Segment labels are
rounded so they add up exactly to the total shown on top.

### Clustered column (`clustered`)

![Clustered](../test/goldens/darwin/deck-3.png)

**Use it for** side-by-side comparison: win rate by region and segment.
**Datasheet:** each row is a series; its bars stand side by side in every category.
**Options:** direction, value axis, column gap, series labels (a legend by default). No totals
or arrows.

### 100% column (`percent`)

![100%](../test/goldens/darwin/percent-1.png)

**Use it for** mix: share of revenue, channel mix.
**Datasheet:** each row is a series; every column is scaled to 100%. Shares are rounded so every
column reads exactly 100%. Negative values are ignored (and reported).
If your data is already in percent and a column may add up to less than 100% (survey answers with
no-replies), tick **Values are %**: each column then keeps its own total.
**Options:** direction, value axis, gap, series labels, totals (off by default; the raw column
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
**Options:** axis titles, axis start; **Size legend** (a reference bubble of a round size, top
right) with an optional caption such as `Revenue ($m)`.

### Pie and doughnut (`pie`, `doughnut`)

![Pie](../test/goldens/darwin/pie-1.png) ![Doughnut](../test/goldens/darwin/pie-2.png)

**Use them** sparingly, for a few shares of one whole.
**Datasheet:** only the first row is drawn: one slice per column. Negative values are ignored, zero
values get no slice.
**Options:** decimals (shares are rounded to add up to exactly 100%), series labels (category names
outside the slices, a legend, or none), totals (doughnut: the grand total in the hole). **Slice
labels** show the share (default), the value, or both (`118m (52%)`); values use the unit and
number options.

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
**Options:** the calendar scale (weeks, months, quarters, years) is picked to fit. Under
*Decorations*: **Today line** (a date; applied on Enter), **Shaded periods** (holidays, freezes)
and **Phase brackets**, each a list of from / to dates with an optional label (**+** adds one,
starting from the first activity's dates; **×** removes it). Separators
(locale) set the date style of milestone labels: US month-first, others day-first.

## Numbers: units, decimals, magnitudes, separators

The *Numbers* section applies to every label, total, arrow and axis tick of the chart:

| Control | Effect | Example |
|---|---|---|
| Unit | Text added to every value | `m` → `42m`; `%` → `42%` |
| Unit before the number | Puts the unit in front | `$` → `$42` |
| Decimals | Fixed digits after the decimal point (0–10); empty = automatic | `1` → `42.0` |
| Magnitude | Divides by a thousand, million or billion and adds k / m / bn; *Auto* picks one for the whole chart | 1,250,000 → `1.25m` |
| Separators | Decimal mark and grouping: US, UK, India (`12,34,567`), DE (`1.234,5`), FR (`1 234,5`) | |
| Number format | An Excel format code; overrides decimals; applied on Enter | `#,##0;(#,##0)` → `(480)` |

Automatic decimals use the fewest digits that show every value exactly (up to 2), so whole numbers
stay whole. Shares in 100%, Mekko and pie charts are rounded together so they always add up to
exactly 100%, and stacked segment labels add up exactly to their total. Negative numbers use a real
minus sign (−). See [DATA-FORMAT.md](DATA-FORMAT.md#numbers) for every detail of format codes.

## Decorations: totals, arrows, value lines, gaps and breaks

- **Totals**: the sum above each column. On by default for stacked, Mekko and Mekko XY charts;
  available on 100%, combination (stacked), stacked area and doughnut (in the hole).
- **CAGR arrow** (stacked and 100%): pick the two categories. The growth rate is computed from the
  column totals, over the number of years read from the category labels (`2021`, `FY24`, `Q1 2023`,
  `Jan 2024`, `2025E` all work), or, if the labels aren't dates, the number of columns between them.
  **Years** overrides that number (e.g. `2.5`); empty = automatic. Both totals must be positive;
  otherwise the arrow is left out and reported.
- **Difference arrow** (stacked and 100%): a bracket between two column totals. **Mode**: *Relative*
  (→, the change in percent; the default), *Absolute* (↔, the signed difference) or *Reverse* (←, the
  change in percent the other way).
- **Level arrows** (stacked, clustered, 100%, line): **+ Level arrow** adds a row: the two ends (a
  category and a series — the top of that segment, bar or point — or *Total* for the column total),
  the mode (relative by default) and where it sits: *Beside* the later column or *Between* the two.
  Dotted guides run from each level to the arrow; a short arrow gets its bubble beside it.
- **Value lines** (stacked, clustered, line): **+ Value line** adds a dashed line at a value or at
  *mean* (the average column total; a clustered chart's bars), labelled on the right or left, e.g.
  `Ø 27.8`, or with your own label (`{value}` inserts the number).
- **100% indicator** (100%): the "← 100%" marker beside the chart, on by default; right, left or none.
- **100%=** (100%): tick **100%= row** to add it to the datasheet: the value each
  column's shares are taken of. With **Values are %**, it is the absolute value behind 100%, and the
  totals show it.
- **Category gaps** and **Category breaks** (column, bar, waterfall, combination): tick the categories
  after which an extra gap of one column slot, or a break symbol on the baseline, goes.
- **Axis break** (stacked, clustered): compress the value range *from*–*to* so a few very tall columns
  don't flatten the rest; *Straight* or *Wiggle*. No segment may end inside the range.
- **Secondary axis** (combination): tick the lines read against a second value axis on the right.

Decorations never cover a label: an arrow, value line label or indicator that finds no free room is
left out and listed under hidden labels.
- **Total columns** (waterfall): which categories stand on the baseline.

## Colours and your deck's theme

Charts use your presentation's theme colours by default: series in Accent 1, 2, 3 … in turn, text in
the theme's text colour, the thin separators between segments in its background colour. A chart
inserted into a client's deck therefore takes the client's colours, and changing the theme (Design ›
Variants › Colors, or a new template) recolours the chart along with PowerPoint's own shapes —
nothing to update. Labels on a dark segment turn white; the add-in decides that with the slide's
actual theme colours.

The greys of axes and secondary text and the green/red of waterfall increases and decreases are fixed
colours that do not change with the theme. A series colour you pick from **Theme colours** follows the
theme too; one you pick as a hex colour stays exactly that colour.

Decks written by the command line come with the open-charts theme (navy to grey, near-black text).
A `style` in the deck can change that theme, the palette, fonts and default number formats for every
chart — see [DATA-FORMAT.md › Style](DATA-FORMAT.md#style).

## Fonts

Charts can use any font. open-charts measures every label in the font PowerPoint will draw it in, so
labels fit and line up in Arial, Calibri, Aptos, Georgia, Helvetica Neue, Segoe UI, Times New Roman,
Verdana or a corporate font alike. Set the font in the editor's style, in a deck's or a style file's
`style.font`, or on the command line:

```sh
open-charts examples/deck.json --font Georgia -o out/deck
```

**In PowerPoint**, a chart you insert takes the deck's body font (the theme's *Body* font, e.g. Aptos or
your corporate font) and follows it if the theme changes; pick another font for the chart or one element
any time.

**Theme fonts.** `+mn-lt` (body) and `+mj-lt` (headings) stand for the presentation's theme fonts,
like PowerPoint's "Body" and "Headings" entries. The add-in reads them from the deck it draws in;
the command line and `fill` take them from the template's theme or from `style.themeFonts`.

**Where fonts are found.** Fonts installed for the computer or for you, and Microsoft Office's own
fonts, which PowerPoint uses but the system does not list:

| | Folders searched |
|---|---|
| Mac | `/System/Library/Fonts` (and `Supplemental`), `/Library/Fonts`, `~/Library/Fonts`, PowerPoint's own fonts (`Microsoft PowerPoint.app/Contents/Resources/DFonts`: Calibri, Cambria, Aptos, …), Office's downloaded cloud fonts (`~/Library/Group Containers/UBF8T346G9.Office/FontCache`) |
| Windows | `C:\Windows\Fonts`, fonts installed for you only (`%LOCALAPPDATA%\Microsoft\Windows\Fonts`), Office's downloaded cloud fonts (`%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts`) |
| Linux | `/usr/share/fonts`, `/usr/local/share/fonts`, `~/.local/share/fonts`, `~/.fonts` |

TrueType and OpenType fonts (`.ttf`, `.otf`), collections (`.ttc`, `.otc`) and variable fonts all
work; a family without a bold file is made bold from its weight axis, or measured slightly wider.
open-charts reads the fonts' names once and keeps an index in its cache folder (`~/.cache/open-charts`),
refreshed when a font folder changes; set `OPEN_CHARTS_FONT_DIRS` to search other folders instead.
Office downloads a cloud font (such as Aptos Display) the first time a deck uses it; until then it is
missing on that computer. macOS may keep other apps out of Office's folder; then cloud fonts are
missing too.

**A missing font.** If a font is not installed, text is measured with a font of exactly the same
widths when there is one (Arial ↔ Liberation Sans ↔ Helvetica, Times New Roman ↔ Liberation Serif,
Calibri ↔ Carlito, Cambria ↔ Caladea, Courier New ↔ Liberation Mono), else with a wider stand-in
that leaves labels room to spare. You are told once: the command line and `fill` print
`warning: font "…" is not installed here`, and the add-in shows the message the first time on that
computer. Install the font (or open the deck once in PowerPoint so Office downloads it) and update
the chart for an exact layout.

**Office.js pane.** The pane measures Arial with its own copy of Liberation Sans. Other fonts are
measured exactly when the pane is served by `npm run officejs:serve` (it reads the fonts on your
computer); otherwise with the browser's copy of the font when it has one (exact widths, estimated
line spacing), else as above. Office's private fonts (Calibri, Aptos) are usually not visible to the
browser.

## A client's look: style files

A **style file** is one JSON file that reproduces a client deck's look for every chart type: fonts and
sizes, the series palette, text and axis colours, number defaults, and the formatting of kinds of
elements, such as dotted gridlines, a darker axis, totals in the brand colour or decoration arrows in
an accent. Give it to the command line with `--style`:

```sh
open-charts examples/deck.json --style examples/styles/ember.json -o out/deck
open-charts fill template.pptx data.json --style brand.json -o out.pptx
```

A style changes how charts look, never what they say: the same data gives the same bars, labels and
arrows in the same places, in other colours. Another font or font size can move labels, because
text is measured in that font. A deck's own `style` and a chart's own formatting override the style
file. `examples/styles/` has two styles: `ember.json` (Georgia, rust to sand) and `brand.json`
(theme accents, so charts follow the client template). The keys are listed in
[DATA-FORMAT.md › Style files](DATA-FORMAT.md#style-files).

## Fill a template

`open-charts fill` puts charts and text into **your own** PowerPoint template. It edits a copy of the
file and does not touch the masters, layouts, pictures or any slide that has nothing to fill.

**1. Mark the template in PowerPoint.**

- **Charts.** Draw a rectangle (or use a content placeholder) where the chart goes. Open the
  Selection Pane (Home › Arrange › Selection Pane) and rename the shape `chart:revenue`. Instead of
  renaming it, you can type `{{chart:revenue}}` as its only text. The chart is laid out to fill
  that shape's frame. A placeholder that you never resized takes its frame from the layout.
- **Text.** Type `{{client}}`, `{{date}}` and so on anywhere: titles, text boxes, table cells and
  speaker notes. A field keeps the formatting of the text it sits in, even when PowerPoint has
  split it into several pieces internally.

**2. Write the data:** values by name, plus an optional list of slides.

```jsonc
{
  "values": {
    "client": "Northwind Traders",
    "date": "October 2026",
    "revenue": { "type": "stacked", "categories": ["2024", "2025"], "series": [{ "name": "Sales", "values": [80, 95] }] }
  },
  "slides": [
    { "slide": 1 },
    { "slide": 3 },
    { "slide": 5, "each": [ { "region": "EMEA", "sales": { "type": "clustered", "...": "..." } },
                           { "region": "Americas", "sales": { "type": "clustered", "...": "..." } } ] }
  ]
}
```

Without `slides`, every template slide is filled once, in template order. With `slides` you choose
the slides and their order. Slides you leave out are dropped, and `each` repeats a slide once per item
(here, one regional slide per region). Sections follow their slides. If the new order would split a
section, the sections are removed and you get a warning.

**3. Check it, then fill it.**

```sh
open-charts fill template.pptx data.json --dry-run       # what would be filled, where, and what is missing
open-charts fill template.pptx data.json -o out.pptx --strict
open-charts fill template.pptx data.json --each clients.json -o "out/{client}.pptx"   # one deck per client
```

The report lists each output slide with its charts (frame and status) and its fields, followed by any
value that no slot used. `--strict` exits with code 1 if a value is missing, a chart is invalid or a label
had to be hidden.

Charts take the template's theme. Series use Accent 1, 2, 3 … of the template's colour scheme, and the
text uses its body font. You can edit each chart with the add-in like any other. Running `fill` again
on a filled deck finds the charts by name and replaces them in place. A full worked example, with
an 8-slide template, is in `examples/template/` (`template.pptx`, `data.json`, `clients.json`).

## Two charts on a slide, same scale

Consultants often put two charts side by side — a revenue bridge next to the revenue history, two
regions — and the reader compares column heights across them. That only works when both use the
same scale. Give the charts the same **scale group** (`"scaleGroup": "revenue"` in their options,
see [DATA-FORMAT.md › scaleGroup](DATA-FORMAT.md#scalegroup)): the chart with the largest values for
its size keeps its scale, and the others draw one unit exactly as long, leaving room above their
columns. Stacked, clustered, waterfall, line, area and combination charts can share a scale.

In PowerPoint, select the charts (click one, Shift-click the others) and click **Same scale** in the
**Chart tools** group: they get one scale group (stored in each chart's data, so it travels with the
deck) and are laid out again at once. If one of them already belongs to a group, the others join
that group. **Remove from same scale** takes the selected charts out; the charts left in the group
are laid out again at their own common scale.

The group stays consistent: **Update chart**, **Apply** in the editor, the Chart tools commands and
resizing one chart of the group re-lay out the other charts of the group on that slide as well. From the command
line, a slide can hold several charts (`"charts": [...]`); without explicit frames they share the
chart area side by side.

## Undo and the right-click menu on a Mac

**Undo.** Each open-charts action is one step of PowerPoint's undo (⌘Z): inserting a chart, Update
chart, Apply from the datasheet, a change in the element toolbar, a direction button and every Chart
tools command (Same scale re-draws several charts and is still one step). Measured in PowerPoint for
Mac 16 with `npm run live -- chart-tools`: one ⌘Z each takes the slide back to exactly where it was.
Resizing is two steps: your drag, then the re-layout open-charts does when you click elsewhere. The
first ⌘Z gives you back the stretched drawing, the second the original size; after one ⌘Z, clicking
elsewhere lays the stretched chart out again.

On Windows the add-in also starts a new undo step before each action (`StartNewUndoEntry`, which
PowerPoint for Windows needs so a macro's changes are not merged into your last edit), so Ctrl+Z
undoes the action and nothing before it. That path has not been measured in PowerPoint for Windows
yet.

**Right-click menu.** The add-in declares its right-click items the standard way for Office add-ins
(`contextMenus` in its ribbon definition). PowerPoint for Mac does not show them: in our check it
never even asked the add-in whether to show them (its right-click menu is PowerPoint's own and takes
no add-in items). On a Mac, use the same commands on the ribbon: **Chart tools** and **Charts**.
PowerPoint for Windows shows the items on charts and their pieces.

## Agenda

An agenda gives the deck a table of contents, chapters and subchapters. Each chapter starts with an
**agenda slide** that shows the whole agenda with that chapter highlighted. When you add, remove,
reorder or rename chapters, every agenda slide follows. The agenda is ordinary PowerPoint text boxes
(one per chapter heading, with bullets and text levels) in a group, so the deck looks the same without
open-charts.

The **Agenda** group on the open-charts tab:

| Control | What it does |
|---|---|
| **Chapter** | Inserts an agenda slide after the current slide: a new chapter. Its heading reads "Chapter" and is selected, so you can type the name straight away. With a heading selected, the new chapter goes after that chapter (after its slides). The slide is titled "Agenda"; change it if you like. |
| **Table of Contents** | Inserts a slide listing the top-level chapters without highlighting, before the first chapter slide (after the current slide if there are no chapters yet). It fills in as you add chapters. |
| **Update agenda** | Brings every agenda slide up to date. Most changes do this by themselves when you click elsewhere (see below). |
| **Move up** / **Move down** | Moves the selected chapter (a selected heading, else the chapter of the current slide) one place; its slides move along: everything from its agenda slide up to the next agenda slide. |
| **Demote** / **Promote** | Makes the chapter a subchapter, or one level less deep. |
| **More › Split Agenda** | The current slide starts a new, independent agenda, e.g. the backup. A grey marker sits just above the slide (never shown in the slide show). |
| **More › Join Agendas** | Removes the split on the current slide. |
| **More › Remove Hidden Agenda Slides** | Deletes the agenda slides that are hidden (see below); their chapters stay in the agenda. |
| **More › Insert Hidden Agenda Slides** | Puts those slides back, before their first subchapter's slide. |
| **More › Agenda Placeholder** | In Slide Master view: marks where the agenda goes on a layout named `open-charts agenda` (see below). |

**Renaming and nesting.** Click into a heading on *any* agenda slide and type: when you click
elsewhere, every agenda slide shows the new name. A heading's text level is its chapter level:
PowerPoint's **Increase/Decrease List Level** buttons (or Shift+Tab / Tab at the start of the heading)
make it a subchapter or bring it back up, and every agenda slide follows. A new chapter takes the
level of the chapter before it.

**What updates by itself.** When you click elsewhere, open-charts notices agenda slides added,
duplicated, moved or deleted (also in the Slide Sorter) and headings you renamed or nested on the
current slide, and updates every agenda slide. A duplicated agenda slide becomes a new chapter of the
same name; a deleted one removes its chapter. A changed fill, outline, colour, weight or size of a
heading is picked up by **Update agenda** (see Styles).

**Which chapters show.** Every agenda slide shows the top-level chapters. Subchapters show on the
slide of their parent chapter and on the slides of the chapters at their own level (their siblings
and, deeper down, the path to the current chapter). On a table of contents only top-level chapters
show.

**Hidden parent slides.** If there are no slides between a chapter and its first subchapter (the
chapter has no content of its own), the chapter's agenda slide is set to **Hide Slide**, so the slide
show does not stop on two agenda slides in a row. Add a slide between them and it is shown again.
Slide numbers in the slide show skip the hidden slide; **Remove Hidden Agenda Slides** deletes such
slides (the chapter stays in the agenda) and **Insert Hidden Agenda Slides** brings them back.

**Several agendas.** Go to the first slide of the second part (typically a slide titled "Backup") and
choose **Split Agenda**. Chapters after the split are not on the agenda slides before it and vice
versa, and moving the last chapter before the split does not take the backup along. Put the split on
that title slide rather than on the backup's first agenda slide: otherwise it moves and duplicates
with that chapter.

**Position and size.** The agenda fills the slide's body area below the title (or the layout's
agenda placeholder, see below), headings at 24 pt, subchapters slightly smaller. With many chapters
the font gets smaller until the agenda fits. Every agenda slide of a section puts it in the same
place. Move it on one agenda slide and the others of that section follow on the next update; new
agenda slides copy the closest agenda slide before them (its layout, master and shapes). A heading
box is never narrower than its text.

**Styles.** Each heading has one of twelve states, and every heading in the same state looks the
same:

| State | When |
|---|---|
| `current-top` | the current chapter, at top level |
| `top-contains-current` | a top-level chapter that contains the current chapter |
| `top` | a top-level chapter that does not contain the current chapter |
| `current-sub` | the current chapter, not at top level |
| `child-of-current-top` | an immediate child of the current chapter, which is top-level |
| `in-subtree` | neither it nor the current chapter is top-level; not on the path to the current chapter, but a sibling of a chapter on that path, or an immediate child of the current chapter |
| `under-current-top` | below the current top-level chapter, not an immediate child of it |
| `same-top` | below the same top-level chapter as the current chapter |
| `other` | any other chapter |
| `toc-top` | in a table of contents, top level |
| `toc-sub` | in a table of contents, not top level |
| `path` | neither it nor the current chapter is top-level, and it is on the path from the top-level chapter to the current chapter |

To restyle a state, format one heading in that state (fill, outline, font colour, bold, size) and
click **Update agenda**: every heading in the same state, on every agenda slide, takes the new look.
The defaults are the *flat* preset: grey bars, the current chapter white on accent 1. From the
command line, `"agenda": {"preset": "plain" | "boxed", "numbering": true, "styles": {...}}` sets the
preset, numbered headings and per-state styles ([DATA-FORMAT.md › Agenda slides](DATA-FORMAT.md#agenda-slides)).

**Default agenda layout.** To give agenda slides your template's look, add a custom layout in Slide
Master view, name it exactly `open-charts agenda`, put on it only what every agenda slide should
show, and click **More › Agenda Placeholder** to mark where the agenda goes (move and size the
invisible rectangle it adds). **Chapter** then uses that layout and position for the first agenda
slide; later ones copy the agenda slide before them.

**From the command line.** In a deck JSON, `{"chapter": "Market"}` is an agenda slide (with
`"level": 2` for a subchapter), `{"toc": true}` a table of contents and `"split": true` starts a new
agenda. `examples/agenda.json` is a complete example; the .pptx it writes behaves in the add-in like
one built there.

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
| Bars and profiles (direction ◀ or ▶) | 200 | 120, or 30 + 12 per category |
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

**Check for updates.** Click **Help › Check for updates**. It asks GitHub for the latest release
(only when you click) and says whether a newer open-charts is available and where to download it.
On a computer managed by IT, updates come from your IT department instead.

**Report a problem.** Click **Help › Report a problem** on the open-charts tab. It saves a zip with
the add-in's log, its settings and the version numbers (open-charts, PowerPoint, your system) and
shows it in Finder or Explorer. If a chart is selected, it asks whether to add that chart's data;
nothing else from your slides goes in. Attach the zip to a new issue on GitHub
(https://github.com/sunajeb/open-charts/issues) and say what you did. From Terminal:
`open-charts addin report` (add `--spec chart.json` to include a chart's data). The zip is in
`~/Library/Application Support/open-charts/reports` on the Mac, `%LOCALAPPDATA%\open-charts\reports` on
Windows.

**The open-charts tab is not on the ribbon.** Run the install again (see the README). If you
installed with `--no-register`, load it yourself: in PowerPoint, *Tools › PowerPoint Add-ins…*,
click **+** and pick `~/Library/Application Support/Microsoft/open-charts/open-charts.ppam`. Make
sure it is ticked in that list.

**PowerPoint asks about macros.** The add-in is a small macro file whose only job is to pass your
click to the chart engine and draw the shapes it returns. Click **Enable Macros** for
`open-charts.ppam`. If macros were disabled and the buttons do nothing, open *PowerPoint ›
Preferences › Security & Privacy*, set *Macro Security* to warn you (or to enable macros), restart
PowerPoint and click **Enable Macros** when asked.

**"open-charts is not responding".** The add-in calls Node through a small script that
the install placed in PowerPoint's scripts folder. Run the install again from Terminal and restart
PowerPoint:

```sh
"$HOME/Library/Application Support/open-charts/bin/open-charts" addin install
```

(`open-charts addin install` if the install linked the command into your PATH; the message in
PowerPoint shows the same line.) This also fixes **Node not found** after you upgraded or moved Node (for example with
nvm): the install records the path of the Node that ran it. Check `node --version` is 22.6 or newer
first.

**"Parts of open-charts do not match".** You updated open-charts but not the installed copy.
Run the command the message shows (the install command above).

**"Select an open-charts chart first".** Click the chart once (select the whole group, not a shape
inside it), then choose Edit data or Update chart.

**"This chart was ungrouped, so it can't be edited as a chart".** The chart's data lives on its
group; ungrouping removes it, and regrouping does not bring it back. Undo (⌘Z) the ungroup, or insert
the chart again (the loose shapes stay as they are).

**"Click a slide (Normal view) first".** Charts are inserted on the current slide in Normal view, not
in Slide Sorter or the outline.

**After Apply, the editor says "Not updated".** The editor only says **Updated on slide** once
PowerPoint has redrawn the chart; otherwise it says why, in a bar under its toolbar:

- *PowerPoint did not respond. If macOS is asking whether PowerPoint may control PowerPoint, click
  Allow…*: see the next item, then click **Apply** again.
- *macOS does not allow PowerPoint to be controlled*: you clicked **Don't Allow** earlier; turn it on
  as the next item says, then **Apply** again.
- *PowerPoint isn't running*: open the presentation, click the chart once, then **Apply** again.
- *PowerPoint could not find the chart*: you deleted it, or moved it to another slide. Click it once,
  then **Apply** again (or click **Update chart**: it picks up the edit).

**"open-charts is still updating a chart".** Wait a moment and click again.

**macOS asks whether "Microsoft PowerPoint" may control "Microsoft PowerPoint".** Click **Allow**: that
is how Apply in the browser updates the slide. While the question is unanswered, the editor shows
"Waiting for PowerPoint…" and, after 20 seconds, that the slide was not updated. If you clicked **Don't Allow**, turn it back on in System Settings › Privacy & Security ›
Automation (Microsoft PowerPoint › Microsoft PowerPoint).

**The chart shows "Chart too small".** It is below its minimum size (see [Resizing](#resizing)). Drag
it bigger; it is laid out again when you click elsewhere.

**A label is missing.** See [Hidden labels and why](#hidden-labels-and-why).

**A cell is red.** The text is not what this chart type expects (a number, `x; y`, `x; y; size` or a
date). Hover or focus the cell to see the message.

**My manual formatting disappeared.** Update chart, Apply and resizing redraw the chart from its
data. Make changes through the datasheet and options instead of editing shapes inside the group.

### Windows

**The open-charts tab is not on the ribbon.** Quit PowerPoint and run `install.ps1` again. It prints
what it did. Then, in PowerPoint, open *File › Options › Add-ins*, choose *Manage: PowerPoint Add-ins
› Go…* and make sure **open-charts** is in the list and ticked. If it is not there, click
**Add New…** and pick `%APPDATA%\Microsoft\AddIns\open-charts.ppam`.

**PowerPoint shows "Microsoft has blocked macros" or a macro warning bar.** The file still carries
the "downloaded from the internet" mark. Run `install.ps1` again, which removes it, and restart
PowerPoint. If the macros stay disabled, open *File › Options › Trust Center › Trust Center
Settings… › Trusted Locations*. `%APPDATA%\Microsoft\AddIns\` should be listed (`install.ps1` adds
it when it is missing). If your organisation disables all macros, the add-in cannot run until IT
allows macros from trusted locations.

**"open-charts is not responding".** The add-in starts
`%LOCALAPPDATA%\open-charts\open-charts.exe` and talks to it on `127.0.0.1:4174`.
- *…open-charts.exe is missing*: run `install.ps1` again.
- *…did not start within 20 seconds*: an antivirus scan or an application-control policy is holding
  it. Run `install.ps1` again; its last step starts the engine and says so if Windows blocks it. Then
  ask IT to allow that file.
- Something else on the PC uses port 4174: close that program.

**After Apply, the editor says "Not updated".**
- *PowerPoint did not respond. Close any open dialog in PowerPoint…*: PowerPoint was showing a dialog
  or you were typing in a text box. Close the dialog or click outside the text, then **Apply** again.
- *PowerPoint isn't running* and *could not find the chart*: as on the Mac (above).

**Where is the log?** `%LOCALAPPDATA%\open-charts\addin.log`, with one line per insert, update,
re-layout and Apply. Include it when you report a problem.

**Uninstall.** `powershell -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\open-charts\uninstall.ps1"`.
Charts already on slides stay as they are.
