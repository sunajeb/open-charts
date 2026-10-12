# open-charts

Consulting-quality charts for PowerPoint, open source. You give it data, it works out the layout (axis scale, label placement, totals, growth arrows) and writes native, editable PowerPoint shapes.

- **Native output.** Each chart is one grouped object of ordinary shapes, so the deck opens and edits in PowerPoint, Keynote and LibreOffice with nothing installed.
- **Layout that checks itself.** Every chart goes through a quality check: labels never overlap and nothing spills outside the chart area. A label that can't fit is hidden and reported, never drawn on top of something else.
- **No build step and no dependencies to install.** Runs on Node 22.6+ directly from TypeScript.

## Chart types

| Type | What it's for |
|---|---|
| `stacked` | Composition over time, with totals and CAGR / difference arrows |
| `percent` | 100% mix; labels are rounded so every column reads exactly 100% |
| `clustered` | Side-by-side comparison, with legend and optional value axis |
| `line` | Trends; labels placed clear of lines, markers and each other |
| `waterfall` | Bridges (e.g. EBITDA walk) with subtotals and connectors |
| `mekko` | Market maps: column width = segment size |

## Use it

**Editor** (datasheet, paste from Excel, live preview, download .pptx):

```sh
node --experimental-strip-types --no-warnings src/cli.ts serve
# open http://localhost:4173
```

**Command line** (JSON in, .pptx out):

```sh
node --experimental-strip-types --no-warnings src/cli.ts examples/deck.json -o out/deck --svg
```

**PowerPoint add-in** (Mac): an **open-charts** ribbon tab with every chart type, Edit data and Update chart:

```sh
node --experimental-strip-types --no-warnings src/cli.ts addin install
```

This builds `open-charts.ppam` (a ribbon plus one small VBA module), copies the engine to `~/Library/Application Support/open-charts`, installs the bridge script PowerPoint's VBA calls (`AppleScriptTask`) and registers the add-in. Charts are drawn as native, grouped PowerPoint shapes with the same `OC_*` tags the command line writes. `addin uninstall` removes it.

## Data format

```json
{
  "slides": [{
    "title": "Revenue grew 2.4x in four years",
    "chart": {
      "type": "stacked",
      "categories": ["2021", "2022", "2023", "2024"],
      "series": [
        { "name": "Enterprise", "values": [42, 58, 75, 96] },
        { "name": "SMB", "values": [18, 21, 22, 26] }
      ],
      "options": { "unit": "m", "cagr": { "from": 0, "to": 3 } }
    }
  }]
}
```

Options: `unit`, `decimals`, `totals`, `valueAxis`, `zeroBased`, `gap`, `seriesLabels` (`right` / `legend` / `none`), `cagr`, `difference`. See `src/core/types.ts` for the full list and the chart-specific options.

## How it works

```
chart spec (JSON) → layout engine (src/core) → scene (exact shapes) → PowerPoint writer / SVG preview
```

The layout engine is pure TypeScript with no Office code. It measures text with real font metrics (Arial, or the metric-compatible Liberation Sans), so what it plans matches what PowerPoint draws. Both renderers read the same scene, so the preview always matches the slide.

## Tests

```sh
npm test          # unit tests, zero dependencies
npm run e2e       # editor end-to-end in Chromium (npm install; npx playwright install chromium)
npm run goldens   # examples → .pptx → LibreOffice PDF → PNG, diffed against test/goldens/<platform>/
npm run e2e:powerpoint  # macOS + PowerPoint: insert/update every chart type through the add-in
```

`npm run goldens` needs `soffice` (LibreOffice) and `pdftoppm` (poppler). Baselines are per platform because fonts render differently; `-- --update` rewrites them after you have reviewed `out/goldens/actual/`, and failures write red-on-grey diffs to `out/goldens/diff/`. CI runs unit tests and E2E on every PR, goldens nightly.

## Licence and provenance

Apache-2.0. Built from public chart-design conventions, published algorithms (Heckbert and Talbot et al. for axis ticks, largest-remainder rounding, Liang–Barsky clipping), the open OOXML standard and study of how existing charting tools behave. All code is original; none is copied from a commercial product. `vendor/opentype.module.js` is opentype.js (MIT).
