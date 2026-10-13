# Changelog

## Unreleased

### Updates and fixes

- `open-charts addin update install`: downloads this platform's installer, checks its SHA-256 and the
  publisher's signature, and only then opens it; refused until open-charts' signing identity exists.
  An IT policy file (`policy.json`, machine-wide) turns the update check off.
- Installing over a running PowerPoint reloads the add-in: no restart needed.
- PowerPoint for Mac: label commands (Convert to plain text, Reset label position) see the selected
  label; PowerPoint no longer starts again by itself after quitting with a chart selected; an engine
  answer that PowerPoint occasionally drops is asked for again.

### Live tests

- `npm run live -- <scenario...>` (`tools/live/`) replaces the 10-minute `e2e:powerpoint` script and the
  separate live scripts: short independent scenarios (each under 2 minutes, a fresh presentation each)
  in one PowerPoint session, a timeout on every call (a blocking PowerPoint dialog is read, dismissed
  and reported), stop at the first failure unless `--keep-going`, one report in `out/live/`. Holds the
  PowerPoint lock, refuses to run beside open presentations unless `--allow-open-docs`, never quits a
  PowerPoint it did not start, reinstalls only with `--install`; `--list` and `--dry-run`.
  `npm run e2e:powerpoint` runs them all.

### Installers (no Node, no Terminal, no admin rights)

- Mac: `open-charts-<version>.pkg` (`npm run build:mac`), installed per user. It contains the engine
  as one universal program, the Edit data window, the add-in and *Uninstall open-charts* (in
  ~/Applications). It registers the add-in in PowerPoint without starting it.
  `npm run check:mac-pkg` installs it into a temporary home folder and removes it again.
- Windows: `open-charts-<version>.msi` (`npm run build:msi`), per user, with `/quiet` for a silent
  install. `open-charts-<version>-machine.msi` is for Intune; it sets up each user at sign-in.
- Signing hooks: Apple Developer ID plus notarization, and Azure Trusted Signing. Both are driven by
  environment variables and secrets, and builds are unsigned without them. A `release` workflow
  drafts a GitHub release from a version tag (docs/release.md).
- `open-charts addin update [on|off]` checks GitHub Releases for a newer version. The daily check
  is off by default and only reports what it finds.

### Any font

- Text is measured in the font it is drawn in: a chart's `style.font`, an element's own font, or a
  theme font (`+mn-lt` / `+mj-lt`) resolved to the deck's theme fonts (the add-in sends them;
  new style key `themeFonts`). Fonts are found among the computer's and the user's fonts, PowerPoint's
  own (Calibri, Aptos, …) and Office's downloaded cloud fonts, on Mac, Windows and Linux; TrueType,
  OpenType/CFF, collections and variable fonts. A missing font uses a metric-compatible substitute
  (Carlito for Calibri, …) or a wider stand-in, and is reported once (USER-GUIDE "Fonts").
- `--font Georgia` sets the charts' font on the command line (`--font file.ttf` still works).
- CLI decks draw a chart's own `style.font` (it was drawn in the deck's font).
- Office.js pane: other fonts are measured exactly through `officejs:serve`, else on a canvas.
- `npm run e2e:fonts`: live check of label widths in PowerPoint (owner-approved runs only).

### Drag handles (PowerPoint for Mac)

- A selected chart shows handles you can drag: column gap width, the top of the value axis (new
  option `axisMax`, headroom only), level-arrow position, and a selected label (pinned where you drop
  it). Live preview while dragging, one update on release, double-click to reset. A transparent,
  click-through window of the pop-up helper; no extra permissions (USER-GUIDE "Dragging handles").

### Templates and style files

- `open-charts fill template.pptx data.json -o out.pptx` fills your own template. Shapes named
  `chart:<name>`, or holding only `{{chart:<name>}}`, become tagged, editable charts laid out in
  that shape's frame; a placeholder without its own frame uses the layout's. `{{name}}` fields in
  text, tables and notes keep their formatting, even when split across runs. `slides` picks,
  reorders and repeats slides (`each`), and sections follow. Charts take the template's theme
  colours and body font. Untouched parts stay byte-identical. Also `--dry-run`, `--strict`,
  `--each items.json -o "out/{field}.pptx"`, `--report` and `--style`. Example in
  `examples/template/`; live check: `npm run live -- fill`.
- Style files (`--style style.json` on build and fill, `schema/style.schema.json`): a client's
  fonts, palette, colours, number defaults and element defaults (`style.elements`, the overrides
  vocabulary) for every chart type, with per-type overrides. Checked as untrusted input. Examples:
  `examples/styles/ember.json` and `examples/styles/brand.json`.

## 0.2.0 (2026-10-13)

The first build meant for people who make slides: charts made in PowerPoint, edited through a
datasheet, laid out by the engine. No AI and no other apps yet.

### PowerPoint add-in (Mac)

- `open-charts addin install` builds `open-charts.ppam`, copies the engine, installs the AppleScript
  bridge, installs the `open-charts` command (linked into your PATH when possible) and registers
  the add-in; `addin uninstall`, `addin build`, `addin types`. Error messages show a reinstall
  command that works whether or not `open-charts` is on your PATH.
- **open-charts** ribbon tab: **Insert chart** menu with every chart type (plus *100% with % axis* and
  *Step line*), **Edit data** and **Update chart**. A chart goes into an empty content placeholder
  when the slide has one.
- **Edit data** opens the datasheet in the browser; **Apply to slide** (⌘↵) redraws the chart in
  place, keeping its position and stacking order.
- Resizing a chart lays it out again at the new size when you click elsewhere; text is never
  stretched.

### Chart types

17 types: stacked, clustered and 100% columns (also as horizontal bars), waterfall, combination,
line (and step line), stacked and 100% area, Mekko, Mekko XY, scatter, bubble, pie, doughnut, Harvey
balls, table and Gantt.

### Layout engine

- Labels never overlap and never leave the chart; a label that cannot fit is hidden and reported with
  the reason.
- Totals, CAGR arrows (years read from labels such as `FY24`, `Q1 2023`, `2025E`) and difference
  arrows; waterfall totals by `"e"` cells or `totalsAt`.
- Shares and segment labels rounded so they add up exactly (largest remainder).
- Numbers: units as prefix or suffix, decimals, k / m / bn magnitudes, Excel number formats, five
  locales (en-US, en-GB, en-IN, de-DE, fr-FR).
- Holds up at any size: category names wrap and shrink (not below 8 pt), crowded time axes thin out,
  series names move to a legend, and below a minimum size the chart shows a "Chart too small"
  placeholder that says how large to make it.
- Text colour on fills chosen by contrast; nice axis ranges.

### Editor (`open-charts serve`)

- Preview on top, Excel-like datasheet below, options in a sidebar grouped by Chart, Numbers, Axes &
  layout, Labels & legend and Decorations.
- Chart type gallery; range selection, copy/cut/paste of Excel ranges, undo/redo, context menu,
  transpose, series colours, per-type hints; invalid cells marked in place; hidden labels listed and
  numbered on the preview.
- Deck mode with slide titles, JSON import/export and .pptx download; session mode for the add-in's
  Edit data window.
- A control for every option (slice labels, bubble size legend, Gantt today line, shaded periods and
  phase brackets, CAGR years included); untouched controls show the engine's default for the chart
  type; switching type drops the options the new type does not use.

### Command line and file format

- `open-charts <deck.json>` writes a .pptx (`--svg`, `--strict`, `--report`, `--font`);
  `inspect` lists the charts in a deck; `import` recovers the deck JSON from a .pptx.
- Spec version 1 with JSON Schemas (`schema/`), path-addressed validation errors and `migrate()` for
  older specs. One table says which chart types read each option; an option on any other type is an
  error (`/options/cagr: not used by pie charts`), and `migrate()` removes such options from
  unversioned specs with a warning.
- Tag contract v1 on every chart group (`OC_V`, `OC_ENGINE`, `OC_SPEC`, `OC_FRAME`, `OC_FONT`), so
  charts stay editable in any deck.

### Documentation and tests

- New user guide (`docs/USER-GUIDE.md`) and data format reference (`docs/DATA-FORMAT.md`); README
  rewritten for consultants and analysts.
- `test/docs.test.ts` validates every JSON example in the data format reference.
- Unit tests, editor end-to-end tests, golden images per platform, a PowerPoint end-to-end script
  and performance budgets (`npm run bench`).
