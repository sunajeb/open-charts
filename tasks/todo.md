# open-charts — plan

## Phase 1 — engine + PowerPoint output
- [x] Core types, scene graph, font-metric text measurement (Arial / Liberation Sans)
- [x] Nice ticks, number formatting, contrast text colour, label spreading
- [x] Stacked, 100%, clustered column charts with totals, CAGR + difference arrows, series labels, legend, value axis
- [x] Dependency-free OOXML writer (grouped native shapes, alt text) + SVG preview
- [x] CLI: deck.json → deck.pptx (+ --svg), layout quality gate (overlaps, out-of-frame)
- [ ] Behaviour spec from manual research (docs/behaviour-spec.md)
- [ ] Waterfall (subtotals, connectors, negatives)
- [ ] Mekko (variable widths, % and unit)
- [ ] Line chart (markers, end labels)
- [ ] Horizontal bar
- [ ] 100% rounding that sums to 100 (largest remainder)
- [ ] Tests + golden checks over all examples
- [ ] Web editor: datasheet + live preview + download .pptx

## Review
