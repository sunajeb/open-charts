# open-charts — live checklist

Plan of record: docs/MASTER-PLAN.md (shareable page: https://claude.ai/artifact/QZPFXp9NmZSUwYaFBhujWH)

## Done (9 Oct 2026)
- [x] Engine core: scene graph, font-metric measurement (no kerning, Arial 1.149 line height), nice ticks, exact-100% rounding, label spreading, segment/box clipping
- [x] Chart types: stacked, 100%, clustered, line, waterfall, Mekko
- [x] Decorations: totals, CAGR (+years), difference arrows, series labels, legend, value axis
- [x] Zero-dependency OOXML writer (grouped shapes, alt text) + SVG preview
- [x] CLI and local editor (datasheet, Excel paste, live preview, pptx download)
- [x] 20 tests passing (`npm test`)
- [x] Plan: docs/plan/01–04, docs/native-windows-design.md, docs/MASTER-PLAN.md

## Next: P1 first 30 days (see MASTER-PLAN §6)
- [x] Public repo + CI on PRs (Node 22, 24): https://github.com/sunajeb/open-charts
- [ ] Week 1: fix B5, B10, B13; CI on PRs; line height from font hhea/OS2; behaviour-spec.md
- [ ] Week 2: numbers/locale service (B14); start spec v1 + JSON Schema + migration (B15)
- [ ] Week 3: finish spec v1; scene ids/roles; writer emits tag contract v1; `inspect`
- [ ] Week 4: reservation ledger; plugin pipeline; label solver v1; LibreOffice golden CI; 10 bench fixtures

## Open decisions (owner)
- [ ] O1 staffing: hybrid recommended
- [ ] O2 native Windows: gate at M7 recommended
- [ ] O3 spec freeze: P1 exit recommended

## Review
- Session result: six chart types at near-consulting quality, rendered and visually checked through LibreOffice; full five-phase plan.
- Process miss: four agents sat idle for ~2h on the folder-trust prompt because they were spawned from a new, untrusted folder (see tasks/lessons.md).
