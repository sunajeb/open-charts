# open-charts — plan to production (v1.0)

Goal: consultants at boutique firms (IT-managed Windows/Mac laptops, think-cell installed) and
prosumers can install open-charts easily, and it works reliably with their decks.
Feature set: frozen at today's (2026-10-13). Earlier checklists are in git history.

Rules for this plan
- Offline checks (unit, browser, goldens, bench, quality) run on every change; they never touch PowerPoint.
- Live PowerPoint runs happen only in a window the owner approves, in short scenarios, never while the owner works.
- One task = one owner (me or one agent), done = tests green + evidence in docs/ACCEPTANCE.md.

## A. Hygiene and safety (no PowerPoint needed)
- [x] A1 Remove analysis notes about the reference product from the repo; keep only public-convention
      and public-manual wording
- [ ] A2 Git history still contains the removed notes — decide: rewrite history + force-push (owner decision, Q1)
- [x] A3 Commit the click-handling fix (user click during background selection work runs at once)
- [x] A4 Security: local engine server accepts only requests carrying a per-install secret token and a
      loopback Host/Origin; editor/pop-up/toolbar/VBA/Office.js pass the token; tests for rejection
- [x] A5 Remove developer-only leftovers from the product (test macros hidden from the ribbon, debug logs off by default)

## B. Any font
- [x] B1 Font resolver: theme fonts (+mn-lt/+mj-lt) → real family from the deck; family → font files on
      Mac and Windows, incl. per-user fonts and Office's downloaded cloud fonts (Aptos etc.)
- [x] B2 Measurer for TTF/OTF/TTC/variable fonts, regular/bold/italic, cached; safe fallback (+margin,
      reported) when the font is missing
- [x] B3 Add-in sends the deck/theme font; engine measures with it; Office.js measures with the browser
- [ ] B4 Font test set (Arial, Calibri, Aptos, Georgia, Helvetica Neue, Segoe UI, Times + owner's corporate
      font): unit + goldens; one live scenario comparing PowerPoint's rendering
      — offline done (test/fonts.test.ts, goldens fonts-*, e2e); live `npm run e2e:fonts` 2026-10-13:
      8 fonts incl. theme fonts, every label within 0.63 % of the engine's width, none overflows.
      Owner's corporate font and Segoe UI (not on this Mac) still to add

## C. Install, sign, deploy, update
- [x] C1 Mac: bundled engine binary (no Node; arm64 + x64) + pop-up helper + .ppam in one per-user .pkg; uninstaller
      (npm run build:mac + check:mac-pkg; first real Installer.app run and PowerPoint registration: live window)
- [x] C2 Windows: per-user installer (no admin) with the exe + .ppam; silent mode; enterprise MSI for Intune
      (npm run build:msi, WiX v5; CI builds both MSIs and installs/uninstalls them silently: green 2026-10-13)
- [ ] C3 Signing pipeline in CI: Apple Developer ID + notarization (Mac), Azure Trusted Signing (Windows exe,
      installer, VBA project) — needs owner accounts (Q3)
- [x] C4 Office.js production: static https hosting, production manifest, M365 admin-center deployment guide;
      AppSource submission later (`build:officejs -- --base https://… --out dir`, docs/OFFICEJS-HOSTING.md;
      a real host and admin-center upload need the owner's domain/tenant)
- [ ] C5 Updates: version check + update download (signed packages only)
      (done: client check behind `addin update on`; `addin update install` downloads, verifies SHA-256 +
      signature, opens; IT policy.json forces it off; Help › Check for updates in the add-in (live PASS).
      Left: in-app install once the signing identity exists (C3))
- [x] C6 IT deployment guide: Intune / Jamf / M365 admin center, coexistence with think-cell, macro trust settings
      (docs/IT-DEPLOYMENT.md; Office policy names and Intune/Jamf details to confirm in the owner's pilot)

## D. Verification (live runs in approved windows only)
- [x] D1 Split tools/e2e-powerpoint.ts into short scenarios (insert-all, edit/apply, resize, chart tools,
      agenda, handles, restart, popup); each < 2 min, fail fast, one report; full set only per release
      (`npm run live -- <scenario...>`, tools/live/; unit-tested and dry-run only: first live run is D2)
- [x] D2 Mac live run of all scenarios: all 12 PASS (2026-10-13; last full run 11/12, the one failure
      was a lost engine answer, fixed and insert-all rerun PASS). Bugs the runs found and fixed:
      Mac VBA `Not` on a PowerPoint True of 1 (label commands saw no label), PowerPoint relaunched by
      a re-measure as it quit, AppleScriptTask returning "" now and then, updates needing a restart
- [ ] D3 Windows live run on the owner's think-cell laptop (script → zip) and fixes
- [ ] D4 Office.js live self-test (Mac) + Windows + PowerPoint on the web
- [ ] D5 Test matrix: older PowerPoint 16.x, Intel Mac, non-English Office, two monitors / scaling,
      a 100-slide deck with 50 charts
      (done: big deck offline perf + live `big-deck` PASS 2026-10-13: opens clean, Update 0.3-0.5 s,
      50/50 charts keep their data. Needs other hardware/Office: older 16.x, Intel, non-English, monitors)
- [x] D6 Robustness: decks from older versions, charts edited by hand without the add-in, copy/paste across decks
      (test/robustness.test.ts, all 19 types; fixes: invalid native edits no longer break Update, damaged
      or newer charts say so plainly, v0 charts upgrade on any command, one damaged chart no longer stops
      a deck import)

## E. Supportability and first-run experience
- [x] E1 "Report a problem": zips add-in log, versions and the chart spec (no slide content unless the user agrees)
- [x] E2 First run: short welcome, sample deck, help links
- [x] E3 Review every user-facing message (plain, actionable): add-in, engine, editor, task pane, installer, bridge

## F. Release 1.0
- [x] F1 Freeze the data format at 1.0 (owner decision, Q5); version, changelog
      (Q5: relative difference arrow by default, format frozen; open-charts 1.0.0, CHANGELOG 1.0.0)
- [ ] F2 Landing page + download for prosumers; licence and legal review (Q2); name/trademark check (Q6)
      (published 2026-10-13 at https://sunajeb.github.io/open-charts/ from docs/site (pages workflow),
      owner's call: name kept, lawyer review later. Left: legal review (Q2))

## Owner questions (open)
- Q1 Rewrite git history to remove the old analysis docs (force-push; anyone who cloned keeps old copies)?
- Q2 Get a lawyer's view before public/commercial distribution?
- Q3 (answered 2026-10-13: later; the pilot ships unsigned) Apple Developer and Azure Trusted Signing
- Q4 Which fonts do target decks use? Share one typical client template if possible.
- Q5 (answered 2026-10-13: relative % default, freeze now)
- Q6 (answered 2026-10-13: keep "open-charts", publish now)
- Q7 When can the Windows laptop run the test script; can it install unsigned software today?

## Questions for the reference session
- How is it installed on Windows in companies (per-user vs per-machine MSI, COM add-in + .ppam, LoadBehavior,
  signed VBA/trusted publisher)? How does IT typically deploy it (Intune/SCCM, Jamf)?
- How does it measure text: system font APIs (DirectWrite/CoreText) at run time, its own metrics, or both?
  How does it handle cloud fonts like Aptos that are not installed as files?
- How does it update itself on Windows (service, scheduled task, in-app)?
