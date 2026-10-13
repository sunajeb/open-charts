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
- [ ] A3 Commit the click-handling fix (user click during background selection work runs at once)
- [ ] A4 Security: local engine server accepts only requests carrying a per-install secret token and a
      loopback Host/Origin; editor/pop-up/toolbar/VBA/Office.js pass the token; tests for rejection
- [ ] A5 Remove developer-only leftovers from the product (test macros hidden from the ribbon, debug logs off by default)

## B. Any font
- [ ] B1 Font resolver: theme fonts (+mn-lt/+mj-lt) → real family from the deck; family → font files on
      Mac and Windows, incl. per-user fonts and Office's downloaded cloud fonts (Aptos etc.)
- [ ] B2 Measurer for TTF/OTF/TTC/variable fonts, regular/bold/italic, cached; safe fallback (+margin,
      reported) when the font is missing
- [ ] B3 Add-in sends the deck/theme font; engine measures with it; Office.js measures with the browser
- [ ] B4 Font test set (Arial, Calibri, Aptos, Georgia, Helvetica Neue, Segoe UI, Times + owner's corporate
      font): unit + goldens; one live scenario comparing PowerPoint's rendering

## C. Install, sign, deploy, update
- [ ] C1 Mac: bundled engine binary (no Node; arm64 + x64) + pop-up helper + .ppam in one per-user .pkg; uninstaller
- [ ] C2 Windows: per-user installer (no admin) with the exe + .ppam; silent mode; enterprise MSI for Intune
- [ ] C3 Signing pipeline in CI: Apple Developer ID + notarization (Mac), Azure Trusted Signing (Windows exe,
      installer, VBA project) — needs owner accounts (Q3)
- [ ] C4 Office.js production: static https hosting, production manifest, M365 admin-center deployment guide;
      AppSource submission later
- [ ] C5 Updates: version check + update download (signed packages only)
- [ ] C6 IT deployment guide: Intune / Jamf / M365 admin center, coexistence with think-cell, macro trust settings

## D. Verification (live runs in approved windows only)
- [x] D1 Split tools/e2e-powerpoint.ts into short scenarios (insert-all, edit/apply, resize, chart tools,
      agenda, handles, restart, popup); each < 2 min, fail fast, one report; full set only per release
      (`npm run live -- <scenario...>`, tools/live/; unit-tested and dry-run only: first live run is D2)
- [ ] D2 Mac live run of all scenarios (one approved window)
- [ ] D3 Windows live run on the owner's think-cell laptop (script → zip) and fixes
- [ ] D4 Office.js live self-test (Mac) + Windows + PowerPoint on the web
- [ ] D5 Test matrix: older PowerPoint 16.x, Intel Mac, non-English Office, two monitors / scaling,
      a 100-slide deck with 50 charts
- [ ] D6 Robustness: decks from older versions, charts edited by hand without the add-in, copy/paste across decks

## E. Supportability and first-run experience
- [ ] E1 "Report a problem": zips add-in log, versions and the chart spec (no slide content unless the user agrees)
- [ ] E2 First run: short welcome, sample deck, help links
- [ ] E3 Review every user-facing message (plain, actionable)

## F. Release 1.0
- [ ] F1 Freeze the data format at 1.0 (owner decision, Q5); version, changelog
- [ ] F2 Landing page + download for prosumers; licence and legal review (Q2); name/trademark check (Q6)

## Owner questions (open)
- Q1 Rewrite git history to remove the old analysis docs (force-push; anyone who cloned keeps old copies)?
- Q2 Get a lawyer's view before public/commercial distribution?
- Q3 Apple Developer account ($99/yr) and Azure Trusted Signing (~$10/mo) — OK to set up when we reach C3?
- Q4 Which fonts do target decks use? Share one typical client template if possible.
- Q5 Difference-arrow default (absolute vs %) and 1.0 freeze timing.
- Q6 Keep the name "open-charts"?
- Q7 When can the Windows laptop run the test script; can it install unsigned software today?

## Questions for the reference session
- How is it installed on Windows in companies (per-user vs per-machine MSI, COM add-in + .ppam, LoadBehavior,
  signed VBA/trusted publisher)? How does IT typically deploy it (Intune/SCCM, Jamf)?
- How does it measure text: system font APIs (DirectWrite/CoreText) at run time, its own metrics, or both?
  How does it handle cloud fonts like Aptos that are not installed as files?
- How does it update itself on Windows (service, scheduled task, in-app)?
