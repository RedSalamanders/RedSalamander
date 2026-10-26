# File Operations last note wrapping

Status: complete for the requested layout fix and scoped x64 Debug validation.

Scope: show the complete Last note paragraph on expanded terminal cards, including localized diagnostics and optional informational notes. Keep progress and actions below the measured text; outer scrolling owns excess card height.

## Checklist

- [x] Identify the fixed one-line terminal note height and paint rectangle.
- [x] Reproduce with deterministic no-activation English/French, 480/960 DIP failed-card captures.
- [x] Reuse cached measured DirectWrite layout for complete wrapped note text and card height.
- [x] Add coverage for explicit line breaks, long tokens, informational-note visibility, diagnostic priority, resize, disclosure lifecycle and notes taller than 4096 pixels.
- [x] Update the normative popup spec and user guide.
- [x] Run scoped x64 Debug validation, inspect native captures and archive baseline/candidate evidence.
- [x] Record qualification limits and move the completed plan to Done.

Protected scenario: repeated paints of narrow expanded terminal cards with localized long notes. Reuse FileOps.Popup.CardHierarchy.FrameUs, FileOps.Popup.Render.TotalUs and FileOps.Popup.FactPair.MeasureUs; retain matched same-machine observations without claiming storage throughput or broad qualification. Tests extend FileOps_ReviewRemediationContracts and retain original scene coverage. Preserve the unrelated existing DxUi pin update; baseline/candidate bind the current pin in build receipts.

## Validation and closeout

The enduring behavior is merged into `Specs/UI/UI_FileOperationsPopup.md` and `docs/FileOperations.md`. The existing cached fact-pair helper now owns the complete terminal note's paragraph, measured height, paint bounds and retained full accessible text; no provider queries, worker changes or new resource strings were added. Note caches clear for hidden/collapsed/removed rows and DPI/font invalidation.

Archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-10-19_last-note-wrap/README.md`.

- Baseline `20261019T083744Z-89496-4bbe7f3ad8fa41719bd0e161ee5df7d8` reproduces the exact reported diagnostic clipping and fails the new full-note predicates in all 16 visible-note combinations. Its four hidden-information-note checks pass.
- Candidate `20261019T084519Z-85416-dd3e0a6f0be34ac493880a6ad57c7b9f` passes the focused ReviewRemediationContracts run (3 runner cases). Its 128 native no-activation captures include 24 note-scene checks, four collapse/reopen/resize/change/clear witnesses and four scrolling witnesses. Long notes span 262..524 lines; their last line and actions remain reachable. English/French, 480/960 DIP, light/dark, actual 144 DPI.
- Both x64 Debug builds report zero warnings/errors. Candidate receipt `6dee06f4856bbbc38f34deb2b948eae5eb1c47b2e751c47af4c5b0ad2d931315` binds source snapshot `88a1405fc95855eb8bc70058f1d1bd695a7165fb4ebfdfa0f518c6562d3ef481`. The unchanged source bytes are recorded in the archive. Both builds use the existing unstaged DxUi pin `f65e4a4b9311370b3e3070088e2a8abe29adcb27`.
- Harness source contracts: 204 pass, 0 fail/skip. Explicit pre-stage archive validation passes all 36 files. The final repository-wide archive inventory passes all 15374 files.
- Matched original-scene warm frame median/p95: baseline 8.152/8.542 ms, candidate 8.158/8.540 ms (960 frames each). Matched note scenes: 8.196/8.536 ms versus 8.172/8.578 ms (200 each). The baseline clips notes and the candidate draws them fully; no equivalence, causality, or throughput improvement is claimed.

This closes the requested layout behavior and scoped x64 Debug verification. Full repository, Release, ARM64, ASan, foreground interaction, other-DPI and storage-throughput qualification were not run. The runner's sandbox audit remains unresolved (272 baseline issues, 273 candidate). The wider affected explanation is retained without claiming a repository/CI pass. Only plan closeout, staging and evidence promotion followed the candidate run; product code/tests, normative spec and user guide remained unchanged.
