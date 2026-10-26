# File Operations counts and throughput placement

Status: complete for the requested functional changes and scoped x64 Debug correctness; performance equivalence and broad qualification are not established.

Owner request: 2026-10-19. Scope: implement the accepted annotated-card proposal, normative spec and user documentation.

## Checklist

- [x] Capture the unchanged x64 Debug baseline without activation.
- [x] Put recursive file/folder counts or an honest counting/unavailable message under the amount, with ETA at the right.
- [x] Put speed at the right of Throughput, including when its plot is collapsed.
- [x] Remove the ordinary Technical details disclosure and selected-root fraction; show the effective automatic transfer limit below the actions.
- [x] Extend deterministic owned-window coverage for known/unknown totals, unavailable completion counts, counting, skipped counting, stalls, narrow widths, localization and disclosure lifecycle.
- [x] Update the authoritative popup spec and user guide together.
- [x] Build, inspect affected/PrePush explanations, execute focused relevant validation, inspect harness captures and archive matched rendering evidence.
- [x] Record qualification limits and move this completed plan to Done.

## Implementation boundaries

Reuse the popup's cached measured fact-pair layout and retained accessibility regions. Do not query providers or traverse folders from the UI. Recursive successfully completed fractions require published type-specific certainty; selected-root counters are never recursive completion. Known found totals remain useful when completion counts are unavailable. Open totals never get a growing denominator, and missing counts never become zero. Existing skipped-counting completed-work summaries, stall/ETA policy, failure actions and throughput preference ownership remain intact.

Protected scenario: warm expanded Copy/Move cards in English/French, 480/960 DIP, light/dark, through counts and plot disclosure changes. Reuse `FileOps.Popup.CardHierarchy.FrameUs`, `FileOps.Popup.Render.*`, and `FileOps.Popup.FactPair.MeasureUs`; compare matched original scenes on the same machine without claiming a storage-throughput improvement. Keep cached layouts bounded by snapshot rows.

## Evidence

Baseline: `20261019T072659Z-70344-398d3ed721f9417f8ade7146a3a27fbc`, x64 Debug build, zero warnings/errors; `FileOps_ReviewRemediationContracts` passed (3 runner cases). Raw run preserved beneath the marked sandbox at `D:\RedSalamander.Perf\evidence\fileops-card-20261019-baseline`. The existing sandbox audit reported 272 unrelated issues; this does not constitute Full qualification.

## Closeout

The normative behavior is in `Specs/UI/UI_FileOperationsPopup.md`; the user guide and native harness capture/provenance are refreshed in `docs/FileOperations.md` and `docs/res/`. Seven new strings are localized in English and all four satellites. The existing cached measured fact-pair helper owns layout, wrapping, tooltips and retained accessible text; no provider queries or worker changes were added.

Archive: `Specs/TestRuns/4cb089111a23/FileOps/2026-10-19_card-counts-throughput/README.md`.

- Final ReviewRemediationContracts: `20261019T074800Z-93124-e2af00322c84424c95b204bdd1774620`, passed, 100 native captures/geometry/accessibility witnesses; x64 Debug build with zero warnings/errors.
- Live FolderLiveProgress: `20261019T075535Z-41660-3ab0bc0920b848c1bbeb0fad0764f626`, passed, 64 assertions; warning-free x64 Debug build and matching source snapshot. Captures prove actual count ink and advancing file meters while completion/finalization/discovery state changes.
- Resource/source contracts: 214 passed. Governed affected and PrePush explanations were inspected; no CI or broad interactive work is claimed passed.
- Same-machine matched 520-frame observations: baseline median/p95 8.167/8.628 ms; first candidate 8.165/8.534 ms; final candidate 8.018/14.056 ms. The higher final tail is retained and unexplained. Renderer source is identical between candidates, but this is not proof of a cause or performance equivalence. No storage-throughput improvement is claimed.

This closes the requested layout/correctness work, not Full, Release, ARM64, ASan, foreground-input, other-DPI or storage-performance qualification. Those gates remain outside this scoped result. The existing test-root audit is unclean (272 unrelated issues). Documentation/provenance closeout is later than the tested snapshots; code/tests are unchanged. The archive retains the first candidate's four test-source warnings and its passing run alongside the clean final build.