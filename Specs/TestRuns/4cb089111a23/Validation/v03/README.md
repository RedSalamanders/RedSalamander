# Footer command flow: Release qualification and rejected confirmation captures

The footer no longer drops Queue when a command row fills. Measured transfer actions
keep their widths, the Queue/Options/Details group moves together, and the window and
task viewport reserve the resulting footer height. The consumer reuses the pinned
DxUi measured-action primitive; it adds no timer, surface or variable-size layout cache.

The test-enabled x64 Release solution builds with **zero warnings/errors**. Fresh PR
run `20260923T094710Z-61284-5d5c33e96d5f4d4b82bb9eff88c94e1e` passes **13 executed and
promoted entries, 824 reported passes, zero failures/skips/audit issues** in 464,257 ms.
The count includes 812 Pester cases and twelve entry outcomes, not every native assertion.
ProductUiTests includes one-, two- and three-row footer cases, fractional-DIP boundaries,
hidden actions, exact measured widths, retained settings, nonoverlap and invalid inputs.

## Identity and reproduction

- Source base: clean baseline `ccebd5af68eb45f2e9deed64a7063780e71eb1ce` in `C:/RsI26Final`.
  The candidate applies the five explicit files in [candidate-inputs.json](raw/candidate-inputs.json).
  Their bytes are included in the build snapshot; the candidate is deliberately dirty.
- Candidate build receipt: `1d40e26a4aacc6c0c2298e744ead9ec101961a611a182011752d508554e63274`;
  snapshot `a8f9a67de89f37a41bda47df69a4c81ce6edbec5572dc18dad9118ee7bbe4c2d`.
- Consumer DxUi pin remains `78b3de389a189c7f86f611787e0489fb6d474218`.
- [Build driver](raw/build-candidate.ps1.receipt.txt) uses process-local CI vcpkg imports,
  tests enabled, x64 Release and four build workers. No user-wide integration changes.
- [PR driver](raw/qualify-pr.ps1.receipt.txt) validates the input hashes and absence of
  foreground entries, then runs `Run-AllTests.ps1 -Suite PR -Configuration Release
  -Platform x64 -SkipBuild -TestRoot E:/RedSalamander.Perf`.
- [Aggregate](raw/run-all-tests-results.json), immutable plans/snapshots and all entry
  evidence/promotions are retained under `raw/`. This is Fresh PR coverage, not Full.

## Gallery comparison and its failure

The unchanged no-activation French gallery runs before and after the five-file change:

| | Baseline | Candidate |
| --- | --- | --- |
| Run | `20260923T092420Z-76296-1bd566068788467fba6b59c4711bd9bf` | `20260923T095606Z-60568-7cdf6164d5d445118eff3507430b21a7` |
| Runner outcome | 1 passed, 0 failed/skipped, audit 0 | 1 passed, 0 failed/skipped, audit 0 |
| Original PNGs | 304 | 304 |
| Physical DPI | 144 and 96 | 144 and 96 |

All 304 dimensions and both scenario/display manifests match. **284 pairs are pixel-identical**,
including the ordinary and compact footers. **Twenty candidate confirmation PNGs are black**.
The runner's success is therefore insufficient: this candidate gallery is rejected as complete
documentation evidence. The originals remain preserved and the capture-readiness/pixel-witness
repair is tracked in the active plan. No black image is an approved illustration of the UI.

The [complete pixel comparison](raw/gallery-comparison.json) binds every original PNG by SHA-256.
The [comparison script](raw/compare-gallery.py.receipt.txt) only reads images. Full original
galleries and metric streams remain at `C:/RedSalamander.Perf/evidence/i26-ui/`, in
`footer-wrap-baseline-20260923` and `footer-wrap-candidate-20260923`.

## Paired measurements and limits

The comparison retains every focused sample, plus hashes and sizes for the full external
metric streams. The unchanged warm delete/verifying fixtures each measure five eight-redraw
batches at 480/760 DIP on both physical displays. Median batches are presentation-paced near
66.6 ms; the largest median increase is 64.916 to 66.713 ms for verifying at 96 DPI/480 DIP.
The footer label cache misses exactly 164 times in each gallery. Its median cold measurement
is **73 to 101 microseconds**, and p95 is **168 to 186 microseconds**. This change is recorded,
not hidden. Measurement now occurs before card layout rather than after it; the pair alone
does not establish why cold timings differ. Whole-gallery render-event counts differ by one,
so aggregate render timings are descriptive rather than a controlled throughput comparison.

One serial gallery pair does not qualify CPU, process memory, retained resources, steady-state
throughput or an accepted regression. The library menu/grid resource gates are unchanged.
The current French vocabulary stays on one footer command row: native wrapped-row interaction,
summary-text wrapping, Debug/Full, native DPI/keyboard/UIA/AT and final integration remain open.

[SHA256SUMS.txt](SHA256SUMS.txt) binds the raw payload. No failed attempt is promoted as a pass.
See the [active plan](../../../../Plans/WIP/UI_FileOperationsReview_2026-09-16.md) and the
[gallery packet](../../../../Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-23-footer-wrap-release/README.md).

## Capture guard and rejected presentation experiment

The stronger confirmation witness correctly fails run
`20260923T100637Z-47608-a7c05d7bad9a48bb9c0b43cb039c09ab` (0/1, audit zero).
The pixel witness requires four close-glyph pixels in the actual production layout;
known-good baseline PNGs satisfy it. The black candidate contains zero.

Added test-only observations in run `20260923T101542Z-41008-a06f4609fee24a169e63a56bc5bf1599`
show three opaque paints, EndDraw S_OK and reported HWND occlusion. This does not prove
why capture fails: CheckWindowState can reflect the previous EndDraw's occlusion.
An isolated ENABLE_TESTS-only `D2D1_PRESENT_OPTIONS_RETAIN_CONTENTS` experiment also
fails, run `20260923T102354Z-29756-f42a5d3934fb499396ee4f2e8f5e93fb`, 0/1 and audit zero.
Its clean Release build receipt is
`e564aff9f96cdfd27748421f28b21e40c0f02c1adf5e1c8376d9247c716415fd`.
It was removed from the isolated checkout and never entered the primary code.
Production render flags remain unchanged; no desktop takeover or screenshot fallback
was used. The guard and diagnostics are retained, and complete confirmation capture
remains a required open gate. Raw attempts are under the three `raw/capture-*` folders;
full artifacts remain in the corresponding external evidence folders.
