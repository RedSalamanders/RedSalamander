# I44 Release sort paint latency

Status: DONE, 2026-10-20. Owner: the user-requested Release Extension Ascending investigation.
Current source: `135cd18c0` plus the retained I43 row-reordering changes.
Branch: `ej/release-extension-sort-latency`.

- [x] Attribute the original five-repeat Extension Ascending timings without rewriting I43 evidence.
- [x] Retain fresh unchanged Debug/Release baselines with no observed background builds.
- [x] Correct measured-request work introduced by the test harness, preserving the Windows comparator and the full 5,000-file request sequence.
- [x] Distinguish successful product presentation from attempted/failed/occluded renders; cover failures, recovery, changed sorts and no-op sorts natively.
- [x] Preserve a separately named legacy warm-render control and version the corrected measurement protocol; never label a protocol change a product speedup.
- [x] Update the performance, FolderView and selftest contracts, source/tool tests and performance-reporting preset.
- [x] Rebuild and validate x64 Debug/Release; retain counterbalanced five-repeat benchmark/control blocks without observed builds or known native test runtimes.
- [x] Independently verify raw evidence, source/output identities and formatting; validate source/tool/spec/archive contracts and archive this completed record.

The historical Release Extension Ascending difference is dominated by presentation:
all five requests execute four frames, with baseline Present totals of 11.885-53.978 ms
and candidate totals of 33.998-64.943 ms. Candidate ordering is 0.530-0.806 ms,
layout about 2 ms, and drawing about 5-6 ms. This establishes the cost attribution,
not a causal claim about ambient compositor scheduling.

The existing wait helper pumps the product's pending paint, injects another full
invalidation, then invokes `DebugWarmRenderingForSelfTest`. That helper requests
swap-chain resize, initializes, renders twice, drains messages and processes icon
work. This is test preparation inside each timed request. The intended correction
measures the product's own successful paint without forcing that work, with a
versioned new metric and an unchanged legacy control. Production comparison,
metadata keys, row ordering, rendering policy and VSync must remain unchanged.

Authority: `Specs/UI/UI_FolderView.md`,
`Specs/Testing/Testing_PerformanceValidation.md`, `Specs/Testing/Testing_SelfTests.md`.
I43 and earlier evidence are frozen history. I2 retains broad runtime/platform
performance qualification. All scratch/evidence is under the marked
`Z:/RedSalamander.Perf/evidence/natural-sort-release-extension-2026-10-20` root.
Coordinate only user-authorized build pauses; never stop independent builds/apps.

The correction removes synthetic work from the primary measured span rather than
changing application sorting or presentation policy. The new
`folder.sort_product_paint_us` / `product_paint_v2` measures the product's paint;
the separate legacy case retains `folder.sort_toggle_us` / `legacy_warm_v1`.
The Windows helper, complete row-ordering implementation, production SetSort and
renderer are unchanged, apart from two test-only successful-Present counters.
Source checks remove exactly those two blocks and compare the remaining renderer
with the baseline. Failed EndDraw/Present and occluded frames cannot satisfy the
paint observer. The original I43 measurements remain immutable.

Both test-enabled x64 solution builds pass with zero warnings/errors. All nine
native cases pass per profile, including current/selection, Windows ordering,
paint failure/occlusion/recovery and device-loss/error-persistence coverage. All
328 focused tooling/source-contract tests pass. A separately preserved Debug
smoke run overlaps the Release build and is functional evidence only.

Each final profile has two five-repeat blocks per protocol: Debug runs
legacy/product/product/legacy and Release the reverse. All eight accepted blocks
have zero observed build or known native-test overlap, with complete process
snapshots and all 400 request timings retained. The v2 observer also blocks known
DxUi runtimes/test drivers; five positive-control observations detect the real
concurrent Accessibility process. Initial unchanged-code baselines use the older
build-only observer and remain informational references with narrower scope.
A gate-only v1 controller is retired before any runner launches. One later block
passes natively but is rejected for a metadata-preflight shell; the metadata run
is deferred and the original attempt remains retained. Selection never depends on
latency. The coordinated build and GUI holds are released after timing completes.

Release Extension Ascending product-paint median is 9.268 ms [8.605-10.994],
against the same executable's legacy control at 32.575 ms [25.266-92.511]. Debug
medians are 41.217 ms [34.838-66.408] and 60.666 ms [57.806-64.029]. Every changed
product request has exactly one successful paint, all ten no-op requests per
profile have zero renders, and every measured product request has zero synthetic
warm calls. This is a measured protocol distinction, not a production speedup or
uniform latency guarantee. Other requests and all outliers remain in the evidence.

The [Release report](../../TestRuns/Local-x64/Navigation/2026-10-20_ReleaseSortPaintLatency-Release/README.md)
and [Debug report](../../TestRuns/Local-x64/Navigation/2026-10-20_ReleaseSortPaintLatency-Debug/README.md)
retain full request/counter rows, rendering/layout/presentation metrics, raw
bindings, receipts and independent verification. Larger raw streams and complete
logs stay in the marked external root with exact digest/size/row bindings.
Artifact-lock readback verifies current runtime outputs and the five compiled C++
inputs in both profiles; main contains those same C++ bytes. Qualification uses
the same frozen checkout as the baseline, excluding unrelated ViewerSpace changes.
Later closeout is documentation only. The running main Release is undisturbed.
All retained native/accepted benchmark audits have the same 41 pre-existing
category/path/reason findings; those remain unresolved. This does not establish
Full, ARM64, ASan, an idle host, a lossless process-start audit or physical scanout
latency. I2 retains broader runtime/performance qualification.
