# I45 Debug sort paint CPU cost

Status: DONE, 2026-10-21. Started 2026-10-20. Owner: the user-requested Debug versus Release investigation.
Base: `a6ced5f41`; retain the I43/I44 implementations and immutable evidence.

- [x] Retain fresh x64 Debug/Release baselines for the unchanged 5,000-file, ten-request `product_paint_v2` scenario, with the same build/test-process observer and no detected overlap.
- [x] Attribute ordering, layout, drawing, presentation and capture overhead; distinguish actual application work from diagnostic observer cost.
- [x] Optimize the demonstrated CPU bottleneck without changing Windows Explorer name/extension ordering, numeric Time/Size keys, stable ties, current/selection or rendering behavior.
- [x] Add meaningful native correctness/resource coverage and retain the benchmark's successful-paint/no-op/failure contract and every metric sample.
- [x] Update the owning contracts, shared-helper catalog if needed, source/tool tests and documentation.
- [x] Rebuild both profiles; run affected native/tooling tests and repeat the same benchmark with no detected background compilation or known native-test overlap.
- [x] Verify source/build/artifact identities, preserve failed/rejected attempts, archive comparable evidence, validate metadata/formatting and move this completed plan to Done.

Authority: `Specs/UI/UI_FolderView.md`, `Specs/Testing/Testing_PerformanceValidation.md`,
`Specs/Testing/Testing_SelfTests.md` and `Specs/Core/Core_SharedHelpers.md` where applicable.
The paint protocol and 5,000-file corpus remain fixed; a faster recorder must be
identified as diagnostic overhead reduction rather than an uncaptured UI speedup.
The first retained I44 product block attributes one Debug Extension Ascending
request to 7.135 ms ordering, 10.335 ms layout and 15.407 ms begin-to-EndDraw,
against Release 0.706/2.207/3.186 ms. These historical samples guide profiling;
fresh matched evidence is required for any optimization claim.

Qualification uses the owned native Git worktree
`E:/RedSalamander.SortValidation-20261020/debug-sort-paint-bottleneck` (baseline) and
`S:/RedSalamander.SortValidation-20261020/debug-sort-writer-validation` (candidate), leaving the running
main Release application untouched. Scratch and full originals remain under
`Z:/RedSalamander.Perf/evidence/debug-sort-product-cost-2026-10-20`.
Only the two user-authorized chats may be messaged for measurement coordination;
independent builds/apps are never stopped. Snapshot observation is not proof of
a completely idle host. Existing sandbox-audit findings and broader I2/Full,
ARM64 and ASan obligations remain separate.

The Z: volume filled during preparation. Preserve the unattested first Debug
build and disk-full Release/dependency-restore failures. Both owned checkouts
were relocated before qualification; canonical builds must re-attest their new
paths. All timed fixtures and JSONL capture use the same explicitly selected marked S: sandbox;
executable loading and warm-up stay outside requests. Source/output paths differ
and remain explicit in receipts. No unrelated files or applications are removed.

## Implementation and completed correctness gates

The candidate changes `Common/Common/PerfJsonl.cpp`, the synchronous diagnostic
writer shared by the application and plugins. It caches escaped run metadata per
configuration, reuses bounded line scratch, handles short proven ASCII without
UTF-16 conversion/allocation, preserves the strict Unicode fallback, escapes
unchanged spans in bulk and formats the fixed numeric/timestamp wire fields
directly. Configuration and complete append share one lock. All original records
remain synchronous and immediately readable; oversized rows are preserved, then
scratch above 64 KiB is released. No sorting, row movement, layout, presentation,
compiler or dependency input is changed.

Identical test-only per-thread writer observations accompany each request in both
roles. The final reader uses TLS and lock-free published resource counters: it
cannot wait for a writer before the request timer and hide that contention. The
superseded prototype's locking reader is retained as preparation history and is
excluded from the final builds/native qualification and all accepted timings.

The final test-enabled x64 Debug and Release builds have zero warnings/errors.
Each final profile passes ten focused native cases: writer schema, Unicode and
ASCII boundaries, concurrency, immediate read, bounded retention, reconfiguration
and live file rotation; Windows ordering; current/selection/focus state; successful
paint/no-op/failure/occlusion; device loss and persistent render errors. The
tooling closeout passes 331 tests with zero failures/skips. Changed native ranges
have no clang-format replacements. Source preservation and current runtime closure
checks bind the exact four compiled diagnostic/test inputs in every profile.

Two chats were authorized for coordination. The user declined messaging
`Create RedTilde Markdown editor` and instructed direct progress; the already-built
launcher therefore waits automatically for its detected compilation/native tests
to finish. No independent operation is stopped. Final claims require accepted
counterbalanced Debug ABBA and Release BAAB blocks, every original timing sample,
and complete observer coverage; preparation passes alone do not close this gate.

Retained mid-study history: three clean Debug blocks were retained (baseline, candidate,
candidate). A later closing-baseline attempt is rejected for fresh RedPrism
compilation; it is not substituted into the comparison. Bounded coordination
intervals interrupt only verified waiting gates, never a launched native runner.
The resumed launcher validates the exact accepted prefix, source/receipt/profile/
dependency identities, current compiled inputs/runtime closure under canonical
profile locks, all captured artifact hashes and the original polling rows and
observer versions. Fifteen positive/negative resume guards pass. Some ephemeral
canonical run paths are absent; the required complete captured copies remain
verified against their originally recorded hashes, and those absent paths are
listed explicitly. No timing data is reconstructed or selected again by latency.
The original pre-resume caller and failed guard-preparation receipts are retained.

## Final matched measurements

All eight first-eligible blocks completed: Debug ABBA and Release BAAB, five
repetitions per block, all ten requests per repetition, ten observations per
request/role/profile and 400 original request samples. Accepted windows have zero
detected compilation/known-native-test overlap; their maximum polling gaps are
below 1,000 ms. Four completed attempts remain excluded (three with real new
build/test activity; one read-only tooling inspection matching the conservative
driver filter), and three pre-run gates were deliberately interrupted. No
observation is rewritten and no timing sample is selected by latency.

| Extension Ascending, median ms | Baseline | Candidate |
| --- | ---: | ---: |
| Debug request/paint | 34.701 | 19.810 |
| Debug counted writer spans | 17.312 | 2.479 |
| Release request/paint | 8.984 | 8.591 |
| Release counted writer spans | 1.591 | 0.883 |

The Debug request reduction is about 43%, and counted writer spans fall about
86%. The per-request arithmetic remainder stays about 17 ms. Candidate Debug
OrderItems/drawing medians remain 6.527/6.728 ms versus Release 0.648/3.110 ms;
normal Debug checks, row movement and rendering are preserved. These scopes
overlap and are not pure CPU samples. Every changed Debug request improves with
non-overlapping observed ranges in these blocks. Release paint ranges overlap;
Name Descending and final None medians are slightly higher, so no uniform Release
gain is claimed. Release Time Ascending includes Present changing from
17.407 to 2.497 ms; the full 16-ms difference cannot be assigned to CPU sorting or
recording. The study does not establish capture-disabled speedup.

Full per-request ranges, stage/preparation distributions and all counters are in
the [Debug packet](../../TestRuns/Local-x64/Navigation/2026-10-20_DebugSortWriter-Debug/README.md)
and [Release packet](../../TestRuns/Local-x64/Navigation/2026-10-20_DebugSortWriter-Release/README.md).
All native code is identical to the qualified candidate, and the protected
ordering/rendering/compiler/pin sources still match `a6ced5f41`. Final native
and tool/source gates are distinct from the broader I2/Full, ARM64, ASan and
foreground-input obligations. Independent verification passes both packets,
including every request/counter, full decompressed observation bytes and compiled
source/artifact identities. The archive validator passes all 173 files within
the default 2-MiB/file and 5-MiB/run limits. Specification/tool inventories have
zero blocking findings, changed native ranges have no formatting replacements,
and final whitespace checks pass. This plan is closed; broader qualification
remains with its existing owners.
