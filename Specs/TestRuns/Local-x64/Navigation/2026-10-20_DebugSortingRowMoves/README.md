# I43 Debug sorting row movement

Fresh same-machine baseline/candidate pairs use the unchanged 5,000-file corpus, ten requests and five repetitions per profile. Every accepted phase passes five cases and retains all 50 request timings. Selection is by case/observer predicates before inspecting latency.

The optimization sorts reusable row indices with the existing comparator, then applies whole-row permutation cycles. Owned strings/layouts and selection/current state follow their original identities. Globally checked Debug iterators remain enabled. The Windows comparison helper, rank preparation, complete production comparator, sort threshold, benchmark fixture and build-property inputs are unchanged; see [source checks](ordering-preservation-source-checks.json).

## Core ordering and complete sort call

Milliseconds, median [minimum - maximum]. Stage records also include setup and fixture restoration; counts are not cold-open trial counts. Separate stage medians must not be added.

| Profile / stage | Baseline ms | Candidate ms | Median change |
| --- | --- | --- | --- |
| Debug OrderItems | 33.385 [6.163 - 38.378] | 5.899 [2.485 - 7.840] | -82.33% |
| Debug ApplyCurrentSort | 33.758 [6.354 - 38.624] | 6.305 [2.743 - 8.125] | -81.32% |
| Release OrderItems | 0.971 [0.228 - 2.413] | 0.529 [0.141 - 0.970] | -45.52% |
| Release ApplyCurrentSort | 1.008 [0.260 - 2.467] | 0.559 [0.163 - 1.037] | -44.54% |

## Total sort-request latency

Milliseconds, median of all five observations for each request. These include layout, rendering, pumping and settling. The first Name Ascending request begins already sorted and can be a SetSort no-op. All samples and ranges remain in [performance-comparison.json](performance-comparison.json).

| Request | Debug baseline / candidate | Debug change | Release baseline / candidate | Release change |
| --- | --- | --- | --- | --- |
| Name Ascending | 31.235 / 28.199 | -9.72% | 9.324 / 6.727 | -27.85% |
| Extension Ascending | 87.761 / 60.042 | -31.59% | 28.266 / 49.861 | 76.40% |
| Time Descending | 103.194 / 80.083 | -22.40% | 66.018 / 64.464 | -2.35% |
| Size Descending | 95.733 / 69.871 | -27.02% | 59.487 / 64.438 | 8.32% |
| None Ascending (first) | 94.617 / 63.588 | -32.79% | 66.338 / 58.132 | -12.37% |
| Name Descending | 68.591 / 62.330 | -9.13% | 58.129 / 57.510 | -1.07% |
| Extension Descending | 66.322 / 36.922 | -44.33% | 58.906 / 58.081 | -1.40% |
| Time Ascending | 81.014 / 60.205 | -25.69% | 57.703 / 58.430 | 1.26% |
| Size Ascending | 81.592 / 55.989 | -31.38% | 59.043 / 57.371 | -2.83% |
| None Ascending (last) | 80.049 / 49.601 | -38.04% | 65.239 / 66.378 | 1.75% |

Debug request medians decreased for 10/10 requests; Release for 6/10. Core ordering and complete request latency are separate claims. Higher request medians remain visible; no sample is removed or rebaselined. Initial worker enumeration sorting is unchanged.

Release Extension Ascending is the largest total-request increase: 28.266 ms to 49.861 ms (+76.399%), with overlapping five-sample ranges of 25.134-69.504 ms and 47.696-79.084 ms. Size Descending, Time Ascending and the last None request also increase. The measured core improves in both profiles, and all Debug request medians improve; this does not qualify Release rendered-latency non-regression. A separate counterbalanced rendering/settling investigation is required before that broader claim. No causal explanation is inferred from these samples alone.

## Pairing, observations and scope

Both Debug and Release pair the same isolated checkout at `Z:/src/RedSalamander-worktrees/debug-sort-row-moves-validation`. Concurrent ViewerSpace edits invalidated the main v2 build, so its unrelated work was preserved and excluded from final pairing. Earlier main builds/tests and the clean main Debug baseline remain supplemental history. Both isolated baselines ran on the unchanged algorithm before the frozen sorting patch was applied there. The final candidate uses one lazy staging row across cycles. Each pair has matching checkout, fixture, toolchain, dependency, build flags and observer identities.

| Profile / phase | Run ID | Receipt | Build overlap polls | Max sample gap | Observer CPU / elapsed |
| --- | --- | --- | --- | --- | --- |
| Debug baseline (paired-baseline-debug-02) | 20261020T131103Z-82268-d50e71b411b94e76bd6174c5191b7bf9 | 9c89f19e0611dbde847909eafc5149ed3969f8adc02060061185790807315c8d | 0 | 770.411 ms | 14.031 s / 97.421 s |
| Debug candidate (paired-candidate-debug-01) | 20261020T132924Z-100560-bea25b3391ac4b81a8c3ef08a932a0a8 | 5a3f7b334066b6c7cd1f36c31aa7b5ec4cbbd4a3e9c89c0eab4a011bc3827003 | 0 | 682.788 ms | 12.484 s / 92.004 s |

Debug candidate counters: 59 movement records, minimum 0 and maximum 7500 payload moves; retained index scratch 40000-40000 bytes. All records satisfy the linear movement and capacity bounds.

| Release baseline (paired-baseline-release-01) | 20261020T131311Z-35412-cc68293073b34bdda808d8f875d88148 | 3ee74cc23947bde2864030423f3b6713cb633d6ad486d2fb749557d83cf79db7 | 0 | 685.069 ms | 12.922 s / 90.005 s |
| Release candidate (paired-candidate-release-01) | 20261020T133126Z-27124-3666245334074cc88107f6e3bb2887e2 | b990d166678ae25d12ab5d321530bcc16453cd775404431d84c7e3c4df03e57b | 0 | 666.264 ms | 12.313 s / 89.587 s |

Release candidate counters: 59 movement records, minimum 0 and maximum 7500 payload moves; retained index scratch 40000-40000 bytes. All records satisfy the linear movement and capacity bounds.

The observer requires a separate 30-second quiet gate and samples the entire canonical runner window, nominally every 200 ms. Exact dead-parent reusable MSBuild nodes are excluded only after 30 seconds of unchanged CPU counters; any resumed CPU activity makes them blockers. The unrelated GIGABYTE GCC executable is excluded by exact verified identity. Compiler detection is otherwise unconditional; unknown orchestration shells are retried and block once persistent for 1,000 ms.

This establishes no observed background build overlap, not a completely idle host or a lossless process-start audit. Short-lived processes can fit between snapshots; short unresolved shells and observer CPU cost are retained. No privileged event subscription was requested. Each rejected benchmark and every gate-only retirement is retained separately. Historical October 17 contended measurements are not the baseline for these pairs.

## Correctness and qualification

The final isolated x64 Debug and Release solution builds are test-enabled and complete with zero warnings/errors. The six focused cases pass in both final executions (functional-final-debug and functional-final-release): natural-name contracts, bounded current memory, navigation roundtrip, refresh/filter selection, visibility/selection restoration, and current/selection/removal state. The expanded real reorder test covers all six keys, both directions, three cache modes and four input permutations; all 720 six-row permutations; owned DirectWrite layouts/strings; empty/singleton models; the actual sequential/parallel boundary; scratch reuse/release; and full-width numeric extremes.

All 289 focused tooling/source-contract Pester tests pass on the final isolated source. Specification/tool inventories, changed-range formatting, archive validation, and independent raw evidence verification are recorded alongside this packet. This is focused sorting qualification; it does not establish Full, ARM64, ASan or process-wide memory/performance qualification.

The shared sandbox audit reports the existing 41 unexpected run directories. Original aggregate results retain every reason. They were preserved; passing native cases and exit codes do not imply a clean sandbox audit.

## Time and Size versus natural sorting

Name and Extension use Windows natural ordering. Time uses signed 64-bit modification timestamps; Size uses unsigned 64-bit byte counts for files, while folders use natural names. Attributes use numeric DWORD flags. Main-pane equal metadata values use natural names, then ordinal spelling and stable source order. Find Size/Modified remain numeric and retain arrival order on equal values. Date formatting and size separators are not sort keys.

## Evidence retention

Original case verdicts, reasons, traces, aggregate results, immutable receipts and claim-bearing metric rows are retained. All 50 request samples and every relevant sorting metric remain in original row order; all benchmark snapshots and the final gate interval are retained. Larger raw streams and complete build logs remain under `Z:/RedSalamander.Perf/evidence/natural-sort-row-reorder-2026-10-20`, with SHA-256, byte/row counts and exact extraction policy in [artifact-bindings.json](artifact-bindings.json). See [independent verification](independent-evidence-verification.json).

The archived scripts describe this measured snapshot; they are not new supported repository commands. Canonical runtime command for each profile:

```powershell
Tools/Run-AllTests.ps1 -Suite Commands -Configuration Debug -SkipBuild -CaseFilter folderView_perf_sort_toggle_stress -SelfTestRepeat 5
Tools/Run-AllTests.ps1 -Suite Commands -Configuration Release -SkipBuild -CaseFilter folderView_perf_sort_toggle_stress -SelfTestRepeat 5
```

The complete focused commands traces and native verdicts/reasons are unabridged. Verbose functional application traces retain the first/last 20 lines and every selftest/assert/error/warning/failure line in original order; the full application traces remain externally digest-bound. Gate-only retirements retain eight available tail lines for diagnosis and make no quiet-window qualification claim. All selected and rejected benchmark measurement snapshots and relevant metrics remain complete.

Final artifact-lock readback verifies the current isolated executable/runtime closure and exact compiled sorting C++ inputs in both profiles; the main checkout contains the same three C++ files. See [current identities](final-identities.json). Main Release was running during closeout and was left undisturbed. Qualified Debug/Release binaries remain in the isolated checkout. Later plan/evidence closeout changes documentation only. The [completed I43 record](../../../../Plans/Done/UI_DebugSortingRowMoves_2026-10-20.md) retains this scope and the unresolved broader Release timing question.
