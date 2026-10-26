# Natural sorting: build-free benchmark rerun

Five repetitions passed in each freshly rebuilt x64 Debug and Release profile, with
all 50 sort-request timings retained per profile. No named compiler/tool process or
identified build launcher was observed during either accepted runner window.

The result is mixed. Against the historical pre-optimization natural-sort Debug
baseline, extension ascending is 35.23% lower, extension descending 49.64% lower,
and name descending 36.52% lower. Five of the ten total-request medians are higher
by 10.77-24.44%, including name ascending and time descending. All ten Debug
request medians are higher than the earlier optimized Debug candidate. This rerun
therefore does not establish a uniform Debug latency improvement or complete
removal of the earlier Debug regression.

The earlier baseline and candidate runs overlapped independent compilation. These
same-machine comparisons remain directional: this is an optimized-candidate rerun,
not a new build-free baseline/candidate A/B. Preserve the
[original I41 evidence](../2026-10-17_NaturalSortingPerformance/README.md).

## Fixture, scope and commands

Machine hash: `4cb089111a23`. Same existing `folderView_perf_sort_toggle_stress`
5,000-file adversarial corpus and ten-request sequence; five repetitions. Each
passing repetition checks one enumeration across the ten requests, 5,000 rows,
settled current/layout state and zero inactive quicksearch-effect updates. No
production code or test input changed in this follow-up: all 23 previously qualified
C++/project inputs still match their original hashes in [tested-code.json](tested-code.json).

Full-solution test-enabled x64 builds ran sequentially before measurement, both
with zero warnings and zero errors. Fresh builds were needed because documentation
and metadata closeout had made the older receipts incompatible. The accepted
runner invocations verified the new receipts with `-SkipBuild`:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -Configuration Debug -SkipBuild `
    -CaseFilter folderView_perf_sort_toggle_stress -SelfTestRepeat 5
.\Tools\Run-AllTests.ps1 -Suite Commands -Configuration Release -SkipBuild `
    -CaseFilter folderView_perf_sort_toggle_stress -SelfTestRepeat 5
```

| Profile | Accepted phase | Cases | Run ID |
| --- | --- | --- | --- |
| Debug | `quiet-debug-03` | 5 passed, 0 failed/skipped | `20261017T193241Z-95132-bb18553c2af84e18ad0255e6b3fc673a` |
| Release | `quiet-release-02` | 5 passed, 0 failed/skipped | `20261017T193836Z-91436-61130a0828da40bd9cd88260029c1763` |

This is focused sorting evidence. The prior correctness/focus/ViewerSpace results
remain in I41; this follow-up does not rerun them or claim Full/ASan/ARM64 qualification.


The canonical runner's sandbox disk audit remains **not clean**: all four attempts
report the same 41 `unexpected-test-run-dir` paths already present in the shared
sandbox (for example, `runs/adhoc-103144`). Their ownership could not be attributed
by the runner; they were preserved. Native benchmark verdicts and runner exit codes
are passing, but this report makes no clean-sandbox or whole-repository claim.
Original `run-all-tests-results.json` files retain every audit reason.
## Total sort-request timings

Milliseconds, **median** [minimum - maximum], five observations per row. Values
include the fixture's layout/render/pump/settling work, not only comparison CPU.
The first name-ascending request starts in the listing's existing ascending order.
None is intentionally requested twice in each repetition. No sample is discarded.

| Request | Debug ms | Release ms |
| --- | --- | --- |
| Name Ascending | **41.607** [29.741 - 49.791] | **9.140** [5.429 - 11.396] |
| Extension Ascending | **117.478** [89.340 - 135.084] | **29.723** [26.287 - 84.542] |
| Time Descending | **142.852** [107.778 - 162.397] | **59.460** [49.929 - 64.838] |
| Size Descending | **106.519** [93.806 - 141.621] | **58.572** [57.129 - 66.667] |
| None Ascending (first) | **120.767** [99.395 - 139.067] | **58.317** [56.636 - 67.355] |
| Name Descending | **76.019** [71.953 - 81.446] | **58.694** [57.347 - 65.322] |
| Extension Descending | **76.230** [68.659 - 82.193] | **58.935** [57.268 - 66.489] |
| Time Ascending | **99.620** [83.374 - 119.430] | **58.515** [56.356 - 66.175] |
| Size Ascending | **99.559** [82.530 - 127.937] | **57.631** [56.322 - 67.910] |
| None Ascending (last) | **94.388** [77.813 - 127.883] | **58.037** [56.431 - 58.145] |

All 100 accepted observations, historical baseline/candidate medians and directional
percentage changes are in [performance-comparison.json](performance-comparison.json).
The complete prior comparison is retained as
[historical-contended-comparison.json](historical-contended-comparison.json).

## Enumeration preparation and sort stage timings

Milliseconds, **median** [minimum - maximum]. These records also include fixture
restoration work; canceled/partial enumeration rows remain in the raw selected
metrics but are excluded from the complete 5,000-row stage summaries below.
Counts are not cold-open trial counts, and medians from separate stages must not
be added as an end-to-end duration.

| Stage | Debug records | Debug ms | Release records | Release ms |
| --- | --- | --- | --- | --- |
| BuildItems (complete 5,000-row records) | 9 | **8.258** [5.063 - 11.885] | 9 | **0.538** [0.299 - 0.821] |
| SortMerge (0 directories, 5,000 files) | 9 | **48.021** [39.948 - 51.466] | 9 | **2.874** [2.697 - 5.705] |
| NaturalSortRanks (5,000 rows) | 9 | **5.694** [4.868 - 9.379] | 9 | **0.627** [0.556 - 0.698] |
| ApplyCurrentSort (5,000 rows) | 59 | **34.602** [6.439 - 62.040] | 59 | **1.014** [0.293 - 1.955] |

Debug `ApplyCurrentSort` median is 34.602 ms versus 50.151 ms in the historical
pre-optimization natural-sort baseline (31.00% lower), and versus 32.596 ms in the
earlier optimized candidate (6.15% higher). Release is 1.014 ms versus the earlier
candidate's 1.068 ms. These comparisons share the historical contention caveat.

Ranks retain 80,000 bytes per 5,000-row listing and use a 40,000-byte temporary index
during preparation. This is a bounded field/index cost, not a measured process-wide
memory claim. `StrCmpLogicalW` ordering and the rank-cache implementation are unchanged.

## Build-activity observation and limits

Each accepted run followed at least 30 seconds without observed build activity.
The observer sampled `System.Diagnostics.Process.GetProcesses` nominally every
200 ms, detected compiler/tool names unconditionally, and classified available
PowerShell/Python/cmd/node/wsl command lines for build launchers. Unresolved shell
command lines were retried; an unresolved shell present for at least 1,000 ms
blocked acceptance. Shorter unresolved launches remain logged explicitly.

| Observation | Debug | Release |
| --- | --- | --- |
| Runner UTC window | 19:32:39.907 - 19:34:20.430 | 19:38:34.553 - 19:40:09.394 |
| Measurement snapshots | 390 | 390 |
| Build-tool/identified launcher polls | 0 | 0 |
| Maximum snapshot gap | 931.328 ms | 575.401 ms |
| Polls with short unresolved shells | 8 | 6 |
| Observer CPU across runner window | 11.516 CPU seconds / 100.523 elapsed seconds | 11.094 CPU seconds / 94.841 elapsed seconds |

This establishes no observed build overlap, not a completely idle host or a
lossless process-start audit. Snapshotting can miss processes shorter than a poll
gap, and short unresolved shells could not be positively classified. The observer
itself adds work; this report records its CPU cost rather than treating it as free.
`Win32_ProcessStartTrace` event subscription was unavailable (access denied).
The verified GIGABYTE Control Center `GCC.exe` process was excluded by its exact
PID/start-time identity after querying its executable image, not by disabling GCC
compiler detection. See [non-build-process-exclusion.json](non-build-process-exclusion.json).

The final observer passed a synthetic build-launcher positive control **after**
both accepted measurements. See [observer-control/probe-result.json](observer-control/probe-result.json).
The orchestration and classification implementation is retained in
[run-quiet-benchmark.ps1](run-quiet-benchmark.ps1).

## Retained attempts and provenance

Acceptance was based on case verdicts and process observations, before timing
analysis; no attempt was selected because it was faster. All four executed
five-repeat benchmark attempts retain every sort sample and original results:

| Attempt | Outcome | Disposition |
| --- | --- | --- |
| `quiet-debug-02` | 5 passed; five short PowerShell observations flagged by the original strict unknown-shell policy | Excluded; observer handling corrected, original observation retained |
| `quiet-debug-03` | 5 passed; zero observed build polls | Accepted Debug |
| `quiet-release-01` | 5 passed; 15 MSBuild process observations during the runner | Excluded; repeated after another quiet interval |
| `quiet-release-02` | 5 passed; zero observed build polls | Accepted Release |

The earlier `quiet-debug-01` gate aborted on a process-exit race before launching
any benchmark. The initial observer control also failed because individual CIM
queries outlived the synthetic process; batching corrected that startup overhead.
These setup failures and all external raw locations remain recorded in
[attempt-history.json](attempt-history.json). Both synthetic controls and rejected
benchmark attempts are separate from claim-bearing performance data.

[build-validation.json](build-validation.json) binds both new successful full-solution
receipts and raw build logs. [runtime-verification.json](runtime-verification.json)
confirms all 384 output digests after the runs, plus the 23 unchanged source inputs.
Original results, runner/Commands traces, immutable receipts, runner logs and phase
bindings are retained per attempt.

High-volume raw metric streams and the complete gate monitor logs remain at
`Z:\RedSalamander.Perf\evidence\natural-sort-quiet-rerun-2026-10-17`.
[artifact-bindings.json](artifact-bindings.json) records their hashes, original row/byte
counts, exact selections and archive digests. Metric extraction retains **every**
row in the named families (`folder.sort_toggle_us`, `FolderView.ApplyCurrentSort`,
and enumeration `BuildItems`, `SortMerge`, `NaturalSortRanks`) in original order.
Monitor extraction retains every benchmark-window sample and the final 30 seconds
of the qualifying gate; earlier busy-gate rows remain in the digest-bound raw file.
The archive disables Git text normalization in `.gitattributes`.

## Validation

The explicit archive contract passed for the base 75-file packet. Independent verification passed for all 73 original/derived bindings, all 200 sort samples across four attempts, every selected metric row in original order, all benchmark-window process snapshots, and the 100 accepted samples. See [independent-evidence-verification.json](independent-evidence-verification.json) and the retained [verification script](verify-evidence.ps1). The final explicit archive gate passed for all 79 files after adding these validation records. `git diff --check` passed.
The whole tracked TestRuns inventory already passed (13,844 files); its original
result is retained in archive-inventory-validation.log.


