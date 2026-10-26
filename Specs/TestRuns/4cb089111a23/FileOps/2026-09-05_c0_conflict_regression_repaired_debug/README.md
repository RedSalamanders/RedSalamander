# C0 conflict regression — repaired candidate

Run `c0-conflict-regression-repaired-20260905`, Debug,
`FileOpsFamily_Phase09_ConflictPrompt`: 11/11 passed, no skips.
Build receipt `a7f0414334e2c51ee67417c2f4383ea81ed6394c3940d0975360d9a41b32cfad`;
zero warnings/errors. Base `58f22d0e` plus working changes is runner-bound.

The prior `2026-09-05_c0_conflict_regression_red_debug` run failed the nested
Keep Both case. The candidate passes Copy and Move at concurrency 1 and 2,
with exactly two prompts each. Intentional child Skip preserves the documented
partial-directory result without authorizing replay of uncertain work.
The existing RetryCap expectation is baseline compatibility, not C2 completion.

`results.json`, `trace.txt`, and the runner aggregate are unchanged raw artifacts.
The raw performance stream remains under the same run's
`artifacts/selftest/last_run/perf/perf_metrics.jsonl` in the owned test root:
8628855 bytes, 24902 nonblank rows, SHA-256
08fc44a522f49e398f14df12d9bf9feadb40c30595b4d5ebbb85571bdc84fd3c.
The compact performance file retains every row whose metric matches
`^FileOps.SelfTest.(KeepBothNestedCacheEligibilityUs|ConflictApplyToAllConvergenceWait)$`,
in original order (raw lines 1451, 1872, 2326, 2784, 3198).
No throughput, percentile improvement, or full-stream timing claim is made.
This is same-machine correctness/prompt-count evidence only.

The disk audit reports 22 retained earlier run directories outside this invocation's
allow-list. They were not deleted; this archive does not claim a clean test-root audit.
Remaining provider coverage, C1–C5, Release and final Fresh Full remain open.
