# File Operations P2.7 atomic Native qualification validation

This governed Debug x64 File Operations run recorded **117 passed / 0 failed / 20 documented
skips** and zero flaky, regression, isolation, or unclassified-failure classifications.
`fileops_results.json` is the complete per-case authority; `fileops_trace.txt` and
`selftest_run_trace.txt` retain execution diagnostics.

The P2.7 claim is behavioral and bounded: Local Native Move performs one provider mutation; a
regular directory merge is routed to Managed; a regular destination directory that races into
existence requalifies the still-untouched item under the same task; a destination link is never a
merge namespace; case-only directory rename remains provider-supported; mixed Local selections
retain per-item Native/Managed strategies; ordinary Copy has its own strategy name; and host/Local
walkers consume the same bounded traversal and discovery calculations. This archive does **not**
claim throughput, latency-percentile, Release-runtime, remote-credential, or production-workload
improvement.

The executable was built from base commit `f96e8b4815c03672f2caa4fe33f67d3f62f41e66` plus the P2.7
working-tree package finalized by the containing work-package commit. The accepted Debug and
Release build receipts are `91723b1dcbfa30eba1a73a28286da46b4c3346e8d36c0638ae15987a51600e8f`
and `ee7fce3e27cfc95772518839f7c60186dbbbe45cff9408b53a67bc048635df48`.

## Metric compaction

The original `perf/perf_metrics.jsonl` was a 348,742,493-byte stream with 956,158 rows. High-frequency
popup, generic operation, and unrelated diagnostics are not load-bearing for P2.7. The compact JSONL
keeps all `FileOps.SelfTest.*`, operation-strategy, Native-directory-race, bridge/Local traversal,
and discovery rows in original order. `perf/perf_metrics_summary.json` binds the original and
compacted digests/counts, complete behavioral result, representative claim-bearing metrics, and the
explicit claim boundary.

## Diagnostic transparency

`Specs/TestRuns/4cb089111a23/Continuation/20260824_113500_fileops_p27_test_hook_diagnostic`
records the abandoned pause-hook experiment; it was a test-control defect, not a product crash or
mutation failure. `Specs/TestRuns/4cb089111a23/Continuation/20260824_113600_fileops_p27_queue_timing_diagnostic`
records one timing-sensitive queue-progress assertion that failed in a prior full run, passed in
isolation, and passed in the subsequent complete runs. Neither diagnostic is omitted from the P2.7
closeout history.
