# File Operations P2.6 integration validation

This governed Debug x64 File Operations run recorded **116 passed / 0 failed / 21 documented skips**
and zero flaky, regression, or isolation classifications. `fileops_results.json` is the complete
per-case authority; `fileops_trace.txt` and `selftest_run_trace.txt` retain execution diagnostics.

The P2.6 claim is behavioral and bounded: every shipped `nativeMove: true` profile is covered by the
provider matrix/source contract, Local direct cross-volume Move cannot copy/delete, host Managed Move
uses exact conditional cleanup, Copy-only retains its source, changed directory membership is
source-kept without a cleanup prompt, legacy atomic-final verification cannot become destructive
authority, and traversal/parallelism ceilings remain executable. The only product-semantic skip is
in-tree link retargeting, explicitly owned by P3.2; the other skips retain their environment reasons.
This archive does **not** claim throughput, latency percentiles, Release runtime performance, or
production-workload representativeness.

## Metric compaction

The original `perf/perf_metrics.jsonl` was a 131,114,106-byte stream with 356,083 rows. High-frequency
popup, progress, identity, discovery, and per-delete diagnostics exceeded repository evidence limits;
their complete ordering is not load-bearing for the P2.6 claim. The compact JSONL keeps all
`FileOps.SelfTest.*`, strategy-selection, native-merge-reclassification, bridge traversal, and
recursive-copy bound rows in original order. `perf/perf_metrics_summary.json` binds the original and
compacted digests/counts, complete behavioral result, retained corpus identity, representative
claim-bearing metrics, and the explicit claim boundary.
