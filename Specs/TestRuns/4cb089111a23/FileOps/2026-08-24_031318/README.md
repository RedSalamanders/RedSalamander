# File Operations P2.5 full-suite validation

This governed Debug x64 File Operations run recorded **110 passed / 0 failed / 27 documented skips** and zero
flaky, regression, or isolation classifications. `fileops_results.json` is the complete per-case authority;
`fileops_trace.txt` and `selftest_run_trace.txt` retain execution diagnostics.

The P2.5 claim is behavioral and bounded: discovery and transfer overlap in one traversal, Skip discovery releases
reserved capacity, cancel remains bounded, selected leaf totals are exact before publication, the discovery queue is
bounded to 256 records, and Local permanent Delete uses the bounded no-follow traversal. This archive does **not**
claim throughput, latency percentiles, Release-build performance, or production-workload representativeness.

## Metric compaction

The original `perf/perf_metrics.jsonl` was a 131,112,178-byte stream with 356,113 rows. High-frequency popup and
identity diagnostics caused it to exceed the repository's 2 MiB file and 5 MiB run limits; their complete ordering is
not load-bearing for the P2.5 claims. The compact JSONL keeps the discovery, traversal-bound, and focused self-test
rows in original order. `perf/perf_metrics_summary.json` binds the original SHA-256, byte and row counts, behavioral
result, retained corpus identity, representative metrics, and the explicit claim boundary.
