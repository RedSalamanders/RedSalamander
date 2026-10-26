# File Operations P2.7 post-close hardening validation

This governed Debug x64 File Operations run recorded **117 passed / 0 failed / 20 documented
skips** and zero flaky, regression, isolation, or unclassified-failure classifications.
`fileops_results.json` is the complete per-case authority; `fileops_trace.txt` and
`selftest_run_trace.txt` retain execution diagnostics.

The hardening claim is narrow: a raced Local directory destination is requalified only after the
provider reports a known non-commit; no redundant pathname shape probe runs before mutation; Local
shape qualification obtains the source/destination I/O interfaces at most once per admitted plan;
and an unavailable race fallback fails without retry while leaving source and destination contents
unchanged. The package preserves the P2.7 one-provider-mutation Native contract. It makes no
throughput, latency-percentile, Release-runtime, remote-credential, or production-workload claim.

The executable was built from commit `d6d44bab3c6f1434cb4076c52e486182723510ae` plus the
post-close hardening working tree finalized by the containing commit. The accepted Debug and
Release build receipts are `9203dc5252039c1893276a0baebdafeaaa755a437801ba033d0f9fade0c3542d`
and `b29dc9356b3d8a89a9e28a788d5fed242aab1752e3f85d7730a26970681ab734`.

## Metric compaction

The original `perf/perf_metrics.jsonl` was a 364,940,655-byte stream with 1,001,068 rows.
High-frequency popup, generic operation, and unrelated diagnostics are not load-bearing for this
hardening claim. The compact JSONL retains plan, strategy, Native-directory-race, traversal,
discovery, and `FileOps.SelfTest.*` rows in original order. `perf/perf_metrics_summary.json` binds
the original and compacted digests/counts, complete behavioral result, representative metrics, and
the explicit claim boundary.
