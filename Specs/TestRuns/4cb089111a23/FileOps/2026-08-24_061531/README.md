# File Operations P3.1 link-policy cutover validation

This governed Debug x64 File Operations run recorded **116 passed / 0 failed / 21 documented skips**
and zero flaky, regression, or isolation classifications. `fileops_results.json` is the complete
per-case authority; `fileops_trace.txt` and `selftest_run_trace.txt` retain execution diagnostics.

The P3.1 claim is behavioral and bounded: the Local provider exposes only Preserve/Skip, legacy
`followTargets` silently migrates to Skip, legacy `copyReparse` silently migrates to Preserve,
the immutable per-call policy overrides the provider default, Skip never follows or mutates the
target, and reserved/invalid ABI policy values fail before provider I/O. Native link relocation,
destination containment, link conflict consent, and bridge reparse handling remain green. The one
product-semantic skip is in-tree semantic retargeting, explicitly owned by P3.2. This archive does
**not** claim throughput, latency-percentile, Release-runtime, remote-credential, or production-
workload improvement.

The executable was built from base commit `a2c6a5d25753e264e8d884d926cf146be719d27a` plus the P3.1
working-tree package finalized by the containing work-package commit. The accepted Debug build
receipt/attestation is `79ab5a7f9c060ddac3d4685d820f2d5c3fd645e35a72e80224c5c140419cf40e`.

## Metric compaction

The original `perf/perf_metrics.jsonl` was a 132,643,720-byte stream with 360,246 rows. High-frequency
popup, progress, identity, discovery, and per-delete diagnostics are not load-bearing for P3.1. The
compact JSONL keeps all `FileOps.SelfTest.*`, operation-strategy, reparse/link-policy, Riptide, and
Causeway rows in original order. `perf/perf_metrics_summary.json` binds the original and compacted
digests/counts, complete behavioral result, representative claim-bearing metrics, and the explicit
claim boundary.
