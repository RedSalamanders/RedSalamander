# File Operations P3.3 typed conflicts and semantic Copy

This governed Debug x64 File Operations run recorded **117 passed / 0 failed / 20 documented skips**
and zero flaky, regression, isolation-suspect, or unclassified-failure results. `fileops_results.json`
is the complete per-case authority; `fileops_trace.txt` and `selftest_run_trace.txt` retain execution
diagnostics.

The P3.3 claim is behavioral. Regular-file, read-only, type-mismatch, exact-destination-link, name,
and retry conflicts have typed action sets. Apply-to-all receipts use the exact conflict scope; plain
Skip is one-shot and Skip all is explicit. Same-pane Copy creates a Keep Both sibling, same-folder
Move remains rejected, and Local recursive Copy applies child-level Keep Both decisions.

Preserved Local links are deferred under the shared traversal ceilings. File symlinks, directory
symlinks, and junctions are reconstructed from no-follow semantic payloads, and actual child Keep
Both decisions populate the sparse transform before forward links publish. A skipped destructive
link/type mismatch reports Partial and never damages the existing destination.

This archive does **not** claim throughput, latency-percentile, Release-runtime, remote-credential,
or production-workload improvement.

The executable was built from base commit `a16b8444f72c77e57c338681a2c6c044f87d6ac2` plus the P3.3
working-tree package finalized by the containing work-package commit. The accepted Debug build
receipt/attestation is `3ff80344402dff244ce699a78b7e4eeb1c88e32235959ea95152272bfda361f2`.

## Metric compaction

The original `perf/perf_metrics.jsonl` was a 359,678,338-byte stream with 986,474 rows. High-frequency
popup, progress, identity, discovery, and per-transfer diagnostics are not load-bearing for P3.3.
The compact JSONL keeps all `FileOps.SelfTest.*`, typed-conflict, operation-strategy, semantic-route,
bridge-link, and recursive-copy rows in original order. `perf/perf_metrics_summary.json` binds the
original and compacted digests/counts, complete behavioral result, representative claim-bearing
metrics, and the explicit claim boundary. The complete original stream remains in the local raw run
`20260824T164412Z-53816-42e1371de8f541ff9c943d50e65cd618`.
