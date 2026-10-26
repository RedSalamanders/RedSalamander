# File Operations P3.2c native semantic-Move qualification

This governed Debug x64 File Operations run recorded **117 passed / 0 failed / 20 documented skips**
and zero flaky, regression, isolation-suspect, or unclassified-failure results. `fileops_results.json`
is the complete per-case authority; `fileops_trace.txt` and `selftest_run_trace.txt` retain execution
diagnostics.

The P3.2c claim is behavioral. Local Native Move remains available for same-volume regular files, but
Local tree and top-level link objects advertise no native semantic transform and are classified as
Managed before mutation. The destination-absent Local tree case contains an absolute in-tree junction;
the host publishes the retargeted link under the destination root and removes the source only through
the retained bound authority. Existing merge, sparse Keep Both, forward/cycle deferral, Preserve/Skip,
Copy-only source retention, and typed unsupported-provider behavior remain green.

Capability profiles that cannot enumerate link objects may advertise the transform vacuously; profiles
that can expose link objects must prove the transform before Native admission. `ACCESS_DENIED` is not a
Native directory-race requalification signal, and Local Native race fallback is Managed-only.

This archive does **not** claim throughput, latency-percentile, Release-runtime, remote-credential, or
production-workload improvement.

The executable was built from base commit `23f14c45cf0ae5707818c4bb8d1f191dfb519a92` plus the P3.2c
working-tree package finalized by the containing work-package commit. The accepted Debug build
receipt/attestation is `c77835b7e26ab9727b866343d9cff390b7ea5347f55de311004a8ec11bdc18b1`.

## Metric compaction

The original `perf/perf_metrics.jsonl` was a 356,765,833-byte stream with 978,112 rows. High-frequency
popup, progress, identity, discovery, and per-transfer diagnostics are not load-bearing for P3.2c.
The compact JSONL keeps all `FileOps.SelfTest.*`, operation-strategy, Native-semantic qualification,
Native-race fail-closed, bridge-link, and recursive-reparse rows in original order.
`perf/perf_metrics_summary.json` binds the original and compacted digests/counts, complete behavioral
result, representative claim-bearing metrics, and the explicit claim boundary.
