# File Operations P3.4 identity-safe mutation phase gate

This governed Debug x64 File Operations run recorded **117 passed / 0 failed / 20 documented skips**
and zero flaky, regression, isolation-suspect, or unclassified-failure results. `fileops_results.json`
is the complete per-case authority; `fileops_trace.txt` and `selftest_run_trace.txt` retain execution
diagnostics.

The P3.4 claim is behavioral. Destructive conflict decisions carry an exact no-follow destination
authority into one conditional mutation. Link replacement, read-only replacement, directory/link
publication, and post-copy metadata writes never authorize mutation from pathname text alone.
Unbound metadata paths emit an authority-unavailable diagnostic and keep provider defaults rather
than reopening the pathname.

The complete tree/link matrix proves that folder-on-folder remains merge, a source link over an
ordinary directory is a typed mismatch, destination links are never implicit merge namespaces, and
semantic Move retargeting publishes mapped links before exact source cleanup. The staged-overwrite
fault cases exercise the live owned-stage `PublishAs` boundary and prove that pre-publication failures
preserve the original destination bytes and attributes.

This archive does **not** claim throughput, latency-percentile, Release-runtime, remote-credential,
or production-workload improvement.

The executable was built from base commit `82c175036a0791768d3d8c2335cde047eb5dcfc3` plus the P3.4
working-tree package finalized by the containing work-package commit. The accepted Debug build
receipt/attestation is `da7f142e3ecb76f28755a7c9b35f7d569c908c8203b55f4b4f8fa3f94a012ef6`.

## Metric compaction

The original `perf/perf_metrics.jsonl` was a 143,518,966-byte stream with 389,485 rows. High-frequency
popup, progress, identity, discovery, and per-transfer diagnostics are not load-bearing for P3.4.
The compact JSONL keeps all `FileOps.SelfTest.*` rows plus aggregate operation, exact-destination,
conditional-link-replacement, semantic-link, and owned-publication fault rows in original order.
`perf/perf_metrics_summary.json` binds the original and compacted digests/counts, complete behavioral
result, representative claim-bearing metrics, and the explicit claim boundary. The complete original
stream remains in local run `20260824T224940Z-54068-2e418dbae5a14677a21adda6aba2b9dc`.
