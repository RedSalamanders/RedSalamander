# File Operations P3.2b sparse semantic-link validation

This governed Debug x64 File Operations run recorded **117 passed / 0 failed / 20 documented skips**
and zero flaky, regression, or isolation classifications. `fileops_results.json` is the complete
per-case authority; `fileops_trace.txt` and `selftest_run_trace.txt` retain execution diagnostics.

The P3.2b claim is behavioral and bounded. Actual nested Keep Both choices feed a sparse component
mapping into the provider semantic transform. Backward references publish during the traversal;
forward references and cycles are held until their required destination names are known. A failed or
skipped target keeps the exact source dependency and containing source directories instead of
publishing a link with false meaning or deleting its source. The host retains only active directory
buffer views, sparse renamed components, deferred exact links, failed prefixes, and held cleanup
records under the shared 4,096-entry / 16-MiB traversal ceilings; it does not construct a whole-tree
manifest or dependency graph.

The moved-tree case proves a backward link, a forward link retargeted through `target (2)`, a cycle,
and a skipped forward target that retains the exact source dependency. The nested-conflict case proves
that a cached top-level Keep Both grant never leaks to a child while an explicit nested Keep Both
choice remains available. Existing Preserve/Skip and typed unsupported-provider behavior remain green.

Native-Move semantic qualification and the complete link-kind/profile matrix remain in P3.2c. This
archive does **not** claim throughput, latency-percentile, Release-runtime, remote-credential, or
production-workload improvement.

The executable was built from base commit `121be19dd6f024a2e8252d0380a8faaba1ee6771` plus the P3.2b
working-tree package finalized by the containing work-package commit. The accepted Debug build
receipt/attestation is `ff450bb548c2593cfb1d0dc5d37ff6b4ac733d243e8cef196457335aede0b468`.

## Metric compaction

The original `perf/perf_metrics.jsonl` was a 363,862,847-byte stream with 998,110 rows. High-frequency
popup, progress, identity, discovery, and per-transfer diagnostics are not load-bearing for P3.2b.
The compact JSONL keeps all `FileOps.SelfTest.*`, operation-strategy, bridge-copy, reparse-enumeration,
and semantic-link bounded-state rows in original order. `perf/perf_metrics_summary.json` binds the
original and compacted digests/counts, complete behavioral result, representative claim-bearing
metrics, and the explicit claim boundary.
