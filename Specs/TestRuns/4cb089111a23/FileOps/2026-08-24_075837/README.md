# File Operations P3.2a semantic-link publication validation

This governed Debug x64 File Operations run recorded **117 passed / 0 failed / 20 documented skips**
and zero flaky, regression, or isolation classifications. `fileops_results.json` is the complete
per-case authority; `fileops_trace.txt` and `selftest_run_trace.txt` retain execution diagnostics.

The P3.2a claim is behavioral and bounded. The Local provider reads symlink and junction semantics
from the exact retained no-follow object, the host retargets selected-root `S -> D` references for
Managed tree Move, creates an exclusive destination link stage, conditionally publishes that exact
stage, and only then performs exact source cleanup. The moved-tree case proves that a junction aimed
inside the selected source tree resolves inside the destination tree after a merge and that the source
tree is removed. Existing Preserve/Skip, destination-containment, native relocation, and typed
unsupported-provider behavior remain green.

Sparse Keep Both mappings, deferred forward dependencies, native-Move semantic qualification, and
the complete link-kind/profile matrix remain in P3.2b/P3.2c. This archive does **not** claim throughput,
latency-percentile, Release-runtime, remote-credential, or production-workload improvement.

The executable was built from base commit `ca911ce7c73e74e90342604106bc044da1e660cb` plus the P3.2a
working-tree package finalized by the containing work-package commit. The accepted Debug build
receipt/attestation is `ad6b105300bc9e20e0a66df93f3628b84dd0dbcbe1fc7a0513798ebcdc5fd6f7`.

## Metric compaction

The original `perf/perf_metrics.jsonl` was a 358,959,132-byte stream with 985,008 rows. High-frequency
popup, progress, identity, discovery, and per-transfer diagnostics are not load-bearing for P3.2a.
The compact JSONL keeps all `FileOps.SelfTest.*`, operation-strategy, reparse/link, Riptide, Causeway,
and Fairstream rows in original order. `perf/perf_metrics_summary.json` binds the original and
compacted digests/counts, complete behavioral result, representative claim-bearing metrics, and the
explicit claim boundary.
