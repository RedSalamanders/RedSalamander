# I66 Create Folder worker admission

The x64 Debug runs below used the no-activation `FileOps_DeferredExitRefusesNewWork` selection on one machine. Both passed setup, the selected case, and cleanup. The canonical runner reported filtered coverage.

| Implementation | Commit | Local run | Build receipt |
|---|---|---|---|
| Synchronous qualification | `9979b38ca4366f13f321aa0358ae04c30f7668a7` | `20261124T203548Z-69376-a1dfaea0b5404b80aabb6d50eed82511` | `c1978acf2d546cf58169ebef4cc73e2909b2ba4ad174aadbecb7c4256f479f4c` |
| Queued Create Folder | `43d1de547e7a4b06c9b81feea6b6502e3081ba17` | `20261125T084628Z-79068-d5df149e2d4c44388f5f13b66eb21f20` | `1bf2ba9bb4011540567747b192dce397a36c671e61707ef4c40f30af7826c3cb` |

The queued build verified no folder creation before dispatch or after cancellation at the preparation barrier. It also preserved an occupied default name and created the numbered sibling. The existing exit refusal and recovery checks passed.

| Source row | Metric | Thread | Duration (µs) | Event |
|---|---|---:|---:|---|
| Synchronous 1587 | `fileops.create_directory.admit_us` | 59504 | 1622 | Accepted Local qualification |
| Queued 249 | `fileops.create_directory.admission.ui_thread_us` | 53484 | 742 | Published the task later canceled |
| Queued 281 | `fileops.create_directory.admission.ui_thread_us` | 53484 | 588 | Published the task that created a sibling |
| Queued 283 | `fileops.create_directory.admit_us` | 86436 | 3562 | Accepted Local qualification |
| Queued 334 | `fileops.create_directory.admit_us` | 86436 | 2245 | Requalified the occupied name |
| Queued 335 | `fileops.create_directory.admit_us` | 86436 | 3611 | Qualified the numbered sibling |
| Queued 336 | `fileops.create_directory.worker_us` | 86436 | 6935 | Two dispatches, one successful publication |

The figures distinguish UI-thread admission from the provider qualification and mutation on the worker. Different fixture paths and concurrent workloads limit a latency comparison. The slow live-provider walkthrough and native Release qualification remain open.

The untrimmed perf streams remain in the two completed-run archives under `artifacts/selftest/last_run/perf/perf_metrics.jsonl`. This record keeps every row from the Create Folder metric family in the original order.

| Source | Original bytes | Original rows | SHA-256 |
|---|---:|---:|---|
| Synchronous | 1383120 | 3769 | `375f473b62d165bdaa2ceecc525e481c8fca45321c6cc7d55aebba82a5f60bb9` |
| Queued | 1573739 | 4215 | `1e54c12bd9d5126fe7824563a2ba71eb8a5eed2ad759a8dd8fea30be7a8f3c11` |
