# File Operations P1.6 integration evidence

This Debug x64 run exercised the complete File Operations suite while the P1.6 working tree was
under validation. `fileops_results.json` is the behavioral authority: **112 passed / 0 failed /
27 skipped**. The claim-bearing cases cover the provider capability matrix, reparse containment,
identity-owned staging, and Inline Rename worker/Queue/interlock behavior.

The retained-authority observation was recorded while an Inline Rename task was paused before
mutation. One active interlock scope retained 13 exact authority nodes, below the explicit
path-depth bound of 1,024. Adjacent `FileOps.Interlock.PrepareUs` rows record the same one-scope,
13-node rename path. This is bounded correctness/retention evidence only; no Release throughput,
percentile, latency-budget, or performance-improvement claim is made.

## Metric compaction

The original `perf/perf_metrics.jsonl` contained 904,572 rows and 331,368,741 bytes because the
Fresh Full process appended the broad diagnostic stream. It exceeded the repository's 2 MiB file
and 5 MiB run limits. The raw stream was moved out of the checked-in archive; the digest-bound
`perf/perf_metrics_summary.json` preserves its SHA-256, byte/row counts, corpus identity,
behavioral result, and the three ordered rows directly supporting the bounded-retention claim.

The raw stream was not load-bearing for a percentile, budget, throughput, or before/after claim,
so retaining unrelated high-frequency rows would add no review value and would violate the archive
contract.
