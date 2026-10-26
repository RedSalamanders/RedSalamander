# File Operations Phase 1 closeout validation

This Debug x64 run used build attestation
`0d942406474f0ea51e7a41a8f44bb688c5df3028d7875961046d910c0ccd5d06` and recorded
**112 passed / 0 failed / 27 skipped**. `fileops_results.json` is the authority for every case and
skip reason; `fileops_trace.txt` preserves execution order. The skips are declared P2/P3 work or
unavailable remote profiles, not failures hidden by this package.

The closeout adds one destination-resolution owner for execution and completion/cache/refresh
consumers. The source path profile derives the leaf, the destination path profile performs the join,
and explicit mappings remain exact. The 256-item interlock-retention scenario retained 270 unique
authority nodes with a maximum chain depth of 14, within the enforced 288-node and 1,024-depth bounds.

## Metric compaction

The original `perf/perf_metrics.jsonl` was a 139,664,718-byte full-suite diagnostic stream with
375,697 rows and 438 metric names. Its SHA-256 is
`83b2dee0767a4aa7a333b05fc1da20241304e1452203e1a9531555176cf4447a`. It exceeded the repository's
2 MiB file and 5 MiB run limits, so it is retained outside Git under
`.build/UnarchivedTestRuns/4cb089111a23/FileOps/2026-08-23_163451/` and replaced here by
`perf/perf_metrics_summary.json`. The summary preserves the original size, digest, row/metric counts,
corpus identity, behavioral authority, claim cases, and the exact representative interlock rows.

No throughput improvement or percentile budget is claimed. The 256-item measurements are a new
same-machine bounded-retention baseline; Release validation is recorded separately in the WIP
closeout evidence.
