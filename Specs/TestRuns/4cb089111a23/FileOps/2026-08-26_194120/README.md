# File Operations Phase 5 integration gate

This Debug x64 run executed the complete File Operations suite for the P5.7 integration gate and recorded
**123 passed / 0 failed / 20 skipped**. The skips are retained environment/delegation outcomes (including optional
7z and unavailable remote-provider profiles), not hidden failures. `fileops_results.json` is the authority for the
complete per-case result; the trace files preserve execution order and terminal status.

The run covers the single per-item policy executor in serial and parallel scheduling, typed conflict and Keep Both
behavior, artifact touch authority, recovery/result projection, provider capability truth, and the former Phase 11
crash transition. Representative passing cases are recorded in `perf/perf_metrics_summary.json`.

The test harness was also hardened after two diagnosed Debug stack-exhaustion access violations. Reentrant UI phase
dispatchers are now split into bounded compiled helpers. The focused former-crash case passed in immutable run
`20260826T170937Z-99416-ef7f90f74ea14583bd14928bb3cba545`; this full run is
`20260826T172428Z-24428-b09b4e4a5d2b48dbb9af0cfb5e2b7b0c`. Crash dumps, hashes, PDB frame evidence, and
the continuation diagnosis are retained under
`Specs/TestRuns/4cb089111a23/Continuation/2026-08-26_p57_fileops_stack-exhaustion/`.

The full-solution Debug rebuild used for this run completed with zero warnings/errors (attestation
`78dd4609fd994c4203a20014af870a92d80850a81ec526bf8b93307ba9476178`; log
`.build/logs/msbuild-20260826_192004_649-pid125316-a2d80176.log`). A later closeout rebuild may supersede this
receipt for the final documentation-only tree; it does not change the immutable run evidence above.

## Metric compaction

The original `perf/perf_metrics.jsonl` was a 205,169,357-byte high-frequency Debug diagnostic stream with 557,980
rows and 507 unique metrics. It exceeded the repository's 2 MiB file and 5 MiB run limits and was removed from the
checked-in archive; the local generated source remains under `.build` and is reproducible by rerunning the suite.
`perf/perf_metrics_summary.json` preserves its SHA-256, corpus identity, row/metric counts, original size violations,
behavioral result, claim cases, and the ordered metrics directly associated with cancellation, scheduler,
artifact-authority, and bridge-pipeline evidence.

This compaction deliberately makes no final Release throughput, latency-budget, or live remote-provider claim.
