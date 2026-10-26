# FileOps Fairstream/Floodgate family revalidation

This Debug x64 run executed `FileOpsFamily_Fairstream` at commit
`ff899c1c14d3d645157e0dcfd1b9b6821054b281` and recorded **47 passed / 0 failed / 0 skipped**.
`fileops_results.json` is the authority for the complete per-case result and `fileops_trace.txt` preserves execution
order.

The load-bearing Floodgate claim is that cross-filesystem MOVE cleanup is suppressed for the complete tree after a
copy-phase partial result: the parallel eight-child case retained all sources while seven valid destinations existed.
The focused serial source/destination size-failure, destination-corruption, directory-mutation, provider-output, and
scheduler/resource cases are also retained in the result and trace files.

Microsoft Drive advertises no same-provider copy, so the host uses this provider-neutral bridge for Microsoft Drive
copy. This archive supports the shared host-bridge failure/cleanup contract applicable to that route; it **does not
claim live Microsoft Graph coverage** or Microsoft Drive server-side move coverage.

## Metric compaction

The original `perf/perf_metrics.jsonl` was a 7,009,000-byte high-frequency diagnostic stream with 17,190 rows and no
percentile, latency-budget, or before/after claim. It exceeded the repository's 2 MiB file and 5 MiB run limits and was
removed during package-16 archive remediation. `perf/perf_metrics_summary.json` preserves its SHA-256, corpus identity,
row/metric counts, original size violations, behavioral result, and the six ordered metrics directly associated with
the Floodgate cleanup claim. The complete behavioral case ordering remains in the compact result/trace artifacts.

This compaction intentionally does not preserve every render, popup, scheduler, and progress sample because those rows
were not load-bearing evidence for this run's behavioral claim.
