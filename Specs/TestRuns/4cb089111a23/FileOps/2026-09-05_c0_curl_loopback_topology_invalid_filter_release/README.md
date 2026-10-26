# Rejected FileOps predecessor filter

This invocation failed before any native case ran. It is **not predecessor
coverage** and must not be treated as a Copy/Move result. The runner reports
COUNT_MISMATCH, COVERAGE_INVALID and NO_CASE_RESULTS; FileOps reports an unknown
case/family filter. No native FileOps `results.json` was emitted, so none is invented.
The root result, native trace, aggregate and fingerprints are retained exactly.

The same candidate Release receipt is recorded in `build-provenance.json`.
The shared matcher supports comma-separated exact names, but the FileOps-specific
`ResolveRunSelection` does not. Until it is repaired and tested, use `C0_`, which
keeps these predecessors and Copy in one process. Do not work around this with
separate isolated processes and claim order-dependent coverage.

The invalid command, retained for diagnosis only:

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter 'C0_CurlDeleteTraversalBounds,C0_CurlDirectorySizeTruth,C0_CurlMovePreflightTruth,C0_CurlCopyTraversalTruth' -RunId c0-curl-loopback-topology-predecessor-release-20260905
```

The following [broad repeat](../2026-09-05_c0_curl_loopback_topology_owned_states_release/README.md)
actually executes the corpus and remains failed. See the [main checkpoint](../2026-09-05_c0_curl_loopback_topology_broad_release/README.md)
for the reverted experiment, provenance and next steps. Raw evidence stays beneath
the named `C:/RedSalamander.Perf/runs/` directory; copied-file hashes are in
`file-provenance.json`. No settings or secrets are archived.
