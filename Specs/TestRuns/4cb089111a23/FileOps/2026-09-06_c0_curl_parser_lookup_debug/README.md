# C0 strict-listing parser: candidate Debug lookup

Base: `3917fef5`; `tested.patch` and `tested-source.json` retain the exact ten-file
candidate before final evidence/plan updates. Full test-enabled solution receipt:
`4ce3b8d38dad1e67204b067e6e2a6d1c594fa2da3c1972e53c9b8d91cab96f69`.
Snapshot: `b59e77718ac62ce6212984e1fa92efc59f2aa92283b4b9f1fd64b123fa44dbf4`.
Build: 187,025 ms, zero warnings/errors. Source/resource checks: 197 passed,
zero failed/skipped.

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild
-TestRoot C:\RedSalamander.Perf -CaseFilter C0_CurlEntryLookupTruth
-RunId c0-curl-parser-lookup-debug-20260906`.

Result: 294/294 native assertions, suite 3 passed / 0 failed / 0 skipped.
The six previously failing fragmented DOS numeric-token assertions now pass.
The production parser preserves only reusable filename storage, resets metadata,
and uses the shared strict UTF converter; it still validates every listing row.

Each parser-only scenario validates 262,208 rows (64 x 4,097), with 64 exact matches
or zero matches. Unix present/missing: 626,350 / 552,519 us; DOS present/missing:
484,903 / 476,961 us. Versus the adjacent pre-repair Debug baseline, these are
approximately 25% / 32% / 35% / 35% faster on this machine. Single-run timings
are directional; neither this microbenchmark nor focused lookup closes broad C0.

Exact-copy provenance, ordered metric selection and build-log normalization are
described by the adjacent JSON manifests. Raw evidence remains beneath the named
owned test run. Standalone shared-converter coverage and broad runs are separate
required witnesses.
