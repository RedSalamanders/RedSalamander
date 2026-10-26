# C0 parser repair: Release lookup and build

Base `3917fef5` plus the exact patch/hashes retained in
`../2026-09-06_c0_curl_parser_lookup_debug/`. Full test-enabled Release receipt:
`4357fa43d22bade20de39b19f7d13e3a47f499fc3f1ff1b6388e600cc56e14fd`.
Snapshot `b59e77718ac62ce6212984e1fa92efc59f2aa92283b4b9f1fd64b123fa44dbf4`
matches Debug. Build: 250,183 ms, zero warnings/errors, including raw log review.

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild
-TestRoot C:\RedSalamander.Perf -CaseFilter C0_CurlEntryLookupTruth
-RunId c0-curl-parser-lookup-release-20260906`.

Passed 294/294 native assertions; suite 3 passed / 0 failed / 0 skipped,
24.3 seconds (30.3 seconds runner wall time).
The four parser-only scenarios validate 262,208 rows each with 64 or zero matches:
Unix present/missing 19,989 / 24,456 us; DOS present/missing 20,367 / 25,127 us.
These are Release observations, not a comparison against the Debug baseline.

Adjacent manifests bind exact raw copies, ordered metric selection, the build
receipt projection and normalized build log. This focused run does not replace
the broad Release case or final Fresh Full.
