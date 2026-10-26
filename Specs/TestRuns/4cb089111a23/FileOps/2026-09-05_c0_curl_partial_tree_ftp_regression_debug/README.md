# C0 Curl partial-tree repair — ordinary FTP regression

- Source: same unchanged implementation/test/spec source as the passing C0 Debug
  run `c0-curl-partial-tree-qualified-debug-20260905`.
- Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete -FailFast -RunId c0-curl-partial-tree-ftp-regression-debug-20260905`.
- Build receipt: `c71f331a24ce0cd23df6ad7708a9984b1fbd9fb689820049d443a13f13a24713`;
  zero warnings/errors, build 2:34.
- Runtime: 3 passed, zero failed/skipped; native harness duration 3,380 ms.

The ordinary host workflow still copies to/from FTP, preserves exact bytes,
creates directories, renames, replaces and deletes. Timing observations: Copy to
FTP 250 ms, Copy from FTP 63 ms, Replace 219 ms, Delete 78 ms; configured provider
watchdog 8,000 ms. These are loopback integration observations, not a throughput
comparison or real-endpoint qualification.

Results, trace, aggregate and full perf stream are copied intact from the run-owned
`artifacts/selftest/last_run` beneath `C:\RedSalamander.Perf\runs`.
This warning-free build does not resolve the separately recorded intermittent
shared-output CoreClean warnings or the Release non-test SQLite staging gate.
