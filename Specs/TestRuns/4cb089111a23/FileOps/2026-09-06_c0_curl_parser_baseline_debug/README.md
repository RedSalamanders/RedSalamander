# C0 strict-listing parser: pre-repair Debug baseline

Base: `3917fef5`; exact six-file test/instrumentation delta is retained in
`tested.patch` and `tested-source.json`. Build receipt:
`70c80befbea745b9f037534b3954de559362c1c452cb7ead63ced329d3e6bbe7`.
No production parser/converter repair is included in this baseline.

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild
-TestRoot C:\RedSalamander.Perf -CaseFilter C0_CurlEntryLookupTruth
-RunId c0-curl-parser-baseline-debug-20260906`.

Expected red result: 288 passed / 6 failed of 294 native assertions; suite
2 passed / 1 failed / 0 skipped. The only failing assertions are rejection of
DOS sizes `4junk` and `4.5`, each under three callback-fragmentation patterns.
The full-solution build exited zero in 186,894 ms. Its summary counted zero
warnings, but the raw log contains an unnumbered vcpkg app-local dependency warning
for ViewerPETests after a file-sharing error. This is not a warning-free build;
C6 retains the staging/diagnostic-accounting issue.

Each parser-only scenario performs 64 complete 4,097-row scans: 262,208 validated
rows, with 64 exact matches or zero matches. Unix present/missing took
834,239 / 812,691 us; DOS present/missing took 744,995 / 732,893 us.
All four correctness controls passed. Timings include the real bounded cursor
callback/framing/parser and construction, but exclude FTP and fixture generation.
These are single-run, same-machine comparison inputs, not final qualification.

Exact-copy provenance, ordered metric selection and build-log normalization are
described by the adjacent JSON manifests. Raw evidence remains beneath the named
owned test run; failed evidence is deliberately preserved.
