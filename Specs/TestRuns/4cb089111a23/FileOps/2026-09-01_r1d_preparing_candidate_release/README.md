# R1d common Preparing candidate

- Machine profile: `4cb089111a23`
- Source commit: `668acbd7f1a94d9a770b8a6fd0f746e214592f47` (R1d implementation)
- Build: test-enabled x64 Release; direct MSBuild completed with `0 warning(s), 0 error(s)`
- Command: `RedSalamander.exe --fileops-selftest --selftest-case=FileOpsFamily_R1dPreparingLifecycle --perf --selftest-no-activate`
- Runs: `r1d-preparing-candidate-668acbd7-01` through `-05`
- Scenario: five sequential independent processes; each executes the common Preparing lifecycle over 4,096 selected top-level roots and 8,192 scope facts without descendant traversal or provider content I/O
- `FileOps.Preparing.BuildSnapshotUs`: `1,378`, `1,377`, `1,555`, `1,440`, and `1,792` us; nearest-rank p95 `1,792` us
- `FileOps.Preparing.SelectedRootReadinessUs`: `6,144`, `6,149`, `6,420`, `6,310`, and `7,192` us; nearest-rank p95 `7,192` us
- `FileOps.Preparing.UiAdmissionUs` process maxima: `2,978`, `3,047`, `2,180`, `2,917`, and `1,926` us; nearest-rank p95 `3,047` us
- Retained immutable preparation state: `2,802,208` bytes in every process
- Ordering counters: mutation, clipboard consumption, and breadcrumb writes before Ready are all zero in all five processes

All five candidate processes passed `3/0/0` with no skips. Snapshot p95 passes the `52,785` us ceiling, readiness p95 passes the `57,433` us ceiling, UI admission p95 is `1,976` us above baseline and passes the `26,071` us ceiling, every protected duration is below `2,000,000` us, and retained state is below 64 MiB. The case also proves the slow 500 ms Preparing presentation threshold, routine clean completion without popup creation, legal phase order, Copy-only acceptance, dense cancel/failure truth, and zero mutation/clipboard/breadcrumb work before Ready. This is top-level lifecycle overhead evidence, not recursive traversal or provider-I/O throughput evidence.
