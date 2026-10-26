# R0a Local direct-final baseline

- Machine profile: `4cb089111a23`
- Build: test-enabled x64 Release, receipt `48f4f24a7cbc57a5f49949c3f68bfd0e6bbd941ff305ec00bfc73f7edb23a11b`
- Variant: pre-R0a `CopyFileExW(..., COPY_FILE_FAIL_IF_EXISTS)` success route
- Scenario: five independent processes running `Floodgate_LocalCopyNewNameConcurrentReplacementSurvives`
- Metric: `FileOps.SelfTest.R0aLocalDirectFinalCopyUs`, 64 MiB ordinary absent-destination Local regular-file Copy
- Samples: `28,374`, `21,200`, `20,891`, `19,495`, `19,079` us
- Nearest-rank p95: `28,374` us

The race assertion is intentionally RED in these archives: the pre-R0a route has no retained exact-object seam and deletes the injected foreign pathname replacement. The successful 64 MiB copy occurs before that RED assertion and is the baseline measurement.
