# C0 native Curl Delete repair — test-enabled Release

- Source: `58f22d0e` plus the current C0 receipt and native Curl Delete changes.
- Build receipt: `889e8e92c09fb17f12e837256413e2b603b12a81426fb596fe7fbdb851853170`.
  `tests_enabled=true` was verified in the receipt; zero warnings/errors.
- Invocation: set process-scoped `RSBuildEnableTests=true`, run
  `Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -CaseFilter C0_ -FailFast -RunId c0-curl-native-delete-verified-release-20260905`,
  then restore the previous environment value. This run followed the Debug work;
  the earlier overlapping Release launch was declined by the shared-dependency lock.
- Harness: 7/7 passed, zero skips. Native Curl guard fixture: 29/29 checks passed.
- Results, trace, aggregate, and the complete 900,410-byte metric stream are unchanged
  copies from the run-owned `artifacts/selftest/last_run` directory.

All eight root requests retain both files with zero DELE/RMD commands. All eight
selected-subtree controls succeed with one DELE, two RMD, and a byte-checked surviving
sibling. Single and bulk Delete each cancel their stalled first LIST in 1,000 ms,
without fixture-forced unblock or mutation. The original Google Drive and Local
mutation-receipt cases remain green.

The Debug red/candidate pair provides same-profile before/after evidence. This
Release run proves the optimized configuration too; it is not a Debug-versus-Release
performance improvement claim. It does not prove real endpoints, deep/wide traversal,
ambiguous mutation replay safety, or final Full qualification. The ordinary non-test
Release staging transition observed earlier remains a separate C6 requirement.

Source/localization checks passed 197/197, spec inventory reported zero blockers,
and whole archive inventory passed 2,083 tracked files; each new ignored curated
archive is separately checked in explicit pre-stage mode. The runner disk audit
reports 32 retained older-run issues, not a clean-root result. No older data was removed.
