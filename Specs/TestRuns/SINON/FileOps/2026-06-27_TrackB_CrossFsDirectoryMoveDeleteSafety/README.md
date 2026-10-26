# Track B Cross-FS Directory Move Delete Safety Evidence

Date: 2026-06-27

Scope:
- `Specs/Plans/WIP/Operation_FSDeepAudit_TrackB_CrossFsDirectoryMoveDeleteSafety_2026-06-26.md`
- Cross-filesystem directory Move source cleanup after bridge copy.
- Regression coverage for source entries changed or created between copy completion and source cleanup.

Build:
- Command: `.\build.ps1 -ProjectName RedSalamander`
- Result: PASS, 0 errors, 0 warnings
- Log: `Build_RedSalamander/msbuild-20260627_100341_653.log`

Focused selftests:
- Command: `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Floodgate_CrossFsDirectoryMoveCleanupPreservesChangedSource`
- Result: PASS, 3 passed, 0 failed, 0 skipped
- Artifacts: `Floodgate_CrossFsDirectoryMoveCleanupPreservesChangedSource/`

- Command: `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Floodgate_CrossFs`
- Result: PASS, 5 passed, 0 failed, 0 skipped
- Artifacts: `Floodgate_CrossFs/`

- Command: `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Fairstream_CrossFsConcurrentMoveUsesBridge`
- Result: PASS, 3 passed, 0 failed, 0 skipped
- Artifacts: `Fairstream_CrossFsConcurrentMoveUsesBridge/`

- Command: `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Riptide_Bridge`
- Result: PASS, 6 passed, 0 failed, 0 skipped
- Artifacts: `Riptide_Bridge/`

- Command: `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter FileOpsFamily_Fairstream`
- Result: PASS, 39 passed, 0 failed, 0 skipped
- Artifacts: `FileOpsFamily_Fairstream/`

Notes:
- Each selftest artifact directory includes selected `last_run` pass/fail JSON and traces.
- The large generated `perf/perf_metrics.jsonl` stream from `%LOCALAPPDATA%\RedSalamander\SelfTest\last_run` was intentionally not archived because this closeout only needs compact pass/fail evidence and FileOps traces.
