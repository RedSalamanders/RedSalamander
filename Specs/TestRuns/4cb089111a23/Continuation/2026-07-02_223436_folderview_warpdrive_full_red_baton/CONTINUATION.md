# FolderView WarpDrive Continuation Baton - Full Suite Red

Date: 2026-07-02 22:34 local
Branch: `codex/folderview-warpdrive`
HEAD at archive time: `275c04034`

## Current State

The FolderView/Commands convergence checkpoint is preserved and still important:

- Debug build is clean with 0 warnings / 0 errors:
  `.build\codex-runs\msbuild_after_nav_ime_failure_diagnostics_20260702.log`
- Broad Commands is green:
  `.build\codex-runs\commands_broad_after_nav_ime_diag_cleanbuild_20260702.log`
  (`776 passed / 0 failed / 2 skipped`, exit `0`, about 14m42s)
- No new RedSalamander crash-handler dump appeared after that final broad Commands run.

The follow-up Full skip-build gate is red:

```text
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

Evidence:

- Raw log: `run-artifacts\full_skipbuild_after_commands_green_20260702.log`
- Parsed summary: `full-failure-summary.txt`
- Selected volatile runner output copied from `%LocalAppData%\RedSalamander\SelfTest\last_run\` into `last-run\`
  (top-level logs/results, Commands/Compare/FileOps results and traces, plus perf summary JSONs).
  Large dumps and generated work folders are intentionally excluded and listed in
  `last-run\excluded-volatile-files.txt`.

Full result:

```text
passed=1091 failed=8 skipped=52 total=1151 duration_ms=2534653 exit_code=1
```

## Red Gates

From `last-run\run-all-tests-results.json` and the raw Full log:

- `CompareDirectories` failed 4 selftests:
  - `search_source_allocation_and_folding_guard`
  - `search_low_hardening_smoke`
  - `search_service_sqlite_cold_start_stale_root_refreshes_before_query`
  - `search_service_query_reports_live_progress` (`hr=0x800700E7`)
- `ViewerPETests` exited `1`; raw log shows `ViewerVLC exposes a visible HUD child that can take focus` failed.
- `RedConfigureTests` failed because `.build\x64\Debug\RedConfigureTests.exe` was not found.
- `RedSalamanderMonitorEtwLatency` exited `1`.
- `ToolsPesterTests` exited `1`.

Important: `Commands` stayed green inside Full (`776 passed / 0 failed / 2 skipped`), and `FileOperations` stayed green inside Full (`102 passed / 0 failed / 20 skipped`).

## Archive Contents

- `last-run\` - selected selftest runner artifacts from `%LocalAppData%\RedSalamander\SelfTest\last_run`; large dumps and generated work folders are listed but not copied.
- `run-artifacts\` - Full red log plus the last green Commands and clean build logs.
- `source-snapshots\` - current WIP plan and the focused source files most relevant to the recent convergence work.
- `patches\` - staged and unstaged diffs excluding `Specs\TestRuns`.
- `crash-sidecars\` - latest RedSalamander crash-handler sidecar manifest and recent `.txt` callstack files. Large `.dmp` files were intentionally not copied.
- `git-*.txt` - branch, HEAD, status, and focused diff stats/name-status snapshots.
- `archive-manifest.txt` - file listing for this baton.

## Next Convergence Plan

1. Start with the Full red gates, not FolderView code:
   - Run focused CompareDirectories failures first.
   - Run `ViewerPETests.exe` directly and inspect the ViewerVLC HUD focus assertion.
   - Decide whether `RedConfigureTests.exe` should be built by the Full prebuild or removed from the current expected Full matrix.
   - Run `RedSalamanderMonitorEtwLatency.exe` directly to capture its stderr/stdout.
   - Run the Tools Pester suite directly to expose the failing assertion rather than only the wrapper exit code.
2. After the red gates are handled, rerun:
   ```text
   .\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
   ```
3. Then rerun the final FolderView perf matrix in Debug and test-enabled Release from the current HEAD, archive/cite both runs, update the acceptance checklist, and only then move the WIP plan to `Specs\Plans\Done`.

Do not move `Specs\Plans\WIP\Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md` to Done from this baton. The Full suite is still red and the final perf matrix has not been refreshed after the latest convergence edits.
