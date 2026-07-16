# Operation Granite Continuation Baton - 2026-07-05

## Final Archive Note - 2026-07-06

Granite is closed. The main Granite plan was moved to
`Specs/Plans/Done/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
after the final closeout audit found no live Granite rows and the latest Full
skip-build artifact reported `1117` passed / `0` failed / `52` skipped.

This baton is archived only as historical resume context. Do not resume Granite
from the older save points below; use the WIP index for the next active owner.

## Latest Save Point - 2026-07-05 22:28 local

User asked to save state before continuing. This section supersedes every older
save point for immediate resume.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **ready for final closeout, not yet moved to Done.**

Authoritative current-state checks from this pause:

- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
  exists.
- `Specs/Plans/Done/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
  does not exist yet.
- Live-row scan returned no matches:
  ```powershell
  rg --pcre2 -n "^\| GR-[^|]+ \| (?!.*(DONE|STALE|SIGNED OFF|DECIDED|ROUTED))|^\| GR-T[0-9A-Za-z-]+ \| (?!.*(DONE|STALE))|^\| GR-TA[0-9A-Za-z-]+ \| (?!.*DONE)" Specs\Plans\WIP\Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
  ```
- Latest Full skip-build artifact remains green:
  `C:\Users\eric\AppData\Local\RedSalamander\SelfTest\last_run\run-all-tests-results.json`
  reports suite `Full`, exit code `0`, `1117` passed / `0` failed / `52`
  skipped / `1169` total, duration `2612504` ms.
- `Specs/Plans/WIP/README.md` already says Granite is **Closeout ready** and
  points the next action at final `git diff --check`, final closeout note,
  move to `Specs/Plans/Done/`, and WIP-index removal.

Work intentionally not done yet:

- Granite has not been moved to `Specs/Plans/Done/`.
- The active goal has not been marked complete.
- The Granite live row has not yet been removed from `Specs/Plans/WIP/README.md`;
  do that only after the move succeeds.

Immediate next move on resume:

1. Run final `git diff --check` for the touched Granite closeout files and the
   three blocker-fix files:
   `Tests/DxUiTests/DxUiTests.NewControls.cpp`, `Tools/Show-PerfRuns.ps1`,
   `Tests/ViewerPETests/ViewerPETests.cpp`,
   `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`,
   `Specs/Plans/WIP/README.md`, and the Granite plan.
2. Add a final Full-green closeout note to the Granite plan if it still lacks
   the exact Full evidence above.
3. Move the Granite plan from `Specs/Plans/WIP/` to `Specs/Plans/Done/`.
4. Remove Granite from the live-plan table in `Specs/Plans/WIP/README.md` after
   the move; keep `Operation_FileOperations_FaultInjectionRatchet_2026-07-05.md`
   as the follow-up owner for the remaining long-term GR-A5 ratchet.
5. Re-run the live-row/file-location checks and mark the active goal complete
   only if the current evidence proves the closeout.

## Latest Save Point - 2026-07-05 22:27 local

This section supersedes every older save point below for immediate resume.
Granite is still active only because the closeout paperwork/move has not been
done yet. Do not re-open the older DxUi/ToolsPester/Monitor/ViewerPE blockers
unless a fresh verification run fails again.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **Granite is ready for final closeout.** All Granite rows are
closed, routed, or stale, and the latest Full skip-build gate is green.

Latest Full gate:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -FailFast -TimeoutMultiplier 3
```

- Result file:
  `C:\Users\eric\AppData\Local\RedSalamander\SelfTest\last_run\run-all-tests-results.json`
- Suite: Full
- Started UTC: `2026-07-05T19:41:26`
- Ended UTC: `2026-07-05T20:24:58`
- Duration: `2612504` ms
- Result: `1117` passed / `0` failed / `52` skipped / `1169` total
- Exit code: `0`

Blockers fixed since the previous save point:

- `Tests/DxUiTests/DxUiTests.NewControls.cpp`: the stale TabControl close-button
  tests now select tab 1 before clicking tab 1's visible close button. Verified
  with a Debug x64 `DxUiTests` build, `DxUiTests.exe --suite=NewControls`, and
  the full `DxUiTests.exe` native run.
- `Tools/Show-PerfRuns.ps1`: `Get-FolderViewBudgetMinimumSamples()` now reads
  both the historical top-level `budgets` shape and the current
  `machines[].budgets[]` shape, so low-sample distribution metrics are again
  quality-gated. Verified with focused `ShowPerfRuns.Tests.ps1` and the full
  `Tools\Tests` Pester bundle.
- `RedSalamanderMonitorEtwLatency`: direct repro of the Full command exited 0;
  the latest Full gate also records exit 0. The empty output log is not itself a
  failure.
- `Tests/ViewerPETests/ViewerPETests.cpp`: the VLC focus-transfer test now waits
  for visible video/HUD surfaces instead of asserting while the loading overlay
  can legitimately hide the HUD. Verified with a Debug x64 `ViewerPETests`
  build, focused `TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly`,
  and the full `ViewerPETests.exe` native run.

Granite live-row scan at this save point:

```powershell
rg --pcre2 -n "^\| GR-[^|]+ \| (?!.*(DONE|STALE|SIGNED OFF|DECIDED|ROUTED))|^\| GR-T[0-9A-Za-z-]+ \| (?!.*(DONE|STALE))|^\| GR-TA[0-9A-Za-z-]+ \| (?!.*DONE)" Specs\Plans\WIP\Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
```

Result: no matches.

Files intentionally touched in this final-gate save slice:

- `Tests/DxUiTests/DxUiTests.NewControls.cpp`
- `Tools/Show-PerfRuns.ps1`
- `Tests/ViewerPETests/ViewerPETests.cpp`
- `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`
- `Specs/Plans/WIP/README.md`

Immediate next move:

1. Run `git diff --check` on the touched files above.
2. Add the final Full-green closeout note to
   `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
   if the plan does not already carry the exact Full evidence above.
3. Move
   `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
   to `Specs/Plans/Done/`.
4. Update `Specs/Plans/WIP/README.md` after the move so Granite no longer appears
   in the live-plan table. Keep
   `Specs/Plans/WIP/Operation_FileOperations_FaultInjectionRatchet_2026-07-05.md`
   as the owner for the remaining long-term GR-A5 ratchet.
5. Mark the active Granite goal complete only after the file move and WIP index
   audit are verified.

The dirty tree is intentionally large. Do not revert unrelated files; inspect
`git status --short` before editing and work with the existing Granite/WarpDrive
changes.

## Latest Save Point - 2026-07-05 20:42 local

This section supersedes every older save point below for immediate resume.
Granite is still active and must stay in WIP. Do not call the goal complete,
and do not move the Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **all Granite rows are closed, routed, or stale, but final
closeout is blocked by the Full gate.** The next work is test-gate triage, not
another Granite row.

Newest archived evidence:

- `Specs/TestRuns/4cb089111a23/FileOps/2026-07-05_203501/` passed 102 / 0 / 20
  in 796442 ms.
- `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_202143/` passed 782 / 0 / 2
  in 1123415 ms.
- `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_194106/` passed
  219 / 0 / 30 in 345113 ms.
- In the Full gate, the Compare slice
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_200259/` failed at
  `search_service_sqlite_status_reports_maintenance_history` with
  `SearchServiceBroker::GetStatus for maintenance-history SQLite service failed.
  hr=0x80070002`.
- Focused repro for that Compare failure is now green:
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_204103/` passed
  `search_service_sqlite_status_reports_maintenance_history` 1 / 0 / 0.

Latest Full-gate blocker list from the terminal capture:

- `CompareDirectories`: `search_service_sqlite_status_reports_maintenance_history`
  failed only in the Full-order run; focused repro passed afterward.
- `DxUiTests`: `FAILED: tab control handles close-button release`.
  Direct repro also fails with:
  `.\.build\x64\Debug\DxUiTests.exe --suite=NewControls`.
- `RedSalamanderMonitorEtwLatency`: exited 1 with an empty
  `C:\Users\eric\AppData\Local\RedSalamander\SelfTest\last_run\RedSalamanderMonitorEtwLatency.output.log`.
- `ToolsPesterTests`: exited 1; no useful output log was found in `last_run`, so
  run Pester directly next.

DxUi root cause already triaged:

- Failing test: `Tests/DxUiTests/DxUiTests.NewControls.cpp`, function
  `TestTabControlCloseButtonRemovesTabsAndInvokesCallback()`.
- The test makes tab 1 closable, fetches its debug close-button rectangle, and
  clicks it while tab 1 is neither selected nor hovered.
- Production `TabControl::IsCloseButtonVisible(...)` only exposes the close
  button when the tab is selected or hovered. So the press starts as a normal
  tab press, selects tab 1, and the release is then over the now-visible close
  button. CW-7's stricter contract correctly refuses to close because the press
  did not begin on the close button.
- This looks like a stale test setup, not a production regression. Minimal next
  edit: add `tabControl->SetSelectedIndex(1u);` after
  `tabControl->SetTabClosable(1u, true);` in that test before the close-button
  geometry/click. The RED is already captured by the direct failing NewControls
  run.

Immediate next move:

1. Patch `Tests/DxUiTests/DxUiTests.NewControls.cpp` so
   `TestTabControlCloseButtonRemovesTabsAndInvokesCallback()` selects tab 1
   before clicking tab 1's close button.
2. Rerun:
   ```powershell
   .\.build\x64\Debug\DxUiTests.exe --suite=NewControls
   .\.build\x64\Debug\DxUiTests.exe
   ```
3. Inspect and run the Tools Pester bundle directly, likely starting with:
   ```powershell
   Invoke-Pester -Script .\Tools\Tests -PassThru -Quiet
   ```
   If that is too broad/noisy, inspect `Tools/TestRunPlan.ps1` for the exact
   Full-gate ToolsPester command.
4. Inspect `Tools/TestRunPlan.ps1` around the
   `RedSalamanderMonitorEtwLatency` entry, run that native command directly,
   and capture why it exits 1 with an empty log.
5. Rerun `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -FailFast
   -TimeoutMultiplier 3`.
6. Only after Full is green, update the Granite plan and WIP index, then move
   Granite to `Specs/Plans/Done/`.

The dirty tree is intentionally large. Do not revert unrelated files; inspect
`git status --short` before editing and work with the existing Granite/WarpDrive
changes.

## Latest Save Point - 2026-07-05 19:17 local

This section supersedes every older save point below for immediate resume. Granite
is still active and must stay in WIP. Do not call the goal complete, and do not
move the Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **all Granite rows are closed, routed, or stale, but final
closeout is blocked by the broad test gate.** GR-12 is fixed, GR-1/GR-T1 are
closed as stale, and GR-A5's long-term ratchet remainder is routed to
`Specs/Plans/WIP/Operation_FileOperations_FaultInjectionRatchet_2026-07-05.md`.
Do not archive Granite until the broad gate failures below are triaged and
re-verified.

Files touched by the latest save slice:

- `RedSalamander/IconCache.cpp`
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
- `Tools/Tests/TestInventory.Tests.ps1`
- `Tests/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/UI/UI_FolderView.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/Operation_Tailwind_FolderViewWarpDrivePreMergeReviewRemediation_2026-06-30.md`
- `Specs/Plans/WIP/README.md`
- `Specs/Plans/WIP/Operation_FileOperations_FaultInjectionRatchet_2026-07-05.md`
- `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`

Latest implementation/documentation state:

- **GR-12 DONE 2026-07-05.** `IconCache` path failure-store races now use
  insert-only `emplace(...)` semantics so a slow failed attribute-mode shell
  query cannot downgrade an existing successful icon-index cache entry to
  `lookupFailed=true`. Losing duplicate races emit
  `iconcache.duplicate_path_query_race`.
- Added source-contract guard
  `keeps IconCache path failure stores from downgrading concurrent successes`.
- Tooling Pester inventory and coverage docs were bumped from 138 to 139 cases,
  and `TestHarnessSourceContracts.Tests.ps1` from 47 to 48 documented cases.
- `Specs/UI/UI_FolderView.md` now carries the durable IconCache failure-store
  contract.
- Granite GR-1 and GR-T1 are closed as stale/current-tree-false: current source
  already resets `_pathEditBlurSuppressActive` after the focus suppression
  window in both relevant path-edit flows.
- Granite GR-A5 is closed for Granite and routed for long-term work to
  `Operation_FileOperations_FaultInjectionRatchet_2026-07-05.md`. That follow-up
  starts from the current baseline: 31 `ForSelfTest` occurrences and 30
  `#ifdef ENABLE_TESTS` blocks in
  `RedSalamander/FolderWindow.FileOperations.State.cpp`, plus 629
  `ForSelfTest` occurrences and 647 `#ifdef ENABLE_TESTS` blocks under
  `RedSalamander/`.

Focused evidence before the broad gate:

- RED source-contract Pester for GR-12: 51 / 1 / 0 against the old
  `_pathToIconIndex.insert_or_assign(...)` failure-store block.
- GREEN source-contract Pester after the GR-12 fix:
  `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` passed 52 / 0 / 0.
- GREEN inventory Pester:
  `Invoke-Pester -Script .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
  -Quiet` passed 5 / 0 / 0.
- GREEN app build:
  `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors; log
  `.build\logs\msbuild-20260705_183746_054.log`.
- Focused live-row scan for Granite returned no matches:
  `rg --pcre2 -n "^\| GR-[^|]+ \| (?!.*(DONE|STALE|SIGNED OFF|DECIDED|ROUTED))|^\| GR-T[0-9A-Za-z-]+ \| (?!.*(DONE|STALE))|^\| GR-TA[0-9A-Za-z-]+ \| (?!.*DONE)" .\Specs\Plans\WIP\Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`.

Broad gate blocker:

- `.\Tools\Run-AllTests.ps1 -SkipBuild` was started after the focused checks
  and app build, but timed out after 30 minutes while running FileOps. The stale
  `RedSalamander.exe --fileops-selftest` process PID 63080 was stopped
  deliberately with `Stop-Process -Id 63080 -Force`.
- Fresh Compare archive:
  `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-05_184740\`
  recorded 218 passed / 1 failed / 30 skipped, exit code 1. Failed case:
  `search_service_sqlite_ntfs_traversal_seed_stays_degraded`. Reason:
  `Forced NTFS traversal seed should persist CURRENTNESS_UNPROVEN, got 2.`
- Fresh Commands archive:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-05_190422\` recorded 781
  passed / 1 failed / 2 skipped, exit code 1. Failed case:
  `cmd_pane_find_dialog_restored_combined_view_state_copy_follows_visible_columns`.
  Reason: `Find copy command should copy the restored combined visible row
  content to the clipboard after reopen.`
- No fresh FileOps archive was produced by that broad run. The latest FileOps
  archive before the timeout remains `2026-07-05_172218`.

Immediate next move:

1. Keep Granite in WIP.
2. Triage the two broad-gate failures with focused repros:
   ```powershell
   .\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_service_sqlite_ntfs_traversal_seed_stays_degraded -FailFast -TimeoutMultiplier 3
   .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_pane_find_dialog_restored_combined_view_state_copy_follows_visible_columns -FailFast -TimeoutMultiplier 3
   ```
3. Fix them if current work caused them, or document pre-existing/flaky status
   with fresh evidence if they reproduce independently of Granite.
4. Re-run FileOps or the full broad gate because the previous run timed out
   before producing a fresh FileOps archive.
5. Only after the broad gate is green or explicitly waived with evidence, move
   Granite to `Specs/Plans/Done/` and update `Specs/Plans/WIP/README.md`.

## Latest Save Point - 2026-07-05 18:33 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-C6 is closed.** Track C is now closed. The WIP index points
to the remaining stale test-row cleanup / completion audit instead of GR-C6.

Files touched by the GR-C6 save slice:

- `Plugins/FileSystemMtp/FileSystemMtp.h`
- `Plugins/FileSystemMtp/Factory.cpp`
- `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp`
- `Plugins/FileSystemMtp/FileSystemMtp.Device.cpp`
- `Plugins/FileSystemMtp/FileSystemMtp.FakeBackend.cpp`
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
- `Tools/Tests/TestInventory.Tests.ps1`
- `Tests/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`

GR-C6 implementation state:

- Each FileSystemMtp `#pragma warning(disable: ...)` site named by Granite now
  has a local rationale comment.
- Header/factory/device suppressions carry the sibling-plugin WIL rationale
  patterns.
- Core/fake-backend yyjson suppressions carry adjacent C6297/C28182 rationale
  comments before the pragma.
- `TestHarnessSourceContracts.Tests.ps1` now has the guard
  `comments FileSystemMtp pragma warning suppressions with local rationale`.
- Test inventory and coverage docs were bumped from 137 to 138 Tooling Pester
  cases, and `TestHarnessSourceContracts.Tests.ps1` from 46 to 47 documented
  cases.

GR-C6 evidence:

- RED source-contract Pester: 50 / 1 / 0 against the uncommented MTP pragma
  suppressions.
- GREEN source-contract Pester after implementation: 51 / 0 / 0.
- GREEN plugin build
  `.\build.ps1 -ProjectName FileSystemMtp -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors; log
  `.build\logs\msbuild-20260705_183102_114.log`.

Closeout verification after this save point:

- `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` passed 51 / 0 / 0.
- `Invoke-Pester -Script .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
  -Quiet` passed 5 / 0 / 0.
- `git diff --check -- Plugins/FileSystemMtp/FileSystemMtp.h
  Plugins/FileSystemMtp/Factory.cpp
  Plugins/FileSystemMtp/FileSystemMtp.Core.cpp
  Plugins/FileSystemMtp/FileSystemMtp.Device.cpp
  Plugins/FileSystemMtp/FileSystemMtp.FakeBackend.cpp
  Tools/Tests/TestHarnessSourceContracts.Tests.ps1
  Tools/Tests/TestInventory.Tests.ps1 Tests/README.md
  Specs/Testing/Testing_TestCoverage.md
  Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
  Specs/Plans/WIP/README.md
  Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md` exited 0
  with only LF-to-CRLF warnings.

Immediate next move:

Run the remaining Granite cleanup/completion audit. The focused scan
`rg --pcre2 -n "^\| GR-[^|]+ \| (?!\*\*.*DONE|.*DONE)|^\| GR-T[0-9A-Za-z-]+ \| (?![^|]*DONE)|STALE|RECHECK|remaining stale|stale test" Specs\Plans\WIP\Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
currently points at:

- `GR-1`, which is already marked **STALE / RECHECK** because current-tree
  evidence refutes the old path-edit premise.
- `GR-T1`, which still reads like a live RED test row and should be reconciled
  with the stale GR-1 disposition.
- `GR-A5`, which is marked **RATCHETED 2026-07-05**; the completion audit must
  decide whether the remaining long-term ratchet belongs in a follow-up plan or
  can stay documented as a non-blocking opportunistic track before Granite moves
  to Done.

Do not re-open GR-C6 unless the source-contract guard fails or the MTP pragma
sites drift.

## Latest Save Point - 2026-07-05 18:29 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-C5 is closed.** GR-S4 and GR-C1 through GR-C4 were already
fully closed in earlier save points. The WIP index now points to **GR-C6**,
then stale test-row cleanup. The next move is GR-C6.

Files touched by the GR-C5 slice:

- `Common/DxUi/DxUi.NativeTextInput.cpp`
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
- `Tools/Tests/TestInventory.Tests.ps1`
- `Tests/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`

GR-C5 implementation state:

- `TraceNativeTextInputTsfStep(...)` now uses `std::array<wchar_t, 768>` plus
  `std::format_to_n(...)` bounded formatting.
- The helper no longer uses the local `wchar_t line[768]` C buffer,
  `StringCchPrintfW`, or `wcslen(...)` to decide write length.
- Test inventory and coverage docs were bumped from 136 to 137 Tooling Pester
  cases, and `TestHarnessSourceContracts.Tests.ps1` from 45 to 46 documented
  cases.

GR-C5 evidence:

- RED source-contract Pester: 49 / 1 / 0 against the old
  `wchar_t line[768]` + `StringCchPrintfW` block.
- GREEN source-contract Pester after implementation: 50 / 0 / 0.
- GREEN inventory Pester after doc/count updates: 5 / 0 / 0.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; log
  `.build\logs\msbuild-20260705_182416_209.log`.
- GREEN `DxUiTests` build
  `.\build.ps1 -ProjectName DxUiTests -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors; log
  `.build\logs\msbuild-20260705_182634_951.log`.
- GREEN focused native test run
  `.\.build\x64\Debug\DxUiTests.exe --suite=NativeTextInput
  --perf-jsonl=Specs\TestRuns\local_scratch\dxui_native_textinput_gr_c5_format_to_n_20260705_1828.jsonl`
  exited 0 with "All DxUi tests passed."

Closeout verification after this save point:

- `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` passed 50 / 0 / 0.
- `Invoke-Pester -Script .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
  -Quiet` passed 5 / 0 / 0.
- `git diff --check -- Common/DxUi/DxUi.NativeTextInput.cpp
  Tools/Tests/TestHarnessSourceContracts.Tests.ps1
  Tools/Tests/TestInventory.Tests.ps1 Tests/README.md
  Specs/Testing/Testing_TestCoverage.md
  Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
  Specs/Plans/WIP/README.md
  Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md` exited 0
  with only LF-to-CRLF warnings.

Immediate next move:

Start **GR-C6**. Read `.github/skills/compiler-warnings/SKILL.md`, inspect the
current MTP `#pragma warning(disable: ...)` sites, and add a RED
source-contract guard before copying the sibling plugin rationale comments (the
row points at `Plugins/FileSystem/FileSystem.h` as the house pattern).

## Latest Save Point - 2026-07-05 18:21 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-C4 is closed.** GR-S4 and GR-C1 through GR-C3 were already
fully closed in earlier save points. The WIP index now points to **GR-C5
through GR-C6**, then stale test-row cleanup. The next move is GR-C5.

Files touched by the GR-C4 slice:

- `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
- `Tools/Tests/TestInventory.Tests.ps1`
- `Tests/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`

GR-C4 implementation state:

- The `sqlite_index_store_load_and_apply_journal_delta` and
  `sqlite_index_store_root_lookup_case_insensitive` `EnumerateVolume(...)`
  callbacks still fail-fast on `std::bad_alloc`.
- Their broader `catch (const std::exception&)` blocks now have mandatory
  `noexcept` callback-boundary comments and log once with `Debug::Error(...)`
  before returning `E_FAIL`.
- Test inventory and coverage docs were bumped from 135 to 136 Tooling Pester
  cases, and `TestHarnessSourceContracts.Tests.ps1` from 44 to 45 documented
  cases.

GR-C4 evidence:

- RED source-contract Pester: 48 / 1 / 0 against the unfixed callback blocks.
- GREEN source-contract Pester after implementation: 49 / 0 / 0.
- GREEN inventory Pester after doc/count updates: 5 / 0 / 0.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; log
  `.build\logs\msbuild-20260705_181849_764.log`.
- GREEN focused Compare archives:
  - `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_182054/`
    `sqlite_index_store_load_and_apply_journal_delta` at 1 / 0 / 0.
  - `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_182059/`
    `sqlite_index_store_root_lookup_case_insensitive` at 1 / 0 / 0.

Closeout verification after this save point:

- `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` passed 49 / 0 / 0.
- `Invoke-Pester -Script .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
  -Quiet` passed 5 / 0 / 0.
- `git diff --check --
  RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp
  Tools/Tests/TestHarnessSourceContracts.Tests.ps1
  Tools/Tests/TestInventory.Tests.ps1 Tests/README.md
  Specs/Testing/Testing_TestCoverage.md
  Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
  Specs/Plans/WIP/README.md
  Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md` exited 0
  with only LF-to-CRLF warnings.

Immediate next move:

Start **GR-C5**. Read `.github/skills/cpp-modern-style/SKILL.md` before
editing, inspect the current `Common/DxUi/DxUi.NativeTextInput.cpp` anchors, and
add a RED source-contract guard before converting the diagnostic buffer from
`wchar_t line[768]` + `StringCchPrintfW` to the project-preferred
`std::array` + `std::format_to_n` style (unless current source already differs
enough to require re-triage).

## Latest Save Point - 2026-07-05 18:15 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-C3 is closed.** GR-S4 and GR-C1/GR-C2 were already fully
closed in earlier save points. The WIP index now points to **GR-C4 through
GR-C6**, then stale test-row cleanup. The next move is GR-C4.

Files touched by the GR-C3 slice:

- `Common/LocalSearchIndexCore.cpp`
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
- `Tools/Tests/TestInventory.Tests.ps1`
- `Tests/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/Core/Core_Search.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`

GR-C3 implementation state:

- `OpenSnapshotTempFile(...)` still fail-fasts on `std::bad_alloc`.
- Its broader `catch (const std::exception&)` now has the mandatory
  `noexcept`-boundary comment and logs once with `Debug::Error(...)` before
  returning `E_FAIL`.
- `Specs/Core/Core_Search.md` now carries the durable snapshot persistence
  exception-handling contract for `noexcept` helpers.
- Test inventory and coverage docs were bumped from 134 to 135 Tooling Pester
  cases, and `TestHarnessSourceContracts.Tests.ps1` from 43 to 44 documented
  cases.

GR-C3 evidence:

- RED source-contract Pester: 47 / 1 / 0 against the unfixed
  `OpenSnapshotTempFile(...)` catch block.
- GREEN source-contract Pester after implementation: 48 / 0 / 0.
- GREEN inventory Pester after doc/count updates: 5 / 0 / 0.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; log
  `.build\logs\msbuild-20260705_181218_425.log`.
- GREEN focused Compare archives:
  - `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_181423/`
    `search_source_allocation_and_folding_guard` at 1 / 0 / 0.
  - `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_181429/`
    `search_low_hardening_smoke` at 1 / 0 / 0.

Closeout verification after this save point:

- `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` passed 48 / 0 / 0.
- `Invoke-Pester -Script .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
  -Quiet` passed 5 / 0 / 0.
- `git diff --check -- Common/LocalSearchIndexCore.cpp
  Tools/Tests/TestHarnessSourceContracts.Tests.ps1
  Tools/Tests/TestInventory.Tests.ps1 Tests/README.md
  Specs/Testing/Testing_TestCoverage.md Specs/Core/Core_Search.md
  Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
  Specs/Plans/WIP/README.md
  Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md` exited 0
  with only LF-to-CRLF warnings.

Immediate next move:

Start **GR-C4**. Inspect the current source and row anchors fresh, then add a
RED source-contract guard before implementing. The row currently calls out
`RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`
callback lambdas that catch `std::exception` at `noexcept` boundaries without
the mandatory explanatory comment and logging review. After GR-C4, continue
GR-C5 and GR-C6 high-to-low.

## Latest Save Point - 2026-07-05 18:07 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-C1 and GR-C2 are closed.** GR-S4 was already fully closed
in the previous save point. The WIP index now points to **GR-C3 through
GR-C6**, then stale test-row cleanup. The next move is GR-C3.

Files touched by the GR-C1/GR-C2 slice:

- `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp`
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
- `Tools/Tests/TestInventory.Tests.ps1`
- `Tests/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/FileSystem/FileSystem_Mtp.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`

GR-C1/GR-C2 implementation state:

- `MtpBufferedWriter` and adjacent `MtpBackendReader` now hold their
  `FileSystemMtp` owner with `wil::com_ptr<FileSystemMtp>` instead of owning
  raw pointers plus manual `_owner->AddRef()` / `_owner->Release()`.
- `FileSystemMtp::ReadDirectoryInfo(...)` now keeps `FilesInformationMtp` in
  `std::unique_ptr<FilesInformationMtp>` ownership until successful
  `*ppFilesInformation = info.release()` handoff, removing the failure-path
  manual `info->Release()`.
- `Specs/FileSystem/FileSystem_Mtp.md` now carries the durable MTP ownership
  contract for helper COM objects and directory-info handoff ownership.
- Test inventory and coverage docs were bumped from 133 to 134 Tooling Pester
  cases, and `TestHarnessSourceContracts.Tests.ps1` from 42 to 43 documented
  cases.

GR-C1/GR-C2 evidence:

- RED source-contract Pester: 46 / 1 / 0 before the MTP RAII owner and
  directory-info handoff guard was satisfied.
- GREEN source-contract Pester after implementation: 47 / 0 / 0.
- GREEN app build `.\build.ps1 -ProjectName FileSystemMtp -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; log
  `.build\logs\msbuild-20260705_180344_864.log`.
- GREEN focused Compare archives:
  - `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_180425/`
    `mtp_fake_backend_enumerate_read_and_capabilities` at 1 / 0 / 0.
  - `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_180431/`
    `mtp_public_writer_stages_until_commit` at 1 / 0 / 0.

Closeout verification after this save point:

- `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` passed 47 / 0 / 0.
- `Invoke-Pester -Script .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
  -Quiet` passed 5 / 0 / 0.
- `git diff --check -- Plugins/FileSystemMtp/FileSystemMtp.Core.cpp
  Tools/Tests/TestHarnessSourceContracts.Tests.ps1
  Tools/Tests/TestInventory.Tests.ps1 Tests/README.md
  Specs/Testing/Testing_TestCoverage.md Specs/FileSystem/FileSystem_Mtp.md
  Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
  Specs/Plans/WIP/README.md
  Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md` exited 0
  with only LF-to-CRLF warnings.

Immediate next move:

Start **GR-C3**. Inspect the current source and row anchors fresh, then add a
RED source-contract guard before implementing. The row currently calls out
`Common/LocalSearchIndexCore.cpp:980-983`: `catch (const std::exception&) {
return E_FAIL; }` needs the mandatory explanatory comment and
`Debug::Error(...)` logging required by AGENTS exception-handling guidance.
After GR-C3, continue GR-C4 through GR-C6 high-to-low.

## Latest Save Point - 2026-07-05 18:01 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-S4 is fully closed**. Subitems (a), (b), (c), (d), (e), and
(f) all have ratchet notes/evidence. The WIP index now points to **GR-C1
through GR-C6** and the remaining stale test-row cleanup. Granite remains active
because those later rows are still open.

Files touched by the GR-S4(f) slice:

- `RedSalamander/FolderView.Icons.cpp`
- `RedSalamander/FolderView.cpp`
- `RedSalamander/FolderView.h`
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
- `Tools/Tests/TestInventory.Tests.ps1`
- `Tests/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/UI/UI_FolderView.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`

GR-S4(f) implementation state:

- `FolderView.Icons.cpp` now has `IncrementThumbnailStat(...)`, and the shell
  provider/cache/success thumbnail stats route through it instead of repeating
  direct `fetch_add(...)` plus matching `PerfEmitCounter(...)` calls.
- `FolderView.h` now has one `PendingToPaintMetric` shape used by both
  `_pendingInputToPaintMetric` and `_pendingRefreshToPaintMetric`.
- `FolderView.cpp` records/emits input-to-paint and refresh-to-paint telemetry
  through that shared shape while keeping separate pending slots.
- The D3D11 creation duplication subitem is already closed by GR-A6:
  `FolderView.Rendering.cpp` uses
  `RedSalamander::DxUi::CreateD3D11DeviceWithWarpFallback(...)`.
- `Specs/UI/UI_FolderView.md` now carries the durable contract for paired
  thumbnail stat/perf telemetry and separate pending input/refresh slots.
- Test inventory and coverage docs were bumped from 131 to 133 Tooling Pester
  cases, and `TestHarnessSourceContracts.Tests.ps1` from 40 to 42 documented
  cases.

GR-S4(f) evidence:

- RED source-contract Pester: 44 / 1 / 0 before
  `IncrementThumbnailStat(...)` existed.
- GREEN after thumbnail-stat helper extraction: 45 / 0 / 0.
- RED source-contract Pester: 45 / 1 / 0 before `PendingToPaintMetric` existed.
- GREEN after pending metric shape extraction: 46 / 0 / 0.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; log
  `.build\logs\msbuild-20260705_175607_052.log`.
- GREEN focused Commands archives:
  - `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_175832/`
    `folderView_thumbnail_cached_only_no_close_stall` at 1 / 0 / 0.
  - `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_175839/`
    `folderView_perf_slow_virtual_provider` at 1 / 0 / 0.
  - `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_175845/`
    `folderView_perf_refresh_preservation` at 1 / 0 / 0.
  - `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_175904/`
    `folderView_perf_scroll_render_stress` at 1 / 0 / 0.

Closeout verification after this save point:

- `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` passed 46 / 0 / 0.
- `Invoke-Pester -Script .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
  -Quiet` passed 5 / 0 / 0.
- `git diff --check -- RedSalamander/FolderView.Icons.cpp
  RedSalamander/FolderView.cpp RedSalamander/FolderView.h
  Tools/Tests/TestHarnessSourceContracts.Tests.ps1
  Tools/Tests/TestInventory.Tests.ps1 Tests/README.md
  Specs/Testing/Testing_TestCoverage.md Specs/UI/UI_FolderView.md
  Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
  Specs/Plans/WIP/README.md
  Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md` exited 0
  with only possible LF-to-CRLF warnings.

Immediate next move:

Continue high-to-low with the remaining Granite rows **GR-C1 through GR-C6**.
Inspect those rows fresh before editing; do not assume their anchors survived
the current worktree churn.

## Latest Save Point - 2026-07-05 17:51 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-S4(a), GR-S4(b), GR-S4(c), GR-S4(d), and GR-S4(e) are
closed as ratchet slices.** The WIP index now points to remaining
**GR-S4(f)**, then GR-C1 through GR-C6 and stale test-row cleanup. Do not mark
the whole Granite plan done yet.

Files touched by the GR-S4(d) slice:

- `Common/Helpers.h`
- `RedSalamander/FolderView.FileOps.cpp`
- `RedSalamander/FolderView.Icons.cpp`
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
- `Tools/Tests/TestInventory.Tests.ps1`
- `Tests/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`

GR-S4(d) implementation state:

- `Common/Helpers.h` now has
  `SubmitOwnedThreadpoolCallback(std::unique_ptr<T>& payload) noexcept`.
- `FolderView::StartPasteShortcutWork(...)` gives `PasteShortcutWork` a
  `void Execute() noexcept` method and queues it with
  `SubmitOwnedThreadpoolCallback(work)`.
- `FolderView::ExtractProviderAllowedThumbnailWithDeadline(...)` gives
  `ProviderAllowedWork` a `void Execute() noexcept` method and queues it with
  `SubmitOwnedThreadpoolCallback(work)`.
- The targeted FolderView blocks no longer carry local
  `TrySubmitThreadpoolCallback` / `work.release()` submit scaffolding.
- `TestHarnessSourceContracts.Tests.ps1` has the guard
  `routes FolderView owned threadpool work through one submit helper`.
- Test inventory and coverage docs were bumped from 130 to 131 Tooling Pester
  cases, and `TestHarnessSourceContracts.Tests.ps1` from 39 to 40 documented
  cases.

GR-S4(d) evidence:

- RED `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` reported 43 passed / 1 failed / 0 skipped because
  `SubmitOwnedThreadpoolCallback(...)` did not exist.
- GREEN after implementation and test-anchor correction:
  `TestHarnessSourceContracts.Tests.ps1` passed 44 / 0 / 0.
- GREEN build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; log
  `.build\logs\msbuild-20260705_174449_942.log`.
- GREEN focused Commands archive
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_174925/` passed
  `cmd_pane_clipboardPasteShortcut_returns_before_worker_complete` with
  1 passed / 0 failed / 0 skipped.
- GREEN focused Commands archive
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_174931/` passed
  `folderView_perf_slow_virtual_provider` with
  1 passed / 0 failed / 0 skipped.

Closeout verification after this save point:

- `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` passed 44 / 0 / 0.
- `Invoke-Pester -Script .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
  -Quiet` passed 5 / 0 / 0.
- `git diff --check -- Common/Helpers.h RedSalamander/FolderView.FileOps.cpp
  RedSalamander/FolderView.Icons.cpp
  Tools/Tests/TestHarnessSourceContracts.Tests.ps1
  Tools/Tests/TestInventory.Tests.ps1 Tests/README.md
  Specs/Testing/Testing_TestCoverage.md
  Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
  Specs/Plans/WIP/README.md
  Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md` exited 0
  with only possible LF-to-CRLF warnings.

Immediate next move:

Continue remaining GR-S4:

- GR-S4(f): WarpDrive parallel-maintenance cleanup for thumbnail counters,
  D3D11 create-device duplication, and pending input/refresh-to-paint metric
  duplication.

## Latest Save Point - 2026-07-05 17:48 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Repo position at save: branch `master`, commit `da6b438a0`. The worktree has many
pre-existing unrelated modified files from other Granite/Clearwater/Tailwind/DxUi
work; do not revert them.

Current state: **GR-S4(a), GR-S4(b), GR-S4(c), and GR-S4(e) are closed as
ratchet slices. GR-S4(d) has code and source-contract/build evidence complete,
but it is not closed in the Granite ledger yet.** Runtime checks, documentation
counts, WIP index text, Granite note updates, and closeout verification still
remain for GR-S4(d). After that, continue with **GR-S4(f)**, then GR-C1 through
GR-C6 and stale test-row cleanup.

Files touched by the in-progress GR-S4(d) slice:

- `Common/Helpers.h`
- `RedSalamander/FolderView.FileOps.cpp`
- `RedSalamander/FolderView.Icons.cpp`
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
- `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`

GR-S4(d) implementation state:

- `Common/Helpers.h` now has
  `SubmitOwnedThreadpoolCallback(std::unique_ptr<T>& payload) noexcept`.
- The helper owns/destroys queued payloads with `std::unique_ptr<T> owned(...)`
  and calls `owned->Execute()` inside a `TrySubmitThreadpoolCallback` callback.
- `FolderView::StartPasteShortcutWork(...)` now gives `PasteShortcutWork` a
  `void Execute() noexcept` method and queues it with
  `SubmitOwnedThreadpoolCallback(work)`.
- `FolderView::ExtractProviderAllowedThumbnailWithDeadline(...)` now gives
  `ProviderAllowedWork` a `void Execute() noexcept` method and queues it with
  `SubmitOwnedThreadpoolCallback(work)`.
- The old local `TrySubmitThreadpoolCallback` / `work.release()` dance is gone
  from both targeted FolderView blocks.
- `TestHarnessSourceContracts.Tests.ps1` has the new guard
  `routes FolderView owned threadpool work through one submit helper`.

GR-S4(d) evidence already collected:

- RED `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` reported 43 passed / 1 failed / 0 skipped because
  `SubmitOwnedThreadpoolCallback(...)` did not exist.
- GREEN after implementation and test-anchor correction:
  `TestHarnessSourceContracts.Tests.ps1` passed 44 / 0 / 0.
- GREEN build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; log
  `.build\logs\msbuild-20260705_174449_942.log`.
- One earlier build invocation timed out at 120 seconds before producing useful
  pass/fail output; ignore it as non-evidence.

Immediate next commands to finish GR-S4(d):

```powershell
$env:REDSALAMANDER_SELFTEST_ROOT = (Join-Path (Resolve-Path .\.build).Path 'codex-runs\gr-s4d-paste-shortcut-worker')
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_pane_clipboardPasteShortcut_returns_before_worker_complete -FailFast -TimeoutMultiplier 3

$env:REDSALAMANDER_SELFTEST_ROOT = (Join-Path (Resolve-Path .\.build).Path 'codex-runs\gr-s4d-provider-allowed-thumbnail')
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_perf_slow_virtual_provider -FailFast -TimeoutMultiplier 3
```

After those runtime checks pass, update:

- `Tools/Tests/TestInventory.Tests.ps1`: bump Tools Pester expected count from
  130 to 131 in both assertions.
- `Tests/README.md`: bump Tooling Script Tests from 130 to 131 and
  `TestHarnessSourceContracts.Tests.ps1` from 39 to 40; add FolderView owned
  threadpool submit helper reuse to the description.
- `Specs/Testing/Testing_TestCoverage.md`: bump Tooling scripts from 130 to 131,
  add a GR-S4(d) checkpoint, and update the tooling row description.
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`:
  mark GR-S4(d) done with the two targeted worker routes, add a GR-S4(d) ratchet
  note, and change prior GR-S4 notes that say remaining `(d) and (f)` to only
  remaining `(f)`.
- `Specs/Plans/WIP/README.md`: change `GR-S4(a/b/c/e)` to
  `GR-S4(a/b/c/d/e)` and remaining `GR-S4(d/f)` to `GR-S4(f)`.
- This baton: prepend a closed GR-S4(d) save point with the runtime archives and
  final verification results.

Final closeout checks to run before moving on:

```powershell
Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru -Quiet
Invoke-Pester -Script .\Tools\Tests\TestInventory.Tests.ps1 -PassThru -Quiet
git diff --check -- Common/Helpers.h RedSalamander/FolderView.FileOps.cpp RedSalamander/FolderView.Icons.cpp Tools/Tests/TestHarnessSourceContracts.Tests.ps1 Tools/Tests/TestInventory.Tests.ps1 Tests/README.md Specs/Testing/Testing_TestCoverage.md Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md Specs/Plans/WIP/README.md Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md
```

Expected after docs/count closeout: source contracts 44 / 0 / 0, inventory
5 / 0 / 0, and `git diff --check` exits 0 except possible LF-to-CRLF warnings.

## Latest Save Point - 2026-07-05 17:38 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-S4(a), GR-S4(b), GR-S4(c), and GR-S4(e) are closed as
ratchet slices.** The WIP index now points to remaining **GR-S4(d/f)**, then
GR-C1 through GR-C6 and stale test-row cleanup. Do not mark the whole GR-S4 row
done yet.

Files touched by the GR-S4(a) slice:

- `RedSalamander/FolderWindow.FileOperationsInternal.h`
- `RedSalamander/FolderWindow.FileOperations.State.Diagnostics.Part.cpp`
- `RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp`
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
- `Tools/Tests/TestInventory.Tests.ps1`
- `Tests/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`

GR-S4(a) implementation state:

- `FolderWindow.FileOperationsInternal.h` now has
  `RestoreActivePaneFolderViewFocusIfWindowHadFocusBeforeHide(...)`.
- `FolderWindow::FileOperationState::ToggleIssuesPane()` captures
  `focusedBeforeHide`, hides the issues pane, and calls the shared helper.
- `FileOperationsIssuesPaneState::OnClose(...)` captures `focusedBeforeHide`,
  hides the pane, checks the live owner lifetime, and calls the same helper.
- `TestHarnessSourceContracts.Tests.ps1` has the new guard
  `routes FileOperations issues-pane focus restore through one helper`.
- Tooling Pester count moved 129 -> 130; `TestHarnessSourceContracts` doc count
  moved 38 -> 39.
- Granite marks GR-S4 as ratcheted with subitems (a), (b), (c), and (e) done
  while (d) and (f) remain open.

GR-S4(a) evidence:

- RED `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` reported 42 passed / 1 failed / 0 skipped because the
  shared helper did not exist.
- GREEN after implementation: `TestHarnessSourceContracts.Tests.ps1` passed
  43 / 0 / 0.
- GREEN build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; log
  `.build\logs\msbuild-20260705_173537_689.log`.
- Focused runtime:

  ```powershell
  $env:REDSALAMANDER_SELFTEST_ROOT = (Join-Path (Resolve-Path .\.build).Path 'codex-runs\gr-s4a-issues-pane-focus')
  .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_pane_fileops_issues_pane_hide_restores_folder_focus -FailFast -TimeoutMultiplier 3
  ```

  passed 1 / 0 / 0 and archived at
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_173744/`.

Closeout verification after this save point:

- `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` passed 43 / 0 / 0.
- `Invoke-Pester -Script .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
  -Quiet` passed 5 / 0 / 0.
- `git diff --check -- RedSalamander/FolderWindow.FileOperationsInternal.h
  RedSalamander/FolderWindow.FileOperations.State.Diagnostics.Part.cpp
  RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp
  Tools/Tests/TestHarnessSourceContracts.Tests.ps1
  Tools/Tests/TestInventory.Tests.ps1 Tests/README.md
  Specs/Testing/Testing_TestCoverage.md
  Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
  Specs/Plans/WIP/README.md
  Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`
  exited 0 with only LF-to-CRLF warnings.

Immediate next move:

Continue remaining GR-S4:

- GR-S4(d): owned-threadpool-submit helper for `PasteShortcutWork` and
  `ProviderAllowedWork`.
- GR-S4(f): larger WarpDrive cleanup for thumbnail counters and pending metrics.

## Latest Save Point - 2026-07-05 17:30 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-S4(b), GR-S4(c), and GR-S4(e) are closed as ratchet
slices.** The WIP index now points to remaining **GR-S4(a/d/f)**, then GR-C1
through GR-C6 and stale test-row cleanup. Do not mark the whole GR-S4 row done
yet.

Additional files touched after the 17:23 save point:

- `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
- `Tools/Tests/TestInventory.Tests.ps1`
- `Tests/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`

GR-S4(e) implementation state:

- `Commands.SelfTest.ViewCommands.cpp` now has one
  `IsViewerSpaceEnvFlagEnabled(const wchar_t* name)` parser for `1` / `true` /
  `yes`.
- `IsViewerSpaceLargePerfEnabled()` and `IsViewerSpace20kLayoutEnabled()` call
  that shared helper with their respective env-var names.
- `TestHarnessSourceContracts.Tests.ps1` has the new guard
  `uses one ViewerSpace env flag parser for opt-in perf scenarios`.
- Tooling Pester count moved 128 -> 129; `TestHarnessSourceContracts` doc count
  moved 37 -> 38.
- Granite marks GR-S4 as ratcheted with subitems (b), (c), and (e) done while
  (a), (d), and (f) remain open.

GR-S4(e) evidence:

- RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`
  reported 41 passed / 1 failed / 0 skipped because the helper did not exist.
- GREEN after implementation: `TestHarnessSourceContracts.Tests.ps1` passed
  42 / 0 / 0.
- GREEN build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; log
  `.build\logs\msbuild-20260705_172622_129.log`.
- Focused runtime:

  ```powershell
  $env:REDSALAMANDER_SELFTEST_ROOT = (Join-Path (Resolve-Path .\.build).Path 'codex-runs\gr-s4e-viewerspace-20k-skip')
  .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_viewer_space_layout_20k_visible_optin -FailFast -TimeoutMultiplier 3
  ```

  passed 0 / 0 / 1 and archived at
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_172847/`.

  ```powershell
  $env:REDSALAMANDER_SELFTEST_ROOT = (Join-Path (Resolve-Path .\.build).Path 'codex-runs\gr-s4e-viewerspace-large-skip')
  .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter viewer_space_perf_large_optin -FailFast -TimeoutMultiplier 3
  ```

  passed 0 / 0 / 1 and archived at
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_172852/`.

Note: an earlier exploratory `-CaseFilter optin` run produced no matching cases
and failed only the wrapper's result-coverage check; it is not product evidence.

Closeout verification after this save point:

- `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
  -PassThru -Quiet` passed 42 / 0 / 0.
- `Invoke-Pester -Script .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
  -Quiet` passed 5 / 0 / 0.
- `git diff --check -- Common/Helpers.h
  RedSalamander/FolderView.Rendering.cpp
  RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp
  RedSalamander/FolderWindow.FileOperations.State.cpp
  Tools/Tests/TestHarnessSourceContracts.Tests.ps1
  Tools/Tests/TestInventory.Tests.ps1 Tests/README.md
  Specs/Testing/Testing_TestCoverage.md
  Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
  Specs/Plans/WIP/README.md
  Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`
  exited 0 with only LF-to-CRLF warnings.

Immediate next move:

Continue remaining GR-S4:

- GR-S4(a): focus-restore helper extraction; old anchors drifted, so inspect
  current blocks first.
- GR-S4(d): owned-threadpool-submit helper for `PasteShortcutWork` and
  `ProviderAllowedWork`.
- GR-S4(f): larger WarpDrive cleanup for thumbnail counters and pending metrics.

## Latest Save Point - 2026-07-05 17:23 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-S4(b) and GR-S4(c) are closed as ratchet slices.** GR-S2
remains closed. The WIP index now points to remaining **GR-S4(a/d/e/f)**, then
GR-C1 through GR-C6 and stale test-row cleanup. Do not mark the whole GR-S4 row
done yet.

Additional files touched after the 17:16 save point:

- `RedSalamander/FolderWindow.FileOperations.State.cpp`
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
- `Tools/Tests/TestInventory.Tests.ps1`
- `Tests/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`

GR-S4(c) implementation state:

- `MaybeInjectBridgeCreateDirectoryRaceForSelfTest(...)` now calls
  `TryReadEnvironmentVariableForSelfTest(kRacePathEnv)` and no longer owns raw
  `GetEnvironmentVariableW(...)` calls.
- `TestHarnessSourceContracts.Tests.ps1` has the new guard
  `routes FileOperations bridge create-directory race env reads through the shared self-test helper`.
- Tooling Pester count moved 127 -> 128; `TestHarnessSourceContracts` doc count
  moved 36 -> 37.
- Granite marks GR-S4 as ratcheted with subitems (b) and (c) done while (a),
  (d), (e), and (f) remain open.

GR-S4(c) evidence:

- RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`
  reported 40 passed / 1 failed / 0 skipped because the race helper still used
  raw `GetEnvironmentVariableW(...)`.
- GREEN after implementation: `TestHarnessSourceContracts.Tests.ps1` passed
  41 / 0 / 0.
- GREEN build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; log
  `.build\logs\msbuild-20260705_172008_860.log`.
- GREEN focused runtime:

  ```powershell
  $env:REDSALAMANDER_SELFTEST_ROOT = (Join-Path (Resolve-Path .\.build).Path 'codex-runs\gr-s4c-create-dir-race')
  .\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Riptide_BridgeCreateDirectoryRaceExistingFilePromptsPartial -FailFast -TimeoutMultiplier 3
  ```

  passed 3 / 0 / 0 and archived at
  `Specs/TestRuns/4cb089111a23/FileOps/2026-07-05_172218/`.

Immediate next move:

1. Run final closeout checks if they were not completed before the next turn:

   ```powershell
   Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
   Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
   git diff --check -- Common/Helpers.h RedSalamander/FolderView.Rendering.cpp RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp RedSalamander/FolderWindow.FileOperations.State.cpp Tools/Tests/TestHarnessSourceContracts.Tests.ps1 Tools/Tests/TestInventory.Tests.ps1 Tests/README.md Specs/Testing/Testing_TestCoverage.md Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md Specs/Plans/WIP/README.md Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md
   ```

2. Continue GR-S4 with the remaining live subitems:
   - GR-S4(a): focus-restore helper extraction; re-inspect drifted anchors first.
   - GR-S4(d): owned-threadpool-submit helper for `PasteShortcutWork` and
     `ProviderAllowedWork`.
   - GR-S4(e): shared selftest env flag reader for ViewerSpace large/20K flags.
   - GR-S4(f): larger WarpDrive cleanup for thumbnail counters and pending metrics.

## Latest Save Point - 2026-07-05 17:16 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-S4(b) is closed as a ratchet slice.** GR-S2 remains closed
from the prior save point. The WIP index now points to remaining
**GR-S4(a/c/d/e/f)**, then GR-C1 through GR-C6 and stale test-row cleanup. Do
not mark the whole GR-S4 row done yet.

Files touched by the GR-S4(b) slice:

- `Common/Helpers.h`
- `RedSalamander/FolderView.Rendering.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
- `Tools/Tests/TestInventory.Tests.ps1`
- `Tests/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`

Implementation state:

- `Common/Helpers.h` now exposes
  `EnvironmentVariables::IsTruthyFlagSet(const wchar_t* name) noexcept`.
- `FolderView.Rendering.cpp` no longer defines local
  `IsTruthySelfTestEnvironmentVariable(...)`; `ShouldForceFolderViewWarpDevice()`
  calls the shared helper with `REDSALAMANDER_FOLDERVIEW_FORCE_WARP`.
- `Commands.SelfTest.ViewCommands.cpp` no longer defines its local
  `IsTruthySelfTestEnvironmentVariable(...)`; `IsFolderViewHugePerfOptInEnabled()`
  and `IsFolderViewForceWarpEnabled()` call the shared helper.
- `TestHarnessSourceContracts.Tests.ps1` has the new guard
  `uses a shared truthy environment flag helper for FolderView perf flags`.
- Tooling Pester count moved 126 -> 127; `TestHarnessSourceContracts` doc count
  moved 35 -> 36.
- Granite marks GR-S4 as `LOW - RATCHETED 2026-07-05` and its row says subitem
  (b) is done while (a), (c), (d), (e), and (f) remain open.

Evidence:

- RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`
  reported 39 passed / 1 failed / 0 skipped before the shared helper existed.
- GREEN after implementation: `TestHarnessSourceContracts.Tests.ps1` passed
  40 / 0 / 0.
- GREEN build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; log
  `.build\logs\msbuild-20260705_171253_615.log`.
- GREEN focused runtime:

  ```powershell
  $env:REDSALAMANDER_SELFTEST_ROOT = (Join-Path (Resolve-Path .\.build).Path 'codex-runs\gr-s4b-warp')
  $env:REDSALAMANDER_FOLDERVIEW_FORCE_WARP = '1'
  .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_perf_huge_folder_scale -FailFast -TimeoutMultiplier 3
  ```

  passed 1 / 0 / 0 and archived at
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_171549/`. The perf artifact
  records `warpRunExecuted: true`.

Immediate next move:

1. Run final closeout checks if they were not completed before the next turn:

   ```powershell
   Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
   Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
   git diff --check -- Common/Helpers.h RedSalamander/FolderView.Rendering.cpp RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp Tools/Tests/TestHarnessSourceContracts.Tests.ps1 Tools/Tests/TestInventory.Tests.ps1 Tests/README.md Specs/Testing/Testing_TestCoverage.md Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md Specs/Plans/WIP/README.md Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md
   ```

2. Continue GR-S4. The smallest next slice is still **GR-S4(c)**:
   add a RED source-contract guard that
   `MaybeInjectBridgeCreateDirectoryRaceForSelfTest(...)` uses
   `TryReadEnvironmentVariableForSelfTest(...)` and has no raw
   `GetEnvironmentVariableW` in that function body, then patch only that
   function. Remaining alternatives are GR-S4(a), GR-S4(d), GR-S4(e), and the
   larger GR-S4(f) WarpDrive cleanup.

## Latest Save Point - 2026-07-05 17:10 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-S2 is closed, and the next concrete Granite move is GR-S4.**
The WIP index already says to continue with `GR-S4 / GR-S4(f)`, then GR-C1
through GR-C6 and stale test-row cleanup. Treat GR-S4 as an incremental
quick-win batch: close one verified slice with RED/GREEN evidence, update the
ledger as a ratchet, and do not mark the whole GR-S4 row done until all live
subitems are handled or explicitly retired as stale.

Do not revert unrelated dirty files. `git status --short` still contains many
pre-existing modified files from accumulated Granite/Clearwater/Tailwind/DxUi
work. The files touched by the completed GR-S2 slice remain the files listed in
the 17:07 save point below. The only new work after that save point was a
read-only GR-S4 drift check plus this baton update.

GR-S4 drift check performed after GR-S2 closeout:

- GR-S4(a) is still plausibly live but its old anchors drifted. Current focus
  restore blocks are in
  `RedSalamander/FolderWindow.FileOperations.State.Diagnostics.Part.cpp:221`
  through `:236` and
  `RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp:889` through at
  least `:903`. Re-inspect before patching because the bodies are no longer the
  exact old `:226-237` / `:886-897` anchors.
- GR-S4(b) is live and compact:
  `RedSalamander/FolderView.Rendering.cpp:33` defines
  `IsTruthySelfTestEnvironmentVariable`, `:40` defines
  `ShouldForceFolderViewWarpDevice`, and
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp:17713`
  defines a second `IsTruthySelfTestEnvironmentVariable` used by the huge-perf
  and force-WARP gates at `:17720` and `:17725`.
- GR-S4(c) is live:
  `RedSalamander/FolderWindow.FileOperations.State.cpp:521` still has
  `MaybeInjectBridgeCreateDirectoryRaceForSelfTest(...)` hand-reading
  `REDSALAMANDER_FILEOPS_BRIDGE_CREATE_DIRECTORY_RACE_PATH` with raw
  `GetEnvironmentVariableW`, while the sibling helper
  `TryReadEnvironmentVariableForSelfTest(...)` exists at `:232` and is already
  used by the later mutation hook at `:559` / `:579`.
- GR-S4(d) is live:
  `RedSalamander/FolderView.FileOps.cpp:901` / `:903` carries
  `StartPasteShortcutWork` / `PasteShortcutWork`, and
  `RedSalamander/FolderView.Icons.cpp:960` carries `ProviderAllowedWork`; both
  still submit owned heap work through `TrySubmitThreadpoolCallback`.
- GR-S4(e) is live:
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp:19499`
  and `:19512` define two viewer-space env-flag readers with identical parsing
  modulo env-var name.
- GR-S4(f) remains the larger WarpDrive parallel-maintenance cleanup bucket from
  the 2026-07-04 Granite addendum; use it after the smaller GR-S4 slices or
  split it into a separate RED-first subplan.

Recommended next slice: **GR-S4(b), shared env truthiness helper for FolderView
WARP/perf flags.** It is small, has two clear duplicate definitions, and directly
protects the WARP-forcing perf/test path. Use TDD:

1. Add a RED source-contract case in
   `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` that requires the chosen
   shared truthy-env helper and rejects local
   `IsTruthySelfTestEnvironmentVariable` definitions in both
   `FolderView.Rendering.cpp` and `Commands.SelfTest.ViewCommands.cpp`.
2. Implement the minimal shared helper in the existing common/helper style after
   inspecting current environment-helper conventions with:

   ```powershell
   rg -n "EnvironmentVariable|GetEnvironmentVariableW|IsTruthy|EnvFlag|TryReadEnvironment" Common RedSalamander -g "*.h" -g "*.cpp"
   ```

3. Re-route `ShouldForceFolderViewWarpDevice`,
   `IsFolderViewHugePerfOptInEnabled`, and `IsFolderViewForceWarpEnabled`
   through the shared helper.
4. Run:

   ```powershell
   Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
   Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
   .\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64
   ```

5. Patch counts/docs if a new Pester `It` is added, then add a GR-S4(b) ratchet
   note to the Granite WIP plan and keep GR-S4 open for the remaining subitems.

Smaller fallback if the next session needs a tiny first step: close **GR-S4(c)**
instead by adding a RED source-contract guard that
`MaybeInjectBridgeCreateDirectoryRaceForSelfTest(...)` uses
`TryReadEnvironmentVariableForSelfTest(...)` and has no raw
`GetEnvironmentVariableW` in that function body, then patch only that function.

## Latest Save Point - 2026-07-05 17:07 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-S2 is closed.** The shared DxUi modal loop exists, archive
Pack/Unpack prompt loops route through it, the quit-repost contract is guarded,
and docs/counts/ledger closeout have been patched. Next high-to-low Granite move:
**GR-S4 / GR-S4(f)**, then GR-C1 through GR-C6 and remaining stale cleanup.

Do not revert unrelated dirty files. `git status --short` still contains many
pre-existing modified files from accumulated Granite/Clearwater/Tailwind/DxUi work.
The files in scope for the completed GR-S2 slice were:

- `Common/DxUi/DxUi.h`
- `Common/DxUi/DxUi.cpp`
- `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp`
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
- `Tools/Tests/TestInventory.Tests.ps1`
- `Tests/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/UI/UI_DxUiSharedGrid.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`

GR-S2 implementation state:

- `Common/DxUi/DxUi.h` declares `DxUiModalLoopResult`,
  `DxUiModalLoopOptions`, and
  `RunDxUiModalLoop(HWND, const DxUiModalLoopOptions&) noexcept`.
- `Common/DxUi/DxUi.cpp` implements the shared nested message pump with one
  `GetMessageW` failure diagnostic, normal `TranslateMessage` /
  `DispatchMessageW` routing, optional quit/continue callbacks, and
  `PostQuitMessage(static_cast<int>(msg.wParam))` before returning
  `DxUiModalLoopResult::Quit`.
- `ArchivePackPromptWindow::ShowModal()` and
  `ArchiveUnpackPromptWindow::ShowModal()` now call
  `RunDxUiModalLoop(_hWnd.get(), ...)`; those class blocks no longer carry local
  `GetMessageW(&msg, ...)` loops.
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` guards the API shape, both
  archive prompt migrations, absence of local archive-prompt `GetMessageW(&msg)`
  loops, and quit propagation.
- `Specs/UI/UI_DxUiSharedGrid.md` now contains the durable modal prompt contract:
  nested DxUi modal pumps must use `RunDxUiModalLoop(...)`, keep the shared
  failure diagnostic, and repost `WM_QUIT` so the outer application loop still
  observes shutdown.

GR-S2 evidence:

- RED #1 `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`
  reported 38 / 1 / 0 before `DxUiModalLoopOptions` existed.
- GREEN #1 after helper + archive prompt migration passed 39 / 0 / 0.
- A focused pack prompt runtime exposed a quit-propagation hazard; preserved
  evidence lives at
  `Specs/TestRuns/4cb089111a23/Continuation/2026-07-05_1700_gr-s2_archive_prompt_hang/`.
- RED #2 after adding the quit-propagation source contract reported 38 / 1 / 0
  before the helper reposted `WM_QUIT`.
- GREEN #2 after adding `PostQuitMessage(static_cast<int>(msg.wParam))` passed
  39 / 0 / 0.
- GREEN app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug
  -Platform x64` passed with 0 warnings / 0 errors; log
  `.build\logs\msbuild-20260705_170405_259.log`.
- Focused pack prompt runtime passed 1 / 0 / 0:
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_170612/`.
- Focused unpack prompt runtime passed 1 / 0 / 0:
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_170618/`.

Closeout docs/counts patched:

- `Tools/Tests/TestInventory.Tests.ps1`: Tooling Script Tests 126.
- `Tests/README.md`: Tooling Script Tests 126 and
  `TestHarnessSourceContracts.Tests.ps1` 35.
- `Specs/Testing/Testing_TestCoverage.md`: Tooling scripts 126 and a GR-S2
  checkpoint with RED/GREEN, build, runtime, and preserved hung-run evidence.
- Granite WIP plan marks GR-S2 `LOW - DONE 2026-07-05`.
- WIP README lists GR-S2 done and points next at GR-S4 / GR-S4(f).

Immediate resume checklist:

1. Run final closeout checks if they were not completed before the next turn:

   ```powershell
   Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
   Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
   git diff --check -- Common/DxUi/DxUi.h Common/DxUi/DxUi.cpp RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp Tools/Tests/TestHarnessSourceContracts.Tests.ps1 Tools/Tests/TestInventory.Tests.ps1 Tests/README.md Specs/Testing/Testing_TestCoverage.md Specs/UI/UI_DxUiSharedGrid.md Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md Specs/Plans/WIP/README.md Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md
   ```

2. Continue high-to-low with `GR-S4` / `GR-S4(f)`.

## Latest Save Point - 2026-07-05 16:52 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-A4 is fully closed; GR-S2 has been started with a deliberate
TDD RED source-contract test, and the next move is to implement the shared DxUi
modal loop helper and route the archive pack/unpack prompt loops through it.**

Do not revert unrelated dirty files. `git status --short` still contains many
pre-existing modified files from accumulated Granite/Clearwater/Tailwind/DxUi work.
The files in scope for the immediate GR-S2 continuation are:

- `Common/DxUi/DxUi.h`
- `Common/DxUi/DxUi.cpp`
- `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp`
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
- `Tools/Tests/TestInventory.Tests.ps1`
- `Tests/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`

GR-A4 closeout completed before this save point:

- `Tools/Tests/TestInventory.Tests.ps1` now expects 125 Tooling Script Tests.
- `Tests/README.md` records Tooling Script Tests at 125 and
  `TestHarnessSourceContracts.Tests.ps1` at 34, including the shared ordinal string
  helper coverage.
- `Specs/Testing/Testing_TestCoverage.md` records the 125 Tooling Script Tests
  total and a GR-A4 shared ordinal string helper checkpoint.
- `Specs/FileSystem/FileSystem_Mtp.md` now contains the durable MTP identity
  comparison contract: intentionally case-insensitive device/session/cache-key
  comparisons use shared `OrdinalString` helpers (`EqualsNoCase`,
  `FoldCaseInvariant`) rather than local CRT `towlower` loops.
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
  marks GR-A4 as `LOW - DONE 2026-07-05` and includes a GR-A4 closeout note.
- `Specs/Plans/WIP/README.md` lists GR-A4 as done and points the next high-to-low
  move at GR-S2.

Fresh GR-A4 closeout evidence:

- `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`
  passed 38 / 0 / 0 before the new GR-S2 RED test was added.
- `Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru` passed
  5 / 0 / 0.
- `git diff --check -- Common/SearchServiceBroker.cpp Plugins/FileSystemMtp/FileSystemMtp.Device.cpp Tools/Tests/TestHarnessSourceContracts.Tests.ps1 Tools/Tests/TestInventory.Tests.ps1 Tests/README.md Specs/Testing/Testing_TestCoverage.md Specs/FileSystem/FileSystem_Mtp.md Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md Specs/Plans/WIP/README.md Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`
  exited 0, with only LF-to-CRLF warnings.

GR-S2 TDD RED now in place:

- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` has a new test:
  `routes archive prompts through the shared DxUi modal loop`.
- The test asserts that `Common/DxUi/DxUi.h` declares:
  - `struct DxUiModalLoopOptions final`
  - `enum class DxUiModalLoopResult`
  - `RunDxUiModalLoop(...)`
- It asserts that `Common/DxUi/DxUi.cpp` defines
  `DxUiModalLoopResult RunDxUiModalLoop(...)`.
- It asserts that the `ArchivePackPromptWindow` and `ArchiveUnpackPromptWindow`
  `ShowModal()` blocks in `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp`
  call `RunDxUiModalLoop(_hWnd.get(), ...)` and no longer contain
  `GetMessageW(&msg`.
- RED command:

  ```powershell
  Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
  ```

  Result: Total 39, Passed 38, Failed 1, Skipped 0. The expected failure is that
  `Common/DxUi/DxUi.h` does not yet contain `struct DxUiModalLoopOptions final`.

Immediate resume checklist:

1. Implement the minimal GR-S2 GREEN slice:

   - Add `DxUiModalLoopResult`, `DxUiModalLoopOptions`, and
     `RunDxUiModalLoop(HWND, const DxUiModalLoopOptions&) noexcept` to the shared
     DxUi API.
   - Implement the helper in `Common/DxUi/DxUi.cpp` with a single diagnostic point
     for `GetMessageW` failure and normal `TranslateMessage` / `DispatchMessageW`
     handling.
   - Replace the local `GetMessageW` modal loops in
     `ArchivePackPromptWindow::ShowModal()` and `ArchiveUnpackPromptWindow::ShowModal()`
     with `RunDxUiModalLoop(_hWnd.get(), ...)`.
   - Preserve existing quit behavior: `WM_QUIT` should set the prompt `_done`
     state and return the current optional result, while `GetMessageW == -1`
     returns `std::nullopt`.

2. Run the GR-S2 source-contract test:

   ```powershell
   Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
   ```

   Expected after implementation: 39 / 0 / 0.

3. Build after the helper and archive prompt migration:

   ```powershell
   .\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64
   ```

4. Find and run focused archive prompt/command selftests if available:

   ```powershell
   rg -n "ArchivePack|ArchiveUnpack|pack prompt|unpack prompt|archive" RedSalamander/SelfTest/Commands
   ```

   Then run the narrowest matching `Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild`
   case filter. If no focused archive prompt selftest exists, record that gap in the
   GR-S2 closeout and rely on Pester plus Debug build for this mechanical helper
   extraction.

5. Patch docs/counts after the GR-S2 GREEN evidence:

   - `Tools/Tests/TestInventory.Tests.ps1`: Tooling Script Tests 125 -> 126.
   - `Tests/README.md`: Tooling Script Tests 125 -> 126 and
     `TestHarnessSourceContracts.Tests.ps1` 34 -> 35; mention shared DxUi modal
     loop/archive prompt coverage.
   - `Specs/Testing/Testing_TestCoverage.md`: Tooling scripts 125 -> 126 and add
     a GR-S2 shared DxUi modal loop checkpoint with RED/GREEN, build log, and any
     runtime archive or explicit selftest gap.
   - Granite WIP plan: mark GR-S2 done if the archive pack/unpack loops and helper
     extraction are complete; note that other modal loops remain opportunistic.
   - WIP README: add GR-S2 to done/current state and set the next high-to-low move
     to GR-S4 / GR-S4(f), then GR-C1 through GR-C6 and stale cleanup.

6. Run final closeout checks after docs/counts:

   ```powershell
   Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
   Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
   git diff --check -- Common/DxUi/DxUi.h Common/DxUi/DxUi.cpp RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp Tools/Tests/TestHarnessSourceContracts.Tests.ps1 Tools/Tests/TestInventory.Tests.ps1 Tests/README.md Specs/Testing/Testing_TestCoverage.md Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md Specs/Plans/WIP/README.md Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md
   ```

After GR-S2, continue high-to-low with `GR-S4`, `GR-S4(f)`, `GR-C1` through
`GR-C6`, and the remaining stale test-row cleanup.

## Latest Save Point - 2026-07-05 16:46 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-A5 is ratcheted/closed as the first bridge-IO decorator slice;
GR-A4 code, source-contract tests, Debug build, and focused runtime checks are
green, but GR-A4 docs/counts/ledger closeout is still pending.** Treat this as a
saved mid-closeout state for GR-A4, not a done state.

Do not revert unrelated dirty files. `git status --short` still contains many
pre-existing modified files from accumulated Granite/Clearwater/Tailwind/DxUi work.
The in-scope files for the current GR-A4 handoff are:

- `Common/SearchServiceBroker.cpp`
- `Plugins/FileSystemMtp/FileSystemMtp.Device.cpp`
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
- `Tools/Tests/TestInventory.Tests.ps1`
- `Tests/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/FileSystem/FileSystem_Mtp.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`

Already completed before this save point:

- GR-A5 bridge IO decorator ratchet docs/counts were updated. Current docs still
  show the Tooling Script Tests count as 124 and `TestHarnessSourceContracts` as
  33 after the GR-A5 source-contract case.
- GR-A5 WIP row was recorded as `MEDIUM (long-term) - RATCHETED 2026-07-05`.
- GR-A5 evidence:
  - RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`
    failed 36 / 1 / 0 before `SelfTestBridgeIoRole` existed.
  - GREEN Pester passed 37 / 0 / 0 after the decorator implementation.
  - GREEN build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
    passed with 0 warnings / 0 errors; log `.build\logs\msbuild-20260705_162741_277.log`.
  - GREEN FileOps focused archive
    `Specs/TestRuns/4cb089111a23/FileOps/2026-07-05_162948/` for
    `Floodgate_CrossFsCopyGetSizeFailureRefusesCommit`, passed 3 / 0 / 0.
  - GREEN FileOps focused archive
    `Specs/TestRuns/4cb089111a23/FileOps/2026-07-05_163017/` for
    `Floodgate_CrossFsMoveGetSizeFailurePreservesSource`, passed 3 / 0 / 0.

GR-A4 code/test state at this save point:

- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` now has
  `uses shared ordinal string helpers instead of local case-folding predicates`.
- That test asserts the targeted GR-A4 local helpers are gone from the named files:
  `EqualsNoCaseLocal` in `ConnectionManagerWindow.cpp`,
  `EqualsOrdinalIgnoreCase` in `FileSystemMtp.Device.cpp`, local
  `StartsWithNoCase` definitions in `FileSystemMtp.Shared.cpp` and
  `Common/SearchServiceBroker.cpp`, plus the local MTP `CaseFoldKey` no longer
  uses `::towlower`.
- `Plugins/FileSystemMtp/FileSystemMtp.Device.cpp` now routes same-device/display
  and hash comparisons through `OrdinalString::EqualsNoCase(...)`.
- `CaseFoldKey(std::wstring_view)` now returns
  `OrdinalString::FoldCaseInvariant(value)`.
- `Common/SearchServiceBroker.cpp` removed the local `StartsWithNoCase(...)` helper
  and now uses `OrdinalString::StartsWithNoCase(...)` for root containment and
  device/extended-prefix rejection checks.
- `ConnectionManagerWindow.cpp` and `FileSystemMtp.Shared.cpp` were already clean in
  the current tree and the new source-contract test keeps them clean.

Fresh GR-A4 evidence already gathered:

- RED #1 `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`
  failed 37 / 1 / 0 because `EqualsOrdinalIgnoreCase` still existed.
- GREEN #1 after removing the named duplicate predicates passed 38 / 0 / 0.
- RED #2 after temporarily restoring the old `CaseFoldKey` manual `::towlower` loop
  failed 37 / 1 / 0 because the helper did not use
  `OrdinalString::FoldCaseInvariant`.
- GREEN #2 after changing `CaseFoldKey` passed 38 / 0 / 0.
- GREEN final Debug build `.\build.ps1 -Configuration Debug -Platform x64` passed
  with 0 warnings / 0 errors; log `.build\logs\msbuild-20260705_164220_979.log`.
- GREEN focused Compare runtime
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_wpd_session_and_path_cache_reuse -FailFast -TimeoutMultiplier 3`
  passed 1 / 0 / 0; final archive
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_164425/`.
- GREEN focused Compare runtime
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_service_rejects_device_root_and_continues -FailFast -TimeoutMultiplier 3`
  passed 1 / 0 / 0; final archive
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_164455/`.

Immediate resume checklist:

1. Finish GR-A4 docs, counts, and ledger closeout:

   - `Tools/Tests/TestInventory.Tests.ps1`: update both ToolsPester expected counts
     from 124 to 125.
   - `Tests/README.md`: update Tooling Script Tests from 124 to 125 and
     `TestHarnessSourceContracts.Tests.ps1` from 33 to 34; mention shared ordinal
     string helper usage in the row description.
   - `Specs/Testing/Testing_TestCoverage.md`: update Tooling scripts from 124 to
     125, add a GR-A4 checkpoint with the RED/GREEN Pester evidence, final build
     log, and the two focused Compare archives above, and update the
     `TestHarnessSourceContracts.Tests.ps1` table row for shared ordinal helpers.
   - `Specs/FileSystem/FileSystem_Mtp.md`: add the durable MTP contract that
     intentionally case-insensitive device/session/cache-key comparisons must use
     shared `OrdinalString` helpers such as `EqualsNoCase` and
     `FoldCaseInvariant`, not local CRT `towlower` loops. Keep this scoped so it
     does not contradict the existing provider path identity
     `ordinalCaseSensitive` contract.
   - `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`:
     mark GR-A4 as `LOW - DONE 2026-07-05`, record the source-contract/build/runtime
     proof, and add a `GR-A4 shared ordinal string helper closeout note`.
   - `Specs/Plans/WIP/README.md`: add GR-A4 to done/current state and set the next
     high-to-low move to GR-S2, unless you intentionally choose another immediate
     GR-A5 decorator migration first.

2. Run final closeout checks after the GR-A4 docs/counts are patched:

   ```powershell
   Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
   Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
   git diff --check -- Common/SearchServiceBroker.cpp Plugins/FileSystemMtp/FileSystemMtp.Device.cpp Tools/Tests/TestHarnessSourceContracts.Tests.ps1 Tools/Tests/TestInventory.Tests.ps1 Tests/README.md Specs/Testing/Testing_TestCoverage.md Specs/FileSystem/FileSystem_Mtp.md Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md Specs/Plans/WIP/README.md Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md
   ```

3. Continue high-to-low after GR-A4 closeout:

   - Next practical row: `GR-S2`.
   - Still open later: `GR-S4`, `GR-S4(f)`, `GR-C1` through `GR-C6`, and remaining
     stale test-row cleanup.

## Latest Save Point - 2026-07-05 16:32 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-P3 and GR-S1 are closed; the first concrete GR-A5
`IFileSystemIO` decorator ratchet slice is implemented and verified, but GR-A5
closeout docs/inventory/ledger updates are still pending.** Treat this as a saved
mid-closeout state, not a done state.

Do not revert unrelated dirty files. `git status --short` at this save point still
contains many pre-existing modified files from accumulated Granite/Clearwater/
Tailwind/DxUi work. The current in-scope Granite files are:

- `RedSalamander/FolderWindow.FileOperations.State.cpp`
- `RedSalamander/FolderWindow.FileOperations.State.Queue.Part.cpp`
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
- `Tools/Tests/TestInventory.Tests.ps1`
- `Tests/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`

GR-A5 ratchet slice implemented:

- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` now has the source-contract
  test `routes FileOperations bridge source-size failures through the IO decorator
  seam`.
- RED for that test failed as expected with 36 passed / 1 failed / 0 skipped because
  `enum class SelfTestBridgeIoRole` did not exist yet.
- `RedSalamander/FolderWindow.FileOperations.State.cpp` now defines
  `SelfTestBridgeIoRole`, `SelfTestBridgeFileReader final : IFileReader`,
  `SelfTestBridgeIoDecorator final : IFileSystemIO`, and
  `DecorateBridgeIoForSelfTest(...)` under `ENABLE_TESTS`.
- Bridge source and destination `IFileSystemIO` objects are decorated at acquisition
  with `DecorateBridgeIoForSelfTest(fileSystemIo, SelfTestBridgeIoRole::Source)`
  and `DecorateBridgeIoForSelfTest(destinationFileSystemIo,
  SelfTestBridgeIoRole::Destination)`.
- `ConsumeBridgeFailNextSourceGetSizeForSelfTest()` moved out of the inline
  `CopyFileWithBuffer(...)` source-size branch and is now consumed by the source
  `SelfTestBridgeFileReader::GetSize(...)` decorator path.
- `CopyFileWithBuffer(...)` now calls `reader->GetSize(&fileTotalBytes)` directly.

Fresh GR-A5 evidence already gathered:

- GREEN `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`
  passed 37 / 0 / 0 after the decorator implementation.
- GREEN `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors; build log
  `.build\logs\msbuild-20260705_162741_277.log`.
- GREEN focused runtime
  `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Floodgate_CrossFsCopyGetSizeFailureRefusesCommit -FailFast -TimeoutMultiplier 3`
  passed 3 / 0 / 0; archive
  `Specs/TestRuns/4cb089111a23/FileOps/2026-07-05_162948/`.
- GREEN focused runtime
  `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Floodgate_CrossFsMoveGetSizeFailurePreservesSource -FailFast -TimeoutMultiplier 3`
  passed 3 / 0 / 0; archive
  `Specs/TestRuns/4cb089111a23/FileOps/2026-07-05_163017/`.

Immediate resume checklist:

1. Finish the GR-A5 ratchet closeout docs and counts:

   - `Tools/Tests/TestInventory.Tests.ps1`: Tooling Script Tests expected counts
     123 -> 124 in both ToolsPester assertions.
   - `Tests/README.md`: Tooling Script Tests 123 -> 124;
     `TestHarnessSourceContracts.Tests.ps1` 32 -> 33; mention the FileOperations
     bridge IO decorator/source-size fault seam.
   - `Specs/Testing/Testing_TestCoverage.md`: Tooling scripts 123 -> 124; add a
     2026-07-05 GR-A5 bridge IO decorator ratchet checkpoint with the RED/GREEN,
     build log, and two FileOps runtime archives above.
   - `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`:
     record GR-A5 as `RATCHETED 2026-07-05` or equivalent unless you decide the
     interface-boundary seam plus first migration is enough to mark the long-term
     row done. Recommended disposition: leave the broader GR-A5 migration open or
     explicitly long-term, with this source-size hook migrated.
   - `Specs/Plans/WIP/README.md`: update Granite's next action after the GR-A5
     ratchet disposition. If GR-A5 remains open, say so; otherwise continue to
     `GR-A4` / `GR-S2`.

2. Run final closeout checks from the current state:

   ```powershell
   Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
   Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
   git diff --check -- RedSalamander/FolderWindow.FileOperations.State.cpp RedSalamander/FolderWindow.FileOperations.State.Queue.Part.cpp Tools/Tests/TestHarnessSourceContracts.Tests.ps1 Tools/Tests/TestInventory.Tests.ps1 Tests/README.md Specs/Testing/Testing_TestCoverage.md Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md Specs/Plans/WIP/README.md Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md
   ```

3. Continue high-to-low:

   - If GR-A5 remains open after the ratchet note, either perform the next decorator
     seam migration or explicitly defer the remaining opportunistic migrations in
     the WIP ledger.
   - Next lower rows after GR-A5 disposition are `GR-A4`, `GR-S2`, `GR-S4`,
     `GR-S4(f)`, then `GR-C1` through `GR-C6` and the remaining stale test-row
     cleanup.

## Latest Save Point - 2026-07-05 16:21 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-P3 and GR-S1 are closed in the working tree; GR-A5 remains
open.** GR-S1 was closed as the first concrete GR-A5 ratchet slice by centralizing
the duplicated FileOperations self-test pause-point machinery. This does **not**
complete the broader GR-A5 interface-boundary `IFileSystemIO` decorator seam.

Do not revert unrelated dirty files. `git status --short` still contains many
pre-existing modified files from accumulated Granite/Clearwater/Tailwind/DxUi work.

GR-S1 closeout completed:

- `RedSalamander/FolderWindow.FileOperations.State.cpp` now defines
  `SelfTestPausePoint final` with enabled/entered/release atomics plus
  `Set(...)`, `HasEntered()`, `Release()`, and `Pause(...)`.
- The post-finished-completion and bridge-move-source-cleanup quiet points are two
  `SelfTestPausePoint` instances instead of two cloned triples of atomics.
- `RedSalamander/FolderWindow.FileOperations.State.Queue.Part.cpp` keeps the
  existing public self-test helper names but delegates them to the pause-point
  instances.
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` now guards the centralized
  pause-point contract.
- `Tools/Tests/TestInventory.Tests.ps1`, `Tests/README.md`,
  `Specs/Testing/Testing_TestCoverage.md`, the Granite plan, and the WIP README
  were updated for the added Pester case and GR-S1 closure.

Fresh GR-S1 evidence:

- RED `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`
  first failed at 35 passed / 1 failed / 0 skipped because
  `struct SelfTestPausePoint final` did not exist.
- GREEN `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`
  passed 36 / 0 / 0.
- GREEN `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors after explicit deleted copy/move members;
  build log `.build\logs\msbuild-20260705_161428_240.log`.
- GREEN focused runtime
  `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Riptide_LiveFinishedSnapshotCarriesDiagnostics -FailFast -TimeoutMultiplier 3`
  passed 3 / 0 / 0; archive
  `Specs/TestRuns/4cb089111a23/FileOps/2026-07-05_161924/`.
- GREEN focused runtime
  `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Floodgate_CrossFsMoveCleanupDetectsDestinationCorruption -FailFast -TimeoutMultiplier 3`
  passed 3 / 0 / 0; archive
  `Specs/TestRuns/4cb089111a23/FileOps/2026-07-05_161955/`.
- GREEN `Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru` passed
  5 / 0 / 0 after the Tooling Script Tests count moved to 123.

Immediate next move:

1. Continue high-to-low with **GR-A5**. Its broad decorator seam is still open; do
   not mark it done from the GR-S1 helper extraction alone.
2. Re-read the current `GR-A5` row and nearby FileOperations hook sites in
   `FolderWindow.FileOperations.State.cpp`.
3. Decide the smallest next decorator-seam ratchet that makes the final GR-A5 state
   more true. New tests should use the interface-boundary/decorator direction rather
   than adding another global/env-var/call-site hook.
4. If GR-A5 needs to remain long-term after a documented ratchet/disposition, then
   continue the low cleanup rows in order: `GR-A4`, `GR-S2`, `GR-S4`, `GR-S4(f)`,
   then `GR-C1` through `GR-C6` and remaining test-row cleanup.

## Latest Save Point - 2026-07-05 16:17 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-P3 is closed; GR-A5 has started; GR-S1 is implemented but not
closed.** Treat GR-S1 as the first concrete GR-A5 ratchet slice. The code/test
change is green at source-contract and Debug-build level, but focused FileOps runtime
verification plus docs/inventory/ledger closeout are still pending.

Do not revert unrelated dirty files. `git status --short` still contains many
pre-existing modified files from accumulated Granite/Clearwater/Tailwind/DxUi work.

### Completed Since The Previous Save Point

- Added a RED source-contract test in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`:
  `keeps FileOperations self-test pause points centralized`.
- Implemented `SelfTestPausePoint final` in
  `RedSalamander/FolderWindow.FileOperations.State.cpp`.
- Replaced the duplicated post-finished-completion and bridge-move-source-cleanup
  pause atomics with two `SelfTestPausePoint` instances.
- Updated
  `RedSalamander/FolderWindow.FileOperations.State.Queue.Part.cpp` so the existing
  public test helper names delegate to `.Set(enabled)`, `.HasEntered()`, and
  `.Release()`.
- Added explicit deleted copy/move special members to silence `/Wall` C4625/C4626
  and C5026/C5027 warnings caused by atomic members.

### Fresh Evidence From This Save Point

RED:

```powershell
Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
```

- Expected failure before implementation: 35 passed / 1 failed / 0 skipped.
- Failure proved the missing `struct SelfTestPausePoint final` contract.

GREEN:

```powershell
Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
```

- Passed 36 / 0 / 0 after the code change.

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64
```

- Passed 0 warnings / 0 errors after the explicit deleted copy/move members.
- Build log: `.build\logs\msbuild-20260705_161428_240.log`.
- Produced `.build\x64\Debug\RedSalamander.exe`.

### Immediate Resume Checklist

1. Run focused FileOps runtime verification against the existing Debug build:

   ```powershell
   .\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Riptide_LiveFinishedSnapshotCarriesDiagnostics -FailFast -TimeoutMultiplier 3
   .\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Floodgate_CrossFsMoveCleanupDetectsDestinationCorruption -FailFast -TimeoutMultiplier 3
   ```

   Optional extra bridge-cleanup coverage if time allows:

   ```powershell
   .\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Floodgate_CrossFsDirectoryMoveCleanupPreservesChangedSource -FailFast -TimeoutMultiplier 3
   ```

2. If runtime is green, update docs/inventory for the added Pester case:

   - `Tools/Tests/TestInventory.Tests.ps1`: Tools Pester expected counts 122 -> 123
     in both ToolsPester assertions. Do not change FileOperations active phase
     counts.
   - `Tests/README.md`: Tooling Script Tests 122 -> 123;
     `TestHarnessSourceContracts.Tests.ps1` 31 -> 32; mention FileOperations
     pause-point centralization.
   - `Specs/Testing/Testing_TestCoverage.md`: Tooling scripts 122 -> 123; add the
     GR-S1/GR-A5 ratchet-slice RED/GREEN/build/runtime checkpoint and the fresh
     FileOps archive paths.
   - `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`:
     mark `GR-S1` DONE with a 2026-07-05 closeout note. Leave `GR-A5` open unless
     a broader decorator-seam disposition/ratchet contract is also established.
   - `Specs/Plans/WIP/README.md`: after GR-S1 closes, move the next action back to
     GR-A5 ratchet/disposition, then continue the low rows (`GR-A4`, `GR-S2`,
     `GR-S4`, `GR-S4(f)`, `GR-C1`...).

3. Final focused checks for the GR-S1 closeout:

   ```powershell
   Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
   Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
   git diff --check -- RedSalamander/FolderWindow.FileOperations.State.cpp RedSalamander/FolderWindow.FileOperations.State.Queue.Part.cpp Tools/Tests/TestHarnessSourceContracts.Tests.ps1 Tools/Tests/TestInventory.Tests.ps1 Tests/README.md Specs/Testing/Testing_TestCoverage.md Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md Specs/Plans/WIP/README.md Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md
   ```

4. After GR-S1 is closed, continue high-to-low with GR-A5 disposition/ratchet work.
   The full GR-A5 decorator seam is long-term and is **not** completed by the
   pause-point extraction alone.

## Latest Save Point - 2026-07-05 16:09 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-P3 is closed in the working tree.** The stale `GR-TA6` test-table
row was also corrected to DONE via the already-closed GR-A6 implementation. The WIP
README now points Granite to **GR-A5** as the next high-to-low decision before the low
cleanup rows.

GR-P3 closeout completed:

- `Tools/TestRunPlan.ps1` and `Tools/Run-AllTests.ps1` thread optional
  `-PerfBudgetPath` / `-RequirePerfBudgets` into native self-test plan entries.
- `RedSalamander.exe` parses `--selftest-require-perf-budgets` into
  `SelfTestOptions::requirePerfBudgets`.
- `CheckFolderViewPerfBudgets` reads `Specs/Testing/FolderViewPerfBudgets.json5`
  as `machines[]`, warns with a scaffold for unknown machines, and strict-fails
  missing path/current-machine/no-hard-build matches.
- `Tools/Tests/RunAllTestsPlan.Tests.ps1` and
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` guard the runner and source
  contracts.
- `Tools/Tests/TestInventory.Tests.ps1`, `Tests/README.md`,
  `Specs/Testing/Testing_TestCoverage.md`,
  `Specs/Testing/Testing_PerformanceValidation.md`, the Granite ledger, and the WIP
  README were updated.

Fresh verification from this save point:

- `Invoke-Pester .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru`: 9 passed / 0
  failed / 0 skipped.
- `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`: 35
  passed / 0 failed / 0 skipped.
- `Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru`: 5 passed / 0
  failed / 0 skipped.
- `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_thumbnail_cached_only_no_close_stall -FailFast -TimeoutMultiplier 3 -PerfBudgetPath Specs\Testing\FolderViewPerfBudgets.json5`:
  passed 1 / 0 / 0; archive
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_160750/`; trace shows the
  multi-machine budget file matched the current machine before Debug skipped the
  Release-only hard budget.
- Targeted `git diff --check` over the GR-P3 touched file set exited 0, with only
  LF-to-CRLF working-copy warnings.

Immediate next move:

1. Re-read `GR-A5` in the Granite plan. It is **MEDIUM (long-term)** and asks for a
   fault-injection seam/ratchet, not a one-shot wholesale migration of every existing
   hook.
2. Decide the smallest closeable GR-A5 slice that makes the final state more true:
   likely establish or document the ratchet/seam contract, then migrate an actually
   touched hook. `GR-S1` is explicitly subsumed by GR-A5 but says to do the pause-point
   extraction now, so it is a plausible first concrete sub-slice if the GR-A5 seam
   stays long-term.
3. After GR-A5 is disposed/ratcheted, continue the low rows in order: `GR-A4`, `GR-S1`,
   `GR-S2`, `GR-S4`, `GR-S4(f)`, then `GR-C1` through `GR-C6` and remaining test-row
   cleanup such as `GR-T1`/stale `GR-1` recheck.

## Latest Save Point - 2026-07-05 16:04 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-P3 is implemented in the working tree and has green focused code
verification, but it is not closed yet.** The remaining work is docs/inventory/ledger
closeout plus final focused verification. The WIP README has been redirected from
`GR-S4(f)` to the GR-P3 closeout step so the next session resumes at the right place.

Why GR-P3 was selected: a fresh open-row scan after GR-24 found `GR-A5` as
**MEDIUM (long-term)** and `GR-P3` as **MEDIUM (tooling)**, followed by low/cleanup
rows (`GR-A4`, `GR-S1`, `GR-S2`, `GR-S4`, `GR-S4(f)`, `GR-C1` through `GR-C6`,
`GR-T1`, and the stale text around `GR-TA6`). Because the active objective is
high-to-low, `GR-P3` is the current concrete actionable item; `GR-A5` remains
long-term/opportunistic by its row text.

Do not revert unrelated dirty files. `git status --short` still shows many modified
files from accumulated Granite/Clearwater/Tailwind/DxUi work.

### GR-P3 Current Code/Test State

GR-P3 bug:

- `--selftest-perf-budget=` was passed by no automated runner.
- `Specs/Testing/FolderViewPerfBudgets.json5` used one top-level `machineHash`.
- `CheckFolderViewPerfBudgets` traced and returned true on machine mismatch.
- Empty budget path returned true silently.

Implementation already present:

- `Tools/TestRunPlan.ps1`:
  - `Get-RSSelfTestArguments` accepts `-PerfBudgetPath` and
    `-RequirePerfBudgets`.
  - Native self-test entries receive `--selftest-perf-budget=<path>` only when a
    path is supplied.
  - Native self-test entries receive `--selftest-require-perf-budgets` only when
    requested.
  - Non-selftest plan entries do not receive the native perf-budget arguments.
- `Tools/Run-AllTests.ps1`:
  - Adds `-PerfBudgetPath` and `-RequirePerfBudgets` parameters and forwards them
    into `Get-RSTestRunPlan`.
- `RedSalamander/SelfTest/Common/SelfTestCommon.h`:
  - `SelfTestOptions` now has `bool requirePerfBudgets = false;`.
- `RedSalamander/RedSalamander.cpp`:
  - Help text documents `--selftest-require-perf-budgets`.
  - Unsupported selftest-arg detection knows the flag for non-test builds.
  - Test builds parse it into `g_selfTestOptions.requirePerfBudgets`.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`:
  - Budget file parsing now expects top-level `"machines": [ ... ]`.
  - Each machine entry has `"machineHash"` plus `"budgets"`.
  - `FindFolderViewPerfBudgetMachine(...)` selects the current machine.
  - Missing budget path, missing current-machine entry, and no hard entry for the
    current build are visible trace/warning paths.
  - With `requirePerfBudgets`, those non-applicable cases fail the focused perf
    selftest instead of silently passing.
  - Unknown machines emit a scaffold message:
    `FolderView perf budget machine entry missing ... add a machines[] entry like { "machineHash": "...", "budgets": [] }.`
- `Specs/Testing/FolderViewPerfBudgets.json5`:
  - Converted from one top-level `machineHash` to a `machines[]` array keyed by
    machine hash.

Tests already added:

- `Tools/Tests/RunAllTestsPlan.Tests.ps1`
  - New case: `threads optional perf budget gates into native self-test entries only`.
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
  - New case: `keeps FolderView perf budgets multi-machine and require-gated`.

### GR-P3 Evidence Already Captured

RED Pester before implementation:

```powershell
Invoke-Pester .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru
```

- Output showed Passed 8, Failed 1.
- Failure: missing `PerfBudgetPath` parameter.

```powershell
Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
```

- Output showed Passed 34, Failed 1.
- Failure: missing `--selftest-require-perf-budgets` source contract.

GREEN Pester after implementation:

```powershell
Invoke-Pester .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru
Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
```

- `RunAllTestsPlan.Tests.ps1`: Passed 9 / Failed 0.
- `TestHarnessSourceContracts.Tests.ps1`: Passed 35 / Failed 0.

GREEN app build:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64
```

- Passed with 0 warnings / 0 errors.
- Log: `.build\logs\msbuild-20260705_155942_502.log`.
- Produced `.build\x64\Debug\RedSalamander.exe`.

GREEN runtime runner/parser exercise:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_thumbnail_cached_only_no_close_stall -FailFast -TimeoutMultiplier 3 -PerfBudgetPath 'Specs\Testing\FolderViewPerfBudgets.json5'
```

- Passed 1 / 0 / 0.
- Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_160204/`.
- `commands_trace.txt` confirms the new route parsed the multi-machine file and
  found the current machine before skipping Debug because the hard budget is Release:
  `FolderView perf budgets skipped for folderView_thumbnail_cached_only_no_close_stall build Debug; no hard entries matched this build.`

### Files Touched For GR-P3

Code/tests/budget shape:

- `Tools/Tests/RunAllTestsPlan.Tests.ps1`
- `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
- `Tools/TestRunPlan.ps1`
- `Tools/Run-AllTests.ps1`
- `RedSalamander/SelfTest/Common/SelfTestCommon.h`
- `RedSalamander/RedSalamander.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- `Specs/Testing/FolderViewPerfBudgets.json5`

Save-state files touched now:

- `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`
- `Specs/Plans/WIP/README.md`

### Immediate Resume Checklist

1. Finish GR-P3 durable docs/inventory:

   - `Tools/Tests/TestInventory.Tests.ps1`: update Tools Pester expected count from
     120 to 122 only for the Tools Pester assertions; do not change FileOperations
     active phase counts.
   - `Tests/README.md`: Tooling Script Tests 120 -> 122; `RunAllTestsPlan` 8 -> 9;
     `TestHarnessSourceContracts` 30 -> 31; mention the perf budget runner/source
     contracts.
   - `Specs/Testing/Testing_TestCoverage.md`: Tooling scripts 120 -> 122; update
     the tool table descriptions for the two edited Pester files; add a 2026-07-05
     GR-P3 checkpoint note.
   - `Specs/Testing/Testing_PerformanceValidation.md`: document `machines[]`,
     visible unknown-machine scaffold warnings, `--selftest-require-perf-budgets`,
     and `Run-AllTests.ps1 -PerfBudgetPath ... -RequirePerfBudgets`.
   - `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`:
     mark `GR-P3` done only after the docs and verification below are green; add a
     GR-P3 closeout note with the RED/GREEN evidence above.
   - `Specs/Plans/WIP/README.md`: after GR-P3 is truly closed, move the Granite
     next action back to the next high-to-low open item; likely `GR-S4(f)` after
     re-checking `GR-A5`/low rows.

2. Run focused closeout verification:

   ```powershell
   Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
   Invoke-Pester .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru
   Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
   git diff --check
   ```

3. Only after those pass, close `GR-P3` in the Granite ledger and prepend a new baton
   save point naming the next row.

## Latest Save Point - 2026-07-05 15:54 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-24 / GR-T24 is closed in the working tree, and the duplicate
Tailwind TW-1 / TW-T1 is closed via Granite GR-24.** The WIP README now points Granite
to **GR-S4(f)** as the next remaining cleanup/addendum item.

Priority caution from a fresh open-row scan: the Granite table still has open rows
including `GR-A5` (**MEDIUM (long-term)**), `GR-P3` (**MEDIUM (tooling)**), `GR-A4`,
`GR-S1`, `GR-S2`, `GR-S4`, `GR-S4(f)`, `GR-C1` through `GR-C6`, `GR-T1`, and
`GR-TA6`. Because the active objective says to work high-to-low, the next session
should first decide whether `GR-A5` and `GR-P3` are intentionally deferred/routed by
their row text before starting the README-routed `GR-S4(f)` cleanup. If following the
README pointer, re-verify `GR-S4(f)` inline: its D3D11 and pending-metric subitems may
already be partially stale after the later `GR-A6` / `GR-22` work.

Do not start a different Granite row without first re-checking the current ledger.

Do not revert unrelated dirty files. `git status --short` still shows many modified
files from accumulated Granite/Clearwater/Tailwind/DxUi work. The files touched for the
GR-24 closeout slice are:

- `RedSalamander/FolderView.h`
- `RedSalamander/FolderView.cpp`
- `RedSalamander/FolderView.Rendering.cpp`
- `RedSalamander/FolderWindow.h`
- `RedSalamander/FolderWindow.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- `Specs/UI/UI_FolderView.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Tests/README.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/Operation_Tailwind_FolderViewWarpDrivePreMergeReviewRemediation_2026-06-30.md`
- `Specs/Plans/WIP/README.md`
- This baton file

### GR-24 Closed State

GR-24 fix:

- `FolderView::DrawItem` still uses cached member brushes for normal selected and
  hovered item backgrounds.
- `FolderView` now has an `ENABLE_TESTS` debug flag and setter:
  `DebugSetForceDrawItemTransientBrushCreateForSelfTest(bool enabled) noexcept`.
- `FolderWindow` exposes
  `DebugSetPaneForceDrawItemTransientBrushCreateForSelfTest(Pane, bool) noexcept`.
- When the forced seam is enabled, `DrawItem` creates a RAII-owned transient
  `ID2D1SolidColorBrush` and increments `_debugDrawItemTransientBrushCreateCount`.
- `folderView_draw_item_brush_reuse_guard` now proves the counter is live: normal
  selected/hovered warm rendering requires `drawItemTransientBrushCreateCount == 0`,
  then the forced transient-brush positive control requires the same counter to become
  nonzero.
- `folderView_perf_huge_folder_scale` remains the adjacent scale guard asserting the
  zero-counter path while select-all scrolling the synthetic 10,000-item folder.

GR-24 evidence:

- RED build before the debug seam existed:
  `.build\logs\msbuild-20260705_153748_205.log`, failed with missing
  `FolderWindow::DebugSetPaneForceDrawItemTransientBrushCreateForSelfTest(...)`.
- GREEN app build:
  `.build\logs\msbuild-20260705_153949_588.log`, produced
  `.build\x64\Debug\RedSalamander.exe`.
- GREEN focused archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_154145/`, passed
  `folderView_draw_item_brush_reuse_guard` with 1 / 0 / 0.
- GREEN adjacent archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_154806/`, passed
  `folderView_perf_huge_folder_scale` with 1 / 0 / 0.
- `Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru` passed 5 / 0 / 0
  after the closeout docs/ledgers were updated.

Immediate next move:

1. Start **GR-S4(f)** from
   `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`.
2. Re-read the GR-S4(f) row before editing. It is cleanup/structural work covering:
   thumbnail shell counters threaded through multiple tiers, duplicated D3D11 creation
   plumbing, and duplicated pending-to-paint metric structs.
3. Before changing code, identify the smallest self-contained sub-slice and follow TDD
   or source-contract validation as appropriate.

## Latest Save Point - 2026-07-05 15:43 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-23 / GR-T23 is closed. GR-24 / GR-T24 is implemented and has a
focused GREEN selftest, but it is not closed yet.** Adjacent validation, durable docs,
Granite/Tailwind/WIP ledgers, stale sweeps, and `git diff --check` are still pending.

Do not revert unrelated dirty files. `git status --short` still shows many modified
files from accumulated Granite/Clearwater/Tailwind/DxUi work. The files touched for the
current GR-24 slice are:

- `RedSalamander/FolderView.h`
- `RedSalamander/FolderView.cpp`
- `RedSalamander/FolderView.Rendering.cpp`
- `RedSalamander/FolderWindow.h`
- `RedSalamander/FolderWindow.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`

### GR-24 Current Code/Test State

GR-24 bug:

- `_debugDrawItemTransientBrushCreateCount` was surfaced in
  `FolderView::RenderingDebugSnapshot`, reset by debug helpers, and asserted by
  selftests, but it had no producer. The `== 0u` assertions therefore passed
  vacuously and would not catch a future per-item `CreateSolidColorBrush` regression in
  `DrawItem`.

Test change already present:

- `TestFolderViewDrawItemBrushReuseGuard` still asserts the normal selected/hovered
  render path keeps `drawItemTransientBrushCreateCount == 0u`.
- The same test now has a positive-control phase:
  - calls `g_folderWindow.DebugSetPaneForceDrawItemTransientBrushCreateForSelfTest(..., true)`,
  - resets the counter,
  - warm-renders the pane,
  - requires `forcedSnapshot.drawItemTransientBrushCreateCount > 0u`,
  - cleanup disables the force seam and resets the counter.

Production/debug seam already present:

- `FolderView` now exposes
  `DebugSetForceDrawItemTransientBrushCreateForSelfTest(bool enabled) noexcept` under
  `ENABLE_TESTS` and stores the flag in
  `_debugForceDrawItemTransientBrushCreateForSelfTest`.
- `FolderWindow` now exposes
  `DebugSetPaneForceDrawItemTransientBrushCreateForSelfTest(Pane, bool) noexcept` and
  forwards to the pane's `FolderView`.
- `FolderView::DrawItem` has an `ENABLE_TESTS` local helper that, only when the force
  seam is enabled, creates a transient `ID2D1SolidColorBrush` through `_d2dContext` and
  increments `_debugDrawItemTransientBrushCreateCount` on success. It is called from
  the selected and hovered background branches. Normal production behavior still uses
  the cached member brushes.
- The forced brush is RAII-owned with `wil::com_ptr`; the seam is `ENABLE_TESTS` only.

### GR-24 Evidence Already Captured

RED build after adding the positive-control test before the debug seam existed:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64
```

- Failed as expected.
- Log: `.build\logs\msbuild-20260705_153748_205.log`.
- Failure:
  - `Commands.SelfTest.ViewCommands.cpp(23634,24): error C2039: 'DebugSetPaneForceDrawItemTransientBrushCreateForSelfTest': is not a member of 'FolderWindow'`
  - `Commands.SelfTest.ViewCommands.cpp(23700,20): error C2039: 'DebugSetPaneForceDrawItemTransientBrushCreateForSelfTest': is not a member of 'FolderWindow'`
- Meaning: the test was exercising a missing producer-control seam.

GREEN build after implementing the seam:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64
```

- Completed and produced `.build\x64\Debug\RedSalamander.exe`.
- Log: `.build\logs\msbuild-20260705_153949_588.log`.

Focused GREEN:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_draw_item_brush_reuse_guard -FailFast -TimeoutMultiplier 3
```

- Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_154145/`.
- Result: passed 1 / 0 / 0.
- `commands_results.json` confirms case `folderView_draw_item_brush_reuse_guard`
  passed in 1126 ms; `selftest.txt` has `exit_code: 0`.

### Immediate Resume Checklist

1. Run the adjacent guard that also asserts the transient-brush counter:

   ```powershell
   .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_perf_huge_folder_scale -FailFast -TimeoutMultiplier 3
   ```

2. If that is green, update durable docs/ledgers:

   - `Specs/UI/UI_FolderView.md`: document that the draw-item transient-brush counter
     is live, normal cached-brush renders must leave it at zero, and the forced
     `ENABLE_TESTS` positive control proves the counter can go non-zero.
   - `Specs/Testing/Testing_TestCoverage.md`: update the coverage text for
     `folderView_draw_item_brush_reuse_guard` / `folderView_perf_huge_folder_scale`
     with the RED compile evidence, GREEN build, focused archive, and adjacent archive.
   - `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`:
     mark GR-24 and GR-T24 done only after adjacent validation and docs are updated.
   - `Specs/Plans/WIP/Operation_Tailwind_FolderViewWarpDrivePreMergeReviewRemediation_2026-06-30.md`:
     mark TW-1 / TW-T1 done via Granite GR-24, since GR-24 restates Tailwind TW-1.
   - `Specs/Plans/WIP/README.md`: keep Granite pointed at GR-24 until it is closed;
     after closeout, inspect the Granite rows before choosing the next remaining
     item (likely GR-S4(f), but verify against the current ledger).
   - This baton: prepend a new save point after GR-24 is truly closed.

3. Run closeout checks after the docs/ledgers are updated:

   ```powershell
   rg -n "GR-24 \| LOW \||GR-T24 \| GR-24 \| Producer test|TW-1\\*? \\| \\*\\*MEDIUM\\*\\*|Continue low addendum work with \\*\\*GR-24|vacuous|counter has no producer" Specs\Plans\WIP\Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md Specs\Plans\WIP\Operation_Tailwind_FolderViewWarpDrivePreMergeReviewRemediation_2026-06-30.md Specs\Plans\WIP\README.md Specs\Testing\Testing_TestCoverage.md Specs\UI\UI_FolderView.md Tests\README.md
   git diff --check
   git status --short
   ```

4. If any selftest inventory counts are changed while documenting GR-24, also run:

   ```powershell
   Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
   ```

## Latest Save Point - 2026-07-05 15:34 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-22 / GR-T22 and GR-23 / GR-T23 are closed in the working tree**.
The WIP README now points Granite to **GR-24 / GR-T24** as the next item. GR-24 is LOW,
but it is the next remaining addendum row after the now-closed medium GR-23 slice.

Do not revert unrelated dirty files. `git status --short` still shows many modified
files from accumulated Granite/Clearwater/Tailwind/DxUi work. The files touched for the
GR-23 closeout slice are:

- `RedSalamander/IconCache.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- `Tools/Tests/TestInventory.Tests.ps1`
- `Tests/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/UI/UI_FolderView.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- This baton file

### GR-23 Closed State

GR-23 fix:

- `IconCache::QuerySysIconIndexForPath(..., useFileAttributes=false)` now treats live
  shell lookup failures as transient: it returns `std::nullopt`, emits
  `iconcache.path_live_lookup_failed_uncached`, and does not store a negative cache
  entry.
- Immediate retry/F5/ForceRefresh can re-enter `SHGetFileInfoW` and recover; a
  successful retry positive-caches the icon index.
- Attribute-mode `SHGFI_USEFILEATTRIBUTES` failures retain the bounded negative cache
  and keep using `iconcache.path_failed_lookup_cached` /
  `iconcache.path_failed_lookup_cache_hit`.
- `SelfTestLatency::Point::IconPathLiveLookup` now supports a forced failure hook for
  deterministic live lookup fault injection.
- `folderView_perf_slow_virtual_provider` now expects one repeated live shell lookup
  for the missing live path (`icons.repeated_failed_lookup_count == 1`) instead of the
  old negative-cache behavior.

GR-23 evidence:

- Test-only build before production code passed:
  `.build\logs\msbuild-20260705_151609_116.log`.
- Initial RED focused archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_151809/`, failed with
  `Forced transient live icon lookup failure should report no icon index.`
- Build after wiring the forced live-lookup failure hook passed:
  `.build\logs\msbuild-20260705_151850_525.log`.
- Second RED focused archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_152050/`, failed with
  `Second live icon lookup should re-query the shell and recover after a transient failure.`
- Final GREEN Debug x64 app build after production and adjacent test updates passed
  with 0 warnings / 0 errors:
  `.build\logs\msbuild-20260705_152904_817.log`.
- Pre-update adjacent slow-provider archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_152816/`, failed only because the
  guard still expected the removed live negative-cache behavior.
- GREEN adjacent slow-provider archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_153107/`, passed 1 / 0 / 0.
- GREEN focused archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_153136/`, passed 1 / 0 / 0.
- GREEN adjacent icon-pipeline archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_153209/`, passed 1 / 0 / 0.

GR-23 closeout verification:

- `Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru` passed 5 / 0 / 0
  after updating Commands static registrations from 681 to 682 and runner-listed cases
  from 783 to 784 in the docs.
- Stale sweep for old GR-23 counts, next-action text, and old live-failure-caching
  wording returned no matches:
  `rg -n "Commands: 783|783 Commands|783 runner-listed|681 static|Expected 681|GR-23 \| MED \||GR-T23 \| GR-23 \| ENABLE_TESTS hook forcing|Continue medium addendum work with \*\*GR-23|Failed live-path lookup caching" ...`
- `git diff --check` exited 0, with only LF-to-CRLF working-copy warnings.

Immediate next move:

1. Start **GR-24 / GR-T24** from
   `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`.
2. Follow TDD: first make the dead transient-brush regression guard fail in the intended
   way, or replace the vacuous guard with an equivalent source/behavior contract if the
   counter is truly obsolete.
3. After GR-24, update Tailwind TW-1 if it is the routed duplicate, then refresh WIP
   README and this baton again.

## Latest Save Point - 2026-07-05 15:26 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-22 / GR-T22 is closed in the working tree**. GR-23 / GR-T23 is
partially implemented and has focused GREEN evidence, but it is **not closed yet**
because adjacent validation, specs, inventory, WIP ledger, README, stale sweeps, and
diff checks are still pending.

Do not revert unrelated dirty files. `git status --short` currently shows many modified
files from accumulated Granite/Clearwater/Tailwind/DxUi work. The files that matter for
the current GR-23 slice are:

- `RedSalamander/IconCache.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`

Docs/specs still need to be updated for GR-23 before the row can be marked done.

### GR-22 Closed State

GR-22 closed after the previous 15:14 save point.

Final GR-22 verification:

- `Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru` passed 5 / 0 / 0
  after the Commands static inventory was updated from 680 to 681.
- The stale-doc sweep for old GR-22 counts/next-action text produced no output.
- `git diff --check` exited 0, with only LF-to-CRLF working-copy warnings.
- `Specs/Plans/WIP/README.md` now points Granite to **GR-23 / GR-T23**.

GR-22 code/test/spec updates already present:

- Added `folderView_refresh_to_paint_metric_clears_after_failed_render`.
- Added `FolderView::ClearReadyPendingRefreshToPaintMetric()`.
- Cleared ready refresh-to-paint metrics on failed/no-present render paths and validated
  generation before emitting.
- Updated `Specs/UI/UI_FolderView.md`, `Specs/Testing/Testing_TestCoverage.md`,
  `Tests/README.md`, `Tools/Tests/TestInventory.Tests.ps1`, the Granite WIP plan, and
  the WIP README.

Useful GR-22 evidence:

- Build after test-only change: `.build\logs\msbuild-20260705_150025_140.log`.
- RED focused: `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_150226/`, failed with
  `Stale refresh-to-paint metric emitted on an unrelated later Present; before=0 after=1.`
- Build after production fix: `.build\logs\msbuild-20260705_150338_519.log`.
- GREEN focused: `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_150541/`.
- Adjacent GREEN:
  - `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_150654/`
    (`folderView_rendering_error_overlay_requires_persistence`)
  - `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_150739/`
    (`folderView_render_device_loss_recovers`)
  - `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_151107/`
    (`folderView_perf_refresh_preservation`)
  - `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_151143/`
    (`folderView_perf_directory_change_storm`)

Do not cite `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_150620/` as product
evidence; it was an invalid broad filter with no matching expected cases.

### GR-23 Current Code/Test State

GR-23 bug:

- Live path icon lookup failures from `QuerySysIconIndexForPath(...,
  useFileAttributes=false)` were cached for `kPathIconFailureTtl`, so a transient shell
  or network failure could survive repeated F5/ForceRefresh attempts for the full TTL.

Test added:

- `folderView_iconcache_live_path_failure_retries_without_negative_cache` in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`.
- The test:
  - creates a temporary live file path,
  - forces the first `IconPathLiveLookup` failure with
    `SelfTestLatency::SetNextFailure(..., E_ACCESSDENIED)`,
  - requires the first live lookup to return `nullopt`,
  - immediately calls the same live lookup again and requires the shell to be re-entered
    and return a real icon index,
  - calls a third time and requires the positive cache to return the same index without
    consuming another live lookup hook.

Production changes currently present in `RedSalamander/IconCache.cpp`:

- The existing `SelfTestLatency::Point::IconPathLiveLookup` hook now supports forced
  failures through `SelfTestLatency::ConsumeFailure(...)`.
- Failed live-path lookups (`useFileAttributes=false`) now return `std::nullopt`
  without storing a negative cache entry.
- Failed uncached live lookups emit
  `iconcache.path_live_lookup_failed_uncached`.
- Attribute-mode lookups (`useFileAttributes=true`) still retain the bounded negative
  cache and continue to emit `iconcache.path_failed_lookup_cached` /
  `iconcache.path_failed_lookup_cache_hit`.

### GR-23 Evidence Already Captured

Build after test-only change:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64
```

- Passed, 0 warnings / 0 errors.
- Log: `.build\logs\msbuild-20260705_151609_116.log`.

First focused RED:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_iconcache_live_path_failure_retries_without_negative_cache -FailFast -TimeoutMultiplier 3
```

- Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_151809/`.
- Failed as expected with:
  `Forced transient live icon lookup failure should report no icon index.`
- Meaning: the new failure hook was not wired yet.

Build after wiring the forced live-lookup failure hook:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64
```

- Passed, 0 warnings / 0 errors.
- Log: `.build\logs\msbuild-20260705_151850_525.log`.

Second focused RED:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_iconcache_live_path_failure_retries_without_negative_cache -FailFast -TimeoutMultiplier 3
```

- Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_152050/`.
- Failed as expected with:
  `Second live icon lookup should re-query the shell and recover after a transient failure.`
- Meaning: this is the actual negative-cache bug.

Build after the live-failure no-cache fix:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64
```

- Passed, 0 warnings / 0 errors.
- Log: `.build\logs\msbuild-20260705_152134_020.log`.

Focused GREEN:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_iconcache_live_path_failure_retries_without_negative_cache -FailFast -TimeoutMultiplier 3
```

- Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_152333/`.
- Result: 1 passed, 0 failed, 0 skipped.

Adjacent icon-pipeline GREEN:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_perf_icon_pipeline_cold_slow -FailFast -TimeoutMultiplier 3
```

- Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_152407/`.
- Result: 1 passed, 0 failed, 0 skipped.

### Immediate Resume Steps

1. Run the adjacent slow-provider case because it contains an older assertion about
   repeated failed live-path lookups hitting the bounded failure cache:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_perf_slow_virtual_provider -FailFast -TimeoutMultiplier 3 2>&1 | Tee-Object -FilePath '.build\logs\gr23_adjacent_slow_virtual_provider.out.txt'
```

If it fails on `icons.repeated_failed_lookup_count == 0`, update the test expectation to
prove live failures are retried instead of negative-cached, then rerun focused GR-23 and
the adjacent cases.

2. Update durable docs/inventory for the new Commands case:

- `Tools/Tests/TestInventory.Tests.ps1`: Commands static registrations `681 -> 682`.
- `Tests/README.md`: Commands runner-listed cases `783 -> 784`, static registrations
  `681 -> 682`, ViewCommands family `108 -> 109`, and mention live icon lookup recovery.
- `Specs/Testing/Testing_TestCoverage.md`: same count updates plus a row for
  `folderView_iconcache_live_path_failure_retries_without_negative_cache` and a recent
  GR-23 focused coverage checkpoint.
- `Specs/UI/UI_FolderView.md`: replace the old live-path negative-cache contract with
  this contract:
  - live-path lookup failures are transient and must not be stored as negative cache
    entries,
  - immediate retry/F5/ForceRefresh may re-enter the shell and recover,
  - attribute-mode computed icon failures may use the bounded negative cache,
  - telemetry uses `iconcache.path_live_lookup_failed_uncached` for uncached live
    failures, `iconcache.path_failed_lookup_cached` for attribute negative-cache stores,
    and `iconcache.path_failed_lookup_cache_hit` for negative-cache hits.
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`:
  mark GR-23 and GR-T23 done only after docs/verification pass.
- `Specs/Plans/WIP/README.md`: advance Granite next action to **GR-24 / GR-T24** only
  after GR-23 is closed.
- This baton file: add a newer save point or closeout note.

Expected count drift before docs are updated:

- Commands runner-listed cases: `783 -> 784`.
- Commands static `SelfTest::RunCase` registrations: `681 -> 682`.
- ViewCommands family cases: `108 -> 109`.

3. Run closeout checks:

```powershell
Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
rg -n "Commands: 783|783 Commands|783 runner-listed|681 static|Expected 681|GR-23 \| MED \||GR-T23 \| GR-23 \| ENABLE_TESTS hook forcing|Continue medium addendum work with \*\*GR-23|Failed live-path lookup caching" Specs\Testing\Testing_TestCoverage.md Tests\README.md Tools\Tests\TestInventory.Tests.ps1 Specs\Plans\WIP\Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md Specs\Plans\WIP\README.md Specs\UI\UI_FolderView.md
git diff --check
git status --short
```

Only after those pass should GR-23 be treated as closed and the next Granite item become
**GR-24 / GR-T24**.

## Latest Save Point - 2026-07-05 15:14 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-22 / GR-T22 closeout edits are now in the working tree**, including
code, the focused selftest, adjacent evidence, the FolderView durable spec, test
coverage inventory docs, WIP README next action, and the Granite WIP rows. Final
post-doc verification is still the immediate gate at this save point.

Fresh GR-22 evidence added after the previous save point:

- `folderView_perf_refresh_preservation` passed 1 / 0 / 0 in
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_151107/`.
- `folderView_perf_directory_change_storm` passed 1 / 0 / 0 in
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_151143/`.
- The pre-doc inventory check failed exactly as expected because the new Commands
  `SelfTest::RunCase` registration raised the static count from 680 to 681 while docs
  still said 680.

Files updated for GR-22 in this slice:

- `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- `RedSalamander/FolderView.h`
- `RedSalamander/FolderView.cpp`
- `RedSalamander/FolderView.Rendering.cpp`
- `Tools/Tests/TestInventory.Tests.ps1`
- `Tests/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/UI/UI_FolderView.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- This baton file

Immediate next commands:

```powershell
Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
git diff --check
git status --short
```

If those pass, GR-22 can be treated as closed and the next Granite item is **GR-23 /
GR-T23**.

## Latest Save Point - 2026-07-05 15:09 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-A2, GR-A3, GR-A6, GR-20, GR-21/TW-8, and the GR-22 code/test fix
are present in the working tree**. GR-22 is **not closed yet** because the remaining
adjacent refresh-preservation check, inventory/docs/spec updates, and stale/diff checks
still need to run. The Granite WIP index still points to GR-22 / GR-T22 until those
closeout steps are completed.

Do not revert unrelated dirty files. `git status --short` currently shows many modified
files from accumulated Granite/Clearwater/Tailwind WIP. The GR-22 files touched in this
slice are:

- `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- `RedSalamander/FolderView.h`
- `RedSalamander/FolderView.cpp`
- `RedSalamander/FolderView.Rendering.cpp`
- This baton file, `Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md`

### GR-22 Current Code/Test State

Focused test added:

- `folderView_refresh_to_paint_metric_clears_after_failed_render`
  in `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`.
- It uses new helper message pumps that can process enumeration completion while
  withholding the target view's `WM_PAINT`, then:
  - loads a two-file local folder fixture,
  - triggers a left-pane refresh,
  - waits until the refresh enumeration completes without painting the target view,
  - forces a synthetic non-device-loss `EndDraw` failure,
  - verifies no `folder.refresh.request_to_paint_us` row is emitted by the failed paint,
  - drives a later unrelated successful Present,
  - verifies no stale request-to-paint metric leaks into that later Present.

Production fix currently present:

- `FolderView::RecordPendingRefreshToPaintStart(...)` resets the pending metric before
  recording a new refresh slot.
- `FolderView::ClearReadyPendingRefreshToPaintMetric() noexcept` was added and only
  clears a pending refresh-to-paint metric whose result is already ready.
- `FolderView::EmitPendingRefreshToPaintMetricAfterPresent()` now drops the pending
  metric if its generation no longer matches the current enumeration generation before
  emitting.
- `FolderView::OnPaint()` clears a ready pending metric on the resource-fallback path.
- `FolderView::Render(...)` clears a ready pending metric on no-target, failed
  `EndDraw`, failed `Present1`, and failed legacy `Present` paths.

### GR-22 Evidence Already Captured

Build after test-only change:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64
```

- Passed, 0 warnings / 0 errors.
- Log: `.build\logs\msbuild-20260705_150025_140.log`.

Focused RED:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_refresh_to_paint_metric_clears_after_failed_render -FailFast -TimeoutMultiplier 3
```

- Failed as expected with:
  `Stale refresh-to-paint metric emitted on an unrelated later Present; before=0 after=1.`
- Archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_150226/`.

Build after production fix:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64
```

- Passed, 0 warnings / 0 errors.
- Log: `.build\logs\msbuild-20260705_150338_519.log`.

Focused GREEN:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_refresh_to_paint_metric_clears_after_failed_render -FailFast -TimeoutMultiplier 3
```

- Passed 1 / 0 / 0.
- Archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_150541/`.

Adjacent rendering checks already run:

- Invalid broad filter attempt:
  `folderView_render` failed before running any case because
  `selftest_result_coverage` found no expected case match.
  Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_150620/`.
  This is not a product/test failure; do not cite it as GR-22 evidence except as a
  discarded command.
- `folderView_rendering_error_overlay_requires_persistence` passed 1 / 0 / 0.
  Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_150654/`.
- `folderView_render_device_loss_recovers` passed 1 / 0 / 0.
  Archive: `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_150739/`.

The captured stdout files named in the previous in-memory note for GR-22 are not present
under `.build\logs` at this save point, but the archived `commands_results.json` files
above are present and contain the pass/fail outcomes.

### Immediate Next Commands

Run the remaining refresh-metric adjacency:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_perf_refresh_preservation -FailFast -TimeoutMultiplier 3 2>&1 | Tee-Object -FilePath '.build\logs\gr22_adjacent_refresh_preservation_green.out.txt'
```

Optional, if time permits, also run the related directory-change storm coverage because it
also consumes `folder.refresh.request_to_paint_us`:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_perf_directory_change_storm -FailFast -TimeoutMultiplier 3 2>&1 | Tee-Object -FilePath '.build\logs\gr22_adjacent_directory_change_storm_green.out.txt'
```

After adding the Commands case, inventory is expected to drift until docs are updated:

```powershell
Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
```

Expected count updates if no other test inventory changes land first:

- Commands runner-listed cases: `782 -> 783`.
- Commands static `SelfTest::RunCase` registrations: `680 -> 681`.

Update at least these files:

- `Tools/Tests/TestInventory.Tests.ps1`
- `Tests/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Specs/UI/UI_FolderView.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- This baton file

Suggested spec text for `Specs/UI/UI_FolderView.md`: refresh-to-paint metrics emit only
after a successful Present for the same current enumeration generation; a ready pending
metric must be discarded on render no-target/no-present, `EndDraw` failure, `Present`
failure, or OnPaint resource fallback so a later unrelated Present cannot emit stale
latency; starting a newer refresh resets any existing pending refresh-to-paint slot.

Then run:

```powershell
Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
git diff --check
git status --short
```

Only after those pass, mark GR-22 and GR-T22 done in the Granite WIP plan, advance the
WIP README next action to GR-23 / GR-T23, and add a new latest save point here.

## Latest Save Point - 2026-07-05 14:55 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-A2, GR-A3, GR-A6, GR-20, and GR-21/TW-8 are closed in the working
tree**. The Granite WIP index points to **GR-22 / GR-T22** as the next medium addendum
item. GR-A4 remains open LOW cleanup and GR-A5 remains medium long-term; do not infer
either is closed from nearby work.

Do not revert unrelated dirty files. `git status --short` currently shows many modified
files from the accumulated Granite/Clearwater/Tailwind WIP, including but not limited to
DxUi, Search, MTP, FolderView, CompareDirectories, SettingsHotReload, specs, tests, and
this untracked baton file.

### Current Execution Plan

- Continue with **GR-22 / GR-T22**.
- Write the focused RED test first in `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`.
- Only after the expected stale refresh-to-paint metric failure is observed, patch
  `FolderView` pending metric ownership/clearing behavior.
- Then run focused GR-22 GREEN, adjacent relevant Commands coverage, inventory/stale
  hygiene, and `git diff --check`.

### GR-22 Starting Point

GR-22 bug:

- `PendingRefreshToPaintMetric` can survive past the paint it belongs to and later emit
  `folder.refresh.request_to_paint_us` against an unrelated later Present.
- `UpdatePendingRefreshToPaintResult(...)` marks the metric ready, but
  `EmitPendingRefreshToPaintMetricAfterPresent()` is only called from successful
  Present branches.
- Render early returns, EndDraw failures, and Present failures bypass the emit path and
  can leave the pending metric stale.

Required GR-T22 proof:

- Drive one refresh so `_pendingRefreshToPaintMetric.resultReady == true`.
- Force a render early-return or EndDraw/Present failure path.
- Assert no `folder.refresh.request_to_paint_us` fires for the failed render and that
  the pending metric is cleared.
- Trigger an unrelated later Present with no new refresh and assert no stale metric
  fires.

Useful search already run:

```powershell
rg -n "PendingRefreshToPaint|UpdatePendingRefreshToPaintResult|EmitPendingRefreshToPaintMetricAfterPresent|CancelPendingRefreshToPaint|RecordPendingRefreshToPaintStart|DebugConsumeNextRenderFailure|Debug.*Render|EndDraw|Present|request_to_paint_us|pendingRefresh" RedSalamander\FolderView.cpp RedSalamander\FolderView.Rendering.cpp RedSalamander\FolderView.Enumeration.cpp RedSalamander\FolderView.h RedSalamander\SelfTest\Commands -S
```

Anchors found:

- `RedSalamander/FolderView.h`
  - `PendingRefreshToPaintMetric` around line 984.
  - `_pendingRefreshToPaintMetric` around line 992.
  - metric helper declarations around lines 1445-1449:
    `RecordPendingRefreshToPaintStart`, `UpdatePendingRefreshToPaintResult`,
    `CancelPendingRefreshToPaint`, `EmitPendingRefreshToPaintMetricAfterPresent`.
  - debug render failure seam around lines 397-400 and 1498-1501:
    `DebugRenderFailurePoint`, `_debugNextRenderFailure`,
    `DebugConsumeNextRenderFailure`.
- `RedSalamander/FolderView.cpp`
  - `RecordPendingRefreshToPaintStart(...)` around lines 124-132.
  - `UpdatePendingRefreshToPaintResult(...)` around lines 135-143.
  - `CancelPendingRefreshToPaint(...)` around lines 146-150.
  - `EmitPendingRefreshToPaintMetricAfterPresent()` around lines 164-174.
  - debug warm/render failure helpers around lines 1029, 1097, 1115, and 1123.
- `RedSalamander/FolderView.Enumeration.cpp`
  - `ForceRefresh` starts the pending metric around line 1121.
  - stale/mismatched results cancel around lines 1493 and 1508.
  - refresh success calls `UpdatePendingRefreshToPaintResult(...)` around line 1695.
  - non-refresh enumeration clears around line 1699.
- `RedSalamander/FolderView.Rendering.cpp`
  - forced EndDraw failure can return before pending metric emit around line 1928.
  - successful `Present1` branch emits around lines 1989-1990.
  - successful legacy `Present` branch emits around lines 2028-2029.
  - failed Present branches return before emit.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
  - `FolderViewPerfMetricCountInRun(L"folder.refresh.request_to_paint_us", ...)`
    around line 21661.
  - existing request-to-paint metric assertion around line 21668.
  - render failure/device-loss selftests around lines 23090-23367.

Suggested next reads:

```powershell
$lines=Get-Content 'RedSalamander\FolderView.cpp'; foreach($i in 112..180){ '{0,5}: {1}' -f $i,$lines[$i-1] }
$lines=Get-Content 'RedSalamander\FolderView.Rendering.cpp'; $ranges=@(); $ranges += 960..1030; $ranges += 1900..2035; foreach($i in $ranges){ '{0,5}: {1}' -f $i,$lines[$i-1] }
$lines=Get-Content 'RedSalamander\FolderView.Enumeration.cpp'; $ranges=@(); $ranges += 1100..1130; $ranges += 1470..1515; $ranges += 1680..1705; foreach($i in $ranges){ '{0,5}: {1}' -f $i,$lines[$i-1] }
$lines=Get-Content 'RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp'; $ranges=@(); $ranges += 21520..21680; $ranges += 23080..23380; foreach($i in $ranges){ '{0,5}: {1}' -f $i,$lines[$i-1] }
```

Likely RED test shape:

- Reuse the existing ViewCommands refresh/perf helpers near the request-to-paint metric
  assertions.
- Capture the current metric scan offset/count for
  `folder.refresh.request_to_paint_us`.
- Trigger a left-pane refresh and wait until enumeration has produced a ready result.
- Force a non-device-loss render failure, probably via
  `DebugForcePaneNextRenderFailure(FolderWindow::Pane::Left,
  FolderView::DebugRenderFailurePoint::EndDraw, E_FAIL)`.
- Drive the render path with the existing warm-render helper.
- Require that the request-to-paint metric count has not increased.
- Drive one more ordinary render/Present with no new refresh and require that the count
  still has not increased. Current buggy behavior is expected to fail here by emitting
  the stale metric on the later successful Present.

Likely production direction after RED:

- Add a small private helper that clears `_pendingRefreshToPaintMetric` unconditionally,
  or adapt the existing cancel helper if the current generation is available.
- Clear a ready pending metric on render early returns and EndDraw/Present failure paths
  that occur after a refresh result is ready.
- Keep successful Present behavior emitting exactly once.
- Preserve the current non-refresh/stale-generation cancellation behavior unless the
  RED test proves it must change.

## Latest Save Point - 2026-07-05 14:53 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-A2, GR-A3, GR-A6, GR-20, and GR-21/TW-8 are closed in the working
tree**. The Granite WIP index now points to **GR-22 / GR-T22** as the next medium
addendum item. GR-A4 remains open LOW cleanup and GR-A5 remains medium long-term; do
not infer either is closed from nearby work.

### GR-21 Closed State

GR-21 fixes Compare Directories "Invert Differences Selection" so it is a one-shot
command again instead of a sticky per-session selection mode.

Code/test changes made:

- `RedSalamander/CompareDirectoriesWindow.cpp`
  - `IDM_COMPARE_INVERT_DIFFERENCES_SELECTION` still applies the inverted decision
    predicate to the current left/right panes, then immediately resets
    `_selectionMode` to `CompareSelectionMode::Default`.
  - Compare-window `IDM_LEFT_REFRESH` / `IDM_RIGHT_REFRESH` re-apply the default
    compare decision selection before forcing `FolderWindow::CommandRefresh(...)`, so
    `FolderView` refresh-preservation cannot carry the inverted selection forward.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp`
  - Extended `cmd_compare_directories_non_file_plugin_path_form_selection_and_empty_state`.
  - The test now proves default left-only compare selection, invokes
    `IDM_COMPARE_INVERT_DIFFERENCES_SELECTION`, observes the inverted selection, sends
    `IDM_LEFT_REFRESH`, and requires the default compare selection to return.
- `Specs/Core/Core_CompareDirectories.md`
  - Documents that Invert Differences Selection is one-shot and that pane refresh,
    navigation, and enumeration callbacks use default decision-model selection unless
    the user invokes Invert again.
- `Specs/Testing/Testing_TestCoverage.md`
  - Adds the GR-21 focused coverage checkpoint and expands the existing case row.
- `Specs/Plans/WIP/Operation_Tailwind_FolderViewWarpDrivePreMergeReviewRemediation_2026-06-30.md`
  - Marks duplicate TW-8 done via Granite GR-21.
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
  - GR-21 and GR-T21 are marked done with a closeout note.
- `Specs/Plans/WIP/README.md`
  - Granite next action advanced to GR-22 / GR-T22.
  - Tailwind next action advanced past TW-8 to TW-1.

GR-21 evidence:

- RED focused:
  `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_compare_directories_non_file_plugin_path_form_selection_and_empty_state -FailFast -TimeoutMultiplier 3`
  failed with `leftSelected=0` after invert + left refresh.
  Archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_143810/`.
  Captured summary:
  `.build\logs\gr21_red.out.txt`.
- GREEN app build:
  `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors.
  Log:
  `.build\logs\msbuild-20260705_144528_275.log`.
- GREEN focused:
  `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_compare_directories_non_file_plugin_path_form_selection_and_empty_state -FailFast -TimeoutMultiplier 3`
  passed 1 / 0 / 0.
  Archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_144738/`.
  Captured summary:
  `.build\logs\gr21_focused_green.out.txt`.
- GREEN adjacent:
  `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_compare_directories_ -FailFast -TimeoutMultiplier 3`
  passed 15 / 0 / 0.
  Archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_145147/`.
  Captured summary:
  `.build\logs\gr21_adjacent_compare_directories_green.out.txt`.

### Immediate Resume After GR-21 Closeout

The next Granite item should be **GR-22 / GR-T22**.

GR-22 bug:

- `PendingRefreshToPaintMetric` can survive past the paint it belongs to and later emit
  `folder.refresh.request_to_paint_us` against an unrelated later Present.
- The leak paths are `Render` early returns / failures after
  `UpdatePendingRefreshToPaintResult(...)` marks the metric ready, plus stale generation
  emit/overwrite behavior.

Required GR-T22 proof:

- Drive one refresh so `_pendingRefreshToPaintMetric.resultReady == true`.
- Force a `Render` early-return or present/end-draw failure path.
- Assert no `folder.refresh.request_to_paint_us` fires and the pending metric is
  cleared.
- Trigger an unrelated later Present with no new refresh and assert no stale metric
  fires.

Likely anchors:

```powershell
rg -n "PendingRefreshToPaint|UpdatePendingRefreshToPaintResult|EmitPendingRefreshToPaintMetricAfterPresent|CancelPendingRefreshToPaint|RecordPendingRefreshToPaintStart|DebugConsumeNextRenderFailure|EndDraw|Present" RedSalamander\FolderView.cpp RedSalamander\FolderView.Rendering.cpp RedSalamander\FolderView.Enumeration.cpp RedSalamander\FolderView.h RedSalamander\SelfTest\Commands
```

Use RED-first. Do not touch production render/metric code until a focused failing test
exists and has failed for the expected stale-metric reason.

Do not revert unrelated dirty files; the working tree intentionally includes multiple
closed Granite/Clearwater/Tailwind slices plus this baton.

## Latest Save Point - 2026-07-05 14:31 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete, and do not move the
Granite plan to `Specs/Plans/Done/` yet.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-A2, GR-A3, GR-A6, and GR-20 are closed in the working tree**. The
Granite WIP index now points to **GR-21 / GR-T21** as the next medium addendum item.
GR-A4 remains open LOW cleanup and GR-A5 remains medium long-term; do not infer either
is closed from nearby work.

### GR-20 Closed State

GR-20 fixes settings hot-reload startup so a transient initial directory-watch arming
failure does not block the UI thread for roughly 2000 ms, does not return
`WAIT_TIMEOUT`, and does not tear down the watcher before the worker retry loop can
self-arm.

Code/test changes made:

- `RedSalamander/SettingsHotReload.h`
  - Added `ENABLE_TESTS`-only
    `SettingsHotReload::DebugSetChangeNotificationOpenFailuresForSelfTest(...)`.
- `RedSalamander/SettingsHotReload.cpp`
  - Added atomic `ENABLE_TESTS` counters for deterministic forced
    `FindFirstChangeNotificationW(...)` open failures.
  - `WatchSettingsDirectoryThread(...)` now consumes the forced failure seam before
    calling `FindFirstChangeNotificationW(...)`; normal retry behavior remains in the
    worker loop.
  - `SettingsHotReload::Start(...)` still creates the stop/ready events and launches the
    `std::jthread`, but no longer waits up to 2000 ms on `readyHandle`.
  - `Start(...)` now returns `S_OK` after launching the watcher, leaving transient
    readiness failures to the asynchronous worker retry loop.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`
  - Added `settings_hot_reload_transient_arm_failure_is_async`.
  - The test forces one transient watcher-arm failure, times `SettingsHotReload::Start`,
    asserts elapsed time is under 200 ms, asserts the result is not `WAIT_TIMEOUT`,
    then proves a later settings write posts the hot-reload message after the worker
    self-arms.
- `Specs/Core/Core_SettingsStore.md`
  - Documents that `SettingsHotReload::Start(...)` must not synchronously wait on
    directory-watch readiness or fail with `WAIT_TIMEOUT` for transient initial
    `FindFirstChangeNotificationW(...)` arming failures.
  - Documents that watcher readiness is worker-internal state and that the worker
    retries asynchronously until it can self-arm.
- `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`, and
  `Tools/Tests/TestInventory.Tests.ps1`
  - Commands inventory updated to 782 runner-listed cases and 680 static
    `SelfTest::RunCase` registrations.
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
  - GR-20 and GR-T20 are marked done with a GR-20 closeout note.
- `Specs/Plans/WIP/README.md`
  - Granite next action advanced to GR-21 / GR-T21.

### GR-20 Evidence Already Collected

- RED compile seam:
  `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  failed before the test seam existed with missing
  `SettingsHotReload::DebugSetChangeNotificationOpenFailuresForSelfTest(...)`.
  Log:
  `.build\logs\msbuild-20260705_141406_457.log`.
- RED behavioral proof:
  `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter settings_hot_reload_transient_arm_failure_is_async -FailFast -TimeoutMultiplier 3`
  failed before the production fix because `Start(...)` blocked/timed out on the forced
  transient watcher-arm failure.
  Archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_141831/`.
- GREEN build after the fix:
  `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors.
  Log:
  `.build\logs\msbuild-20260705_142433_667.log`.
- GREEN focused GR-20:
  `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter settings_hot_reload_transient_arm_failure_is_async -FailFast -TimeoutMultiplier 3`
  passed 1 / 0 / 0.
  Archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_142640/`.
  Captured summary:
  `.build\logs\gr20_focused_green.out.txt`.
- GREEN adjacent hot-reload family:
  `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter settings_hot_reload_ -FailFast -TimeoutMultiplier 3`
  passed 5 / 0 / 0.
  Archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_142715/`.
  Captured summary:
  `.build\logs\gr20_adjacent_settings_hot_reload_green.out.txt`.
- Note: an earlier adjacent run at
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_142318/` failed because the new test
  overasserted `TryLoadChangedSettings()` after message delivery. The test was refined to
  prove only the required GR-T20 contract: async start, no `WAIT_TIMEOUT`, watcher
  self-arm, and later message post.

GR-20 closeout verification:

```powershell
Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
git diff --check
```

Results:

- Pester passed 5 / 0.
- `git diff --check` exited 0 with only LF-to-CRLF warnings from the dirty tree.

### Immediate Resume After GR-20 Closeout

The next Granite item should be **GR-21 / GR-T21**.

GR-21 bug:

- Compare Directories "Invert selection" currently appears sticky because
  `_selectionMode = Inverted` re-applies on every pane refresh/navigation.
- It should behave like a one-shot command: invert the current compare selection, then a
  pane re-enumeration or navigation should restore the default compare selection.

Required GR-T21 proof:

- Start an active compare with default differences selected.
- Invoke `IDM_COMPARE_INVERT_DIFFERENCES_SELECTION`.
- Assert the current compare selection is inverted.
- Trigger pane re-enumeration, such as F5 or navigating away and back.
- Assert selection restored to the default compare selection rather than staying
  inverted.

Likely anchors:

```powershell
rg -n "_selectionMode|ApplySelectionForFolder|IDM_COMPARE_INVERT_DIFFERENCES_SELECTION|PrepareCompareRun|CancelCompareMode" RedSalamander
rg -n "compare.*invert|Invert.*selection|selectionMode|CompareDirectories" RedSalamander\SelfTest Tests Specs
```

Do not revert unrelated dirty files; the working tree intentionally includes multiple
closed Granite/Clearwater/Tailwind slices plus this baton.

## Latest Save Point - 2026-07-05 14:12 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-A2, GR-A3, and GR-A6 are closed in the working tree**. The WIP
index now points to **GR-20 / GR-T20** as the next medium addendum item. GR-20
reconnaissance is started, but no GR-20 files have been edited yet. GR-A4 remains open
LOW cleanup and GR-A5 remains medium long-term; do not infer either is closed from the
GR-A3/GR-A6 work.

### GR-A6 Closed State

GR-A6 hoisted shared DxUi/FolderView Direct2D/DXGI device-loss classification and
D3D11 hardware-to-WARP creation fallback.

Code changes:

- `Common/DxUi/DxUi.h`
  - Declares `IsDeviceLossHResult(...)`.
  - Declares `CreateD3D11DeviceWithWarpFallback(...)`.
- `Common/DxUi/DxUi.WindowHost.cpp`
  - Implements `IsDeviceLossHResult(...)` with
    `D2DERR_RECREATE_TARGET`, `DXGI_ERROR_DEVICE_REMOVED`,
    `DXGI_ERROR_DEVICE_RESET`, and `DXGI_ERROR_DEVICE_HUNG`.
  - Implements `CreateD3D11DeviceWithWarpFallback(...)`.
  - WindowHost shared graphics creation uses the shared D3D helper.
  - Both WindowHost render overloads classify `hrDraw` and `hrPresent` with the
    shared predicate.
- `RedSalamander/FolderView.Rendering.cpp`
  - Removed private `IsFolderViewDeviceLoss`.
  - Device creation uses `CreateD3D11DeviceWithWarpFallback(...)`, including the
    existing `REDSALAMANDER_FOLDERVIEW_FORCE_WARP` selftest path.
  - EndDraw/Present recovery uses `RedSalamander::DxUi::IsDeviceLossHResult(...)`.
- `Tests/DxUiTests/DxUiTests.WindowHost.cpp`
  - Added `TestDxUiDeviceLossAndD3dCreationAreShared`.
  - The test feeds `DXGI_ERROR_DEVICE_HUNG` through the shared predicate, checks
    non-device-loss cases, and source-guards WindowHost/FolderView routing through the
    shared predicate/helper.

Durable docs and ledgers updated:

- `Specs/UI/UI_DxUiWinUIDesign.md`
  - Documents DxUi-owned shared device-loss classification and D3D fallback helper.
- `Specs/UI/UI_FolderView.md`
  - Documents FolderView must use the DxUi shared predicate/helper and must not carry
    private copies.
- `Specs/Testing/Testing_TestCoverage.md`
  - Adds the GR-A6 focused coverage checkpoint.
- `Tests/README.md`
  - DxUiTests total updated 636 -> 637.
  - WindowHost family updated 51 -> 52.
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
  - GR-A6 marked `MED - DONE 2026-07-05`.
  - Added GR-A6 closeout note.
- `Specs/Plans/WIP/README.md`
  - Granite next action advanced to GR-20 / GR-T20.

GR-A6 evidence:

- RED:
  `.\build.ps1 -ProjectName DxUiTests -Configuration Debug -Platform x64`
  failed before the shared predicate existed with
  `error C3861: 'IsDeviceLossHResult': identifier not found`.
  Log:
  `.build\logs\msbuild-20260705_140035_875.log`.
- GREEN DxUiTests build:
  `.\build.ps1 -ProjectName DxUiTests -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors.
  Log:
  `.build\logs\msbuild-20260705_140302_574.log`.
- GREEN focused DxUi:
  `.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost`
  exited 0 and ran `TestDxUiDeviceLossAndD3dCreationAreShared`.
- GREEN app build:
  `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors.
  Log:
  `.build\logs\msbuild-20260705_140513_048.log`.
- GREEN focused FolderView:
  `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_render_device_loss_recovers -FailFast -TimeoutMultiplier 3`
  passed 1 / 0 / 0.
  Archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_140722/`.
- Source guard:

```powershell
rg -n "IsFolderViewDeviceLoss|D3D11CreateDevice\(" Common\DxUi\DxUi.WindowHost.cpp RedSalamander\FolderView.Rendering.cpp
```

  returned only the shared helper's `D3D11CreateDevice(...)` call in
  `Common\DxUi\DxUi.WindowHost.cpp`.
- `git diff --check` exited 0 with only LF-to-CRLF warnings.
- `Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru` passed 5 / 0.

### GR-20 Reconnaissance Started

No GR-20 source/test/docs edits have been made yet. The next agent can continue from
these already-inspected anchors without relying on chat history:

- `RedSalamander/SettingsHotReload.h`
- `RedSalamander/SettingsHotReload.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`

Live code findings:

- `WatchSettingsDirectoryThread(...)` retries `FindFirstChangeNotificationW(...)` after
  failure by waiting up to 1000 ms on the stop event, then looping.
- It signals `readyEvent` only after `FindFirstChangeNotificationW(...)` succeeds.
- `SettingsHotReload::Start(...)` currently creates `stopEvent` and `readyEvent`,
  launches the `std::jthread`, then blocks on `WaitForSingleObject(readyHandle, 2000)`.
- If that wait times out, `Start(...)` calls `Stop()` and returns
  `HRESULT_FROM_WIN32(WAIT_TIMEOUT)`, which tears down the watcher instead of letting it
  self-arm on the next retry.
- Existing Commands selftest helpers already provide a hot-reload test window,
  `WndMsg::kSettingsFileChanged` payload handling via `TakeMessagePayload`, and
  `WaitForAtomicAtLeast(...)`-style polling.

Likely RED-first path for GR-T20:

1. Add a focused Commands selftest, probably in
   `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`, that injects one
   transient `FindFirstChangeNotificationW(...)` arming failure, calls
   `SettingsHotReload::Start(...)`, and asserts:
   - elapsed time is under 200 ms,
   - the HRESULT is not `WAIT_TIMEOUT`,
   - a later settings-file change still posts `WndMsg::kSettingsFileChanged`.
2. If no existing seam is found, make the first RED a compile/source RED by calling a
   missing `_DEBUG`/`ENABLE_TESTS`-only seam such as
   `SettingsHotReload::DebugSetChangeNotificationOpenFailuresForSelfTest(...)`.
3. Implement only that deterministic test seam, then run the focused test again before
   changing `Start(...)`; the expected behavioral RED is the current ~2000 ms wait plus
   `WAIT_TIMEOUT`/watcher teardown.
4. Fix `Start(...)` so watcher arming is asynchronous and transient readiness timeout is
   non-fatal. The likely implementation is to stop waiting on `readyEvent` in `Start(...)`
   and return `S_OK` after launching the watcher, leaving the worker retry loop alive.

After green, update the durable hot-reload/settings contract in
`Specs/Core/Core_SettingsStore.md` if that remains the authoritative settings spec, then
close GR-20/GR-T20 in the Granite WIP ledger and refresh this baton.

### Immediate Resume Steps

1. Start GR-20 / GR-T20 from the Granite addendum:
   - Bug: `SettingsHotReload::Start()` blocks the UI thread up to 2000 ms waiting for
     watcher readiness and disables hot reload if the settings directory is initially
     slow/locked/unavailable.
   - Required RED: point settings hot-reload at a path whose
     `FindFirstChangeNotificationW` transiently fails for one arming cycle; assert
     `Start()` returns quickly, does not hard-fail/tear down the watcher, and later arms
     when the directory becomes reachable.
2. Inspect live anchors before editing:

```powershell
rg -n "class SettingsHotReload|SettingsHotReload::Start|FindFirstChangeNotificationW|readyEvent|WaitForSingleObject|OnMainWindowCreate|WM_CREATE|settings.*changed|hot.reload" -S .
```

3. Use RED-first. Do not touch production hot-reload code until a focused failing test
   exists and has failed for the expected reason.
4. Do not revert unrelated dirty files; the worktree intentionally includes multiple
   closed Granite/Clearwater/Tailwind slices plus this baton.

## Latest Save Point - 2026-07-05 13:58 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-A2 and GR-A3 are closed in the working tree**. The next active
Granite item is **GR-A6** (Track A, medium): hoist shared DxUi/FolderView device-loss
classification and hardware-to-WARP device creation helpers so WindowHost and FolderView
classify `DXGI_ERROR_DEVICE_HUNG` identically and stop carrying divergent D3D11 creation
logic. **GR-A4 remains open LOW cleanup**; do not mark it done just because GR-A3 removed
one adjacent `StartsWithNoCase` clone in `FileSystemMtp.Shared.cpp`.

### GR-A3 Closed State

GR-A3 removed the duplicated MTP identity/hash/suffix/sanitizer/JSON helper clones that
remained in the MTP plugin after GR-A2 eliminated host-side picker identity derivation.

Durable docs and ledgers updated:

- `Specs/FileSystem/FileSystem_Mtp.md`
  - Documents that device-root suffixes, persistent/session object suffixes,
    duplicate-name suffixes, connection-root hashes, overwrite-journal
    device-state directory hashes, path sanitization, and plugin JSON escaping must use
    the shared `FileSystemMtp.Shared.cpp` helpers declared in
    `FileSystemMtp.Internal.h`.
- `Specs/Testing/Testing_TestCoverage.md`
  - CompareDirectories count is now 249 listed cases.
  - Compare static `SelfTest::RunCase` count is now 241.
  - MTP/PTP family count is now 52 and includes
    `mtp_identity_helpers_are_shared`.
- `Tests/README.md`
  - Compare listed/static counts and MTP/PTP family count were updated.
- `Tools/Tests/TestInventory.Tests.ps1`
  - Compare static registration expectation is now 241.
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
  - GR-A3 is marked `MEDIUM - DONE 2026-07-05`.
  - Added a GR-A3 closeout note with RED/GREEN/adjacent evidence.
- `Specs/Plans/WIP/README.md`
  - Granite next action now points to GR-A6.

GR-A3 evidence:

- RED:
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_identity_helpers_are_shared -FailFast -TimeoutMultiplier 3`
  failed before shared identity declarations existed with
  `MTP identity source guard: shared stable-hash declaration is missing.`
  Archive:
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_134733/`.
- GREEN build:
  `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors.
  Log:
  `.build\logs\msbuild-20260705_135001_254.log`.
- GREEN focused:
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_identity_helpers_are_shared -FailFast -TimeoutMultiplier 3`
  passed 1 / 0 / 0.
  Archive:
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_135203/`.
- GREEN adjacent:
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -FailFast -TimeoutMultiplier 3`
  passed 51 / 0 / 1.
  Archive:
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_135647/`.
  The skipped case was the expected no-device `mtp_live_device_smoke`.
- Static clone cleanup guard returned no matches:

```powershell
rg -n "StableDeviceHash|std::uint64_t StableHash|SanitizePathComponent\(|std::string JsonEscape|DuplicateSuffix\(" Plugins\FileSystemMtp\FileSystemMtp.Core.cpp Plugins\FileSystemMtp\FileSystemMtp.Device.cpp Plugins\FileSystemMtp\FileSystemMtp.FakeBackend.cpp
```

- Stale inventory search returned no matches:

```powershell
rg -n "CompareDirectories: 248|248 runner-listed cases|240 static|MTP/PTP File System \(51 cases\)|\| MTP/PTP \| 51|248 Compare listed" Specs\Testing\Testing_TestCoverage.md Tests\README.md
```

- `git diff --check` exited 0 with only LF-to-CRLF warnings.
- `Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru` passed 5 / 0.

### Immediate Resume Steps

1. Start GR-A6 with RED-first coverage:
   - Row:
     `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
     addendum GR-A6 / GR-TA6.
   - Required proof:
     feed `DXGI_ERROR_DEVICE_HUNG` to the hoisted shared
     `DxUi::IsDeviceLossHResult` and assert true; assert WindowHost render paths and
     FolderView route through the same predicate so all classify `DEVICE_HUNG`
     identically.
2. Inspect current anchors before editing:

```powershell
rg -n "DXGI_ERROR_DEVICE_HUNG|DXGI_ERROR_DEVICE_REMOVED|DXGI_ERROR_DEVICE_RESET|D2DERR_RECREATE_TARGET|D3D11CreateDevice|IsFolderViewDeviceLoss|recoverFromDeviceLoss|EnsureDeviceResources|CreateDevice" Common\DxUi RedSalamander\FolderView.Rendering.cpp RedSalamander\SelfTest Tests
```

3. Read relevant local skill docs before code:
   - `.github/skills/cpp-build/SKILL.md`
   - `.github/skills/direct2d-rendering/SKILL.md`
   - `.github/skills/cpp-modern-style/SKILL.md` if the helper API shape is not obvious.
4. Do not revert unrelated dirty files. The current working tree intentionally contains
   multiple closed Granite/Clearwater/Tailwind slices plus this baton.

## Latest Save Point - 2026-07-05 13:54 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: **GR-A2 is closed in the working tree**. **GR-A3 code and focused
RED/GREEN proof are in the working tree**, but GR-A3 is not closed yet because durable
docs, inventory counts, ledger/index updates, adjacent verification, and final hygiene
still need to run.

### GR-A2 Closed State

GR-A2 moved Connection Manager MTP picker browse out of host WPD code and into the MTP
plugin factory contract. Durable docs and ledgers have already been updated:

- `Specs/FileSystem/FileSystem_Mtp.md`
  - Connection Manager uses `RedSalamanderBrowseConnectionTargets(...)` through the
    plugin factory contract keyed by `pluginId`.
  - The MTP plugin owns live WPD browse, fake-backend browse, result JSON, and metrics.
- `Specs/Testing/Testing_TestCoverage.md`
  - Connection Manager Commands coverage is now 30 cases.
  - `cmd_connection_manager_window_mtp_picker_populates_profile` documents plugin-backed
    MTP picker coverage.
- `Tests/README.md`
  - Connections coverage mentions the plugin-backed MTP picker.
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
  - GR-A2 is marked `MEDIUM - DONE 2026-07-05`.
- `Specs/Plans/WIP/README.md`
  - Granite next action was advanced to GR-A3.

GR-A2 evidence:

- Build:
  `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors.
- Build log:
  `.build\logs\msbuild-20260705_133324_146.log`.
- Focused Commands archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-05_133530/`
  passed 1 / 0 / 0 for
  `cmd_connection_manager_window_mtp_picker_populates_profile`.
- Adjacent MTP Compare archive:
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_133656/`
  passed 50 / 0 / 1; skipped `mtp_live_device_smoke` because no approved live device
  was configured.
- `PluginContractTests.exe` passed.
- `git diff --check` passed with only LF-to-CRLF warnings.
- `Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru` passed 5 / 0.
- Host cleanup guard returned no matches:

```powershell
rg -n "PortableDevice|IPortable|WPD_|MtpPickerComInitialization|EnumerateMtpPicker|g_mtpPickerFixture|DebugSetConnectionManagerMtpPickerFixture|DebugClearConnectionManagerMtpPickerFixture|DebugSetMtpPickerFixture|DebugClearMtpPickerFixture|StableMtpPickerHash|SanitizeMtpPathComponent|EqualsNoCaseLocal|MtpPickerStringProperty|MtpPickerGuidPropertyEquals" RedSalamander\ConnectionManagerWindow.cpp RedSalamander\ConnectionManagerWindow.h RedSalamander\SelfTest
```

### GR-A3 Current Slice

Goal: remove duplicated MTP identity/hash/suffix/JSON helper clones inside the plugin.
The host-side copies were already removed by GR-A2; GR-A3 is about the remaining plugin
copies in Core, live Device, and FakeBackend.

RED already captured:

- Added source-contract Compare case `mtp_identity_helpers_are_shared`.
- Command:
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_identity_helpers_are_shared -FailFast -TimeoutMultiplier 3`
- Archive:
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_134733/`
- Expected failure:
  `MTP identity source guard: shared stable-hash declaration is missing.`

GREEN already captured:

- Build:
  `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors.
- Build log:
  `.build\logs\msbuild-20260705_135001_254.log`.
- Focused command:
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_identity_helpers_are_shared -FailFast -TimeoutMultiplier 3`
- Archive:
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_135203/`
- Result:
  1 passed / 0 failed / 0 skipped.

GR-A3 code changes currently in tree:

- `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.cpp`
  - Added `mtp_identity_helpers_are_shared` to `kCompareCaseNames`.
  - Added small repo-root/source-read helpers for source-contract tests.
- `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp`
  - Added `SelfTest::RunCase` for `mtp_identity_helpers_are_shared`.
  - The guard requires shared declarations/implementations in
    `FileSystemMtp.Internal.h` and `FileSystemMtp.Shared.cpp`.
  - The guard rejects old local clones in `FileSystemMtp.Core.cpp`,
    `FileSystemMtp.Device.cpp`, and `FileSystemMtp.FakeBackend.cpp`.
- `Plugins/FileSystemMtp/FileSystemMtp.Internal.h`
  - Added shared declarations:
    `StableMtpIdentityHash`, `FormatMtpIdentityHash`, `SanitizeMtpPathComponent`,
    `MtpDeviceIdentitySuffix`, `MtpPersistentObjectIdentitySuffix`,
    `MtpObjectIdentitySuffix`, and `MtpDuplicateObjectSuffix`.
- `Plugins/FileSystemMtp/FileSystemMtp.Shared.cpp`
  - Added shared implementations for the hash, formatting, path sanitization, identity
    suffix, duplicate suffix, and JSON-escape helpers.
  - Removed the local `StartsWithNoCase` clone in `NormalizeMtpPath` by using
    `OrdinalString::StartsWithNoCase`.
- `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp`
  - Removed local `StableDeviceHash`, `SanitizePathComponent`, and `JsonEscape`.
  - Device-root, journal path, and overwrite-journal JSON code now use shared helpers.
- `Plugins/FileSystemMtp/FileSystemMtp.Device.cpp`
  - Removed local `StableHash`, `SanitizePathComponent`, `JsonEscape`, and
    `DuplicateSuffix`.
  - Device display names, object property JSON, duplicate disambiguation, and hash
    matching now use shared helpers.
- `Plugins/FileSystemMtp/FileSystemMtp.FakeBackend.cpp`
  - Removed local `StableHash`, `JsonEscape`, and `DuplicateSuffix`.
  - Fake duplicate paths and object-property JSON now use shared helpers.
- `Tools/Tests/TestInventory.Tests.ps1`
  - Compare static `SelfTest::RunCase` count already changed from 240 to 241.

Important caveat: GR-A4 is still not done. GR-A3 removed one adjacent clone
(`StartsWithNoCase`) from `FileSystemMtp.Shared.cpp`, but `FileSystemMtp.Device.cpp`
still has local case-folding/equality helpers and `Common/SearchServiceBroker.cpp` still
has its own `StartsWithNoCase`-style helper. Do not mark GR-A4 done unless those are
actually addressed and verified.

### Immediate Resume Steps

1. Finish GR-A3 count/docs updates:
   - `Specs/Testing/Testing_TestCoverage.md`
     - CompareDirectories total: 248 -> 249.
     - Compare static registrations: 240 -> 241.
     - MTP/PTP family count: 51 -> 52.
     - Add row for `mtp_identity_helpers_are_shared`.
   - `Tests/README.md`
     - Compare listed cases: 248 -> 249.
     - Compare static fallback count: 240 -> 241.
     - MTP/PTP family count: 51 -> 52.
   - `Specs/FileSystem/FileSystem_Mtp.md`
     - Document that plugin identity suffixes, duplicate suffixes, overwrite-journal
       identity hashes, path sanitization, and JSON escaping are shared helpers in
       `FileSystemMtp.Shared.cpp`; Core, live WPD Device, FakeBackend, and factory JSON
       paths must not carry local clones.
2. Update execution ledgers:
   - Mark GR-A3 done in
     `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
     only after adjacent verification passes.
   - Update `Specs/Plans/WIP/README.md` to the next high-to-low Granite item. Likely
     next Track A item is GR-A6, while GR-A4 remains a LOW cleanup unless deliberately
     pulled forward.
   - Update this baton again after GR-A3 closeout.
3. Run source-guard cleanup search:

```powershell
rg -n "StableDeviceHash|std::uint64_t StableHash|SanitizePathComponent\(|std::string JsonEscape|DuplicateSuffix\(" Plugins\FileSystemMtp\FileSystemMtp.Core.cpp Plugins\FileSystemMtp\FileSystemMtp.Device.cpp Plugins\FileSystemMtp\FileSystemMtp.FakeBackend.cpp
```

Expected result: no matches.

4. Run adjacent and hygiene verification:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -FailFast -TimeoutMultiplier 3
git diff --check
Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
```

Expected adjacent MTP direction after the new case: 51 passed / 0 failed / 1 skipped,
where `mtp_live_device_smoke` skips without `REDSALAMANDER_SELFTEST_MTP_DEVICE`.

### Current Dirty Tree Snapshot

`git status --short` at this save point:

```text
 M Common/DxUi/DxUi.Accessibility.cpp
 M Common/DxUi/DxUi.Internal.h
 M Common/DxUi/DxUi.Menu.cpp
 M Common/DxUi/DxUi.WindowHost.cpp
 M Common/LocalSearchIndexCore.cpp
 M Common/LocalSearchIndexCore.h
 M Common/PlugInterfaces/Factory.h
 M Common/SearchServiceBroker.cpp
 M Common/SearchServiceBroker.h
 M Common/SqliteIndexStore.cpp
 M Common/SqliteIndexStore.h
 M Plugins/FileSystem/FileSystem.Internal.h
 M Plugins/FileSystem/FileSystem.Search.cpp
 M Plugins/FileSystem/FileSystem.cpp
 M Plugins/FileSystemMtp/Factory.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.Core.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.Device.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.FakeBackend.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.Internal.h
 M Plugins/FileSystemMtp/FileSystemMtp.Shared.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.h
 M RedSalamander/ConnectionManagerWindow.cpp
 M RedSalamander/ConnectionManagerWindow.h
 M RedSalamander/FileSystemPluginManager.cpp
 M RedSalamander/FileSystemPluginManager.h
 M RedSalamander/FolderView.Icons.cpp
 M RedSalamander/FolderView.h
 M RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp
 M RedSalamander/FolderWindow.h
 M RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp
 M RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.cpp
 M RedSalamanderSearchService/Main.cpp
 M Specs/Core/Core_Search.md
 M Specs/FileSystem/FileSystem_Mtp.md
 M Specs/Plans/WIP/Operation_Clearwater_TwoDayMasterReviewRemediation_2026-06-28.md
 M Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
 M Specs/Plans/WIP/Operation_Tailwind_FolderViewWarpDrivePreMergeReviewRemediation_2026-06-30.md
 M Specs/Plans/WIP/README.md
 M Specs/Testing/Testing_TestCoverage.md
 M Specs/UI/UI_DxUiSharedGrid.md
 M Specs/UI/UI_DxUiWinUIDesign.md
 M Specs/UI/UI_FolderView.md
 M Tests/DxUiTests/DxUiTests.Accessibility.cpp
 M Tests/DxUiTests/DxUiTests.Menu.cpp
 M Tests/README.md
 M Tools/Tests/TestInventory.Tests.ps1
?? Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md
```

Do not revert unrelated dirty files. This workspace intentionally contains multiple
completed Granite/Clearwater/Tailwind slices plus the in-progress GR-A3 slice.

## Latest Save Point - 2026-07-05 13:39 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: Granite **GR-A2** code implementation is in the working tree and the
focused/adjacent code verification is green. The remaining work is documentation,
ledger/index closeout, evidence archiving if desired, and final hygiene checks before
moving on to the next Granite row.

### GR-A2 Scope Implemented

Goal implemented:

- Remove the duplicate host-side WPD MTP picker stack from
  `RedSalamander/ConnectionManagerWindow.cpp`.
- Route Connection Manager MTP device/storage browse through an optional MTP plugin
  factory export keyed by `pluginId`.
- Keep live WPD browse and deterministic fake-backend browse inside
  `Plugins/FileSystemMtp`.
- Delete host MTP picker fixture globals and debug fixture APIs.

Important code changes now present:

- `Common/PlugInterfaces/Factory.h`
  - Added `FactoryConnectionBrowseKind`.
  - Added `FactoryConnectionBrowseRequest`.
  - Added `FactoryConnectionBrowseResult`.
  - Declared optional `RedSalamanderBrowseConnectionTargets(...)` export.
- `RedSalamander/FileSystemPluginManager.h/.cpp`
  - Added `ConnectionBrowseDevice` and `ConnectionBrowseStorage`.
  - Added `EnumerateConnectionBrowseDevices(...)` and
    `EnumerateConnectionBrowseStorages(...)`.
  - Added JSON validation/parsing helpers and transient plugin export dispatch.
- `Plugins/FileSystemMtp/FileSystemMtp.Internal.h`
  - Added MTP connection-browse DTOs and browse helper declarations.
- `Plugins/FileSystemMtp/FileSystemMtp.Shared.cpp`
  - Added shared `JsonEscapeUtf8(...)`.
- `Plugins/FileSystemMtp/FileSystemMtp.Device.cpp`
  - Added live WPD-backed browse helpers and fake-backend browse helpers.
  - Device `pnpId` remains the WPD PnP id; `devicePuid` prefers the WPD persistent
    unique id and falls back to the PnP id.
  - Storage browse enumerates device root children and emits root-like initial paths.
- `Plugins/FileSystemMtp/Factory.cpp`
  - Added `RedSalamanderBrowseConnectionTargets(...)`.
  - Added `_DEBUG` export
    `RedSalamanderMtpSetPickerFakeBackendForSelfTest(...)`.
  - Added perf scopes:
    `mtp.connection_browse.devices_us` and
    `mtp.connection_browse.storages_us`.
  - Added emitted values:
    `mtp.connection_browse.devices` and
    `mtp.connection_browse.storages`.
- `RedSalamander/ConnectionManagerWindow.cpp/.h`
  - Removed direct `PortableDevice`/WPD picker implementation from the host.
  - Worker context now carries `pluginId`.
  - Worker calls `FileSystemPluginManager::EnumerateConnectionBrowseDevices` and
    `EnumerateConnectionBrowseStorages`.
  - Removed host fixture globals and `DebugSetConnectionManagerMtpPickerFixture` /
    `DebugClearConnectionManagerMtpPickerFixture`.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`
  - Migrated `cmd_connection_manager_window_mtp_picker_populates_profile` to seed
    the MTP plugin fake picker backend instead of host fixtures.

Host-side cleanup search already passed:

```powershell
rg -n "PortableDevice|IPortable|WPD_|MtpPickerComInitialization|EnumerateMtpPicker|g_mtpPickerFixture|DebugSetConnectionManagerMtpPickerFixture|DebugClearConnectionManagerMtpPickerFixture|DebugSetMtpPickerFixture|DebugClearMtpPickerFixture|StableMtpPickerHash|SanitizeMtpPathComponent|EqualsNoCaseLocal|MtpPickerStringProperty|MtpPickerGuidPropertyEquals" RedSalamander\ConnectionManagerWindow.cpp RedSalamander\ConnectionManagerWindow.h RedSalamander\SelfTest
```

That command returned no matches.

### Verification Captured

RED before implementation:

- Build:
  `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors.
- Build log:
  `.build\logs\msbuild-20260705_132106_929.log`.
- Focused RED:
  `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_connection_manager_window_mtp_picker_populates_profile -FailFast -TimeoutMultiplier 3`.
- Expected failure:
  `Failed to seed the MTP plugin fake picker browse backend. hr=0x8007007F`
  (`ERROR_PROC_NOT_FOUND`).

GREEN after implementation:

- Build:
  `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors.
- Build log:
  `.build\logs\msbuild-20260705_133324_146.log`.
- Focused Commands:
  `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_connection_manager_window_mtp_picker_populates_profile -FailFast -TimeoutMultiplier 3`
  passed, 1 passed / 0 failed / 0 skipped.
- Plugin contract:
  `.\.build\x64\Debug\PluginContractTests.exe` passed and printed
  `PluginContractTests passed.`
- Adjacent MTP family:
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter mtp_ -FailFast -TimeoutMultiplier 3`
  passed, 50 passed / 0 failed / 1 skipped. The skipped case was
  `mtp_live_device_smoke` because `REDSALAMANDER_SELFTEST_MTP_DEVICE` was not set.

### Remaining Work To Resume

1. Optionally archive focused Commands evidence if no automatic archive exists. The
   Compare run likely overwrote `last_run`, so rerun the focused Commands case before
   copying artifacts if archived GR-A2 evidence is required.
2. Update durable docs:
   - `Specs/FileSystem/FileSystem_Mtp.md`: state that the Connection Manager MTP
     picker browses through the plugin factory/export and that WPD/fake enumeration
     stays inside the MTP plugin.
   - `Tests/README.md` and `Specs/Testing/Testing_TestCoverage.md`: update the
     existing MTP picker selftest description if it still says host fixture/WPD.
3. Update execution ledgers:
   - `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`:
     mark GR-A2 done with the verification above.
   - `Specs/Plans/WIP/README.md`: advance Granite to the next high-to-low row after
     GR-A2.
   - This baton: either keep this save point or add a closeout note after docs pass.
4. Run final hygiene:
   - `git diff --check`
   - `Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru`
   - final Debug x64 RedSalamander build after docs, if code changed again.
5. Do not revert unrelated dirty files. The broader workspace includes prior
   Granite/Clearwater/Tailwind work.

### GR-A2 Dirty Files

Treat these as part of the GR-A2 slice:

- `Common/PlugInterfaces/Factory.h`
- `RedSalamander/FileSystemPluginManager.h`
- `RedSalamander/FileSystemPluginManager.cpp`
- `Plugins/FileSystemMtp/FileSystemMtp.Internal.h`
- `Plugins/FileSystemMtp/FileSystemMtp.Shared.cpp`
- `Plugins/FileSystemMtp/FileSystemMtp.Device.cpp`
- `Plugins/FileSystemMtp/Factory.cpp`
- `RedSalamander/ConnectionManagerWindow.cpp`
- `RedSalamander/ConnectionManagerWindow.h`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`

Potential caveat to keep in mind: the selftest seeds fake picker JSON through a
debug export on the MTP plugin DLL. The focused test passed, so the current load path
works, but if the helper is refactored later ensure the setter reaches the same loaded
module instance that the picker browse export uses.

## Latest Save Point - 2026-07-05 13:15 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: Track U **GR-18/GR-T17** and sibling **TW-4** are closed in the
working tree. The next Granite move is **GR-A2**. A GR-A2 reconnaissance pass started
before context compaction; no GR-A2 code/test implementation has been made yet.

### GR-A2 Row To Resume

Granite row:

- **GR-A2 | MEDIUM**: `ConnectionManagerWindow` hosts a second full WPD stack and
  test fixtures in production source.
- Anchors from the plan:
  - `RedSalamander/ConnectionManagerWindow.cpp:918-1415`
  - `EnumerateMtpPickerDevices`
  - `EnumerateMtpPickerStorages`
  - `ReadMtpPickerDevicePersistentId`
  - host-side COM init/client-info/disambiguation
  - `#ifdef ENABLE_TESTS` globals `g_mtpPickerFixtureDevices` and related storage
    fixture state
  - `MtpPickerWorkerCallback` never consults the plugin.
- Plan direction: expose device/storage browse through the plugin contract
  (factory-level enumerate API keyed by `pluginId`), have the host picker call the
  plugin, have `FakeMtpBackend` serve selftests, and delete the host fixture globals.

### Reconnaissance Already Done

Important symbols found in `RedSalamander/ConnectionManagerWindow.cpp`:

- Protocol table row around line 808:
  `builtin/file-system-mtp` / `MTP`.
- MTP picker data structures around line 878:
  `MtpPickerDeviceEntry`, `MtpPickerStorageEntry`, `MtpPickerRequestKind`,
  `MtpPickerResultPayload`, `MtpPickerWorkerContext`, and
  `MtpPickerComInitialization`.
- Host-side WPD/helper clone functions:
  - `StableMtpPickerHash` around line 941.
  - `MtpPickerStringProperty` around line 983.
  - `MtpPickerGuidPropertyEquals` around line 995.
  - `CreateMtpPickerClientInfo` around line 1001.
  - `ReadMtpPickerDevicePersistentId` around line 1038.
  - `EnumerateMtpPickerDevices` around line 1098.
  - `EnumerateMtpPickerObjectIds` around line 1173.
  - `DisambiguateMtpPickerStorageNames` around line 1222.
  - `EnumerateMtpPickerStorages` around line 1250.
- Test fixture globals under `#ifdef ENABLE_TESTS`:
  - `g_mtpPickerFixtureMutex`
  - `g_mtpPickerFixtureEnabled`
  - `g_mtpPickerFixtureDevices`
  - `g_mtpPickerFixtureStoragesByDevice`
  - `TryGetMtpPickerFixtureDevices`
  - `TryGetMtpPickerFixtureStorages`
- `MtpPickerWorkerCallback` around line 1377 currently:
  - uses fixture devices/storages if enabled,
  - otherwise calls host-side WPD enumeration helpers,
  - posts `WndMsg::kConnectionManagerMtpPickerComplete`.
- Picker UI/control functions found:
  - `_mtpPickerDevices`, `_mtpPickerStorages` around line 1712.
  - `RefreshMtpDevicePicker` around line 2844.
  - `RefreshMtpStoragePicker` around line 2883.
  - `OnMtpPickerResult` around line 2900.
- Existing exported fixture/debug plumbing:
  - `DebugSetMtpPickerFixture` around line 5383.
  - `DebugClearMtpPickerFixture` around line 5429.
  - `DebugSetConnectionManagerMtpPickerFixture` around line 6435.
  - `DebugClearConnectionManagerMtpPickerFixture` around line 6446.

Existing selftest to migrate:

- File:
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`
- Test:
  `TestConnectionManagerWindowMtpPickerPopulatesProfile`, starting around line 2059.
- Registered case:
  `cmd_connection_manager_window_mtp_picker_populates_profile`, around line 7057.
- Current issue:
  the test uses `DebugSetConnectionManagerMtpPickerFixture(...)` /
  `DebugClearConnectionManagerMtpPickerFixture(...)`, which exercises the host fixture
  path that GR-A2 intends to delete. Convert this coverage to plugin-backed fake MTP
  browse data instead of preserving another host fixture.

Spec/doc context:

- `Specs/FileSystem/FileSystem_Mtp.md` currently documents MTP profile keys:
  `pluginId=builtin/file-system-mtp`, `host` as WPD PnP id, `path` as storage/subfolder
  path, `extra.devicePuid`, and `extra.readOnly`.
- That spec still says picker workers create their own WPD COM objects and marshal plain
  picker data back. Update it after GR-A2 so the durable contract says the host picker
  uses the plugin-backed browse API.

Search caveat from the interrupted pass:

- One `rg` attempt used PowerShell glob arguments like
  `Plugins\FileSystemMtp\*.cpp`; `rg` treated them as literal paths and failed.
- Re-run searches using directory paths, for example:

```powershell
rg -n "TryCreateFakeMtpFileSystemInstance|Debug|Fake|CreateInstance|CreateFileSystem|GetConfigurationSchema|GetCapabilities|extern|__declspec|IFileSystem" Plugins\FileSystemMtp Plugins\FileSystem RedSalamander\FileSystemPluginManager.* Common
```

### Immediate Resume Steps

1. Reconfirm the dirty tree with `git status --short`.
2. Re-read small bounded slices instead of huge full-file output:

```powershell
$p = ".\RedSalamander\ConnectionManagerWindow.cpp"
(Get-Content $p)[877..1420]
(Get-Content $p)[2740..2915]
(Get-Content $p)[5360..5450]
(Get-Content $p)[6420..6455]
(Get-Content ".\RedSalamander\SelfTest\Commands\Commands.SelfTest.Connections.cpp")[2058..2210]
```

3. Inspect the plugin manager and MTP plugin/fake backend contract before designing the
   API:

```powershell
rg -n "TryCreateFakeMtpFileSystemInstance|FakeMtp|Debug|CreateInstance|CreateFileSystem|IFileSystem|PluginManager|GetConfigurationSchema|GetCapabilities|__declspec|extern" Plugins\FileSystemMtp Plugins\FileSystem RedSalamander Common
```

4. Add RED coverage first. Preferred shape:
   migrate/extend `cmd_connection_manager_window_mtp_picker_populates_profile` so it
   depends on plugin-backed fake MTP browse data and fails while the host fixture/WPD
   clone remains the only picker path.
5. Implement the plugin-backed picker path, remove host-side WPD clone helpers and
   `g_mtpPickerFixture*` globals, then update `Specs/FileSystem/FileSystem_Mtp.md`, the
   Granite ledger, WIP README, and this baton.
6. Run focused Commands selftest, adjacent MTP picker/profile checks discovered during
   implementation, `git diff --check`, inventory checks if counts change, and a Debug
   x64 RedSalamander build.

### Current Dirty Tree

`git status --short` at this save point:

```text
 M Common/DxUi/DxUi.Accessibility.cpp
 M Common/DxUi/DxUi.Internal.h
 M Common/DxUi/DxUi.Menu.cpp
 M Common/DxUi/DxUi.WindowHost.cpp
 M Common/LocalSearchIndexCore.cpp
 M Common/LocalSearchIndexCore.h
 M Common/SearchServiceBroker.cpp
 M Common/SearchServiceBroker.h
 M Common/SqliteIndexStore.cpp
 M Common/SqliteIndexStore.h
 M Plugins/FileSystem/FileSystem.Internal.h
 M Plugins/FileSystem/FileSystem.Search.cpp
 M Plugins/FileSystem/FileSystem.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.Core.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.FakeBackend.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.h
 M RedSalamander/FolderView.Icons.cpp
 M RedSalamander/FolderView.h
 M RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp
 M RedSalamander/FolderWindow.h
 M RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.cpp
 M RedSalamanderSearchService/Main.cpp
 M Specs/Core/Core_Search.md
 M Specs/FileSystem/FileSystem_Mtp.md
 M Specs/Plans/WIP/Operation_Clearwater_TwoDayMasterReviewRemediation_2026-06-28.md
 M Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
 M Specs/Plans/WIP/Operation_Tailwind_FolderViewWarpDrivePreMergeReviewRemediation_2026-06-30.md
 M Specs/Plans/WIP/README.md
 M Specs/Testing/Testing_TestCoverage.md
 M Specs/UI/UI_DxUiSharedGrid.md
 M Specs/UI/UI_DxUiWinUIDesign.md
 M Specs/UI/UI_FolderView.md
 M Tests/DxUiTests/DxUiTests.Accessibility.cpp
 M Tests/DxUiTests/DxUiTests.Menu.cpp
 M Tests/README.md
 M Tools/Tests/TestInventory.Tests.ps1
?? Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md
```

Do not revert this dirty tree. It contains prior completed Granite/Clearwater slices,
the closed GR-18/TW-4 slice, and this continuation baton.

## Latest Save Point - 2026-07-05 13:12 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: Track U **GR-18/GR-T17** is closed in the working tree, and sibling
**TW-4** is closed in the Tailwind ledger. The next Granite move is **GR-A2**.

### GR-18 / TW-4 Closed This Turn

Implemented behavior:

- `FolderView::OnCreateThumbnailBitmap(...)` now returns on stale thumbnail batch or
  stale enumeration generation before registering the pending-count decrement.
- `FolderView::ThumbnailBitmapRequest` now carries `countsPending`, defaulting true for
  normal worker posts.
- Abandoned late provider-allowed thumbnail delivery sets `countsPending=false`, so a
  late current bitmap can apply without consuming another request's pending slot.
- The test-only thumbnail message helper was renamed to
  `DebugSeedThumbnailPendingAndPostThumbnailBitmapMessagesForTest(...)` and can inject
  stale-batch, stale-generation, and unaccounted-current bitmap payloads.

Selftest added/extended:

- `folderView_thumbnail_stale_bitmap_messages_preserve_pending_count` in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`.

Evidence:

- GR-18 stale-message RED:
  - Build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
    passed with 0 warnings / 0 errors; log `.build\logs\msbuild-20260705_124514_955.log`.
  - Focused test failed with
    `pending=0 expected=2 staleDrops=1`.
- GR-18 stale-message GREEN before TW-4 extension:
  - Build log `.build\logs\msbuild-20260705_124759_677.log`, 0 warnings / 0 errors.
  - Focused test passed 1/0.
- TW-4 RED:
  - Build log `.build\logs\msbuild-20260705_125655_971.log`, 0 warnings / 0 errors.
  - Focused test failed with
    `pending=1 expected=2 completed=1`.
- Final GREEN:
  - Build log `.build\logs\msbuild-20260705_125952_183.log`, 0 warnings / 0 errors.
  - Rebuild log `.build\logs\msbuild-20260705_130500_471.log`, 0 warnings / 0 errors.
  - `folderView_thumbnail_stale_bitmap_messages_preserve_pending_count` passed; log
    `.build\logs\gr18_tw4_folderView_thumbnail_stale_bitmap_messages_preserve_pending_count.out.txt`.
  - Adjacent `folderView_thumbnail_size_change_while_pending` passed; log
    `.build\logs\gr18_tw4_folderView_thumbnail_size_change_while_pending.out.txt`.
  - Adjacent `folderView_thumbnail_cached_only_no_close_stall` passed; log
    `.build\logs\gr18_tw4_folderView_thumbnail_cached_only_no_close_stall.out.txt`.
  - Optional provider guard `folderView_perf_slow_virtual_provider` passed; log
    `.build\logs\gr18_tw4_folderView_perf_slow_virtual_provider.out.txt`.
  - `git diff --check` exited 0 with only LF-to-CRLF warnings.
  - `Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru` passed 5/0.
  - Final app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
    passed with 0 warnings / 0 errors; log `.build\logs\msbuild-20260705_131031_435.log`.

Docs/ledger updated:

- `Specs/UI/UI_FolderView.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Tests/README.md`
- `Tools/Tests/TestInventory.Tests.ps1`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/Operation_Tailwind_FolderViewWarpDrivePreMergeReviewRemediation_2026-06-30.md`
- `Specs/Plans/WIP/README.md`

Inventory after this slice:

- Commands runner-listed cases: 781.
- Commands static `SelfTest::RunCase` registrations: 679.
- CompareDirectories runner-listed cases/static registrations: 248 / 240.
- FileOperations runner-listed phases/active ordered phases: 122 / 120.

### Current Dirty Tree

`git status --short` at this save point:

```text
 M Common/DxUi/DxUi.Accessibility.cpp
 M Common/DxUi/DxUi.Internal.h
 M Common/DxUi/DxUi.Menu.cpp
 M Common/DxUi/DxUi.WindowHost.cpp
 M Common/LocalSearchIndexCore.cpp
 M Common/LocalSearchIndexCore.h
 M Common/SearchServiceBroker.cpp
 M Common/SearchServiceBroker.h
 M Common/SqliteIndexStore.cpp
 M Common/SqliteIndexStore.h
 M Plugins/FileSystem/FileSystem.Internal.h
 M Plugins/FileSystem/FileSystem.Search.cpp
 M Plugins/FileSystem/FileSystem.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.Core.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.FakeBackend.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.h
 M RedSalamander/FolderView.Icons.cpp
 M RedSalamander/FolderView.h
 M RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp
 M RedSalamander/FolderWindow.h
 M RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.cpp
 M RedSalamanderSearchService/Main.cpp
 M Specs/Core/Core_Search.md
 M Specs/FileSystem/FileSystem_Mtp.md
 M Specs/Plans/WIP/Operation_Clearwater_TwoDayMasterReviewRemediation_2026-06-28.md
 M Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
 M Specs/Plans/WIP/Operation_Tailwind_FolderViewWarpDrivePreMergeReviewRemediation_2026-06-30.md
 M Specs/Plans/WIP/README.md
 M Specs/Testing/Testing_TestCoverage.md
 M Specs/UI/UI_DxUiSharedGrid.md
 M Specs/UI/UI_DxUiWinUIDesign.md
 M Specs/UI/UI_FolderView.md
 M Tests/DxUiTests/DxUiTests.Accessibility.cpp
 M Tests/DxUiTests/DxUiTests.Menu.cpp
 M Tests/README.md
 M Tools/Tests/TestInventory.Tests.ps1
?? Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md
```

Do not revert this dirty tree. It contains prior completed Granite/Clearwater slices,
the closed GR-18/TW-4 slice, and this continuation baton.

### Resume Checklist

1. Confirm the dirty tree with `git status --short`.
2. Read the current Granite row **GR-A2**.
3. Continue Track A with **GR-A2**: `ConnectionManagerWindow` hosts a second full WPD
   picker stack and test fixtures in production source; the plan direction is to route
   device/storage browse through the plugin contract and remove the host-side WPD clone.
4. Keep TDD/spec closeout discipline and do not move Granite to Done until every remaining
   row is closed and verified.

## Latest Save Point - 2026-07-05 12:53 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: Track U **GR-18/GR-T17** is partially closed in code and focused proof,
but sibling **TW-4** still needs to be folded into the same pass before docs, counts,
ledger, and WIP index closeout. The Granite routing table says GR-18 and TW-4 share the
thumbnail pending-counter accounting path and should be fixed together.

### Current Slice Status

Completed in the working tree:

- Added a Commands selftest,
  `folderView_thumbnail_stale_bitmap_messages_preserve_pending_count`, in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`.
- Added test-only helpers on `FolderView` and `FolderWindow` to seed thumbnail pending
  state and post stale thumbnail bitmap messages.
- Fixed GR-18 by moving the thumbnail pending decrement in
  `FolderView::OnCreateThumbnailBitmap(...)` after stale batch and stale enumeration
  generation checks.

Focused GR-18 evidence already collected:

- RED build:
  `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors; log `.build\logs\msbuild-20260705_124514_955.log`.
- RED test:
  `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_thumbnail_stale_bitmap_messages_preserve_pending_count -FailFast -TimeoutMultiplier 3`
  failed as intended with
  `pending=0 expected=2 staleDrops=1`.
- GREEN build:
  `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors; log `.build\logs\msbuild-20260705_124759_677.log`.
- GREEN focused test:
  same `Run-AllTests.ps1` command exited 0 with 1 passed, 0 failed.

Important: this only proves the stale-message half of the shared accounting issue.
Do not update docs or mark GR-18/GR-T17 done until TW-4 is covered and green too.

### Current Dirty Tree

`git status --short` at this save point:

```text
 M Common/DxUi/DxUi.Accessibility.cpp
 M Common/DxUi/DxUi.Internal.h
 M Common/DxUi/DxUi.Menu.cpp
 M Common/DxUi/DxUi.WindowHost.cpp
 M Common/LocalSearchIndexCore.cpp
 M Common/LocalSearchIndexCore.h
 M Common/SearchServiceBroker.cpp
 M Common/SearchServiceBroker.h
 M Common/SqliteIndexStore.cpp
 M Common/SqliteIndexStore.h
 M Plugins/FileSystem/FileSystem.Internal.h
 M Plugins/FileSystem/FileSystem.Search.cpp
 M Plugins/FileSystem/FileSystem.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.Core.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.FakeBackend.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.h
 M RedSalamander/FolderView.Icons.cpp
 M RedSalamander/FolderView.h
 M RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp
 M RedSalamander/FolderWindow.h
 M RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.cpp
 M RedSalamanderSearchService/Main.cpp
 M Specs/Core/Core_Search.md
 M Specs/FileSystem/FileSystem_Mtp.md
 M Specs/Plans/WIP/Operation_Clearwater_TwoDayMasterReviewRemediation_2026-06-28.md
 M Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
 M Specs/Plans/WIP/README.md
 M Specs/Testing/Testing_TestCoverage.md
 M Specs/UI/UI_DxUiSharedGrid.md
 M Specs/UI/UI_DxUiWinUIDesign.md
 M Tests/DxUiTests/DxUiTests.Accessibility.cpp
 M Tests/DxUiTests/DxUiTests.Menu.cpp
 M Tests/README.md
 M Tools/Tests/TestInventory.Tests.ps1
?? Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md
```

Do not revert this dirty tree. It contains prior completed Granite/Clearwater slices,
the partial GR-18 slice, and this continuation baton.

### Files Touched For GR-18 So Far

- `RedSalamander/FolderView.Icons.cpp`
- `RedSalamander/FolderView.h`
- `RedSalamander/FolderWindow.h`
- `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`

### Exact Next Move: Fold TW-4

Continue TDD before production changes:

1. Extend the existing helper and same selftest rather than adding another registered
   case unless a separate count bump is desired.
2. Recommended helper rename:
   `DebugSeedThumbnailPendingAndPostThumbnailBitmapMessagesForTest(...)`.
3. Add a fourth helper parameter:
   `uint64_t unaccountedCurrentMessageCount`.
4. Existing stale messages should still post:
   - stale batch, current generation
   - current batch, stale generation
5. New TW-4 test messages should post:
   - current batch, current generation
   - without incrementing `pendingBitmapCreates`
   - representing late provider delivery after the fallback consumed the only counted
     pending increment.
6. Extend the existing selftest after the stale-message assertion:
   - reset pending to 0
   - capture `beforeLate`
   - seed `kCurrentBatchPendingApplies`
   - post `unaccountedCurrentMessageCount = 1`
   - wait until `thumbnailCompletedCount >= beforeLate.thumbnailCompletedCount + 1`
   - assert `thumbnailPendingCount == kCurrentBatchPendingApplies`
7. Build and run the focused case. Expected RED before the TW-4 fix:
   pending is decremented by the unaccounted current message, likely
   `pending=1 expected=2`.

Commands for the TW-4 RED:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_thumbnail_stale_bitmap_messages_preserve_pending_count -FailFast -TimeoutMultiplier 3
```

After RED, implement the accounting fix:

- Add `bool countsPending = true;` to `FolderView::ThumbnailBitmapRequest` in
  `RedSalamander/FolderView.h`.
- In the late provider callback inside
  `FolderView::ExtractProviderAllowedThumbnailWithDeadline(...)`, set
  `payload->countsPending = false;`.
- In the helper, set `countsPending = false` for the TW-4 unaccounted current messages.
  Stale messages may keep the default `true` because stale checks return before the
  decrement.
- In `FolderView::OnCreateThumbnailBitmap(...)`, have the decrement `scope_exit` return
  without subtracting when `requestPtr->countsPending` is false.

Then rerun the same build and focused test for GREEN.

### Inventory And Docs Waiting On TW-4

Current inventory after adding the one Commands case:

```text
commands.runCaseRegistrations = 679
compareDirectories.runCaseRegistrations = 240
fileOperations.activePhases = 120
```

`RedSalamander.exe --selftest-list-cases --commands-selftest` reported 781 listed
Commands cases and included
`folderView_thumbnail_stale_bitmap_messages_preserve_pending_count`.

After TW-4 is green, update:

- `Tools/Tests/TestInventory.Tests.ps1`
  - Commands static fallback expected count from 678 to 679 in both assertions.
- `Tests/README.md`
  - Commands listed cases: 781.
  - Commands static fallback: 679.
  - ViewCommands family: 107.
- `Specs/Testing/Testing_TestCoverage.md`
  - Commands listed/static counts to 781/679.
  - ViewCommands static registrations to 107.
  - Add coverage row for
    `folderView_thumbnail_stale_bitmap_messages_preserve_pending_count`.
  - Add a recent coverage bullet for GR-18/TW-4.
- `Specs/UI/UI_FolderView.md`
  - Durable thumbnail pending-accounting contract for stale batch/generation drops and
    late unaccounted provider delivery.
- Granite WIP plan
  - Mark GR-18 and GR-T17 DONE 2026-07-05 only after TW-4 proof is green.
  - Mention TW-4 folded into the same fix.
- `Specs/Plans/WIP/README.md`
  - Advance to the next Granite row after re-reading the ledger.
- This baton.

### Closeout Checks After TW-4

Minimum focused and adjacent checks before marking the slice done:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_thumbnail_stale_bitmap_messages_preserve_pending_count -FailFast -TimeoutMultiplier 3
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_thumbnail_size_change_while_pending -FailFast -TimeoutMultiplier 3
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_thumbnail_cached_only_no_close_stall -FailFast -TimeoutMultiplier 3
git diff --check
Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64
```

Optional but useful if time allows:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_perf_slow_virtual_provider -FailFast -TimeoutMultiplier 3
```

### Resume Checklist

1. Run `git status --short` and confirm the dirty tree.
2. Read this newest baton section plus the GR-18/GR-T17 and TW-4 rows.
3. Add the TW-4 assertion to the existing Commands selftest and verify RED.
4. Implement the `countsPending` accounting flag and verify GREEN.
5. Only then update docs, counts, Granite ledger, WIP README, and this baton.

## Latest Save Point - 2026-07-05 12:40 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: Track U **GR-16/GR-T16** is closed in the working tree. The next
Granite move is **GR-18/GR-T17**.

### GR-16 Closed This Turn

Implemented behavior:

- Published accessibility snapshots now carry `hasCollapsedSemanticRoot`.
- The snapshot collapse flag is computed with
  `TryResolveSingleSemanticRootControlPath(...)`, so it uses the same
  `ShouldExposeSingleSemanticRootControl(...)` predicate as the retained path.
- Snapshot navigation/properties no longer collapse any arbitrary single semantic
  control. Label-only roots remain a root provider with a label child, matching
  retained-path semantics.

Selftest added:

- `TestAccessibilityLabelOnlyRootDoesNotUseDirectSemanticRootCollapse()` in
  `Tests/DxUiTests/DxUiTests.Accessibility.cpp`.

Evidence:

- RED `.\build.ps1 -ProjectName DxUiTests -Configuration Debug -Platform x64`:
  `.build/logs/msbuild-20260705_123529_940.log`, 0 warnings / 0 errors.
- RED `.\.build\x64\Debug\DxUiTests.exe --suite=Accessibility`: exited 1 at
  `label-only root exposes the label as a child instead of collapsing it into the root`.
- GREEN `.\build.ps1 -ProjectName DxUiTests -Configuration Debug -Platform x64`:
  `.build/logs/msbuild-20260705_123625_707.log`, 0 warnings / 0 errors.
- GREEN `.\.build\x64\Debug\DxUiTests.exe --suite=Accessibility`: exited 0.
- `git diff --check` exited 0 with only LF-to-CRLF warnings.
- `Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru` passed 5/0.
- Final app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors; log `.build/logs/msbuild-20260705_123839_070.log`.

Docs/ledger updated:

- `Specs/UI/UI_DxUiSharedGrid.md`
- `Specs/UI/UI_DxUiWinUIDesign.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Tests/README.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`

Known verification caveat:

- Broad `.\.build\x64\Debug\DxUiTests.exe` was not claimed green in this slice. The
  previously recorded focused `NewControls` failure at `tab control handles close-button
  release` remains a separate caveat unless a later resume proves otherwise.

### Current Dirty Tree

`git status --short` at this save point:

```text
 M Common/DxUi/DxUi.Accessibility.cpp
 M Common/DxUi/DxUi.Internal.h
 M Common/DxUi/DxUi.Menu.cpp
 M Common/DxUi/DxUi.WindowHost.cpp
 M Common/LocalSearchIndexCore.cpp
 M Common/LocalSearchIndexCore.h
 M Common/SearchServiceBroker.cpp
 M Common/SearchServiceBroker.h
 M Common/SqliteIndexStore.cpp
 M Common/SqliteIndexStore.h
 M Plugins/FileSystem/FileSystem.Internal.h
 M Plugins/FileSystem/FileSystem.Search.cpp
 M Plugins/FileSystem/FileSystem.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.Core.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.FakeBackend.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.h
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.cpp
 M RedSalamanderSearchService/Main.cpp
 M Specs/Core/Core_Search.md
 M Specs/FileSystem/FileSystem_Mtp.md
 M Specs/Plans/WIP/Operation_Clearwater_TwoDayMasterReviewRemediation_2026-06-28.md
 M Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
 M Specs/Plans/WIP/README.md
 M Specs/Testing/Testing_TestCoverage.md
 M Specs/UI/UI_DxUiSharedGrid.md
 M Specs/UI/UI_DxUiWinUIDesign.md
 M Tests/DxUiTests/DxUiTests.Accessibility.cpp
 M Tests/DxUiTests/DxUiTests.Menu.cpp
 M Tests/README.md
 M Tools/Tests/TestInventory.Tests.ps1
?? Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md
```

Do not revert this dirty tree. It contains prior completed Granite/Clearwater slices,
GR-9, GR-10, GR-6, GR-16, and this continuation baton.

### Resume Checklist

1. Confirm the dirty tree with `git status --short`.
2. Read the current `GR-18` and `GR-T17` rows in the Granite plan.
3. Continue Track U with **GR-18/GR-T17**:
   - Bug: `OnCreateThumbnailBitmap` decrements `pendingBitmapCreates` before stale
     batch/generation checks.
   - Expected fix direction: move the stale batch/generation early returns above the
     decrement `scope_exit`, mirroring `OnCreateIconBitmap`.
   - Expected test direction: extend the thumbnail snapshot selftest with
     cancel+requeue plus stale messages in flight, asserting `pendingCount` equals the
     queued new-batch applies.

## Latest Save Point - 2026-07-05 12:34 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: Track U **GR-6/GR-T6** is closed in the working tree, with folded
**GR-17** also closed. The next Granite move is **GR-16/GR-T16**, followed by
GR-18/GR-T17.

### GR-6 Closed This Turn

Implemented behavior:

- Cross-thread UIA actions now post a heap-owned shared dispatch block through
  `PostMessagePayload(...)` and receive it with
  `TakeMessagePayload<std::shared_ptr<AccessibilityUiActionDispatch>>(...)`.
- The dispatch block owns request/output storage, keeps element and text-range
  providers alive with `wil::com_ptr_nothrow`, signals completion through an event,
  and returns `ERROR_TIMEOUT` without reading late output after a sender timeout.
- `WindowHost` initializes the posted-payload registry on attach and drains queued
  payloads during `WM_NCDESTROY`.
- The test-only `DebugGetContextMenuPopupState` path no longer uses timed-out
  `SendMessageTimeoutW(...)` with caller stack output; it uses synchronous
  `SendMessageW(...)`.

Selftests added:

- `TestAccessibilityUiActionDispatchOwnsTimedOutRequestStorage()`.
- `TestAccessibilityTextRangeBoundingRectanglesTimeoutKeepsLateHandlerStorageAlive()`.
- `TestContextMenuDebugStateCrossThreadQueryDoesNotUseTimedOutStackStorage()`.

Evidence:

- RED Accessibility source-contract guard after build log
  `.build/logs/msbuild-20260705_121657_689.log`: suite failed at
  `accessibility UIA action shared dispatch helper is found`.
- RED Menu source-contract guard after build log
  `.build/logs/msbuild-20260705_122109_640.log`: suite failed at
  `context menu debug state cross-thread query does not time out while using caller stack output storage`.
- GREEN `.\build.ps1 -ProjectName DxUiTests -Configuration Debug -Platform x64`:
  `.build/logs/msbuild-20260705_122720_807.log`, 0 warnings / 0 errors.
- GREEN `.\.build\x64\Debug\DxUiTests.exe --suite=Accessibility`: exited 0 and
  includes the late-handler timeout regression.
- GREEN `.\.build\x64\Debug\DxUiTests.exe --suite=Menu`: exited 0.
- GREEN `.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost`: exited 0.
- `git diff --check` exited 0 with only LF-to-CRLF warnings.
- `Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru` passed 5/0.
- Final app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors; log `.build/logs/msbuild-20260705_123200_423.log`.

Docs/ledger updated:

- `Specs/UI/UI_DxUiSharedGrid.md`
- `Specs/UI/UI_DxUiWinUIDesign.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Tests/README.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`

Known verification caveat:

- Broad `.\.build\x64\Debug\DxUiTests.exe` was not claimed green in this slice. The
  previously recorded focused `NewControls` failure at `tab control handles close-button
  release` remains a separate caveat unless a later resume proves otherwise.

### Current Dirty Tree

`git status --short` at this save point:

```text
 M Common/DxUi/DxUi.Accessibility.cpp
 M Common/DxUi/DxUi.Internal.h
 M Common/DxUi/DxUi.Menu.cpp
 M Common/DxUi/DxUi.WindowHost.cpp
 M Common/LocalSearchIndexCore.cpp
 M Common/LocalSearchIndexCore.h
 M Common/SearchServiceBroker.cpp
 M Common/SearchServiceBroker.h
 M Common/SqliteIndexStore.cpp
 M Common/SqliteIndexStore.h
 M Plugins/FileSystem/FileSystem.Internal.h
 M Plugins/FileSystem/FileSystem.Search.cpp
 M Plugins/FileSystem/FileSystem.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.Core.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.FakeBackend.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.h
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.cpp
 M RedSalamanderSearchService/Main.cpp
 M Specs/Core/Core_Search.md
 M Specs/FileSystem/FileSystem_Mtp.md
 M Specs/Plans/WIP/Operation_Clearwater_TwoDayMasterReviewRemediation_2026-06-28.md
 M Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
 M Specs/Plans/WIP/README.md
 M Specs/Testing/Testing_TestCoverage.md
 M Specs/UI/UI_DxUiSharedGrid.md
 M Specs/UI/UI_DxUiWinUIDesign.md
 M Tests/DxUiTests/DxUiTests.Accessibility.cpp
 M Tests/DxUiTests/DxUiTests.Menu.cpp
 M Tests/README.md
 M Tools/Tests/TestInventory.Tests.ps1
?? Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md
```

Do not revert this dirty tree. It contains prior completed Granite/Clearwater slices,
GR-9, GR-10, the closed GR-6/GR-T6 slice, and this continuation baton.

### Resume Checklist

1. Confirm the dirty tree with `git status --short`.
2. Read the current `GR-16` / `GR-T16` rows in the Granite plan.
3. Continue Track U with **GR-16/GR-T16**.

## Latest Save Point - 2026-07-05 12:23 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: Track U **GR-6/GR-T6** is implemented in the working tree, including the
folded **GR-17** test-only menu timeout path, but the slice is **not closed out yet**.
Do not move on to GR-16 until the documentation/count updates and final focused checks
below are completed.

### GR-6 Implemented This Turn

Implemented behavior:

- `Common/DxUi/DxUi.Accessibility.cpp` now marshals cross-thread UIA action requests by
  posting a heap-owned `std::shared_ptr<AccessibilityUiActionDispatch>` through
  `PostMessagePayload(...)`, instead of sending a caller-stack request through
  `SendMessageTimeoutW(...)`.
- The dispatch block owns the moved request, holds `wil::com_ptr_nothrow` references to
  the provider/text-range provider, and carries a completion event.
- Sender threads wait on the event for 5s and return `HRESULT_FROM_WIN32(ERROR_TIMEOUT)`
  on timeout without reading late handler output. Late handlers still write into live
  heap storage and signal the event.
- `TryHandleWindowHostAccessibilityMessage(...)` now receives the posted shared dispatch
  with `TakeMessagePayload<std::shared_ptr<AccessibilityUiActionDispatch>>(...)`.
- `AccessibilityUiActionRequest::textRangeBoundsDip` is owned storage, not a caller stack
  pointer.
- `Common/DxUi/DxUi.WindowHost.cpp` now calls `InitPostedPayloadWindow(_hwnd)` on attach
  and drains posted payloads in `WM_NCDESTROY`.
- `Common/DxUi/DxUi.Menu.cpp` folded GR-17 by changing
  `DebugGetContextMenuPopupState` from timed-out `SendMessageTimeoutW(...)` with stack
  output to synchronous `SendMessageW(...)`.

Selftests added:

- `TestAccessibilityUiActionDispatchOwnsTimedOutRequestStorage()` in
  `Tests/DxUiTests/DxUiTests.Accessibility.cpp`.
- `TestContextMenuDebugStateCrossThreadQueryDoesNotUseTimedOutStackStorage()` in
  `Tests/DxUiTests/DxUiTests.Menu.cpp`.

Evidence already collected:

- RED accessibility source-contract guard: after build log
  `.build/logs/msbuild-20260705_121657_689.log`,
  `.\.build\x64\Debug\DxUiTests.exe --suite=Accessibility` exited 1 at
  `accessibility UIA action shared dispatch helper is found`.
- GREEN accessibility after UIA dispatch implementation: final clean build log
  `.build/logs/msbuild-20260705_121949_531.log` passed with 0 warnings / 0 errors, then
  `.\.build\x64\Debug\DxUiTests.exe --suite=Accessibility` exited 0 and printed
  `[DONE] Accessibility`.
- RED menu source-contract guard: after build log
  `.build/logs/msbuild-20260705_122109_640.log`,
  `.\.build\x64\Debug\DxUiTests.exe --suite=Menu` exited 1 at
  `context menu debug state cross-thread query does not time out while using caller stack output storage`.
- GREEN menu after changing the debug query to `SendMessageW`: build log
  `.build/logs/msbuild-20260705_122133_179.log` passed with 0 warnings / 0 errors, then
  `.\.build\x64\Debug\DxUiTests.exe --suite=Menu` exited 0 and printed `[DONE] Menu`.

Important caveat:

- GR-T6's original text asked for a stalled-handler/AppVerifier-style dynamic harness.
  The current working tree has deterministic source-contract guards plus the ownership
  rewrite, not an AppVerifier/ASan run. Decide during closeout whether that proof is
  acceptable for this codebase or whether to add the dynamic stalled-handler harness
  before marking GR-T6 done.

### Remaining Closeout Before GR-16

Run/check these before marking GR-6/GR-T6 done:

1. Re-run `.\.build\x64\Debug\DxUiTests.exe --suite=Accessibility` after all current
   edits.
2. Run `.\.build\x64\Debug\DxUiTests.exe --suite=WindowHost` because `WindowHost`
   now initializes/drains posted payloads.
3. Run `git diff --check`.
4. Update test inventory docs/counts for the new Accessibility and Menu tests:
   `Tests/README.md`, `Specs/Testing/Testing_TestCoverage.md`, and
   `Tools/Tests/TestInventory.Tests.ps1`.
5. Update the durable UI/DxUi contract, most likely `Specs/UI/UI_DxUiSharedGrid.md` or
   `Specs/UI/UI_DxUiWinUIDesign.md`, with the UIA cross-thread dispatch ownership and
   posted-payload init/drain rule.
6. Update the Granite ledger row `GR-6` and test row `GR-T6` to done only after the
   proof decision above is settled.
7. Update `Specs/Plans/WIP/README.md` so the next Granite row becomes **GR-16/GR-T16**.
8. Run `Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru`.
9. Run final `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`.

Known verification caveat:

- Broad `.\.build\x64\Debug\DxUiTests.exe` was not claimed green in this slice. The
  previously recorded focused `NewControls` failure at `tab control handles close-button
  release` remains a separate caveat unless a later resume proves otherwise.

### Current Dirty Tree

`git status --short` at this save point:

```text
 M Common/DxUi/DxUi.Accessibility.cpp
 M Common/DxUi/DxUi.Menu.cpp
 M Common/DxUi/DxUi.WindowHost.cpp
 M Common/LocalSearchIndexCore.cpp
 M Common/LocalSearchIndexCore.h
 M Common/SearchServiceBroker.cpp
 M Common/SearchServiceBroker.h
 M Common/SqliteIndexStore.cpp
 M Common/SqliteIndexStore.h
 M Plugins/FileSystem/FileSystem.Internal.h
 M Plugins/FileSystem/FileSystem.Search.cpp
 M Plugins/FileSystem/FileSystem.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.Core.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.FakeBackend.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.h
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.cpp
 M RedSalamanderSearchService/Main.cpp
 M Specs/Core/Core_Search.md
 M Specs/FileSystem/FileSystem_Mtp.md
 M Specs/Plans/WIP/Operation_Clearwater_TwoDayMasterReviewRemediation_2026-06-28.md
 M Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
 M Specs/Plans/WIP/README.md
 M Specs/Testing/Testing_TestCoverage.md
 M Specs/UI/UI_DxUiSharedGrid.md
 M Tests/DxUiTests/DxUiTests.Accessibility.cpp
 M Tests/DxUiTests/DxUiTests.Menu.cpp
 M Tests/README.md
 M Tools/Tests/TestInventory.Tests.ps1
?? Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md
```

Do not revert this dirty tree. It contains prior completed Granite/Clearwater slices,
GR-9, GR-10, the current partial GR-6/GR-T6 slice, and this continuation baton.

### Resume Checklist

1. Confirm the dirty tree with `git status --short`.
2. Read this newest baton section and the current `GR-6` / `GR-T6` rows in the Granite
   plan before editing.
3. Complete the GR-6 closeout list above.
4. Only after GR-6/GR-T6 is marked done, continue Track U with **GR-16/GR-T16**.

## Latest Save Point - 2026-07-05 12:09 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: Track U **GR-10/GR-T10** is closed in the working tree. The next Granite
move is **GR-6/GR-T6**, followed by GR-16/GR-T16 per the WIP index.

### GR-10 Closed This Turn

Implemented behavior:

- Snapshot point-hit records now carry an accumulated window-space translation plus an
  optional ancestor viewport clip.
- `DxUi::ScrollPanel` descendants inherit the panel viewport clip and `-GetScrollOffset()`
  translation, so content-space child bounds are recorded in clipped window DIP space.
- Tree, grid, control, and password-reveal point-hit rectangles are transformed/clipped
  before they are appended; empty clipped rectangles are skipped.
- A root fallback point-hit record is appended after child records, so specific child
  fragments still win, while clipped empty viewport areas resolve the root provider.

Selftest:

- Added `TestAccessibilityProviderPointHitsClipAndTranslateScrollPanelChildren()` in
  `Tests/DxUiTests/DxUiTests.Accessibility.cpp`.

Evidence:

- RED: after build log `.build/logs/msbuild-20260705_120301_879.log`,
  `.\.build\x64\Debug\DxUiTests.exe --suite=Accessibility` exited 1 at
  `scrolled panel empty viewport provider is the root provider hr=0x80004002`.
- GREEN: after build log `.build/logs/msbuild-20260705_120534_990.log`,
  `.\.build\x64\Debug\DxUiTests.exe --suite=Accessibility` exited 0 and printed
  `[DONE] Accessibility`.
- Adjacent: `.\.build\x64\Debug\DxUiTests.exe --suite=Control` exited 0 and printed
  `[DONE] Control`.
- `git diff --check` exited 0 with only LF-to-CRLF warnings.
- `Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru` passed 5/0.
- Final app build `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors; log `.build/logs/msbuild-20260705_120717_281.log`.

Known verification caveat:

- Broad `.\.build\x64\Debug\DxUiTests.exe` was not claimed green in this slice. The
  previously recorded focused `NewControls` failure at `tab control handles close-button
  release` remains a separate caveat unless a later resume proves otherwise.

Docs/ledger updated:

- `Specs/UI/UI_DxUiSharedGrid.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Tests/README.md`

### Current Dirty Tree

`git status --short` at this save point:

```text
 M Common/DxUi/DxUi.Accessibility.cpp
 M Common/LocalSearchIndexCore.cpp
 M Common/LocalSearchIndexCore.h
 M Common/SearchServiceBroker.cpp
 M Common/SearchServiceBroker.h
 M Common/SqliteIndexStore.cpp
 M Common/SqliteIndexStore.h
 M Plugins/FileSystem/FileSystem.Internal.h
 M Plugins/FileSystem/FileSystem.Search.cpp
 M Plugins/FileSystem/FileSystem.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.Core.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.FakeBackend.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.h
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.cpp
 M RedSalamanderSearchService/Main.cpp
 M Specs/Core/Core_Search.md
 M Specs/FileSystem/FileSystem_Mtp.md
 M Specs/Plans/WIP/Operation_Clearwater_TwoDayMasterReviewRemediation_2026-06-28.md
 M Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
 M Specs/Plans/WIP/README.md
 M Specs/Testing/Testing_TestCoverage.md
 M Specs/UI/UI_DxUiSharedGrid.md
 M Tests/DxUiTests/DxUiTests.Accessibility.cpp
 M Tests/README.md
 M Tools/Tests/TestInventory.Tests.ps1
?? Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md
```

Do not revert this dirty tree. It contains prior completed Granite/Clearwater slices,
GR-9, GR-10, and this continuation baton.

### Resume Checklist

1. Confirm the dirty tree with `git status --short`.
2. Read `GR-6` and `GR-T6` in the Granite plan before editing.
3. Re-read the relevant UIA cross-thread dispatch code in `Common/DxUi/DxUi.Accessibility.cpp`
   and the test-only `DebugGetContextMenuPopupState` path in `Common/DxUi/DxUi.Menu.cpp`.
4. Add the deterministic stalled-handler regression first and verify RED.
5. Implement the heap-owned request/event timeout shape, update durable specs/ledger/index,
   and run focused GREEN plus adjacent build/checks.

## Latest Save Point - 2026-07-05 11:58 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP. Do not call the goal complete.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: Track U **GR-9/GR-T9** is closed in the working tree. Stop point was a
user-requested save before starting the next Granite row.

Next move: Track U **GR-10/GR-T10**.

- `GR-10`: snapshot point-hit records must be clipped/transformed by ancestor viewports
  and scroll offsets.
- `GR-T10`: add a DxUi accessibility ScrollPanel regression where an off-viewport
  content-space child must not be returned by `ElementProviderFromPoint`, while a
  visible scrolled child still resolves correctly.
- After GR-10/GR-T10, continue with GR-6/GR-T6 per `Specs/Plans/WIP/README.md`.

### GR-9 Closed At This Save Point

Implemented behavior:

- `AccessibilityControlNavigationSnapshot` now keeps `gridVisibleColumns` for visible
  headers/point-hit behavior and `gridAccessibleColumns` for structural row exposure.
- Grid row accessible names, row-owned `GridCell` records, snapshot cell lookup, and row
  cell sibling navigation use the full model column set independent of horizontal scroll.
- Horizontally off-view cells remain present in the UIA tree and report offscreen state;
  hit testing and visible header exposure remain viewport-clipped.

Selftest:

- Added `TestAccessibilityProviderExposesHorizontallyScrolledGridRowStructure()` in
  `Tests/DxUiTests/DxUiTests.Accessibility.cpp`.

Evidence:

- RED: after build log `.build/logs/msbuild-20260705_115326_983.log`,
  `.\.build\x64\Debug\DxUiTests.exe --suite=Accessibility` exited 1 at
  `scrolled grid row name includes all model columns, including horizontally off-view cells`.
- GREEN: after build log `.build/logs/msbuild-20260705_115509_020.log`,
  `.\.build\x64\Debug\DxUiTests.exe --suite=Accessibility` exited 0 and printed
  `[DONE] Accessibility`.
- `git diff --check` exited 0 with only LF-to-CRLF warnings.
- `Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru` passed 5/0.

Known verification caveat:

- Broad `.\.build\x64\Debug\DxUiTests.exe` currently exits 1 in `NewControls` at
  `FAILED: tab control handles close-button release`.
- Focused `.\.build\x64\Debug\DxUiTests.exe --suite=NewControls` reproduces the same
  failure.
- Search/diff evidence showed no `TabControl`, `NewControls`, or `DxUi.Controls.cpp`
  edits in this GR-9 slice. Treat this as an existing unrelated broad-suite failure
  unless a later resume finds contrary evidence. Do not claim the full DxUi suite is
  green.

Docs/ledger updated:

- `Specs/UI/UI_DxUiSharedGrid.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Tests/README.md`

### Current Dirty Tree

`git status --short` at this save point:

```text
 M Common/DxUi/DxUi.Accessibility.cpp
 M Common/LocalSearchIndexCore.cpp
 M Common/LocalSearchIndexCore.h
 M Common/SearchServiceBroker.cpp
 M Common/SearchServiceBroker.h
 M Common/SqliteIndexStore.cpp
 M Common/SqliteIndexStore.h
 M Plugins/FileSystem/FileSystem.Internal.h
 M Plugins/FileSystem/FileSystem.Search.cpp
 M Plugins/FileSystem/FileSystem.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.Core.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.FakeBackend.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.h
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.cpp
 M RedSalamanderSearchService/Main.cpp
 M Specs/Core/Core_Search.md
 M Specs/FileSystem/FileSystem_Mtp.md
 M Specs/Plans/WIP/Operation_Clearwater_TwoDayMasterReviewRemediation_2026-06-28.md
 M Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
 M Specs/Plans/WIP/README.md
 M Specs/Testing/Testing_TestCoverage.md
 M Specs/UI/UI_DxUiSharedGrid.md
 M Tests/DxUiTests/DxUiTests.Accessibility.cpp
 M Tests/README.md
 M Tools/Tests/TestInventory.Tests.ps1
?? Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md
```

Do not revert this dirty tree. It contains prior completed Granite/Clearwater slices,
the completed GR-9 slice, and this continuation baton.

### Resume Checklist

1. Confirm the dirty tree with `git status --short`.
2. Read `GR-10` and `GR-T10` in the Granite plan before editing.
3. Add the ScrollPanel/content-space point-hit accessibility selftest first and verify it
   fails RED.
4. Implement snapshot point-hit clipping and scroll translation.
5. Run focused GREEN, adjacent checks/build, then update the UI spec, Granite ledger,
   WIP index, test inventory docs if counts change, and this baton.

## Latest Save Point - 2026-07-05 11:56 local

This section supersedes every older save point below for immediate resume. Granite is
still active and must stay in WIP.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: Track U **GR-9/GR-T9** is closed in the working tree. The next Granite
move is **GR-10/GR-T10** (snapshot point-hit records must be clipped/transformed by
ancestor viewports/scroll offsets), followed by GR-6/GR-T6 per the WIP index.

### GR-9 Closed This Turn

Implemented behavior:

- `AccessibilityControlNavigationSnapshot` now keeps `gridVisibleColumns` for
  visible headers/point-hit behavior and `gridAccessibleColumns` for structural row
  exposure.
- Grid row accessible names, row-owned `GridCell` records, snapshot cell lookup, and
  row cell sibling navigation now use the full model column set independent of
  horizontal scroll.
- Horizontally off-view cells remain present in the UIA tree and report offscreen state;
  hit testing and visible header exposure remain viewport-clipped.

Selftest:

- Added `TestAccessibilityProviderExposesHorizontallyScrolledGridRowStructure()` in
  `Tests/DxUiTests/DxUiTests.Accessibility.cpp`.

Evidence:

- RED: after build log `.build/logs/msbuild-20260705_115326_983.log`,
  `.\.build\x64\Debug\DxUiTests.exe --suite=Accessibility` exited 1 at
  `scrolled grid row name includes all model columns, including horizontally off-view cells`.
- GREEN: after build log `.build/logs/msbuild-20260705_115509_020.log`,
  `.\.build\x64\Debug\DxUiTests.exe --suite=Accessibility` exited 0 and printed
  `[DONE] Accessibility`.

Closeout/docs updated:

- `Specs/UI/UI_DxUiSharedGrid.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Tests/README.md`

### Next Move On Resume

1. Confirm the dirty tree with `git status --short`; do not revert existing dirty files.
2. Read Track U row `GR-10` and test row `GR-T10` in the Granite plan.
3. Add the DxUi accessibility RED test for ScrollPanel/content-space point-hit clipping.
4. Implement snapshot point-hit clipping/scroll translation, then run focused GREEN,
   adjacent checks/build, and update specs/ledger/baton.

## Latest Save Point - 2026-07-05 11:51 local

This section supersedes every older save point below for immediate resume. The 11:50
section remains the full GR-S3/GR-T15 closeout evidence. Granite is still active and
must stay in WIP.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current next move: Track U **GR-9/GR-T9**. No GR-9 production or test files have been
edited yet at this save point.

### Resume Checklist

1. Confirm dirty tree with `git status --short`; do not revert existing dirty files.
2. Add the GR-T9 DxUi accessibility selftest first, then run it RED.
3. Implement the snapshot-builder fix for full-column row names and GridCell fragments.
4. Run the focused GREEN/adjacent DxUi checks, build, then update specs/ledger/index.

### GR-9 Problem Summary

Plan row:

- `GR-9`: grid row UIA names and per-cell navigation drop columns that are horizontally
  scrolled out of view.
- `GR-T9`: `DxUiTests.Accessibility` should create a grid wider than its viewport,
  scroll right, then assert the row UIA `Name` contains all column texts and GridCell
  fragments exist for off-view columns.

Likely production anchors:

- `Common/DxUi/DxUi.Accessibility.cpp`
  - Snapshot record currently stores `gridVisibleColumns`.
  - Builder around the GR-9 plan anchors fills `gridVisibleColumns` from
    `Grid::GetVisibleColumnAt(...)`, then builds `gridCells` and
    `gridRowAccessibleName` by iterating only those visible columns.
  - `FindSnapshotGridCellRecord(...)` also gates lookups on `gridVisibleColumns`.
  - Row UIA `Name` is served from `gridRowAccessibleName`.
- `Common/DxUi/DxUi.Grid.cpp`
  - `Grid::GetVisibleColumnCount()` / `Grid::GetVisibleColumnAt(...)` use
    `ComputeVisibleColumnSpan(...)`.
  - `ComputeVisibleColumnSpan(...)` is viewport-clipped by `_horizontalScrollDip`.
- `Common/DxUi/DxUi.h`
  - Test hook: `Grid::DebugSetScrollOffsets(float verticalScrollDip, float horizontalScrollDip)`.

Expected implementation shape:

- Keep `gridVisibleColumns` for visible headers and point-hit behavior.
- Add/use a separate full-column list for structural row/cell exposure.
- Build `gridRowAccessibleName`, `gridCells`, and row child navigation from the full
  model column set, independent of horizontal scroll.
- Keep viewport clipping only for point-hit rectangles; do not make hit-testing expose
  off-view cells.

Test breadcrumbs:

- Test file: `Tests/DxUiTests/DxUiTests.Accessibility.cpp`.
- Existing sibling pattern:
  `TestAccessibilityProviderExposesGridRowSelectionPatterns()`.
- Add a nearby test with a one-row, three-column grid whose total column width exceeds
  the viewport, then call `DebugSetScrollOffsets(0.0f, <rightward offset>)` before
  creating/querying the accessibility provider.
- Navigate from the grid provider to the first row and assert:
  - row `UIA_NamePropertyId` is the full joined row text, e.g.
    `Alpha | Ready | Archived`;
  - row first/next GridCell children include the off-view first column rather than only
    visible columns;
  - the off-view cell reports `UIA_IsOffscreenPropertyId == true`.
- Add the new test call in `RunAccessibilityTests()`.

Likely commands:

```powershell
.\build.ps1 -ProjectName DxUiTests -Configuration Debug -Platform x64
.\.build\x64\Debug\DxUiTests.exe --suite=Accessibility
```

## Latest Save Point - 2026-07-05 11:50 local

This section supersedes every older save point below. Granite is still active and must
stay in WIP. Do not call the goal complete.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: GR-S3/GR-T15 is closed in the working tree. The next Granite move is
Track U **GR-9/GR-T9** (grid row UIA names/cells must include horizontally scrolled-out
columns), followed by GR-10/GR-T10 and GR-6/GR-T6 per the WIP index.

### GR-S3 Closed This Turn

Implemented behavior:

- SQLite `meta.store_generation` is initialized by bootstrap/v2 migration.
- `ReplaceVolume`, `ApplyJournalDelta`, and existing-volume `DeleteVolume` increment the
  generation inside their write transaction.
- `StoreInfo` and cached `PersistentStoreInfo` carry the inspected generation.
- Query paths validate cached generation with `SqliteIndexStore::ReadStoreGeneration(...)`
  before direct SQLite use and refresh cached store info on mismatch/read failure.
- Stale/missing/invalid/cutover retry is no longer restricted to zero emitted rows.
- Candidate path keys are deduped across retry/live fallback before forwarding.

Selftest:

- Added `search_service_sqlite_external_rotation_refreshes_without_retry`.

Evidence:

- RED: `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_113058/`
  failed because the post-rotation query used `search.backend.sqlite.retry_query_ms`.
- GREEN before docs: `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_113650/`
  passed 1/0.
- Final GREEN after fresh build:
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_114819/` passed 1/0.
- Adjacent passed:
  - `2026-07-05_114114/` =
    `search_service_sqlite_cold_start_stale_root_refreshes_before_query`
  - `2026-07-05_114131/` =
    `search_service_sqlite_invalid_store_falls_back_live_scan`
  - `2026-07-05_114218/` =
    `sqlite_index_store_load_and_apply_journal_delta`
- Adjacent skipped because live journal currentness cursor was unavailable:
  - `2026-07-05_114043/` =
    `local_index_core_sqlite_cutover_blocked_by_pending_legacy_import`
  - `2026-07-05_114147/` =
    `search_service_sqlite_query_failure_falls_back_live_scan`
  - `2026-07-05_114203/` =
    `search_service_sqlite_midquery_failure_restarts_live_scan_without_duplicates`
- Misleading non-evidence to ignore:
  `2026-07-05_113725/` used a pipe-separated `-CaseFilter`, matched no cases, and
  failed only the result-coverage guard.

Closeout/docs updated:

- `Specs/Core/Core_Search.md`
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
- `Specs/Plans/WIP/README.md`
- `Specs/Testing/Testing_TestCoverage.md`
- `Tests/README.md`
- `Tools/Tests/TestInventory.Tests.ps1`

Verification:

- `git diff --check` passed; output contained only existing LF-to-CRLF warnings.
- `Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru` passed 5/0.
- `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors; log:
  `.build/logs/msbuild-20260705_114612_121.log`.
- Final focused GR-T15 rerun passed against that build; archive:
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_114819/`.

### Current Dirty Tree

`git status --short` at this save point:

```text
 M Common/LocalSearchIndexCore.cpp
 M Common/LocalSearchIndexCore.h
 M Common/SearchServiceBroker.cpp
 M Common/SearchServiceBroker.h
 M Common/SqliteIndexStore.cpp
 M Common/SqliteIndexStore.h
 M Plugins/FileSystem/FileSystem.Internal.h
 M Plugins/FileSystem/FileSystem.Search.cpp
 M Plugins/FileSystem/FileSystem.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.Core.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.FakeBackend.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.h
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.cpp
 M RedSalamanderSearchService/Main.cpp
 M Specs/Core/Core_Search.md
 M Specs/FileSystem/FileSystem_Mtp.md
 M Specs/Plans/WIP/Operation_Clearwater_TwoDayMasterReviewRemediation_2026-06-28.md
 M Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
 M Specs/Plans/WIP/README.md
 M Specs/Testing/Testing_TestCoverage.md
 M Tests/README.md
 M Tools/Tests/TestInventory.Tests.ps1
?? Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md
```

Do not revert this dirty tree. It includes multiple completed Granite/Clearwater slices
plus the continuation baton.

### Next Move On Resume

1. Confirm the dirty tree with `git status --short`.
2. Read Track U row `GR-9` and test row `GR-T9` in the Granite plan.
3. Inspect `Common/DxUi/DxUi.Accessibility.cpp` and relevant grid model APIs before
   editing.
4. Add/adjust the DxUi accessibility selftest first so the horizontally scrolled grid
   row-name/cell regression goes RED, then implement the snapshot-builder fix.
5. Keep durable spec updates in the authoritative UI/accessibility spec, update the
   Granite ledger and WIP index, and re-run focused DxUi/related checks plus build.

## Latest Save Point - 2026-07-05 11:40 local

This section supersedes the older save point sections below. Granite is still active and
must stay in WIP. Do not call this goal complete.

Active objective:

> fix all `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
> from high to low until the plan can move to done

Current state: GR-S3/GR-T15 has code and focused RED/GREEN evidence in the
working tree, but adjacent checks, docs, inventory updates, and WIP closeout are still
pending.

### Current Dirty Tree

`git status --short` at this save point:

```text
 M Common/LocalSearchIndexCore.cpp
 M Common/LocalSearchIndexCore.h
 M Common/SearchServiceBroker.cpp
 M Common/SearchServiceBroker.h
 M Common/SqliteIndexStore.cpp
 M Common/SqliteIndexStore.h
 M Plugins/FileSystem/FileSystem.Internal.h
 M Plugins/FileSystem/FileSystem.Search.cpp
 M Plugins/FileSystem/FileSystem.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.Core.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.FakeBackend.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.h
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.cpp
 M RedSalamanderSearchService/Main.cpp
 M Specs/Core/Core_Search.md
 M Specs/FileSystem/FileSystem_Mtp.md
 M Specs/Plans/WIP/Operation_Clearwater_TwoDayMasterReviewRemediation_2026-06-28.md
 M Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
 M Specs/Plans/WIP/README.md
 M Specs/Testing/Testing_TestCoverage.md
 M Tests/README.md
 M Tools/Tests/TestInventory.Tests.ps1
?? Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md
```

Do not revert any of this as "unrelated"; it carries prior completed Granite/Clearwater
slices plus the in-progress GR-S3 slice.

### GR-S3 / GR-T15 Work In Flight

Plan row:

- `GR-S3 | LOW (architecture)`: stale cached `PersistentStoreInfo` can mask external
  SQLite store rotation and rely on the one-shot query retry path.
- `GR-T15`: rotate the store externally mid-session and assert the next query hits the
  refreshed store without the `search.backend.sqlite.retry_query_ms` path.

New selftest added:

- `search_service_sqlite_external_rotation_refreshes_without_retry`
- File: `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`
- Test name also added to `kCompareCaseNames` in
  `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.cpp`.

Focused RED evidence:

- Build before RED: `.build/logs/msbuild-20260705_112906_274.log`
- Command:
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_service_sqlite_external_rotation_refreshes_without_retry -FailFast -TimeoutMultiplier 3`
- Archive:
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_113058/`
- Expected failure:
  `External store rotation should refresh cached PersistentStoreInfo before querying instead of using retry_query_ms.`

Production changes already made:

- `Common/SqliteIndexStore.h/.cpp`
  - Added `StoreInfo::storeGeneration`.
  - Added SQLite meta key `store_generation`.
  - Added `ReadStoreGeneration(...)`.
  - Schema/bootstrap/migration ensure the generation exists.
  - `ReplaceVolume`, `ApplyJournalDelta`, and existing-volume `DeleteVolume` increment
    generation inside their write transaction.
- `Common/LocalSearchIndexCore.h/.cpp`
  - Added `PersistentStoreInfo::storeGeneration`.
  - Added `Repository::GetValidatedCachedPersistentStoreInfoForQuery()`.
  - `Enumerate` and `EnumerateNoWait` validate cached SQLite store generation before
    opening/querying the store.
  - Generation mismatch/read failure refreshes cached store info and emits
    `search.backend.sqlite.store_generation_refreshes`.
  - Stale/missing/invalid/cutover retry is no longer restricted to zero emitted rows.
  - `TrackCandidateAndForward` dedupes emitted path keys across retry before forwarding.

Build after production fix:

- Command: `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
- Log: `.build/logs/msbuild-20260705_113444_025.log`
- Result observed in terminal: 0 warnings, 0 errors.

Focused GREEN evidence:

- Command:
  `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_service_sqlite_external_rotation_refreshes_without_retry -FailFast -TimeoutMultiplier 3`
- Archive:
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_113650/`
- Result: 1 passed, 0 failed.
- Note: `perf/perf_metrics.jsonl` still contains one `search.backend.sqlite.retry_query_ms`
  from the priming/cutover-blocked query. The selftest records a file-size offset after
  external rotation and asserts no retry metric appears after that offset. That is the
  intended contract.

Misleading adjacent-check attempt:

- Archive:
  `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_113725/`
- Command used a pipe-separated `-CaseFilter`.
- Result was not a real adjacent test run:
  `selftest_result_coverage` failed with `no expected cases matched the requested filter`.
- Treat all adjacent GR-S3 checks as still pending. Run them one by one.

### Next Move On Resume

1. Confirm the dirty tree with `git status --short`.
2. Run adjacent cases one by one, redirecting output to logs to avoid huge terminal output:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter local_index_core_sqlite_cutover_blocked_by_pending_legacy_import -FailFast -TimeoutMultiplier 3
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_service_sqlite_cold_start_stale_root_refreshes_before_query -FailFast -TimeoutMultiplier 3
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_service_sqlite_invalid_store_falls_back_live_scan -FailFast -TimeoutMultiplier 3
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_service_sqlite_query_failure_falls_back_live_scan -FailFast -TimeoutMultiplier 3
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_service_sqlite_midquery_failure_restarts_live_scan_without_duplicates -FailFast -TimeoutMultiplier 3
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter sqlite_index_store_load_and_apply_journal_delta -FailFast -TimeoutMultiplier 3
```

Potential expected follow-up:

- `local_index_core_sqlite_cutover_blocked_by_pending_legacy_import` may have an outdated
  expectation that cached-ready state continues using SQLite until explicit invalidation.
  Under GR-S3, generation validation should detect external pending-legacy changes before
  query/open, refresh cache, and block SQLite cutover. If it fails for that reason, update
  the test expectation rather than undoing the generation validation.

After adjacent checks pass:

1. Re-run the focused GR-T15 case if any code/test changed.
2. Re-run the Debug x64 build.
3. Update durable docs:
   - `Specs/Core/Core_Search.md`: SQLite `store_generation`, generation increments,
     cached store-info validation before query/open, retry/dedupe fallback contract, and
     GR-T15 coverage.
   - `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`:
     mark GR-S3 and GR-T15 done with RED/GREEN evidence.
   - `Specs/Plans/WIP/README.md`: advance Granite next action after GR-S3; inspect the
     WIP ledger before choosing the next high-to-low item.
   - `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`, and
     `Tools/Tests/TestInventory.Tests.ps1`: bump Compare selftest counts for the new case
     if not already done. Expected direction from adding one Compare case: Compare total
     247 -> 248, static `RunCase` count 239 -> 240, SearchService family 38 -> 39.
4. Run:

```powershell
git diff --check
Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
```

Only after all remaining Granite rows are done should the plan move to `Specs/Plans/Done/`.

Purpose: resume the active Granite remediation without relying on chat history. Treat the
current worktree as authoritative and inspect it before editing.

## Save Point

Refreshed on 2026-07-05 after closing GR-P2, GR-7, GR-11, GR-13, GR-14, and
Clearwater-owned CW-10 in the working tree. Granite is not complete and must stay in WIP.

Next move: continue Granite Track S with GR-S3/GR-T15.

Resume from this file plus `git status --short`; do not rely on terminal scrollback.

## Required Startup Context

- Read root `AGENTS.md` before touching code.
- Follow the repo skills referenced by `AGENTS.md` as needed, especially `cpp-build`,
  `cpp-modern-style`, `wil-raii`, `async-threading`, `error-handling`,
  `compiler-warnings`, and `perf-validation`.
- Keep TDD discipline: add or update the selftest that fails first, then implement.
- Use `apply_patch` for manual edits.
- Do not revert unrelated dirty work. The dirty tree intentionally carries multiple
  completed Granite slices that have not been committed yet.

## Current Dirty Tree

`git status --short` at this save point:

```text
 M Common/LocalSearchIndexCore.h
 M Common/SearchServiceBroker.cpp
 M Common/SearchServiceBroker.h
 M Plugins/FileSystem/FileSystem.Internal.h
 M Plugins/FileSystem/FileSystem.Search.cpp
 M Plugins/FileSystem/FileSystem.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.Core.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.FakeBackend.cpp
 M Plugins/FileSystemMtp/FileSystemMtp.h
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp
 M RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.cpp
 M RedSalamanderSearchService/Main.cpp
 M Specs/Core/Core_Search.md
 M Specs/FileSystem/FileSystem_Mtp.md
 M Specs/Plans/WIP/Operation_Clearwater_TwoDayMasterReviewRemediation_2026-06-28.md
 M Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md
 M Specs/Plans/WIP/README.md
 M Specs/Testing/Testing_TestCoverage.md
 M Tests/README.md
 M Tools/Tests/TestInventory.Tests.ps1
?? Specs/Plans/WIP/Operation_Granite_ContinuationBaton_2026-07-05.md
```

The untracked baton file is intentional. Do not delete it during cleanup.

## WIP Ledger State

- `Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`
  has CW-10, GR-13/GR-T13, and GR-14/GR-T14 marked DONE 2026-07-05.
- `Operation_Clearwater_TwoDayMasterReviewRemediation_2026-06-28.md` has
  CW-10/CW-T10 marked DONE 2026-07-05.
- `Specs/Plans/WIP/README.md` says Granite's next action is GR-S3/GR-T15.

## Verification Snapshot

Latest relevant evidence at this save point:

- CW-10 final code build:
  - `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  - log: `.build\logs\msbuild-20260705_111223_980.log`
  - result: 0 warnings, 0 errors.
- CW-10 intended RED:
  - archive: `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_111008/`
  - failure: `search_service_transient_authorization_failure_is_incomplete_not_cached`
    reported `warnings=0x00000000`.
- CW-10 GREEN:
  - archive: `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_111440/`
  - result: 1 passed, 0 failed.
- CW-10 adjacent SearchService checks passed:
  - `2026-07-05_111529/` = `search_service_filters_cached_descendants_denied_to_client`
  - `2026-07-05_111546/` = `search_service_candidate_impersonation_failure_is_incomplete_warning`
  - `2026-07-05_111603/` = `search_service_query_reports_live_progress`
  - `2026-07-05_111619/` = `search_service_status_and_query_roundtrip`
- GR-14 final code build:
  - `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  - log: `.build\logs\msbuild-20260705_105305_813.log`
  - result: 0 warnings, 0 errors.
- GR-14 intended RED:
  - archive: `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_104801/`
  - failure: `search_service_candidate_impersonation_failure_is_incomplete_warning`
    failed the whole query with `hr=0x80070558`.
- GR-14 GREEN:
  - archive: `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_105520/`
  - result: 1 passed, 0 failed.
- GR-14 adjacent SearchService checks passed:
  - `2026-07-05_105549/` = `search_service_query_reports_live_progress`
  - `2026-07-05_105550/` = `search_service_filters_cached_descendants_denied_to_client`
  - `2026-07-05_105551/` = `search_service_status_and_query_roundtrip`
  - `2026-07-05_105551_001/` = `search_service_multi_client_and_rebuild_control`
  - `2026-07-05_105552/` = `search_service_rebuild_deleted_root_purges_index`
- GR-13 final PluginContract validation:
  - build log: `.build\logs\msbuild-20260705_105915_861.log`
  - output: `.build\logs\plugin_contract_gr13_gr14_final.out.txt`
  - result: `FileSystem.dll` debug selftests passed with `passed=42, failed=0,
    hr=0x00000000`; `PluginContractTests passed.`
- Test inventory:
  - `Invoke-Pester .\Tools\Tests\TestInventory.Tests.ps1 -PassThru`
  - result after the CW-10 count bump: 5 passed, 0 failed.

Earlier closed slices still have useful archived evidence:

- GR-P2 streaming MTP reader:
  - `Specs/TestRuns/4cb089111a23/Mtp/2026-07-05_090500_gr_p2_streaming_reader/`
- GR-7 completed-swap overwrite journal:
  - `Specs/TestRuns/4cb089111a23/Mtp/2026-07-05_094500_gr_7_completed_swap_journal/`
- GR-11 deleted-root rebuild purge:
  - RED log: `.build\logs\search_service_rebuild_deleted_root_gr11_red.txt`
  - GREEN archive: `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-07-05_095517/`

## Closed Work Summary

### GR-P2

MTP `CreateFileReader` is now streaming-first. Open/GetSize do not download file bytes;
byte transfer accounting is emitted on actual reads. Durable contract is in
`Specs/FileSystem/FileSystem_Mtp.md`.

### GR-7

Completed temp-swap overwrite journals no longer replay forever when the final
destination exists and the temp is missing. The replay path now infers completed swaps,
clears matching journals, has a stale-journal quarantine fallback, and caches absent
journals per identity. Durable contract is in `Specs/FileSystem/FileSystem_Mtp.md`.

### GR-11

SearchService rebuild authorization now permits deleted indexed roots by ACL-checking the
deepest existing ancestor while keeping the requested deleted root as the purge target.
Exact-root authorization is still required for queries. Durable contract is in
`Specs/Core/Core_Search.md`.

### GR-13

`Plugins/FileSystem/FileSystem.Search.cpp::IsServiceFallbackCandidate(...)` now treats
server-side root-validation errors as local fallback candidates:

- `HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND)`
- `HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME)`
- `E_INVALIDARG`

The FileSystem debug selftest coverage is wired through PluginContractTests.

### GR-14

Candidate authorization impersonation failures no longer abort the whole SearchService
query. A failed per-candidate authorization now skips that candidate, ORs
`FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED` into query stats/runtime warning flags,
and lets the query complete with incomplete-warning semantics.

Important implementation details:

- `Common/LocalSearchIndexCore.h`
  - `QueryStats::warningFlags` carries server warning bits.
- `Common/SearchServiceBroker.cpp`
  - debug/ASan-only foreground service flag:
    `--test-fail-client-auth-impersonation-once`
  - `ServerCandidateBatchState` carries mutable stats and accumulated warning flags.
  - `StreamServerCandidate(...)` returns `LocalSearchIndexCore::kSkipCandidateHr` for
    candidate authorization failure.
  - query completion runtime payload propagates warning flags to the client.
- `RedSalamanderSearchService/Main.cpp`
  - parses the debug/ASan-only test fault flag.
- `Plugins/FileSystem/FileSystem.Search.cpp`
  - `ApplyIndexedQueryStats(...)` merges `stats.warningFlags` into plugin runtime
    warnings.
- Selftest added:
  - `search_service_candidate_impersonation_failure_is_incomplete_warning`

### CW-10

`Common/SearchServiceBroker.cpp::CheckClientCanListDirectory(...)` no longer caches
`false` for transient directory authorization open failures. It now caches `false` only
for durable negative results:

- `ERROR_ACCESS_DENIED`
- `ERROR_FILE_NOT_FOUND`
- `ERROR_PATH_NOT_FOUND`

Transient failures such as `ERROR_SHARING_VIOLATION` are logged, returned through the
per-candidate authorization failure path, skipped with
`FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED`, and not inserted into
`clientDirectoryAccessCache`.

Selftest added:

- `search_service_transient_authorization_failure_is_incomplete_not_cached`

Durable contract is in `Specs/Core/Core_Search.md`. Clearwater and Granite WIP rows are
updated.

## Next Move: GR-S3 / GR-T15

Owner: Granite Track S.

Bug shape from Granite:

- `LocalSearchIndexCore.cpp` currently retries a direct SQLite query only after a stale,
  missing, invalid, or cutover-blocked store failure when zero rows have been emitted.
- If external store rotation is detected after rows have already been emitted, the query
  falls to slow soft fallback instead of proactively validating refreshed store metadata
  and retrying with dedup.
- `InvalidateCachedPersistentStoreInfo(...)` exists but does not cover external rotation.

Expected GR-T15 RED-first coverage:

- Rotate the persistent store externally mid-session.
- Assert the next query uses the refreshed store without relying on the
  `search.backend.sqlite.retry_query_ms` path.
- The expected current RED is that the retry path is taken.

Useful starting points:

- `Common/LocalSearchIndexCore.cpp`
  - direct SQLite store-info cache and query retry handling around the GR-S3 anchor.
- `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`
  - existing direct-SQLite and SearchService SQLite store rotation/failure tests.
- `Specs/Core/Core_Search.md`
  - update if the store generation/currentness contract changes.

## Resume Checklist

1. Run `git status --short` and confirm the dirty-file list has not been unexpectedly
   changed.
2. Read the GR-S3/GR-T15 rows in Granite and inspect current
   `LocalSearchIndexCore.cpp` anchors before editing.
3. Start GR-T15 with a failing selftest before production changes.
4. After GR-S3 is green, update `Specs/Core/Core_Search.md`, the Granite ledger,
   `Specs/Plans/WIP/README.md`, test inventory docs if counts change, and this baton.

## 2026-07-05 13:26 Pause Point - GR-A2

This is the current active Granite item and supersedes the stale "Next Move:
GR-S3 / GR-T15" section above. `Specs/Plans/WIP/README.md` already points Granite
to Track A / GR-A2.

Goal:

- Fix GR-A2: remove the duplicate host-side WPD MTP picker stack from
  `ConnectionManagerWindow` and route device/storage picker browse through the MTP
  plugin contract instead.
- The host should call a factory/plugin browse path keyed by plugin id.
- The MTP plugin should own both live WPD browse and fake-backend browse for
  deterministic selftests.
- Delete the old host fixture globals and debug fixture APIs after the plugin path
  is wired.

Current implementation state:

- RED coverage has been migrated in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`.
- New plugin-contract surface has been started in
  `Common/PlugInterfaces/Factory.h`:
  `FactoryConnectionBrowseRequest`, `FactoryConnectionBrowseResult`, and optional
  `RedSalamanderBrowseConnectionTargets(...)`.
- `RedSalamander/FileSystemPluginManager.h` now declares
  `ConnectionBrowseDevice`, `ConnectionBrowseStorage`,
  `EnumerateConnectionBrowseDevices(...)`, and
  `EnumerateConnectionBrowseStorages(...)`.
- `RedSalamander/FileSystemPluginManager.cpp` has JSON parsing helpers for the
  plugin browse response. The earlier concern about accidental `RETURN_IF_FAILED`
  usage has already been checked; no `RETURN_IF_FAILED` remains in the new parser
  block.
- Production manager methods, MTP plugin export, MTP fake-backend browse export,
  live WPD browse helpers inside the plugin, and the `ConnectionManagerWindow`
  host cleanup are not implemented yet.

RED evidence already captured:

- Build before implementation:
  `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64`
  passed with 0 warnings / 0 errors.
- Build log: `.build\logs\msbuild-20260705_132106_929.log`.
- Focused RED command:
  `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_connection_manager_window_mtp_picker_populates_profile -FailFast -TimeoutMultiplier 3`.
- Expected RED failure:
  `Failed to seed the MTP plugin fake picker browse backend. hr=0x8007007F`
  (`ERROR_PROC_NOT_FOUND`). This proves the test now requires a plugin debug
  export and no longer uses the host fixture.

Dirty files that are part of this GR-A2 slice so far:

- `Common/PlugInterfaces/Factory.h`
- `RedSalamander/FileSystemPluginManager.h`
- `RedSalamander/FileSystemPluginManager.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`

Other dirty files already existed in the broader Granite/WarpDrive workspace. Do not
revert them.

Resume from here:

1. Inspect the partial manager implementation:
   `rg -n "RETURN_IF_FAILED|Utf16FromUtf8|ParseConnectionBrowse|EnumerateConnectionBrowse|BrowseConnectionTargets|ConnectionBrowse" RedSalamander\FileSystemPluginManager.cpp RedSalamander\FileSystemPluginManager.h Common\PlugInterfaces\Factory.h RedSalamander\SelfTest\Commands\Commands.SelfTest.Connections.cpp`.
2. Add the `FileSystemPluginManager.cpp` production methods that load or reuse the
   plugin DLL, discover `RedSalamanderBrowseConnectionTargets`, call it with
   `__uuidof(IFileSystem)`, copy the `CoTaskMemAlloc` JSON result, and parse
   devices/storages.
3. Add MTP plugin browse support:
   - shared UTF-8 JSON escaping/copy-to-CoTaskMem helper;
   - live WPD-backed device/storage browse inside the MTP plugin, not the host;
   - fake-backend browse path;
   - `_DEBUG` export `RedSalamanderMtpSetPickerFakeBackendForSelfTest(...)`;
   - optional factory export `RedSalamanderBrowseConnectionTargets(...)`.
4. Update `ConnectionManagerWindow` to call
   `FileSystemPluginManager::EnumerateConnectionBrowseDevices/Storages`, carry the
   plugin id in the worker context, replace local no-case helpers with
   `OrdinalString::EqualsNoCase`, and delete:
   - host `PortableDevice` includes;
   - host WPD helper functions;
   - host MTP fixture globals;
   - `DebugSetConnectionManagerMtpPickerFixture` /
     `DebugClearConnectionManagerMtpPickerFixture` declarations and definitions.
5. Verification search before building:
   `rg -n "PortableDevice|IPortable|WPD_|MtpPickerComInitialization|EnumerateMtpPicker|g_mtpPickerFixture|DebugSetConnectionManagerMtpPickerFixture|DebugClearConnectionManagerMtpPickerFixture|StableMtpPickerHash|SanitizeMtpPathComponent|EqualsNoCaseLocal" RedSalamander\ConnectionManagerWindow.cpp RedSalamander\ConnectionManagerWindow.h`.
6. Rebuild and run the focused test, then adjacent Commands/MTP tests, inventory,
   and docs/spec closeout.

Docs still needed after green:

- `Specs/FileSystem/FileSystem_Mtp.md`: durable contract that the Connection
  Manager MTP picker browses through the plugin factory/export and that WPD/fake
  enumeration stays inside the MTP plugin.
- `Specs/Plans/WIP/Operation_Granite_FiveDayMasterBranchAndWorktreeReviewRemediation_2026-07-02.md`:
  mark GR-A2 done with evidence only after verification.
- `Specs/Plans/WIP/README.md`: advance Granite to the next live row after GR-A2.
- `Tests/README.md` / `Specs/Testing/Testing_TestCoverage.md`: update the existing
  `cmd_connection_manager_window_mtp_picker_populates_profile` description if needed;
  no new test count is expected unless additional cases are added.
