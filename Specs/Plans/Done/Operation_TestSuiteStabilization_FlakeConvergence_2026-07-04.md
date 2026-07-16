# Operation Test-Suite Stabilization — Flake Convergence

**Created:** 2026-07-04
**Status:** Done (closed 2026-07-10 22:47 +02:00). Final proof root was `C:\RSPerf`.
The accepted `Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2` run exited 0 with build enabled,
1173 total cases, 1121 passed, 0 failed, 52 expected skips, clean disk audit with 0 issues, and
0 flaky/regression/isolation-suspect/unclassified-failure classifications. Final evidence is
archived at
`Specs\TestRuns\8fa8954fe9ff\Full\2026-07-10_2247_full_green\last_run`.
**Goal:** Make every test iteration deterministic and self-classifying, so the suite stops
spending days to converge. A flaky or broad-only failure must be identified, root-caused, and fixed
or replaced with a deterministic equivalent across *every* gated harness; it must not be made
non-blocking by default. Every case must be proven isolated.
**Related:** [[Operation_CommandsSelfTestInputIsolation_2026-06-24]] (prior Commands input-isolation work),
`Tests/README.md`, `Specs/Testing/Testing_SelfTests.md`, `.github/workflows/ci.yml`,
`Tools/Run-AllTests.ps1`, `Specs/Plans/WIP/Operation_PerfMeasurementContract_2026-07-06.md`.

## Final closeout checklist - 2026-07-10 22:47 +02:00

- [x] Final root was the required short NTFS TestSandbox root `C:\RSPerf`.
- [x] Final Full gate passed with build enabled:
  `Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2`, run id
  `20260710T200626Z-43532-7869ee949a30441c8d586e476fd428d3`, 1173 total, 1121 passed,
  0 failed, 52 expected skips.
- [x] Final disk proof was clean: `test_sandbox_audit.is_clean=true`, `issue_count=0`.
- [x] Failure classification stayed clean: 0 flaky, 0 regression, 0 isolation-suspect,
  0 unclassified failures.
- [x] Final accepted evidence was archived before volatile root cleanup:
  `Specs\TestRuns\8fa8954fe9ff\Full\2026-07-10_2247_full_green\last_run`.
- [x] Source contracts guard the new fixes: SearchService foreground readiness must use
  process-aware diagnostics, and FileOperations queued-cancel proof must stage active/queued task
  ownership instead of assuming same-tick queueing.
- [x] Durable requirements were merged into authoritative specs:
  `Specs/Testing/Testing_SelfTests.md` and `Specs/Testing/Testing_TestCoverage.md`.
- [x] No flaky test was made non-blocking and no anonymous permanent quarantine was introduced.
- [x] This plan has been moved to `Specs/Plans/Done/`.

## Resume checklist - 2026-07-10 22:05 +02:00 (superseded)

Historical WIP checkpoint kept for traceability. It is superseded by the final closeout checklist
above; do not use it as the current resume point.

Process state at this save point: no `RedSalamander.exe`, `RedSalamanderSearchService.exe`, `MSBuild`,
`cl`, or `link` process was left running after the broad FileOps, Compare, and Tools gates.

### Done

- [x] Final proof root remains `C:\RSPerf`.
- [x] Source contracts passed after the SearchService and FileOps queued-cancel fixes:
  `Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`, 121 passed /
  0 failed.
- [x] Full Debug build passed after the SearchService and FileOps edits:
  `.build\logs\msbuild-20260710_213834_178.log`, 0 warnings / 0 errors.
- [x] RedSalamander Debug rebuild passed after the staged FileOps queued-cancel correction:
  `.build\logs\msbuild-20260710_214341_182.log`, 0 warnings / 0 errors.
- [x] Compare SearchService default-store and seeded default-store cases now start foreground service
  output capture and wait through `WaitForSearchServiceStatusWithProcessDiagnostics(...)`, so early
  service exits include exit code and captured output instead of timing out blind.
- [x] `local_search_service_indexed_name_latency_and_parity` now waits for matching service status
  before its warmup query and appends foreground service exit/output diagnostics to warmup query
  failures.
- [x] Focused Compare proofs passed under clean `C:\RSPerf` before later proof-root cleaning:
  `search_service_sqlite_default_store_uses_build_specific_programdata_root` run
  `20260710T194101Z-72196-73ef7a9dc81f4bc78e148d9285f3e944`, 1 passed / 0 failed;
  `local_search_service_indexed_name_latency_and_parity` run
  `20260710T194107Z-77260-a2498f8d407d403e92557279c61185a9`, 1 passed / 0 failed; both disk audits
  clean.
- [x] FileOperations `Phase5_CancelQueuedTask` no longer waits for a queued state it did not prove.
  The phase starts task A first, waits until A has entered operation and owns the active slot, then
  starts B as queued, cancels B while `IsWaitingInQueue()`, and verifies cancel HRESULTs for A/B/C.
  The first focused rerun exposed the real race (`taskB` entered operation before being queued);
  this staged fix replaces that timing dependency.
- [x] Focused FileOps queued-cancel proof passed under clean `C:\RSPerf` before later proof-root
  cleaning: run `20260710T194554Z-15568-ffd489ab056346e4bf704fe3b092f283`, 3 passed / 0 failed,
  disk audit clean.
- [x] Broad FileOps is green from the current binary under `C:\RSPerf`: run
  `20260710T194612Z-58152-b9b945319fde4af0aa98ab4ee046ba92`, 102 passed / 0 failed / 20 expected
  skips, disk audit clean.
- [x] Broad Compare is green from the current binary under `C:\RSPerf`: run
  `20260710T195933Z-692-456eceeb76ed4481bbeb1dbf4f874987`, 219 passed / 0 failed / 30 expected
  skips, disk audit clean.
- [x] Tools Pester is green from the current tree:
  `Invoke-Pester -Path .\Tools\Tests -PassThru`, 232 passed / 0 failed / 0 skipped.

### Completed after this checkpoint

- [x] Safely clean stale `C:\RSPerf` contents before final Suite Full so the disk audit starts from
  the required root and has no stale focused-run residue.
- [x] Rerun `.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2` under `C:\RSPerf` with build
  enabled.
- [x] If Suite Full exposes any new broad-only blocker, root-cause and fix it directly. Do not make
  flaky tests non-blocking and do not create anonymous permanent quarantine.
- [x] Archive final accepted evidence under `Specs/TestRuns/<current-head>/...` before cleaning
  volatile `C:\RSPerf` evidence.
- [x] Merge durable requirements/lessons into authoritative specs or repo guidance. The Tools Pester
  inventory count is already merged into `Specs/Testing/Testing_TestCoverage.md` and `Tests/README.md`;
  still review whether SearchService process-aware foreground readiness, FileOps queue-staging,
  Commands page-host route, Preferences readiness, Shortcuts live-search message-pump, Compare
  Options global close/reopen, prompt ValuePattern/model settling, Quick Search focus repair, and
  menu-bar stale-pointer lessons need durable testing-spec text.
- [x] Move this plan to `Specs/Plans/Done/`.
- [x] Commit and push the intended stabilization files after final Suite Full proof, unless the
  maintainer asks for another WIP checkpoint first.

## Resume checklist - 2026-07-10 19:49 +02:00 (previous)

Use this section first when resuming. This is a WIP checkpoint only; do not treat it as final proof.
The operation is not done until Suite Full is green under the final root `C:\RSPerf`, final evidence
is archived, durable requirements are merged out of WIP, this plan is moved to
`Specs/Plans/Done/`, and the intended stabilization files are committed and pushed. Do not make flaky
tests non-blocking. Do not create anonymous permanent quarantine.

Process state at this save point: no `RedSalamander.exe`, `RedSalamanderSearchService.exe`, `MSBuild`,
`cl`, or `link` process was left running after preserving the green broad Commands run.

### Done

- [x] Final proof root remains `C:\RSPerf`.
- [x] Source contracts passed after the latest Quick Search and menu-bar fixes:
  `Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`, 121 passed /
  0 failed.
- [x] Current tree built after the latest menu-bar stale-pointer seed:
  `.build\logs\msbuild-20260710_192905_867.log`, 0 warnings / 0 errors.
- [x] Compare Options close/reopen churn is broad-order hardened in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp` by synchronously closing
  and waiting for all Compare Directories option windows before reopening or reusing live UIA.
- [x] Prompt churn cases are broad-order hardened by condition-based `ValuePattern` and debug-model
  waits in `Commands.SelfTest.Settings.cpp`, `Commands.SelfTest.FileOps.cpp`, and
  `Commands.SelfTest.Dialogs.cpp`.
- [x] Preferences category-tree churn is broad-order hardened by focusing the Dx category host before
  synthetic clicks and by focusing the host on `WM_LBUTTONDOWN` in `Preferences.Dialog.cpp`.
- [x] Quick Search no-match reactivation is broad-order hardened by reusing
  `activateQuickSearchDirect(...)`, waiting for the pane focus repair, and verifying the query is
  still active/empty before typing the no-match character. Focused proof passed under `C:\RSPerf`:
  run `20260710T172639Z-26208-9c4420d5f7404b7ca79bd0051db7e15b`, 1 passed / 0 failed, disk audit
  clean.
- [x] Temporary menu-bar hover-switching is broad-order hardened by opening the already-selected
  temporary menu through keyboard activation and seeding the keyboard-open stale-pointer guard before
  hovering the neighboring top-level item. The first attempted keyboard-open patch exposed a real
  focused failure, archived at
  `Specs/TestRuns/SINON/Continuation/2026-07-10_1926_menu_keyboard_open_hover_no_switch_failure/`;
  the fixed focused proof passed under `C:\RSPerf`: run
  `20260710T173114Z-40944-2a7b2dc63a3b4e869bc1f173588f7f57`, 1 passed / 0 failed, disk audit clean.
- [x] The paired Quick Search + menu-bar focused order passed under `C:\RSPerf`: run
  `20260710T173121Z-80128-23ad3f6a4b1447bb92c22984978aebdd`, 2 passed / 0 failed, disk audit
  clean.
- [x] Broad Commands is green under `C:\RSPerf`: run
  `20260710T173129Z-96456-246d4743c7e74a77b76dad5c04de946e`, 786 passed / 0 failed / 2 opt-in
  skipped, disk audit clean. Evidence was archived at
  `Specs/TestRuns/SINON/Continuation/2026-07-10_1931_commands_broad_green_after_quicksearch_menubar_fixes/`.

### Remaining

- [ ] Safely clean stale `C:\RSPerf` contents before final Suite Full so the disk audit starts from
  the required root and has no stale focused-run residue.
- [ ] Rerun `.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2` under `C:\RSPerf` with build
  enabled.
- [ ] If Suite Full exposes any new broad-only blocker, root-cause and fix it directly. Do not make
  flaky tests non-blocking and do not create anonymous permanent quarantine.
- [ ] Archive final accepted evidence under `Specs/TestRuns/<current-head>/...` before cleaning
  volatile `C:\RSPerf` evidence.
- [ ] Merge durable requirements/lessons into authoritative specs or repo guidance. The Tools Pester
  inventory count is already merged into `Specs/Testing/Testing_TestCoverage.md` and `Tests/README.md`;
  still review whether the Compare service-recovery, Commands page-host route, Preferences readiness,
  Shortcuts live-search message-pump, Compare Options global close/reopen, prompt ValuePattern/model
  settling, Quick Search focus repair, and menu-bar stale-pointer lessons need durable testing-spec
  text.
- [ ] Move this plan to `Specs/Plans/Done/`.
- [ ] Commit and push the intended stabilization files after final Suite Full proof, unless the
  maintainer asks for another WIP checkpoint first.

## Resume checklist - 2026-07-10 17:54 +02:00 (previous)

Use this section first when resuming. This is a WIP checkpoint only; do not treat it as final proof.
The operation is not done until Suite Full is green under the final root `C:\RSPerf`, final evidence
is archived, durable requirements are merged out of WIP, this plan is moved to
`Specs/Plans/Done/`, and the intended stabilization files are committed and pushed. Do not make flaky
tests non-blocking. Do not create anonymous permanent quarantine.

Process state at this save point: no `RedSalamander.exe`, `RedSalamanderSearchService.exe`, `MSBuild`,
`cl`, or `link` process was left running after preserving and stopping the broad Commands failure
run.

### Done

- [x] Final proof root remains `C:\RSPerf`.
- [x] Shortcuts live UIA search broad-order hang is fixed in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Shortcuts.cpp` by routing live search
  `ValuePattern` reads, `SelectionPattern` reads, and search `SetValue` calls through
  message-pumped UIA helpers instead of raw same-thread UIA calls.
- [x] The reusable message-pumped UIA helpers are saved in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`:
  `CollectVisibleDescendantSelectionPatternStateWithMessagePump(...)` and
  `SetVisibleDescendantValueWithMessagePump(...)`.
- [x] Source-contract regression coverage is saved in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` to keep the Shortcuts live-search case on the
  message-pumped helpers and prevent raw UIA calls from returning under broad order.
- [x] Source contracts passed after the Shortcuts fix:
  `Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`, 117 passed /
  0 failed.
- [x] Current tree built after the Shortcuts fix:
  `.build\logs\msbuild-20260710_172236_346.log`, 0 warnings / 0 errors.
- [x] Focused Commands proof for `cmd_app_shortcuts_live_dx_search_interaction` passed under
  `C:\RSPerf`: run `20260710T152502Z-81556-19fd17d4f8824147829db3ce4478e017`, 1 passed /
  0 failed, disk audit clean.
- [x] Broad Commands rerun after the Shortcuts fix progressed past
  `cmd_app_shortcuts_live_dx_search_interaction`, proving the Shortcuts hang is no longer the active
  blocker. Evidence was archived at
  `Specs/TestRuns/SINON/Continuation/2026-07-10_1750_commands_broad_five_failures_after_shortcuts_fix/`.
  That run used `C:\RSPerf`, exited after 19m2s with 781 passed / 5 failed / 2 skipped, and had a
  clean disk audit.
- [x] The five broad-run failures all passed when rerun focused, so they remain broad-order/order
  interaction suspects rather than accepted flakes:
  `cmd_compare_directories_options_live_dx_body_interaction`
  (`20260710T154526Z-100960-8f78f64f09094c4ea6630d2c96c6081d`),
  `cmd_pane_navigation_rename_prompt_keeps_navigation_shell_stable`
  (`20260710T154721Z-79340-037b88cd590a43bc8c58bd2292bc26ed`),
  `cmd_pane_navigation_go_to_path_from_other_pane_keeps_navigation_shell_stable`
  (`20260710T154723Z-79340-0e58ae8d327644208ada5e38d9ae2cb0`),
  `cmd_pane_changeCase_prompt_long_run_open_close_stays_stable`
  (`20260710T154725Z-79340-1c8e9b05bca44ede8922539bd824496d`), and
  `cmd_pane_itemProperties_window_long_run_scrolling_stays_bounded`
  (`20260710T154729Z-79340-a244e1f246ea42729e46074536a58545`). The later focused disk-audit counts
  were stale-run accumulation in the same root, so clean `C:\RSPerf` before final proof.
- [x] A reduced Compare Options slice passed after the broad failure, which narrows the first blocker
  to broader prior-state interaction rather than the local Compare Options sequence alone:
  run `20260710T154908Z-57116-3aafd7dd9a8b49229b246a93493ac7ce`, 4 passed / 0 failed, disk audit
  clean.

### Remaining

- [ ] Root-cause and fix the first broad Commands failure:
  `cmd_compare_directories_options_live_dx_body_interaction`. The archived trace shows
  `FolderWindow::Destroy` running while the test is waiting on the Compare Options UIA read, leaving
  the diagnostic snapshot empty. Treat this as a broad-order teardown/state-leak bug, not an accepted
  flake.
- [ ] After the first broad-order blocker is fixed, rerun broad Commands proof under `C:\RSPerf` and
  continue fixing any newly exposed broad-order failures directly.
- [ ] Safely clean stale `C:\RSPerf` contents before final Suite Full so the disk audit starts from
  the required root and has no stale focused-run residue.
- [ ] Rerun `.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2` under `C:\RSPerf` with build
  enabled.
- [ ] If Suite Full exposes any new broad-only blocker, root-cause and fix it directly. Do not make
  flaky tests non-blocking and do not create anonymous permanent quarantine.
- [ ] Archive final accepted evidence under `Specs/TestRuns/<current-head>/...` before cleaning
  volatile `C:\RSPerf` evidence.
- [ ] Merge durable requirements/lessons into authoritative specs or repo guidance. The Tools Pester
  inventory count is already merged into `Specs/Testing/Testing_TestCoverage.md` and `Tests/README.md`;
  still review whether the Compare service-recovery, Commands page-host route, Preferences readiness,
  Shortcuts live-search message-pump, and broad teardown/state-leak lessons need durable testing-spec
  text.
- [ ] Move this plan to `Specs/Plans/Done/`.
- [ ] Commit and push the intended stabilization files as a WIP checkpoint if the maintainer asks for
  immediate publication before final Suite Full proof.

## Resume checklist - 2026-07-10 17:15 +02:00 (previous)

Use this section first when resuming. This is a WIP checkpoint only; do not treat it as final proof.
The operation is not done until Suite Full is green under the final root `C:\RSPerf`, final evidence
is archived, durable requirements are merged out of WIP, this plan is moved to
`Specs/Plans/Done/`, and the intended stabilization files are committed and pushed. Do not make flaky
tests non-blocking. Do not create anonymous permanent quarantine.

Process state at this save point:

```powershell
Get-Process -Id 84304 -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding,MainWindowTitle,Path
```

returned `RedSalamander.exe` PID 84304, responding, with `MainWindowTitle=Shortcuts`, launched from
`.build\x64\Debug\RedSalamander.exe`. The continuation evidence was archived before stopping it:
`Specs/TestRuns/SINON/Continuation/2026-07-10_1715_commands_full_timeout_after_preferences_fixes/`.

### Done

- [x] Final proof root remains `C:\RSPerf`.
- [x] Source contracts passed after the latest Preferences edits:
  `Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`, 116 passed /
  0 failed.
- [x] Current tree built after the latest Preferences edits:
  `.build\logs\msbuild-20260710_164439_792.log`, 0 warnings / 0 errors.
- [x] Focused Commands proof for `cmd_preferences_dialog_category_tree_handles_reverse_keyboard_navigation`
  passed under `C:\RSPerf`: run `20260710T142535Z-90124-b91c47be2db94afb9890a271b0b45f2d`, 1 passed /
  0 failed, disk audit clean.
- [x] Broad Preferences slice passed under `C:\RSPerf`:
  `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2 -CaseFilter
  cmd_preferences_dialog_`, run `20260710T144702Z-80496-d3046b994f3b43938dae7832d8af583f`, 173 passed /
  0 failed / 0 skipped, disk audit clean.
- [x] Preferences Compare Directories round-trip source-contract cleanup removed the obsolete
  always-true Phase 8 field assertions.
- [x] Preferences reverse keyboard-navigation readiness is broad-order hardened: setup waits can
  reselect and refocus the expected category, pre-key checks cannot hide a navigation defect by
  reselecting, readiness requires tree focus, active Dx focus control, a selected item, nonzero tree
  render count, expected category, and expected title.
- [x] `PreferencesDialog::DebugSelectCategory` now drives the DxUi category tree through
  `RequestSelectVisibleItem(...)` and rejects transient plugin-item selection when selecting a
  category.
- [x] Category churn validation no longer fails on a brittle three-render ceiling; it validates
  settled state, one visible current page host, one rendered current-page DxUi host, focus/selection
  retention, no resize failures, and a bounded nonzero repaint delta.
- [x] Full Commands run evidence for the next blocker was archived:
  `Specs/TestRuns/SINON/Continuation/2026-07-10_1715_commands_full_timeout_after_preferences_fixes/`.
  That run used `C:\RSPerf`, timed out after 20 minutes, had 496 passed / 0 failed / 0 skipped in
  `commands_results.json`, and stopped advancing at `cmd_app_shortcuts_live_dx_search_interaction`
  with the Shortcuts dialog still open.
- [x] Stale `RedSalamander.exe` PID 84304 from that timed-out Commands run was stopped after the
  continuation evidence was archived.

### Remaining

- [ ] Root-cause and fix the broad Commands hang at
  `cmd_app_shortcuts_live_dx_search_interaction`; do not quarantine or make it non-blocking.
- [ ] Run focused Commands proof for `cmd_app_shortcuts_live_dx_search_interaction` under
  `C:\RSPerf`.
- [ ] Rerun broad Commands proof under `C:\RSPerf`.
- [ ] Safely clean stale `C:\RSPerf` contents before final Suite Full so the disk audit starts from
  the required root and has no stale focused-run residue.
- [ ] Rerun `.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2` under `C:\RSPerf` with build
  enabled.
- [ ] If Suite Full exposes any new broad-only blocker, root-cause and fix it directly. Do not make
  flaky tests non-blocking and do not create anonymous permanent quarantine.
- [ ] Archive final accepted evidence under `Specs/TestRuns/<current-head>/...` before cleaning
  volatile `C:\RSPerf` evidence.
- [ ] Merge durable requirements/lessons into authoritative specs or repo guidance. The Tools Pester
  inventory count is already merged into `Specs/Testing/Testing_TestCoverage.md` and `Tests/README.md`;
  still review whether the Compare service-recovery, Commands page-host route, Preferences readiness,
  and Shortcuts live-search lessons need durable testing-spec text.
- [ ] Move this plan to `Specs/Plans/Done/`.
- [ ] Commit and push the intended stabilization files as a WIP checkpoint if the maintainer asks for
  immediate publication before final Suite Full proof.

## Resume checklist - 2026-07-10 11:47 +02:00 (previous)

Use this section first when resuming. The operation is not done until Suite Full is green under the
final root `C:\RSPerf`, final evidence is archived, durable requirements are merged out of WIP, this
plan is moved to `Specs/Plans/Done/`, and the intended stabilization files are committed. Do not make
flaky tests non-blocking. Do not create anonymous permanent quarantine.

Process state at this save point:

```powershell
Get-Process RedSalamander,RedSalamanderSearchService,MSBuild,cl,link -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding
```

returned no matching processes after stopping the stale pre-patch Suite Full parent process
`powershell.exe` PID 86572 and child `RedSalamander.exe` PID 82476. That stale run was not final proof:
it used the binary from before the latest Commands page-host fix, had already failed
`cmd_preferences_dialog_scroll_host_preserves_retained_page_state`, and then stopped advancing in
FileOps `Phase11_BridgePipelineDummyToDummyPerf`.

### Done

- [x] Final proof root remains `C:\RSPerf`.
- [x] Tools Pester inventory drift is fixed and documented at 223 Tools cases in
  `Tools/Tests/TestInventory.Tests.ps1`, `Specs/Testing/Testing_TestCoverage.md`, and
  `Tests/README.md`. Focused proof passed: `Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1
  -PassThru`, 5 passed / 0 failed. Full Tools proof passed without build contention:
  `Invoke-Pester -Path .\Tools\Tests -PassThru`, 223 passed / 0 failed.
- [x] Compare slow-partial service recovery is fixed in
  `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`
  by replacing the single post-timeout status sample with condition-based
  `WaitForSearchServiceStatus(...)`. Focused proof passed under `C:\RSPerf`:
  `20260710T085323Z-83860-37d26b385dab4494b6a9ac3589892711`, 1 passed / 0 failed.
- [x] Compare live compact-request ordering is fixed in
  `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`:
  the explicit compact test now prepares a fragmented SQLite store with reclaimable free pages below
  the automatic idle-maintenance threshold and preflights that the foreground service is live,
  SQLite-backed, and not queued/running maintenance before issuing `--request-compact`.
- [x] Compare proof under `C:\RSPerf` is green after the latest Compare fixes: focused compact-request
  run `20260710T090301Z-90068-b0c7ca5380fb40fb98e62d8391245f32`, 1 passed / 0 failed; full
  CompareDirectories run `20260710T090312Z-92912-87274c73db35490e9bf3ca7940830f86`, 219 passed /
  0 failed / 30 skipped, disk audit clean.
- [x] Source-contract guards cover the slow-partial condition wait and live compact-request preflight
  in `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`. Latest source-contract proof passed:
  `Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`, 112 passed /
  0 failed.
- [x] Commands Preferences page-host route hardening is saved in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp`:
  `cmd_preferences_dialog_scroll_host_preserves_retained_page_state` now sends the synthetic wheel to
  an unambiguous interior point of the page host, asserts the hit-test target is the page host or one
  of its children before routing, and emits actual category/title/page-scroll/route diagnostics if the
  retained page-state assertion fails. This is a fix/diagnostic hardening path, not quarantine.
- [x] Latest C++ build before the Commands page-host route patch passed:
  `.build\logs\msbuild-20260710_110046_781.log`, 0 warnings / 0 errors.
- [x] Stale Suite Full run `20260710T091013Z-86572-be90d4c2943d4e679672dd54eae1ad6d` is explicitly
  rejected as final evidence. It ran the old binary, Compare completed green, Commands failed
  `cmd_preferences_dialog_scroll_host_preserves_retained_page_state`, and FileOps stopped advancing at
  `Phase11_BridgePipelineDummyToDummyPerf` before the run was stopped.

### Remaining

- [ ] Build the current tree after the Commands page-host route patch.
- [ ] Run focused Commands proof under `C:\RSPerf`:
  `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2 -CaseFilter
  cmd_preferences_dialog_scroll_host_preserves_retained_page_state`. Prefer `-SelfTestRepeat 3` after
  the first pass if time allows.
- [ ] If the focused Commands proof still fails, use the richer diagnostics to root-cause and fix the
  page-host/category routing defect directly. Do not make the test non-blocking.
- [ ] Rerun broad Commands proof, then safely clean stale `C:\RSPerf` contents before final Suite Full
  so the disk audit starts from the required root and has no stale focused-run residue.
- [ ] Rerun `.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2` under `C:\RSPerf` with build
  enabled.
- [ ] If Suite Full exposes any new broad-only blocker, root-cause and fix it directly. Do not make
  flaky tests non-blocking and do not create anonymous permanent quarantine.
- [ ] Archive final accepted evidence under `Specs/TestRuns/<current-head>/...` before cleaning
  volatile `C:\RSPerf` evidence.
- [ ] Merge durable requirements/lessons into authoritative specs or repo guidance. The Tools Pester
  inventory count is already merged into `Specs/Testing/Testing_TestCoverage.md` and `Tests/README.md`;
  still review whether the Compare service-recovery and Commands page-host route lessons need durable
  testing-spec text.
- [ ] Move this plan to `Specs/Plans/Done/`.
- [ ] Commit and push the intended stabilization files as a WIP checkpoint if the maintainer asks for
  immediate publication before final Suite Full proof.

## Resume checklist - 2026-07-10 11:08 +02:00 (previous)

Use this section first when resuming. The operation is not done until Suite Full is green under the
final root `C:\RSPerf`, final evidence is archived, durable requirements are merged out of WIP, this
plan is moved to `Specs/Plans/Done/`, and the intended stabilization files are committed. Do not make
flaky tests non-blocking. Do not create anonymous permanent quarantine.

Process state at this save point:

```powershell
Get-Process RedSalamander,RedSalamanderSearchService,MSBuild,cl,link -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding
```

returned no matching processes. No build, compiler, linker, test-runner, application, or search-service
process was active.

### Done

- [x] Final proof root remains `C:\RSPerf`.
- [x] The 2026-07-10 Suite Full repro under `C:\RSPerf`
  (`20260710T075954Z-78660-c468ce7d36cc4c04a97f7529317bfd11`) had a clean disk audit and proved the
  previous Commands menu cursor-isolation and Compare delete-burst maintenance blockers fixed, but
  exposed two remaining blockers: `search_service_slow_partial_client_does_not_block_next_client` and
  aggregate `ToolsPesterTests`.
- [x] Tools Pester root cause is fixed in `Tools/Tests/TestInventory.Tests.ps1`,
  `Specs/Testing/Testing_TestCoverage.md`, and `Tests/README.md`: the source-derived Tools Pester
  inventory is now 223 cases after the two new source-contract guards. Focused inventory proof passed:
  `Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru`, 5 passed / 0 failed. Full
  tooling proof passed when run without a concurrent build:
  `Invoke-Pester -Path .\Tools\Tests -PassThru`, 223 passed / 0 failed. A prior 222/1 Pester result
  was validation self-contention from running full Tools Pester in parallel with `.\build.ps1`; its
  targeted deployment test hit `.obj` permission-denied errors while the main build owned the same
  intermediate files.
- [x] Compare slow-partial service recovery is fixed in
  `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`:
  after the intentionally partial frame times out, the case uses condition-based
  `WaitForSearchServiceStatus(...)` instead of a single `GetStatus` sample during the pipe-recreation
  window. Focused proof passed under `C:\RSPerf`:
  `20260710T085323Z-83860-37d26b385dab4494b6a9ac3589892711`, 1 passed / 0 failed.
- [x] Compare live compact-request ordering is fixed in
  `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`:
  `search_service_compact_request_roundtrip` now prepares a fragmented store that retains reclaimable
  free pages but stays below the automatic idle-maintenance threshold, then preflights that the
  foreground service is SQLite-backed, live, and not queued/running maintenance before launching the
  explicit `--request-compact` process. This removes the race where automatic idle maintenance could
  consume the pipe window and make the explicit request fail with `0x80070002`.
- [x] Source-contract guards cover the slow-partial condition wait and live compact-request preflight
  in `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`. Latest source-contract proof passed:
  `Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`, 112 passed /
  0 failed.
- [x] Latest C++ build after the Compare and inventory fixes passed:
  `.build\logs\msbuild-20260710_110046_781.log`, 0 warnings / 0 errors.
- [x] Compare proof under `C:\RSPerf` is green after the latest fixes: focused compact-request run
  `20260710T090301Z-90068-b0c7ca5380fb40fb98e62d8391245f32`, 1 passed / 0 failed; full
  CompareDirectories run `20260710T090312Z-92912-87274c73db35490e9bf3ca7940830f86`, 219 passed /
  0 failed / 30 skipped, disk audit clean.

### Remaining

- [ ] Safely clean stale `C:\RSPerf` contents before final proof so the disk audit starts from the
  required root and has no stale focused-run residue.
- [ ] Rerun `.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2` under `C:\RSPerf` with build
  enabled.
- [ ] If Suite Full exposes any new broad-only blocker, root-cause and fix it directly. Do not make
  flaky tests non-blocking and do not create anonymous permanent quarantine.
- [ ] Archive final accepted evidence under `Specs/TestRuns/<current-head>/...` before cleaning
  volatile `C:\RSPerf` evidence.
- [ ] Merge durable requirements/lessons into authoritative specs or repo guidance. The Tools Pester
  inventory count is already merged into `Specs/Testing/Testing_TestCoverage.md` and `Tests/README.md`;
  still review whether the Compare service-recovery lessons need durable testing-spec text.
- [ ] Move this plan to `Specs/Plans/Done/`.
- [ ] Commit only the intended stabilization files.

## Resume checklist - 2026-07-10 09:58 +02:00 (previous)

Use this section first when resuming. The operation is not done until Suite Full is green under the
final root `C:\RSPerf`, final evidence is archived, durable requirements are merged out of WIP, this
plan is moved to `Specs/Plans/Done/`, and the intended stabilization files are committed. Do not make
flaky tests non-blocking. Do not create anonymous permanent quarantine.

Process state at this save point:

```powershell
Get-Process RedSalamander,RedSalamanderSearchService,MSBuild,cl,link -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding
```

returned no matching processes. No build, compiler, linker, test-runner, application, or search-service
process was active.

### Done

- [x] Final proof root remains `C:\RSPerf`.
- [x] The 2026-07-10 Suite Full repro under `C:\RSPerf` had a clean disk audit but failed three cases:
  `search_service_sqlite_delete_burst_maintenance_preserves_query_parity`,
  `cmd_app_menuBar_mouse_open_keeps_popup_selection_clear`, and
  `cmd_app_menuBar_submenu_placement_matches_spec`.
- [x] Commands menu popup root cause is fixed in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`: menu selftests now park the
  real OS cursor at a neutral main-window client point before DxUi popup creation, while synthetic
  menu-bar input and stale-hover suppression continue to use the intended click/suppression point.
  This prevents broad-order cursor residue from seeding popup `hoveredIndex`.
- [x] Commands menu cursor isolation is source-contract guarded in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` by
  `parks real cursor away from DxUi popup initial hover in menu selftests`.
- [x] Compare delete-burst root cause is fixed in
  `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`:
  the test now accepts either the queued pre-compaction state or an already completed first idle
  compaction window, then continues proving bounded maintenance progress and query parity. This fixes
  the legitimate race where foreground-service readiness had already performed one status request and
  the idle grace expired before the test's second status observation.
- [x] Compare delete-burst maintenance observation is source-contract guarded in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` by
  `lets delete-burst maintenance tests accept an already completed first idle window`.
- [x] Latest source contracts passed:
  `Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`,
  112 passed / 0 failed.
- [x] Latest C++ build after the Commands and Compare fixes passed:
  `.build\logs\msbuild-20260710_095530_806.log`, 0 warnings / 0 errors.
- [x] Commands proof under `C:\RSPerf` is green:
  `cmd_app_menuBar_` run `20260710T075039Z-86120-24171c433e8549a193e8e30bf6c7b6d0`,
  17 passed / 0 failed; `cmd_app_` run
  `20260710T075053Z-79088-2502139d1752427aab21ccc0d5b5342e`, 83 passed / 0 failed.
- [x] Compare proof under `C:\RSPerf` is green:
  focused delete-burst run `20260710T075743Z-69360-aacf723cdd1c4bdf8af070709ca44e4a`,
  1 passed / 0 failed; SQLite prefix run
  `20260710T075757Z-24224-2c554e7531064784b3583863f7a08a32`, 15 passed / 0 failed / 3 skipped.

### Remaining

- [ ] Safely clean stale `C:\RSPerf` contents before final proof so the disk audit starts from the
  required root and has no stale focused-run residue.
- [ ] Rerun `.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2` under `C:\RSPerf` with build
  enabled.
- [ ] If Suite Full exposes any new broad-only blocker, root-cause and fix it directly. Do not make
  flaky tests non-blocking and do not create anonymous permanent quarantine.
- [ ] Archive final accepted evidence under `Specs/TestRuns/<current-head>/...` before cleaning
  volatile `C:\RSPerf` evidence.
- [ ] Merge durable requirements/lessons into authoritative specs or repo guidance.
- [ ] Move this plan to `Specs/Plans/Done/`.
- [ ] Commit only the intended stabilization files.

## Resume checklist - 2026-07-08 22:54 +02:00 (previous)

Use this section first when resuming. The operation is not done until Suite Full is green under the
final root `C:\RSPerf`, final evidence is archived, durable requirements are merged out of WIP, this
plan is moved to `Specs/Plans/Done/`, and the intended stabilization files are committed. Do not make
flaky tests non-blocking. Do not create anonymous permanent quarantine.

Process state at this save point:

```powershell
Get-Process RedSalamander,RedSalamanderSearchService,MSBuild,cl,link -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding
```

returned only the user's normal RedSalamander process:

```text
Id=17284 ProcessName=RedSalamander CPU=220.36 StartTime=2026-07-08 17:23:54 Responding=True
```

Leave PID 17284 alone. No build, compiler, linker, test-runner, or search-service process was active.

### Done

- [x] Final proof root remains `C:\RSPerf`.
- [x] SettingsStore TestSandbox settings-path redirect is kept, but guarded behind
  `#ifdef ENABLE_TESTS` in `Common/Common/SettingsStore.cpp`, so a normal non-test build with
  `REDSALAMANDER_TEST_ROOT` / `REDSALAMANDER_TEST_RUN_ID` in its environment keeps using the normal
  settings contract instead of the test scratch tree.
- [x] SettingsStore source-contract regression was added in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`: `keeps SettingsStore TestSandbox path
  overrides behind ENABLE_TESTS`.
- [x] Commands broad-only blockers are fixed, not quarantined. The retained fixes cover
  Preferences category/page readiness, debug category selection, Find split-menu stationary hover,
  Quick Search visible-order expectations, Find result shortcut selection, and UIA cache
  reinitialization after explicit case listing.
- [x] Broad Commands is green under `C:\RSPerf`:
  `20260708T172023Z-57328-e062a30ebde34df9b64c7b04aa419713`, 786 passed / 0 failed / 2 skipped.
- [x] Compare SQLite traversal-seed blocker is fixed in `Common/LocalSearchIndexCore.cpp`:
  metadata-completeness probes no longer downgrade `kVolumeStateCurrentnessUnproven`.
- [x] Compare foreground Search service broad-only races are fixed in
  `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.cpp` and
  `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`:
  status probes now use `WaitForSearchServiceStatus(...)`, and foreground readiness requires a live
  next pipe instance instead of accepting a cleanly exited process as ready.
- [x] Broad CompareDirectories is green under `C:\RSPerf`:
  `20260708T194609Z-68980-82608334128e472badc953a559736b07`, 219 passed / 0 failed / 30 skipped.
- [x] FileOperations old Suite Full blocker `Phase11_ConnectionOverrideClamp` is not currently
  reproducing:
  focused `20260708T195349Z-70488-a49824a32a1a41ff92f2d2a049042852`, 3 passed / 0 failed;
  Phase 11 family `20260708T195417Z-76032-9294a024ab6c42829cf4c2c721e538fb`, 9 passed / 0 failed;
  full FileOperations `20260708T195955Z-100504-82260c4a41ac4463b9b51898e747383e`, 102 passed /
  0 failed / 20 skipped.
- [x] Latest source contracts passed:
  `Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`,
  110 passed / 0 failed.
- [x] Latest C++ build after the Compare foreground service readiness fix passed:
  `.build\logs\msbuild-20260708_214401_797.log`, 0 warnings / 0 errors.
- [x] Fresh Suite Full was run under `C:\RSPerf` after Commands, Compare, and FileOperations proof:
  `20260708T201246Z-81460-de7dd23211f0406894a73994ffdbc419`, 1119 passed / 2 failed /
  52 skipped, exit code 1.
- [x] In that Suite Full run, `CompareDirectories`, `Commands`, `FileOperations`, `ViewerPETests`,
  and the other standalone executable/CppUnitTest lanes passed. The prior ViewerVLC HUD failure did
  not reproduce: `ViewerPETests` passed on the initial attempt.
- [x] `ToolsPesterTests` root cause has a focused repro: `Invoke-Pester -Path
  .\Tools\Tests\TestInventory.Tests.ps1 -PassThru` fails because the Tools Pester count is now 221
  while `Tools/Tests/TestInventory.Tests.ps1`, `Specs/Testing/Testing_TestCoverage.md`, and
  `Tests/README.md` still expect/document 212.
- [x] `RedSalamanderMonitorEtwLatency` runner failure was checked once outside the runner with the
  same executable arguments; the direct command exited 0 with empty stdout/stderr, so the current
  failure needs runner-context or environment/order reproduction before patching.

### Remaining

- [ ] Patch the deterministic Tools Pester inventory drift:
  update the expected Tools Pester count from 212 to 221 in `Tools/Tests/TestInventory.Tests.ps1`,
  `Specs/Testing/Testing_TestCoverage.md`, and `Tests/README.md`; rerun
  `Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru`, then full
  `Invoke-Pester -Path .\Tools\Tests -PassThru`.
- [ ] Root-cause `RedSalamanderMonitorEtwLatency` in the runner context. Suite Full recorded
  `Process exited with code 1` with a zero-byte
  `C:\RSPerf\runs\20260708T201246Z-81460-de7dd23211f0406894a73994ffdbc419\artifacts\selftest\last_run\RedSalamanderMonitorEtwLatency.output.log`,
  while the direct command below exited 0:

  ```powershell
  $env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
  Remove-Item Env:\REDSALAMANDER_TEST_RUN_ID -ErrorAction SilentlyContinue
  .\.build\x64\Debug\RedSalamanderMonitor.exe --chrome-selftest --wait-instance --perf `
    --monitor-etw-burst-mode=latency --monitor-etw-burst-count=60 --monitor-etw-burst-size=260
  ```

  Next investigation should inspect `Tools/Run-AllTests.ps1` executable invocation/capture and rerun
  the monitor entry through the runner path with richer stdout/stderr/exit diagnostics.
- [ ] Rebuild if any C++ changes are made.
- [ ] Rerun source contracts if any source guard changes.
- [ ] Rerun Suite Full under `C:\RSPerf` after the Pester inventory patch and monitor ETW root-cause
  fix.
- [ ] Resolve or explicitly archive/explain the TestSandbox disk audit issues. The latest Suite Full
  reports `test_sandbox_audit.is_clean=false` with `issue_count=214`, mostly old manual run
  directories and root-level scratch files under `C:\RSPerf`.
- [ ] Archive final accepted evidence under `Specs/TestRuns/<current-head>/...` before cleaning
  volatile `C:\RSPerf` evidence.
- [ ] Merge durable requirements/lessons into authoritative specs or repo guidance.
- [ ] Move this plan to `Specs/Plans/Done/`.
- [ ] Commit only the intended stabilization files. Do not stage the unrelated
  `Specs/Plans/WIP/Product_WhimFilesGapAnalysisAndImprovementPlan_2026-07-08.md` unless explicitly
  requested.

### Current blocker details

Latest Suite Full evidence:

```text
C:\RSPerf\runs\20260708T201246Z-81460-de7dd23211f0406894a73994ffdbc419\
suite=Full
duration_ms=2479509
passed=1119
failed=2
skipped=52
fail_fast=false
timeout_scale=2
exit_code=1
test_sandbox_audit.issue_count=214
```

Current failures:

```text
suite=RedSalamanderMonitorEtwLatency
kind=Executable
exit_code=1
classification=UNCLASSIFIED_FAILURE
reason=Process exited with code 1.
output_log_path=C:\RSPerf\runs\20260708T201246Z-81460-de7dd23211f0406894a73994ffdbc419\artifacts\selftest\last_run\RedSalamanderMonitorEtwLatency.output.log
note=output log is 0 bytes; direct command with same arguments exited 0 once.

suite=ToolsPesterTests
kind=Pester
exit_code=1
classification=UNCLASSIFIED_FAILURE
reason=Process exited with code 1.
focused_repro=Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
focused_failures=Tools Pester expected/documented count 212, actual count 221.
```

### Worktree notes

Modified files intentionally related to stabilization:

```text
Common/Common/SettingsStore.cpp
Common/LocalSearchIndexCore.cpp
RedSalamander/Preferences.Dialog.cpp
RedSalamander/Preferences.Dialog.h
RedSalamander/Preferences.cpp
RedSalamander/Preferences.h
RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp
RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp
RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp
RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp
RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp
RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.cpp
Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md
Tools/Tests/TestHarnessSourceContracts.Tests.ps1
```

Unrelated dirty file present in the worktree:

```text
Specs/Plans/WIP/Product_WhimFilesGapAnalysisAndImprovementPlan_2026-07-08.md
```

## Resume checklist - 2026-07-07 22:42 +02:00 (previous)

Use this section first when resuming. The operation is not done until broad Commands and Suite Full
are green under the final root `C:\RSPerf`, final evidence is archived, durable requirements are
merged out of WIP, this plan is moved to `Specs/Plans/Done/`, and the intended stabilization files
are committed. Do not make flaky tests non-blocking. Do not create anonymous permanent quarantine.

No GUI/build process was active at this save point:

```powershell
Get-Process RedSalamander,MSBuild,cl,link -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding
```

returned no matching processes.

### Done

- [x] Final proof root specified and used for accepted local proof: `C:\RSPerf`.
- [x] Broad Commands fail-fast is green under `C:\RSPerf`:
  `20260707T150904Z-93028-a711e3f4adc54622bb6671a8ab3b907c`, 786 passed / 0 failed / 2 skipped.
- [x] Earlier broad-only Find/menu blockers are diagnosed and repeat-green:
  `manual-20260707Tresume-repeat-find-menu-after-revert`, 40 passed / 0 failed across 20 repeats.
- [x] Connection Manager long-run open/close blocker is root-caused, patched, source-contract
  guarded, rebuilt, and focused/predecessor/prefix-green:
  `manual-20260707Tconn-longrun-focused-after-settle-patch`,
  `manual-20260707Tconn-longrun-predecessor13-after-settle-patch`,
  `manual-20260707Tconn-longrun-prefix208-after-settle-patch`.
- [x] Preferences reverse-navigation broad blocker is root-caused as missing deterministic DxUi tree
  focus, not timing. The retained fix adds `DebugFocusPreferencesCategoryTree()`, snapshots the
  category-tree DxUi focus-control state, and requires that focus before each synthetic tree key.
- [x] Preferences reverse-navigation proof ladder is green:
  source contracts 99 passed / 0 failed, build
  `.build\logs\msbuild-20260707_204450_579.log` with 0 warnings / 0 errors,
  `manual-20260707Tprefs-reverse-focused-after-focus-routing-fix`, 1 passed / 0 failed, and
  `manual-20260707Tprefs-reverse-predecessor-window-after-focus-routing-fix`, 13 passed / 0 failed.
- [x] Compare Options scroll blocker is root-caused as a bad test setup contract: the old test
  shrank height only and then demanded a scrollable body even when the real snapshot stayed wide,
  two-column, and non-scrollable.
- [x] Compare Options scroll fix is patched in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp`: it now tries bounded
  scrollable resize candidates, waits for the real snapshot to prove `bodyScrollMax > 0`, and logs
  each requested/final size on failure.
- [x] Compare Options scroll source guard is patched in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`: it requires explicit resize candidates,
  attempt diagnostics, and no stale `reducedHeightPx` height-only path.
- [x] Compare Options scroll proof ladder is green:
  source contracts 99 passed / 0 failed, build
  `.build\logs\msbuild-20260707_211551_526.log` with 0 warnings / 0 errors,
  `manual-20260707Tcompare-options-scroll-focused-after-scrollable-resize-fix`, 1 passed / 0
  failed, and `manual-20260707Tcompare-options-scroll-predecessor-window-after-scrollable-resize-fix`,
  13 passed / 0 failed.
- [x] Plugins reordered/resized/search blocker is root-caused as a broad-only readiness race: the
  test waited for the selected plugin row, then captured header rects before the Plugins main-list
  header/visible columns were guaranteed ready.
- [x] Plugins fix is patched in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`: the test now
  uses `waitForInitialMainHeaderRects` readiness and requires category, selection/details state, row
  count, visible column/cell counts, no page-resize failures, and valid Name/Type header ordering
  before interactions.
- [x] Plugins source guard is patched in `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`: it now
  covers both sort/search and reordered-resized/search variants, requiring Name/Type rect readiness
  and rejecting the old single-header failure path.
- [x] Plugins proof ladder is green:
  source contracts 99 passed / 0 failed, build
  `.build\logs\msbuild-20260707_214304_335.log` with 0 warnings / 0 errors,
  `manual-20260707Tplugins-reordered-resized-search-focused-after-header-wait-fix`, 1 passed / 0
  failed, and
  `manual-20260707Tplugins-reordered-resized-search-predecessor-window-after-header-wait-fix`,
  13 passed / 0 failed.
- [x] NavigationView full-path popup failures were reproduced only under a wider broad-order window,
  not as standalone failures:
  `manual-20260707Tnav-fullpath-popup-focused-before-fix`, 2 passed / 0 failed;
  `manual-20260707Tnav-fullpath-popup-predecessor-window-before-fix`, 14 passed / 0 failed;
  `manual-20260707Tnav-fullpath-popup-window640-690-before-fix`, 49 passed / 2 failed.
- [x] NavigationView root cause was isolated by shrinking the predecessor window:
  windows 671..690 and 640..690 failed, while 672..690 and 673..690 passed. The mutating predecessor
  is `dispatch_smoke_all_commands` at broad index 671.
- [x] NavigationView root cause is fixed, not quarantined: `dispatch_smoke_all_commands` dispatched
  `cmd/pane/zoomPanel` and did not restore zoomed-pane state. With the left pane left zoomed, the
  NavigationView path region became too wide, so the short path no longer needed an ellipsis and the
  full-path popup tests timed out capturing their baseline ellipsis snapshot.
- [x] NavigationView isolation fix is patched in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`: `TestDispatchAllCommandsSmoke`
  captures baseline zoomed pane, zoom-restore split ratio, and split ratio, then restores zoom state
  and split ratio during smoke cleanup.
- [x] NavigationView isolation source guard is patched in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`: `restores dispatch smoke zoom state after
  broad command probes` requires the baseline capture and cleanup restore.
- [x] NavigationView proof ladder is green:
  `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1`, 100 passed / 0 failed;
  build `.build\logs\msbuild-20260707_221712_041.log` with 0 warnings / 0 errors;
  `manual-20260707Tnav-fullpath-popup-window671-690-after-dispatch-zoom-restore`, 20 passed / 0
  failed; and short-ID window `m-nav640a`, 51 passed / 0 failed.
- [x] Broad Commands after the NavigationView isolation fix cleared the NavigationView popup
  failures and exposed the next blockers:
  `20260707T202249Z-91432-3955222bb2574f109adad865ecd53d02`,
  784 passed / 2 failed / 2 skipped.
- [x] Worktree state is intentionally dirty with stabilization edits. Untracked review/spec files
  listed below are unrelated to this operation and must not be staged unless explicitly requested.

### Remaining

- [ ] Root-cause and fix `cmd_compare_directories_options_live_dx_body_interaction` first. Start
  with focused repro, then a predecessor window around broad index 526. Do not quarantine it and do
  not make it non-blocking.
- [ ] If the Compare live-Dx focused repro is red, fix the deterministic setup/product bug directly.
  If focused is green but predecessor fails, identify the broad-order state leak before patching.
- [ ] Add or adjust a source contract if the Compare fix introduces a durable invariant, for example
  cancel/reopen edit restore must use the actually named visible edit after scroll/focus settles and
  must not compare against a stale/currently hidden UIA edit.
- [ ] After Compare is fixed, root-cause and fix
  `cmd_pane_navigation_find_window_keeps_navigation_shell_stable`, currently failing because pane
  contents are not ready after switching to the Find shell-stability root.
- [ ] Rebuild if C++ changed.
- [ ] Rerun source contracts if any source guard changed.
- [ ] Rerun focused proof for each patched failing case under `C:\RSPerf`.
- [ ] Rerun the predecessor window around each patched broad index under `C:\RSPerf`.
- [ ] Rerun broad Commands without `-FailFast` under `C:\RSPerf`.
- [ ] If broad Commands exposes a new blocker, root-cause/fix it the same way: focused repro,
  predecessor proof, deterministic patch, source guard where useful, then broad rerun.
- [ ] Run Suite Full under `C:\RSPerf` only after broad Commands is green.
- [ ] Archive final accepted evidence under `Specs/TestRuns/<current-head>/...` before cleaning
  volatile `C:\RSPerf` evidence.
- [ ] Merge durable requirements/lessons into authoritative specs or repo guidance.
- [ ] Move this plan to `Specs/Plans/Done/`.
- [ ] Commit only the intended stabilization files. Do not stage unrelated untracked review/spec
  files unless explicitly instructed.

### Current blocker details

Latest broad Commands evidence:

```text
C:\RSPerf\runs\20260707T202249Z-91432-3955222bb2574f109adad865ecd53d02\
suite=Commands
duration_ms=1105851
passed=784
failed=2
skipped=2
fail_fast=false
classification=
```

Current failures:

```text
index=526
name=cmd_compare_directories_options_live_dx_body_interaction
duration_ms=51489
reason=Compare Directories options DX edit 'Ignore files' should discard the canceled mutation and reopen at its original value; phase='reopened edit restore' edit='Ignore files' expected='estion ' initial='estion ' edited='selftest-ignore-pattern' snapshot=visible=1 dxStatics=1 dxButtons=1 dxToggles=1 dxEdits=1 legacyStatics=0 legacyFooterButtons=0 legacyToggles=0 legacyEdits=0 nativeBody=0 bodyRendered=1 bodyResizeFailures=0 bodyPresentFailures=0 bodySize=1306x579 bodyContentHeight=745 bodyScroll=53/166 bodyCards=10 bodyHeaders=4 dxFooterHosts=2 visibleDxFooterHosts=2 themeDark=1 themeHighContrast=0 themeRainbow=0 focusTarget=0 lastObserved=name='Ignore files' value='estion ' readonly=false controlType=50004 currentNamed=name='Ignore files' value='' readonly=false controlType=50004 visibleEdits=[0] name='Ignore files' value='' hasValue=true readonly=false bounds=(1181,1153,1777,1192) rawEdits=[0] name='Ignore files' value='' hasValue=true readonly=false controlType=50004; [1] name='Command:' value='' hasValue=true readonly=false controlType=50004.

index=594
name=cmd_pane_navigation_find_window_keeps_navigation_shell_stable
duration_ms=6193
reason=Pane contents not ready for Find shell-stability test.
```

Blunt next diagnosis: the first current blocker is not the already-fixed Compare Options scroll
case. It is the live-Dx body interaction cancel/reopen restore check. The failure text shows the
initial value as truncated (`estion `) and the reopened visible edit as empty, while `lastObserved`
still has the truncated value. That points at either stale edit selection/focus after reopen,
scroll-position dependent UIA targeting, or a real cancel/reopen persistence bug. Prove which one
with focused and predecessor runs before touching code.

### Investigation notes to keep

- Compare first blocker source:
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp`.
- Compare live-Dx function:
  `TestCompareDirectoriesOptionsLiveDxBodyInteraction`, starts around line 923.
- Compare failure assertions:
  focus/reopen restore logic around lines 1531..1592; case registration around lines 5113..5114.
- Search anchors:
  `live_dx_body_interaction`, `reopened edit restore`, `Ignore files`, and
  `discard the canceled mutation`.
- Navigation second blocker source:
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Navigation.cpp`.
- Navigation Find function:
  `TestPaneNavigationFindWindowKeepsNavigationShellStable`, pane setup around lines 6406..6451 and
  case registration around line 9645.
- Navigation popup blocker is fixed. Do not restart there unless it reappears in a new broad run.
  The retained root-cause summary is: `dispatch_smoke_all_commands` leaked zoomed-pane state after
  `cmd/pane/zoomPanel`, widening the NavigationView path region and removing the baseline ellipsis
  required by the popup tests.
- Avoid long manual run IDs for path-stress windows. A long NavigationView proof run ID caused
  unrelated setup failures in long-name pane tests by pushing classic Win32 path length budget. Use
  short IDs such as `m-compare-live-pred` and `m-navfind-pred`.
- Run only one native GUI selftest at a time.

### Worktree notes

Modified files that are part of this stabilization operation include:

```text
Common/Common/SettingsStore.cpp
Common/DxUi/DxUi.Menu.cpp
Common/DxUi/DxUi.h
RedSalamander/ManagePluginsDialog.cpp
RedSalamander/Preferences.Dialog.cpp
RedSalamander/Preferences.Dialog.h
RedSalamander/Preferences.cpp
RedSalamander/Preferences.h
RedSalamander/RedSalamander.cpp
RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp
RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp
RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp
RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp
RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp
RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp
RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ThemesGeneralPanes.cpp
RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp
RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.cpp
RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp
RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp
RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp
Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md
Specs/Plans/WIP/README.md
Tests/DxUiTests/DxUiTests.Menu.cpp
Tools/Tests/TestHarnessSourceContracts.Tests.ps1
Tools/Tests/WingetValidation.Tests.ps1
```

Untracked files observed before this checkpoint are unrelated to this operation; do not stage them
unless the user explicitly asks:

```text
Specs/Core/Core_FileSystemBridge.md
Specs/Plans/WIP/FileSystem_CrossFsBridgeImplementationAndPerformanceRedesign_2026-07-07.md
Specs/Plans/Done/Operation_Causeway_CrossFsBridgeSecurityAndPerformanceRemediation_2026-07-07.md
Specs/Plans/WIP/UI_FileOperationsPopupUxRefinementPlan_2026-07-07.md
Specs/Reviews/FileSystemBridge-2026-07-07-Findings.md
Specs/Reviews/ThreeDayDiff-2026-07-06-Findings.md
```

### Exact resume commands

```powershell
# Confirm no app/build process is holding the interactive desktop.
Get-Process RedSalamander,MSBuild,cl,link -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding

# Inspect the current broad failures.
$run='C:\RSPerf\runs\20260707T202249Z-91432-3955222bb2574f109adad865ecd53d02'
$root="$run\artifacts\selftest\last_run\commands"
$d=Get-Content "$root\results.json" -Raw | ConvertFrom-Json
$d | Select-Object suite,duration_ms,passed,failed,skipped,fail_fast,repeat_count,classification | Format-List
$failed = @()
for ($i=0; $i -lt $d.cases.Count; $i++) {
  if ($d.cases[$i].status -eq 'failed') {
    $failed += [pscustomobject]@{
      index=$i
      name=$d.cases[$i].name
      duration_ms=$d.cases[$i].duration_ms
      reason=$d.cases[$i].reason
    }
  }
}
$failed | Format-List
510..535 | Where-Object { $_ -ge 0 -and $_ -lt $d.cases.Count } |
  ForEach-Object { [pscustomobject]@{index=$_; name=$d.cases[$_].name; status=$d.cases[$_].status; duration_ms=$d.cases[$_].duration_ms} } |
  Format-Table -AutoSize
585..600 | Where-Object { $_ -ge 0 -and $_ -lt $d.cases.Count } |
  ForEach-Object { [pscustomobject]@{index=$_; name=$d.cases[$_].name; status=$d.cases[$_].status; duration_ms=$d.cases[$_].duration_ms} } |
  Format-Table -AutoSize

# Locate the Compare live-Dx failure source.
rg -n "TestCompareDirectoriesOptionsLiveDxBodyInteraction|live_dx_body_interaction|Ignore files|reopened edit restore|discard the canceled mutation" `
  RedSalamander\SelfTest\Commands\Commands.SelfTest.CompareOptions.cpp

# Focused Compare repro.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='m-compare-live-before'
$case='cmd_compare_directories_options_live_dx_body_interaction'
$args=@('--commands-selftest',"--selftest-case=$case",'--selftest-timeout-multiplier=2')
$p=Start-Process -FilePath (Resolve-Path '.\.build\x64\Debug\RedSalamander.exe') -ArgumentList $args -PassThru -Wait
$p.ExitCode
Get-Content "C:\RSPerf\runs\m-compare-live-before\artifacts\selftest\last_run\commands\results.json" -Raw |
  ConvertFrom-Json |
  Select-Object passed,failed,skipped

# Predecessor window around broad index 526. Keep the run ID short.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='m-compare-live-pred'
$run='C:\RSPerf\runs\20260707T202249Z-91432-3955222bb2574f109adad865ecd53d02'
$d=Get-Content "$run\artifacts\selftest\last_run\commands\results.json" -Raw | ConvertFrom-Json
$start=[Math]::Max(0,526-12)
$windowCases = (($start)..526 | ForEach-Object { $d.cases[$_].name }) -join ','
$args=@('--commands-selftest',"--selftest-case=$windowCases",'--selftest-timeout-multiplier=2')
$p=Start-Process -FilePath (Resolve-Path '.\.build\x64\Debug\RedSalamander.exe') -ArgumentList $args -PassThru -Wait
$p.ExitCode
Get-Content "C:\RSPerf\runs\m-compare-live-pred\artifacts\selftest\last_run\commands\results.json" -Raw |
  ConvertFrom-Json |
  Select-Object passed,failed,skipped

# After patching, run contracts/build/proof/broad in this order.
Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1
.\build.ps1 -ProjectName RedSalamander
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
Remove-Item Env:\REDSALAMANDER_TEST_RUN_ID -ErrorAction SilentlyContinue
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2

# Only after broad Commands is green:
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

## Superseded resume checkpoint - 2026-07-07 22:06 +02:00

Use this section first when resuming. The operation is not done until broad Commands and Suite Full
are green under the final root `C:\RSPerf`, final evidence is archived, durable requirements are
merged out of WIP, this plan is moved to `Specs/Plans/Done/`, and the intended stabilization files
are committed. Do not make flaky tests non-blocking. Do not create anonymous permanent quarantine.

No GUI/build process was active at this save point:

```powershell
Get-Process RedSalamander,MSBuild,cl,link -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding
```

returned no matching processes.

### Done

- [x] Final proof root specified and used for accepted local proof: `C:\RSPerf`.
- [x] Broad Commands fail-fast is green under `C:\RSPerf`:
  `20260707T150904Z-93028-a711e3f4adc54622bb6671a8ab3b907c`, 786 passed / 0 failed / 2 skipped.
- [x] Earlier broad-only Find/menu blockers are diagnosed and repeat-green:
  `manual-20260707Tresume-repeat-find-menu-after-revert`, 40 passed / 0 failed across 20 repeats.
- [x] Connection Manager long-run open/close blocker is root-caused, patched, source-contract
  guarded, rebuilt, and focused/predecessor/prefix-green:
  `manual-20260707Tconn-longrun-focused-after-settle-patch`,
  `manual-20260707Tconn-longrun-predecessor13-after-settle-patch`,
  `manual-20260707Tconn-longrun-prefix208-after-settle-patch`.
- [x] Swap Panes focused and predecessor probes are green, and the retained diagnostic patch now
  records enough pane/navigation context if that case reappears:
  `manual-20260707Tswap-panes-focused-after-diagnostics`,
  `manual-20260707Tswap-panes-predecessor-after-diagnostics`.
- [x] Preferences reverse-navigation broad blocker was root-caused as missing deterministic DxUi
  tree focus, not timing. The retained fix adds `DebugFocusPreferencesCategoryTree()`, snapshots the
  category-tree DxUi focus-control state, and requires that focus before each synthetic tree key.
- [x] Source contracts after the Preferences focus-routing fix are green:
  `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1`, 99 passed / 0 failed.
- [x] Build after the Preferences focus-routing fix is green:
  `.build\logs\msbuild-20260707_204450_579.log`, 0 warnings / 0 errors.
- [x] Preferences reverse-navigation proof ladder after the focus-routing fix is green:
  `manual-20260707Tprefs-reverse-focused-after-focus-routing-fix`, 1 passed / 0 failed;
  `manual-20260707Tprefs-reverse-predecessor-window-after-focus-routing-fix`,
  13 passed / 0 failed.
- [x] Broad Commands after the Preferences fix completed under `C:\RSPerf` and removed the
  Preferences blocker:
  `20260707T184755Z-102276-c00ff4ec02eb45f6afc0715b10565031`, 785 passed / 1 failed / 2 skipped.
- [x] Compare Options scroll blocker was reproduced focused before fixing:
  `manual-20260707Tcompare-options-scroll-focused-before-fix`, 0 passed / 1 failed.
- [x] Compare Options root cause was confirmed as test setup, not product scrolling: the old test
  shrank height only and then demanded a scrollable body even when the real snapshot stayed wide,
  two-column, and non-scrollable (`host=2738x680`, `contentHeight=625..667`, `scrollMax=0`,
  `twoColumns=true`).
- [x] Compare Options fix is patched in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp`: it now tries bounded
  scrollable resize candidates, waits for the real snapshot to prove `bodyScrollMax > 0`, and logs
  each requested/final size on failure.
- [x] Compare Options source guard is patched in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`: it requires explicit resize candidates,
  attempt diagnostics, and no stale `reducedHeightPx` height-only path.
- [x] Build after the Compare Options fix is green:
  `.build\logs\msbuild-20260707_211551_526.log`, 0 warnings / 0 errors.
- [x] Compare Options source contracts and proof ladder are green:
  `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1`, 99 passed / 0 failed;
  `manual-20260707Tcompare-options-scroll-focused-after-scrollable-resize-fix`,
  1 passed / 0 failed;
  `manual-20260707Tcompare-options-scroll-predecessor-window-after-scrollable-resize-fix`,
  13 passed / 0 failed.
- [x] Broad Commands after the Compare Options fix removed that blocker and exposed the next one:
  `20260707T192240Z-99972-778b30694d14437d8affa4ef018097c8`, 785 passed / 1 failed / 2 skipped.
- [x] Plugins reordered/resized/search blocker was root-caused as a broad-only readiness race: the
  test waited for the selected plugin row, then captured header rects before the Plugins main-list
  header/visible columns were guaranteed ready.
- [x] Plugins fix is patched in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`: the test now
  uses the established `waitForInitialMainHeaderRects` readiness pattern and requires category,
  selection/details state, row count, visible column/cell counts, no page-resize failures, and valid
  Name/Type header ordering before interactions.
- [x] Plugins source guard is patched in `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`: it now
  covers both sort/search and reordered-resized/search variants, requiring Name/Type rect readiness
  and rejecting the old single-header failure path.
- [x] Build after the Plugins fix is green:
  `.build\logs\msbuild-20260707_214304_335.log`, 0 warnings / 0 errors.
- [x] Plugins source contracts and proof ladder are green:
  `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1`, 99 passed / 0 failed;
  `manual-20260707Tplugins-reordered-resized-search-focused-after-header-wait-fix`,
  1 passed / 0 failed;
  `manual-20260707Tplugins-reordered-resized-search-predecessor-window-after-header-wait-fix`,
  13 passed / 0 failed.
- [x] Broad Commands after the Plugins fix removed that blocker and exposed the current
  NavigationView blocker:
  `20260707T194611Z-102464-8db1b1234b6c409dbfc17f5da2b51b9e`,
  784 passed / 2 failed / 2 skipped.
- [x] Worktree state is intentionally dirty with stabilization edits. Untracked review/spec files
  listed below are unrelated to this operation and must not be staged unless explicitly requested.

### Remaining

- [ ] Root-cause the current NavigationView full-path popup failures. Start with focused repro,
  then predecessor window. Do not quarantine these tests and do not make them non-blocking.
- [ ] If focused repro is red, fix the deterministic setup/product bug directly. If focused is
  green but predecessor fails, root-cause the broad-order dependency before patching.
- [ ] Likely investigation path: prove whether the tests deterministically create a baseline
  ellipsis state. If the main window/path width is not forced tightly enough, the test should drive
  the UI to an actual ellipsis snapshot before validating popup edit/ancestor routes.
- [ ] Add/adjust a source contract if the fix introduces a durable test-harness invariant, for
  example "full-path popup tests must force and verify a visible ellipsis snapshot before opening
  the popup".
- [ ] Rebuild if C++ changed.
- [ ] Rerun source contracts if the guard changed.
- [ ] Rerun focused proof for both current NavigationView popup cases under `C:\RSPerf`.
- [ ] Rerun the predecessor window around broad indexes 677..690 under `C:\RSPerf`.
- [ ] Rerun broad Commands without `-FailFast` under `C:\RSPerf`.
- [ ] If broad Commands exposes a new blocker, root-cause/fix it the same way: focused repro,
  predecessor proof, deterministic patch, source guard where useful, then broad rerun.
- [ ] Run Suite Full under `C:\RSPerf` only after broad Commands is green.
- [ ] Archive final accepted evidence under `Specs/TestRuns/<current-head>/...` before cleaning
  volatile `C:\RSPerf` evidence.
- [ ] Merge durable requirements/lessons into authoritative specs or repo guidance.
- [ ] Move this plan to `Specs/Plans/Done/`.
- [ ] Commit only the intended stabilization files. Do not stage unrelated untracked review/spec
  files unless explicitly instructed.

### Current blocker details

Latest broad Commands evidence:

```text
C:\RSPerf\runs\20260707T194611Z-102464-8db1b1234b6c409dbfc17f5da2b51b9e\
suite=Commands
duration=about 18m54.8s
passed=784
failed=2
skipped=2
fail_fast=false
```

Current failures:

```text
index=689
name=cmd_pane_navigationView_full_path_popup_edit_route
reason=Failed to capture baseline navigation-view ellipsis state for full-path popup test.

index=690
name=cmd_pane_navigationView_full_path_popup_ancestor_click_navigates_to_ancestor
reason=Failed to capture baseline ellipsis state before full-path popup ancestor-click validation.
```

Nearby cases 677..688 passed, including the path double-click, keyboard activation, ancestor-click,
and unfocused-pane NavigationView coverage. After the two failures, later NavigationView
region/dropdown/edit tests also passed. That makes the current failure narrower than "NavigationView
is broken": the weak point is the full-path popup baseline ellipsis setup or its broad-order
precondition.

Blunt diagnosis to verify next: this smells like another test setup contract hole. The tests appear
to need a visible ellipsis in the breadcrumb path before opening the full-path popup, but the current
broad failure only says the baseline ellipsis snapshot could not be captured. Do not patch a timeout
blindly. First prove whether the path/window geometry actually produces an ellipsis, whether a prior
case leaves width/focus/navigation state different, or whether the product stops exposing the
ellipsis debug snapshot when it should.

### Investigation notes to keep

- Failing test file:
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`.
- First failing function:
  `TestPaneNavigationViewFullPathPopupEditRoute`, starts around line 4171; baseline ellipsis failure
  around line 4290.
- Second failing function:
  `TestPaneNavigationViewFullPathPopupAncestorClickNavigatesToAncestor`, starts around line 4545;
  baseline ellipsis failure around line 4629.
- Case registrations are around lines 24073..24077.
- Product/debug helpers:
  `FolderWindow::DebugGetNavigationViewHwnd` and `DebugGetNavigationViewSnapshot` in
  `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp` around lines 11046..11052.
- Snapshot type:
  `NavigationViewDebugSnapshot` in `RedSalamander/NavigationView.h` around line 75.
- Snapshot implementation:
  `NavigationView::DebugGetSnapshot` in `RedSalamander/NavigationView.cpp`.
- Breadcrumb/ellipsis layout:
  `RedSalamander/NavigationView.Breadcrumb.cpp`.
- Full-path popup product code:
  `RedSalamander/NavigationView.FullPathPopup.cpp`.
- Current trace around the failures is thin; expect to add diagnostics if focused/predecessor proof
  cannot distinguish "no ellipsis because setup did not force one" from "product should expose one
  and does not".
- Run only one native GUI selftest at a time.

### Worktree notes

Modified files that are part of this stabilization operation include:

```text
Common/Common/SettingsStore.cpp
Common/DxUi/DxUi.Menu.cpp
Common/DxUi/DxUi.h
RedSalamander/ManagePluginsDialog.cpp
RedSalamander/Preferences.Dialog.cpp
RedSalamander/Preferences.Dialog.h
RedSalamander/Preferences.cpp
RedSalamander/Preferences.h
RedSalamander/RedSalamander.cpp
RedSalamander/SelfTest/Commands/*.cpp
Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md
Specs/Plans/WIP/README.md
Tests/DxUiTests/DxUiTests.Menu.cpp
Tools/Tests/TestHarnessSourceContracts.Tests.ps1
Tools/Tests/WingetValidation.Tests.ps1
```

Untracked files observed before this checkpoint are unrelated to this operation; do not stage them
unless the user explicitly asks:

```text
Specs/Core/Core_FileSystemBridge.md
Specs/Plans/WIP/FileSystem_CrossFsBridgeImplementationAndPerformanceRedesign_2026-07-07.md
Specs/Plans/Done/Operation_Causeway_CrossFsBridgeSecurityAndPerformanceRemediation_2026-07-07.md
Specs/Plans/WIP/UI_FileOperationsPopupUxRefinementPlan_2026-07-07.md
Specs/Reviews/FileSystemBridge-2026-07-07-Findings.md
Specs/Reviews/ThreeDayDiff-2026-07-06-Findings.md
```

### Exact resume commands

```powershell
# Confirm no app/build process is holding the interactive desktop.
Get-Process RedSalamander,MSBuild,cl,link -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding

# Inspect the current broad failure.
$run='C:\RSPerf\runs\20260707T194611Z-102464-8db1b1234b6c409dbfc17f5da2b51b9e'
$root="$run\artifacts\selftest\last_run\commands"
$d=Get-Content "$root\results.json" -Raw | ConvertFrom-Json
$d | Select-Object suite,duration_ms,passed,failed,skipped,fail_fast,repeat_count | Format-List
$failed = for ($i=0; $i -lt $d.cases.Count; $i++) {
  if ($d.cases[$i].status -eq 'failed') {
    [pscustomobject]@{index=$i; name=$d.cases[$i].name; duration_ms=$d.cases[$i].duration_ms; reason=$d.cases[$i].reason}
  }
}
$failed | Format-List
677..700 | Where-Object { $_ -ge 0 -and $_ -lt $d.cases.Count } |
  ForEach-Object {
    [pscustomobject]@{
      index=$_
      name=$d.cases[$_].name
      status=$d.cases[$_].status
      duration_ms=$d.cases[$_].duration_ms
      reason=$d.cases[$_].reason
    }
  } | Format-Table -AutoSize -Wrap
Select-String -LiteralPath "$root\trace.txt" `
  -Pattern 'full_path_popup|baseline navigation-view ellipsis|baseline ellipsis|cmd_pane_navigationView' `
  -Context 4,12 | ForEach-Object { $_.ToString() }

# Read the failing tests and product/debug helpers before patching.
$path='RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp'
$lines=Get-Content $path
for($i=4170;$i -le 4310;$i++){ '{0,5}: {1}' -f $i,$lines[$i-1] }
for($i=4545;$i -le 4645;$i++){ '{0,5}: {1}' -f $i,$lines[$i-1] }
rg -n "WaitForNavigationViewSnapshot|ellipsis|visibleEllipsis|hasEllipsis|baseline navigation-view ellipsis|baseline ellipsis|full_path_popup" `
  RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp `
  RedSalamander\NavigationView*.cpp RedSalamander\NavigationView*.h
rg -n "DebugGetNavigationViewHwnd|DebugGetNavigationViewSnapshot|NavigationViewDebugSnapshot|DebugGetSnapshot" `
  RedSalamander\FolderWindow.FileSystem.Commands.Part.cpp `
  RedSalamander\NavigationView*.cpp RedSalamander\NavigationView*.h

# Focused repro/proof for the current blocker.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-20260707Tnav-fullpath-popup-focused-before-fix'
$case='cmd_pane_navigationView_full_path_popup_edit_route,cmd_pane_navigationView_full_path_popup_ancestor_click_navigates_to_ancestor'
$args=@('--commands-selftest',"--selftest-case=$case",'--selftest-timeout-multiplier=2')
$p=Start-Process -FilePath (Resolve-Path '.\.build\x64\Debug\RedSalamander.exe') -ArgumentList $args -PassThru -Wait
Start-Sleep -Seconds 3
"processExit=$($p.ExitCode)"
$proof="C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands"
$res=Get-Content "$proof\results.json" -Raw | ConvertFrom-Json
$res | Select-Object suite,duration_ms,passed,failed,skipped,fail_fast,repeat_count | Format-List
$res.cases | Where-Object status -ne 'passed' | Format-Table -AutoSize -Wrap

# If focused is green, run the broad predecessor window before patching.
$failedRun='C:\RSPerf\runs\20260707T194611Z-102464-8db1b1234b6c409dbfc17f5da2b51b9e\artifacts\selftest\last_run\commands'
$d=Get-Content "$failedRun\results.json" -Raw | ConvertFrom-Json
$windowCases = (677..690 | ForEach-Object { $d.cases[$_].name }) -join ','
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-20260707Tnav-fullpath-popup-predecessor-window-before-fix'
$args=@('--commands-selftest',"--selftest-case=$windowCases",'--selftest-timeout-multiplier=2')
$p=Start-Process -FilePath (Resolve-Path '.\.build\x64\Debug\RedSalamander.exe') -ArgumentList $args -PassThru -Wait
Start-Sleep -Seconds 3
"processExit=$($p.ExitCode)"
$proof="C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands"
$res=Get-Content "$proof\results.json" -Raw | ConvertFrom-Json
$res | Select-Object suite,duration_ms,passed,failed,skipped,fail_fast,repeat_count | Format-List
$res.cases | Where-Object status -ne 'passed' | Format-Table -AutoSize -Wrap

# After the NavigationView fix/proof ladder is green, rerun broad Commands.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
Remove-Item Env:\REDSALAMANDER_TEST_RUN_ID -ErrorAction SilentlyContinue
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
```

## Superseded resume checkpoint - 2026-07-07 21:09 +02:00

Use this section first when resuming. The operation is not done until broad Commands and Suite Full
are green under the final root `C:\RSPerf`, final evidence is archived, durable requirements are
merged out of WIP, this plan is moved to `Specs/Plans/Done/`, and the intended stabilization files
are committed. Do not make flaky tests non-blocking. Do not create anonymous permanent quarantine.

No GUI/build process was active at this save point:

```powershell
Get-Process RedSalamander,MSBuild,cl,link -ErrorAction SilentlyContinue
```

returned no processes.

### Done

- [x] Final proof root specified and used for accepted local proof: `C:\RSPerf`.
- [x] Broad Commands fail-fast is green under `C:\RSPerf`:
  `20260707T150904Z-93028-a711e3f4adc54622bb6671a8ab3b907c`, 786 passed / 0 failed / 2 skipped.
- [x] Earlier broad-only Find/menu blockers are diagnosed and repeat-green:
  `manual-20260707Tresume-repeat-find-menu-after-revert`, 40 passed / 0 failed across 20 repeats.
- [x] Connection Manager long-run open/close blocker is root-caused, patched, source-contract guarded,
  rebuilt, and focused/predecessor/prefix-green:
  `manual-20260707Tconn-longrun-focused-after-settle-patch`,
  `manual-20260707Tconn-longrun-predecessor13-after-settle-patch`,
  `manual-20260707Tconn-longrun-prefix208-after-settle-patch`.
- [x] Swap Panes focused and predecessor probes are green, and the retained diagnostic patch now
  records enough pane/navigation context if that case reappears:
  `manual-20260707Tswap-panes-focused-after-diagnostics`,
  `manual-20260707Tswap-panes-predecessor-after-diagnostics`.
- [x] Preferences reverse-navigation broad blocker was root-caused as missing deterministic DxUi tree
  focus, not timing. The retained fix adds `DebugFocusPreferencesCategoryTree()`, snapshots the
  category-tree DxUi focus-control state, and requires that focus before each synthetic tree key.
- [x] Source contracts after the Preferences focus-routing fix are green:
  `Invoke-Pester .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1`, 99 passed / 0 failed.
- [x] Build after the Preferences focus-routing fix is green:
  `.build\logs\msbuild-20260707_204450_579.log`, 0 warnings / 0 errors.
- [x] Preferences reverse-navigation proof ladder after the focus-routing fix is green:
  `manual-20260707Tprefs-reverse-focused-after-focus-routing-fix`, 1 passed / 0 failed;
  `manual-20260707Tprefs-reverse-predecessor-window-after-focus-routing-fix`,
  13 passed / 0 failed.
- [x] Broad Commands after the Preferences fix completed under `C:\RSPerf` and removed the Preferences
  blocker:
  `20260707T184755Z-102276-c00ff4ec02eb45f6afc0715b10565031`, 785 passed / 1 failed / 2 skipped.
- [x] Current broad failure was reproduced as a focused failure before any fix:
  `manual-20260707Tcompare-options-scroll-focused-before-fix`, 0 passed / 1 failed.
- [x] Current root-cause hypothesis is saved below: the Compare Options scroll test assumes shrinking
  height alone will expose a scroll range, but the live options body remains wide/two-column and
  non-scrollable (`host=2738x680`, `contentHeight=625..667`, `scrollMax=0`, `twoColumns=true`).
- [x] Worktree state is intentionally dirty with stabilization edits. Untracked review/spec files
  are unrelated to this operation and must not be staged unless explicitly requested.

### Remaining

- [ ] Fix `cmd_compare_directories_options_scroll_to_lower_cards_stays_stable` by making its setup
  deterministically produce a real scrollable one-host DX options body, or by fixing a real product
  layout bug if investigation proves one. Do not quarantine it and do not make it non-blocking.
- [ ] Inspect Compare window min-size/layout constraints before patching. The important question is
  whether the test can force one-column or smaller-height layout using normal window resize.
- [ ] Patch the deterministic root cause. Preferred direction: replace the current single
  height-only resize with a bounded resize-candidate loop that checks actual snapshot state
  (`bodyScrollMax > 0`) after each candidate and logs requested/final sizes on failure.
- [ ] Add/adjust a source contract if the fix introduces a durable test-harness invariant, for
  example "scroll lower-cards test must try explicit scrollable resize candidates and assert the
  post-resize snapshot is actually scrollable".
- [ ] Rebuild if C++ changed.
- [ ] Rerun source contracts if the guard changed.
- [ ] Rerun focused proof for
  `cmd_compare_directories_options_scroll_to_lower_cards_stays_stable` under `C:\RSPerf`.
- [ ] Rerun the predecessor window around broad index 529 under `C:\RSPerf`.
- [ ] Rerun broad Commands without `-FailFast` under `C:\RSPerf`.
- [ ] If broad Commands exposes a new blocker, root-cause/fix it the same way: focused repro,
  predecessor proof, deterministic patch, source guard where useful, then broad rerun.
- [ ] Run Suite Full under `C:\RSPerf` only after broad Commands is green.
- [ ] Archive final accepted evidence under `Specs/TestRuns/<current-head>/...` before cleaning
  volatile `C:\RSPerf` evidence.
- [ ] Merge durable requirements/lessons into authoritative specs or repo guidance.
- [ ] Move this plan to `Specs/Plans/Done/`.
- [ ] Commit only the intended stabilization files. Do not stage unrelated untracked review/spec
  files unless explicitly instructed.

### Current blocker details

Latest broad Commands evidence:

```text
C:\RSPerf\runs\20260707T184755Z-102276-c00ff4ec02eb45f6afc0715b10565031\
suite=Commands
passed=785
failed=1
skipped=2
fail_fast=false
first failure index=529
first failure name=cmd_compare_directories_options_scroll_to_lower_cards_stays_stable
```

Broad failure text:

```text
Compare Directories options did not expose a scrollable one-host DX body after the validation resize.
optionsVisible=true usesDxStatics=true usesDxButtons=true usesDxToggles=true usesDxEdits=true
legacyStatics=0 legacyButtons=0 legacyToggles=0 legacyEdits=0
host=2738x680 contentHeight=667 scrollMax=0 twoColumns=true visibleDxHosts=1
resizeFailures=0 presentFailures=0 runSnapshot=true runVisible=true runOptionsVisible=true
compareStarted=false compareActive=false runPending=false.
```

Focused repro evidence:

```text
C:\RSPerf\runs\manual-20260707Tcompare-options-scroll-focused-before-fix\
suite=Commands
passed=0
failed=1
skipped=0
failure name=cmd_compare_directories_options_scroll_to_lower_cards_stays_stable
host=2738x680 contentHeight=625 scrollMax=0 twoColumns=true visibleDxHosts=1
```

Blunt diagnosis: the test is red because its setup did not create the state it claims to validate.
It shrinks the Compare window height once, then demands a scrollable options body. In this desktop,
the options host remains `2738x680`, content is shorter than the viewport, and two-column layout is
still active. `scrollMax=0` is correct for that state. Treat this as a deterministic test setup bug
unless min-size/layout inspection proves the product is refusing a valid user resize.

### Investigation notes to keep

- Failing test:
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp`,
  `TestCompareDirectoriesOptionsScrollToLowerCardsStaysStable`.
- Current brittle code path: if `snapshot.bodyScrollMax <= 0`, the test computes
  `reducedHeightPx = (currentHeightPx * 2) / 3`, calls `SetWindowPos(...)` once with the current
  width, then waits for `bodyScrollMax > 0`. Width is never reduced, so two-column layout can stay
  active and content can remain shorter than the host.
- Product layout to inspect:
  `RedSalamander/CompareDirectoriesWindow.Options.cpp`.
  `measureLayout(contentW)` chooses two columns when the content width is large enough and computes
  `scrollMax = max(0, contentHeight - viewportH2)`.
- Snapshot fields already contain enough proof data:
  `bodyDxHostWidth`, `bodyDxHostHeight`, `bodyContentHeight`, `bodyScrollMax`,
  `bodyUsesTwoColumns`, and visible DX/legacy control counts.
- The replacement test setup should be evidence-driven: request a candidate size, pump/wait, read
  the real snapshot, and proceed only when the snapshot itself proves the body is scrollable.
- Candidate implementation idea:
  try several bounded `(width,height)` pairs such as current width with smaller heights, then
  narrower widths to force one-column layout, then lower heights if the product min size allows it.
  Record each requested size and resulting snapshot in the final failure text.
- Do not make the product artificially scrollable just for the test unless a real layout contract
  says lower cards must remain reachable at that size.
- Run only one native GUI selftest at a time.

### Exact resume commands

```powershell
# Confirm no app/build process is holding the interactive desktop.
Get-Process RedSalamander,MSBuild,cl,link -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding

# Inspect the current broad failure.
$root='C:\RSPerf\runs\20260707T184755Z-102276-c00ff4ec02eb45f6afc0715b10565031\artifacts\selftest\last_run\commands'
$d=Get-Content "$root\results.json" -Raw | ConvertFrom-Json
$d | Select-Object suite,duration_ms,passed,failed,skipped,fail_fast,repeat_count | Format-List
$failed = for ($i=0; $i -lt $d.cases.Count; $i++) {
  if ($d.cases[$i].status -eq 'failed') {
    [pscustomobject]@{index=$i; name=$d.cases[$i].name; duration_ms=$d.cases[$i].duration_ms; reason=$d.cases[$i].reason}
  }
}
$failed | Format-List
$idx=$failed[0].index
($idx-10)..($idx+10) | Where-Object { $_ -ge 0 -and $_ -lt $d.cases.Count } |
  ForEach-Object {
    [pscustomobject]@{
      index=$_
      name=$d.cases[$_].name
      status=$d.cases[$_].status
      duration_ms=$d.cases[$_].duration_ms
      reason=$d.cases[$_].reason
    }
  } | Format-Table -AutoSize -Wrap
Select-String -LiteralPath "$root\trace.txt" `
  -Pattern 'cmd_compare_directories_options_scroll_to_lower_cards_stays_stable|scrollable one-host DX body|bodyScrollMax|twoColumns' `
  -Context 4,12 | ForEach-Object { $_.ToString() }

# Inspect the test and product layout constraints before patching.
rg -n "TestCompareDirectoriesOptionsScrollToLowerCardsStaysStable|bodyScrollMax|bodyUsesTwoColumns|SetWindowPos|reducedHeightPx" `
  RedSalamander\SelfTest\Commands\Commands.SelfTest.CompareOptions.cpp
rg -n "WM_GETMINMAXINFO|ptMinTrackSize|MINMAXINFO|MinTrack|minimum.*Compare|kMin" `
  RedSalamander -g "*.cpp" -g "*.h"
rg -n "measureLayout|bodyScrollMax|bodyUsesTwoColumns|scrollMax|twoColumns" `
  RedSalamander\CompareDirectoriesWindow.Options.cpp RedSalamander -g "*.cpp" -g "*.h"

# Rebuild after the Compare Options fix if C++ changed.
.\build.ps1 -ProjectName RedSalamander

# Focused proof after patching the scrollable-resize setup.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-20260707Tcompare-options-scroll-focused-after-scrollable-resize-fix'
$case='cmd_compare_directories_options_scroll_to_lower_cards_stays_stable'
$args=@('--commands-selftest',"--selftest-case=$case",'--selftest-timeout-multiplier=2')
$p=Start-Process -FilePath (Resolve-Path '.\.build\x64\Debug\RedSalamander.exe') -ArgumentList $args -PassThru -Wait
Start-Sleep -Seconds 3
"processExit=$($p.ExitCode)"
$proof="C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands"
$res=Get-Content "$proof\results.json" -Raw | ConvertFrom-Json
$res | Select-Object suite,duration_ms,passed,failed,skipped,fail_fast,repeat_count | Format-List
$res.cases | Where-Object status -ne 'passed' | Format-Table -AutoSize -Wrap

# Predecessor window around the broad failure after the focused proof is green.
$failedRun='C:\RSPerf\runs\20260707T184755Z-102276-c00ff4ec02eb45f6afc0715b10565031\artifacts\selftest\last_run\commands'
$d=Get-Content "$failedRun\results.json" -Raw | ConvertFrom-Json
$case='cmd_compare_directories_options_scroll_to_lower_cards_stays_stable'
$idx=-1
for ($i=0; $i -lt $d.cases.Count; $i++) {
  if ($d.cases[$i].name -eq $case) {
    $idx=$i
    break
  }
}
if ($idx -lt 0) { throw "Could not find $case in $failedRun" }
$start=[Math]::Max(0, $idx - 12)
$windowCases = (($start)..$idx | ForEach-Object { $d.cases[$_].name }) -join ','
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-20260707Tcompare-options-scroll-predecessor-window-after-scrollable-resize-fix'
$args=@('--commands-selftest',"--selftest-case=$windowCases",'--selftest-timeout-multiplier=2')
$p=Start-Process -FilePath (Resolve-Path '.\.build\x64\Debug\RedSalamander.exe') -ArgumentList $args -PassThru -Wait
Start-Sleep -Seconds 3
"processExit=$($p.ExitCode)"
$proof="C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands"
$res=Get-Content "$proof\results.json" -Raw | ConvertFrom-Json
$res | Select-Object suite,duration_ms,passed,failed,skipped,fail_fast,repeat_count | Format-List
$res.cases | Where-Object status -ne 'passed' | Format-Table -AutoSize -Wrap

# Broad gate only after focused/predecessor are green.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
Remove-Item Env:\REDSALAMANDER_TEST_RUN_ID -ErrorAction SilentlyContinue
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
```

## Superseded resume checkpoint - 2026-07-07 20:34 +02:00

Use this section first when resuming. The operation is not done until broad Commands and Suite Full
are green under the final root `C:\RSPerf`, final evidence is archived, durable requirements are
merged out of WIP, this plan is moved to `Specs/Plans/Done/`, and the intended stabilization files
are committed. Do not make flaky tests non-blocking. Do not create anonymous permanent quarantine.

### Done

- [x] Final proof root specified and used for accepted local proof: `C:\RSPerf`.
- [x] Broad Commands fail-fast is green under `C:\RSPerf`:
  `20260707T150904Z-93028-a711e3f4adc54622bb6671a8ab3b907c`, 786 passed / 0 failed / 2 skipped.
- [x] Find/menu broad-only failures were diagnosed and kept repeat-green after reverting the
  over-broad menu rewrite:
  `manual-20260707Tresume-repeat-find-menu-after-revert`, 40 passed / 0 failed across 20 repeats.
- [x] Preferences cluster fixes are retained: explicit focus/header waits, render-count settle for
  reverse category navigation, and source-contract guards. Last source-contract run after those
  guards was green: 99 passed / 0 failed.
- [x] Connection Manager long-run open/close broad blocker is root-caused, patched, source-contract
  guarded, rebuilt, and proven focused/predecessor/prefix-green:
  `manual-20260707Tconn-longrun-focused-after-settle-patch`,
  `manual-20260707Tconn-longrun-predecessor13-after-settle-patch`,
  `manual-20260707Tconn-longrun-prefix208-after-settle-patch`.
- [x] Broad Commands non-fail-fast after the Connection Manager patch removed that blocker:
  `20260707T175039Z-101860-05ee6a54a92b4ab8af658f176d7f21e5`, 785 passed / 1 failed / 2 skipped.
- [x] Swap Panes was investigated next. Focused and immediate predecessor proofs were green before
  and after adding diagnostic detail:
  `manual-20260707Tswap-panes-focused-continue`,
  `manual-20260707Tswap-panes-predecessor-window-continue`,
  `manual-20260707Tswap-panes-focused-after-diagnostics`,
  `manual-20260707Tswap-panes-predecessor-after-diagnostics`.
- [x] Diagnostic-only Swap Panes failure text is retained in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp` around
  `TestSwapPanesKeepsNavigationShellStable`: it now records `snapshotOk`, nav hwnd/visibility,
  expected path, pane model path, current nav path, plugin id/short id/instance context, focus/edit
  state, child windows, history counts, expected item visibility, name filter, and focused-folder
  match.
- [x] Rebuild after the Swap Panes diagnostic patch was green:
  `.build\logs\msbuild-20260707_201636_621.log`, 0 warnings / 0 errors.
- [x] The accidentally launched broad run without `REDSALAMANDER_TEST_ROOT=C:\RSPerf` was stopped
  and must not be accepted as evidence.
- [x] A corrected broad Commands diagnostic run was launched with
  `REDSALAMANDER_TEST_ROOT=C:\RSPerf` and without `-FailFast`:
  `C:\RSPerf\runs\20260707T182113Z-107496-68d5704fd5b64eef914e1da9ff068460\`.

### Remaining

- [ ] Treat broad Commands as red. The current first blocker is
  `cmd_preferences_dialog_category_tree_handles_reverse_keyboard_navigation`, not Swap Panes.
- [ ] Check whether the corrected diagnostic run is still active. At this save point, parent
  PowerShell PID `107496` and app PID `95116` were still running and responding.
- [ ] If the diagnostic run finished, parse its final `results.json`. If it is still running, either
  let it finish for extra clues or stop it intentionally before starting focused GUI probes.
- [ ] Root-cause and fix the Preferences reverse-navigation failure. Do not increase timeouts as the
  fix; the captured state shows VK_END did not route to a focused DxUi tree item at all.
- [ ] Add/adjust a source contract for the durable category-tree focus/key-routing invariant.
- [ ] Rebuild if C++ changed.
- [ ] Rerun source contracts if the guard changed.
- [ ] Rerun focused proof, predecessor window, and prefix or broad proof for the Preferences reverse
  navigation case under `C:\RSPerf`.
- [ ] Rerun broad Commands without `-FailFast` under `C:\RSPerf`.
- [ ] If broad Commands reaches Swap Panes again, use the enriched diagnostics from the retained
  Swap Panes patch; otherwise keep the earlier Swap failure as a secondary open risk.
- [ ] Run Suite Full under `C:\RSPerf` only after broad Commands is green.
- [ ] Archive final accepted evidence under `Specs/TestRuns/<current-head>/...` before cleaning
  volatile `C:\RSPerf` evidence.
- [ ] Merge durable requirements/lessons into authoritative specs or repo guidance.
- [ ] Move this plan to `Specs/Plans/Done/`.
- [ ] Commit only the intended stabilization files. Do not stage unrelated untracked review/spec
  files unless explicitly instructed.

### Current blocker details

Corrected diagnostic run:

```text
C:\RSPerf\runs\20260707T182113Z-107496-68d5704fd5b64eef914e1da9ff068460\
```

Partial state at 2026-07-07 20:34 +02:00:

```text
suite=Commands
passed=525
failed=1
skipped=0
fail_fast=false
first failure index=386
first failure name=cmd_preferences_dialog_category_tree_handles_reverse_keyboard_navigation
```

Failure text:

```text
VK_END should move Preferences reverse-navigation setup to the focused Advanced category;
capturedSnapshot=yes category=0 pageTitle='General' treeFocused=yes treeSelected=yes
selectedVisibleIndex=0 firstVisibleIndex=0 treeScrollDip=0 treeRender=1 pageScroll=0/0
visibleChildren=3 visibleCurrentPageChildren=1 currentPageDxHosts=1 pageResizeFailures=0
shellResizeFailures=0 categoryTreeResizeFailures=no
focus=0x29B31486 isWindow=yes class='Static' title='' ctrlId=2420 visible=yes childOfPrefs=yes
categoryTree=0x29B31486 isWindow=yes class='Static' title='' ctrlId=2420 visible=yes childOfPrefs=yes
activePage=0x71918A8 isWindow=yes class='RedSalamanderPrefsPageHost' title='' ctrlId=2421 visible=yes childOfPrefs=yes
shellHost=0xA9C17AC isWindow=yes class='RedSalamanderPrefsDxShellHost' title='' ctrlId=0 visible=yes childOfPrefs=yes
focusTargets(general=0, panes=0, hotPaths=0, advanced=0, monitor=0, compare=0, fileOps=0, themes=0, shell=0).
```

The important signal is blunt: Win32 focus is on the category-tree host, the selected category is
still General, and all page/shell focus-target probes are false. VK_END did nothing. This is not a
timing problem unless a real DxUi focus-control transition is missing from the wait predicate. Root
cause likely lives in category-tree DxUi focus/key routing or in the test's setup assumption that
`FocusWindowAndWait(categoryTreeHost)` is enough to give the DxUi `Tree` control keyboard focus.

### Investigation notes to keep

- Reverse-navigation test location:
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp`,
  `TestPreferencesDialogCategoryTreeReverseKeyboardNavigation`.
- That test sends `WM_KEYDOWN/WM_KEYUP` for `VK_END`, `VK_UP`, and `VK_HOME` directly to
  `IDC_PREFS_CATEGORY_LIST`. The failure occurs on the first `VK_END`.
- Nearby category-tree tests pass after this failure, including page navigation, wheel scrolling,
  expand/collapse, and boundary navigation from scrolled state. That means the tree is not globally
  dead; the reverse-navigation setup is missing a deterministic focus/key-routing precondition.
- `Preferences.Dialog.cpp` routes category host messages through
  `PreferencesDxCategoryHostWndProc` / `WindowHost::HandleMessage(...)`.
- `Common/DxUi/DxUi.Tree.cpp` already handles `VK_HOME` and `VK_END` in `Tree::OnKeyDown`. If key
  routing misses the `Tree`, fixing `Tree::OnKeyDown` is probably the wrong layer.
- Inspect `Common/DxUi` `WindowHost::HandleMessage`, `WM_SETFOCUS`, `SetFocusControl`, and any
  existing `DebugFocus...` pattern before patching.
- Preferred fix direction: make Preferences category-tree focus deterministic, either by adding a
  debug helper such as `DebugFocusPreferencesCategoryTree()` that sets Win32 focus and the DxUi
  focus control to the category `Tree`, or by fixing product `WM_SETFOCUS` handling if real keyboard
  navigation should always work when the host receives focus. Then use that helper/path before every
  synthetic category-tree key in the reverse-navigation test.
- Do not hide the case behind quarantine or non-blocking behavior.

### Exact resume commands

```powershell
# Confirm the previous diagnostic run is not still using the desktop.
Get-Process -Id 107496,95116 -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding

# Inspect the latest corrected broad diagnostic run.
$root='C:\RSPerf\runs\20260707T182113Z-107496-68d5704fd5b64eef914e1da9ff068460\artifacts\selftest\last_run\commands'
$d=Get-Content "$root\results.json" -Raw | ConvertFrom-Json
$d | Select-Object suite,duration_ms,passed,failed,skipped,fail_fast,repeat_count | Format-List
$failed = for ($i=0; $i -lt $d.cases.Count; $i++) {
  if ($d.cases[$i].status -eq 'failed') {
    [pscustomobject]@{index=$i; name=$d.cases[$i].name; duration_ms=$d.cases[$i].duration_ms; reason=$d.cases[$i].reason}
  }
}
$failed | Format-List
$idx=$failed[0].index
($idx-10)..($idx+10) | Where-Object { $_ -ge 0 -and $_ -lt $d.cases.Count } |
  ForEach-Object {
    [pscustomobject]@{
      index=$_
      name=$d.cases[$_].name
      status=$d.cases[$_].status
      duration_ms=$d.cases[$_].duration_ms
      reason=$d.cases[$_].reason
    }
  } | Format-Table -AutoSize -Wrap
Select-String -LiteralPath "$root\trace.txt" `
  -Pattern 'cmd_preferences_dialog_category_tree_handles_reverse_keyboard_navigation|reverse-navigation|VK_END|focusTargets' `
  -Context 4,12 | ForEach-Object { $_.ToString() }

# Read the suspected focus/key-routing layer.
rg -n "WindowHost::HandleMessage|WM_SETFOCUS|WM_KEYDOWN|SetFocusControl|GetFocusControl|OnKeyDown" `
  Common\DxUi -g "*.cpp" -g "*.h"
rg -n "PreferencesDxCategoryHostWndProc|DebugSelectPreferencesCategory|DebugGetPreferencesDialogSnapshot|IDC_PREFS_CATEGORY_LIST" `
  RedSalamander -g "*.cpp" -g "*.h"

# Focused proof after patching the Preferences reverse-navigation root cause.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-20260707Tprefs-reverse-focused-after-focus-routing-fix'
$case='cmd_preferences_dialog_category_tree_handles_reverse_keyboard_navigation'
$args=@('--commands-selftest',"--selftest-case=$case",'--selftest-timeout-multiplier=2')
$p=Start-Process -FilePath (Resolve-Path '.\.build\x64\Debug\RedSalamander.exe') -ArgumentList $args -PassThru -Wait
Start-Sleep -Seconds 3
"processExit=$($p.ExitCode)"
$proof="C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands"
$res=Get-Content "$proof\results.json" -Raw | ConvertFrom-Json
$res | Select-Object suite,duration_ms,passed,failed,skipped,fail_fast,repeat_count | Format-List
$res.cases | Where-Object status -ne 'passed' | Format-Table -AutoSize -Wrap

# Predecessor window around the broad failure after the focused proof is green.
$idx=-1
for ($i=0; $i -lt $d.cases.Count; $i++) {
  if ($d.cases[$i].name -eq $case) {
    $idx=$i
    break
  }
}
if ($idx -lt 0) { throw "Could not find $case in $root" }
$start=[Math]::Max(0, $idx - 12)
$windowCases = (($start)..$idx | ForEach-Object { $d.cases[$_].name }) -join ','
$env:REDSALAMANDER_TEST_RUN_ID='manual-20260707Tprefs-reverse-predecessor-window-after-focus-routing-fix'
$args=@('--commands-selftest',"--selftest-case=$windowCases",'--selftest-timeout-multiplier=2')
$p=Start-Process -FilePath (Resolve-Path '.\.build\x64\Debug\RedSalamander.exe') -ArgumentList $args -PassThru -Wait
Start-Sleep -Seconds 3
"processExit=$($p.ExitCode)"
$proof="C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands"
$res=Get-Content "$proof\results.json" -Raw | ConvertFrom-Json
$res | Select-Object suite,duration_ms,passed,failed,skipped,fail_fast,repeat_count | Format-List
$res.cases | Where-Object status -ne 'passed' | Format-Table -AutoSize -Wrap
```

## Superseded resume checkpoint - 2026-07-07 20:11 +02:00

Use this section first when resuming. The operation is not done until broad Commands and Suite Full
are green under the final root `C:\RSPerf`, final evidence is archived, durable requirements are
merged out of WIP, this plan is moved to `Specs/Plans/Done/`, and the intended stabilization files
are committed. Do not make flaky tests non-blocking. Do not create anonymous permanent quarantine.

### At-a-glance checklist

- [x] Final proof root specified and used for accepted local proof: `C:\RSPerf`.
- [x] Broad Commands fail-fast is green under `C:\RSPerf`.
- [x] Preferences cluster is root-caused, fixed, source-contract guarded, rebuilt, and local tail
  replay-green.
- [x] Connection Manager long-run open/close broad blocker is root-caused, fixed, source-contract
  guarded, rebuilt, and focused/predecessor/prefix-green.
- [x] Source contracts are green after current retained guards: 99 passed / 0 failed.
- [ ] Broad Commands without `-FailFast` is not green yet. Current first failure is
  `cmd_app_swapPanes_keeps_navigation_shell_stable`.
- [ ] Swap Panes first failure still needs focused/predecessor/prefix proof, deterministic root
  cause, patch, source guard if applicable, rebuild, and broad rerun.
- [ ] Suite Full under `C:\RSPerf` has not been run after the latest blockers because Commands is
  still red.
- [ ] Final accepted evidence has not yet been archived under `Specs/TestRuns/<current-head>/...`.
- [ ] Durable requirements still need to be merged into authoritative specs/repo guidance before
  closeout.
- [ ] This WIP plan has not been moved to `Specs/Plans/Done/`.
- [ ] Intended files have not been committed. Do not stage unrelated untracked review/spec files.

### Proven now

- [x] Final proof root is specified and unchanged: use `REDSALAMANDER_TEST_ROOT=C:\RSPerf` for all
  accepted local proof runs.
- [x] Checkout at save point: `cb2478689bc1b37afc6066a6a5994d8f135a56cc` plus uncommitted
  stabilization edits.
- [x] Broad Commands fail-fast is green under `C:\RSPerf`:
  `20260707T150904Z-93028-a711e3f4adc54622bb6671a8ab3b907c`, 786 passed / 0 failed / 2 skipped.
- [x] The earlier Find/menu broad-only failures are captured and repeat-green after retained Find
  diagnostic work and reverted menu rewrites:
  `manual-20260707Tresume-repeat-find-menu-after-revert`, 40 passed / 0 failed across 20 repeats.
- [x] Broad Commands non-fail-fast was rerun after the Find/menu checkpoint and exposed a new
  Preferences cluster:
  `20260707T160926Z-47116-80ce577e8ced42129611faa09aaaa513`,
  782 passed / 4 failed / 2 skipped.
- [x] The four Preferences failures from that broad run reproduced as focused-green before patch:
  `manual-20260707Tresume-focused-four-prefs-failures-before-patch`, 4 passed / 0 failed.
- [x] Retained Preferences patches are saved:
  Viewers sort/search waits for visible Match/F3 header rects before reorder/resize interaction;
  Themes retained-search and Compare Directories paths use `FocusWindowAndWait(...)`;
  Editors/Mouse tab traversal uses the shared focus helper, emits pane/focus diagnostics, and uses a
  5000 ms snapshot wait; reverse category-tree navigation waits for render-count settle before
  accepting category/page state.
- [x] Source contracts were extended to guard the new Preferences focus/header/wait contracts. The
  temporary reverse-navigation anchor typo was fixed later in this checkpoint.
- [x] Focused proof after the first Preferences patch is green:
  `manual-20260707Tresume-focused-four-prefs-after-focus-header-patch`, 4 passed / 0 failed.
- [x] Local predecessor Preferences windows after the first patch are green:
  `manual-20260707Tresume-prefs-failure-predecessor-windows-after-patch`, 37 passed / 0 failed.
- [x] Latest completed build before the reverse-navigation render-settle patch passed:
  `.build\logs\msbuild-20260707_185426_971.log`, 0 warnings / 0 errors.
- [x] Broad Commands non-fail-fast after the first Preferences patch is archived at
  `C:\RSPerf\runs\20260707T163528Z-10892-33089aa6332846628e5fda42dde57ad7\`,
  768 passed / 18 failed / 2 skipped. Treat later failures as potentially contaminated until
  independently reproduced.
- [x] Local tail replay after the Editors/Mouse wait change is archived at
  `C:\RSPerf\runs\manual-20260707Tresume-prefs-tail-first-failures-after-tab-wait\`,
  17 passed / 1 failed. Editors/Mouse and category expand passed; the remaining local failure was
  `cmd_preferences_dialog_category_tree_handles_reverse_keyboard_navigation`.
- [x] Reverse category-tree render-settle patch is saved after that local tail failure.
- [x] Source-contract anchor typo was fixed:
  `TestPreferencesDialogCategoryTreeReverseKeyboardNavigation`.
- [x] Source contracts after the typo fix were green: 98 passed / 0 failed.
- [x] Rebuild after adding `std::chrono_literals` scope for the reverse-navigation settle patch was
  green: `.build\logs\msbuild-20260707_190348_705.log`, 0 warnings / 0 errors.
- [x] Local Preferences tail after reverse render-settle was green:
  `manual-20260707Tresume-prefs-tail-after-reverse-render-settle`, 18 passed / 0 failed.
- [x] Source contracts after the C++ literal-scope fix were green: 98 passed / 0 failed.
- [x] Broad Commands non-fail-fast after the Preferences fixes exposed a new first blocker:
  `20260707T170806Z-105112-adbb86e4e6d8436c9970a5ac8b990fcc`,
  778 passed / 8 failed / 2 skipped. First failure:
  `cmd_connection_manager_window_long_run_open_close_stays_stable`.
- [x] Connection Manager focused, predecessor-13, and prefix-208 probes before the patch were green,
  proving the broad failure needed deterministic open/close-settle hardening rather than quarantine.
- [x] Connection Manager long-run patch is saved in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`: per-cycle singleton cleanup,
  shortcut-command dispatch, stronger close-clear assertion, and diagnostic-rich open failure text.
- [x] Rebuild after the Connection Manager patch was green:
  `.build\logs\msbuild-20260707_194351_614.log`, 0 warnings / 0 errors.
- [x] Connection Manager post-patch proof ladder is green:
  `manual-20260707Tconn-longrun-focused-after-settle-patch`, 1 passed / 0 failed;
  `manual-20260707Tconn-longrun-predecessor13-after-settle-patch`, 13 passed / 0 failed;
  `manual-20260707Tconn-longrun-prefix208-after-settle-patch`, 208 passed / 0 failed.
- [x] Source contract added for the Connection Manager long-run open/close invariants.
- [x] Source contracts after the Connection Manager guard were green: 99 passed / 0 failed.
- [x] Broad Commands non-fail-fast after the Connection Manager patch removed that blocker and left
  one first failure:
  `20260707T175039Z-101860-05ee6a54a92b4ab8af658f176d7f21e5`,
  785 passed / 1 failed / 2 skipped.

### Current blocker details

The current first failure is:

```text
cmd_app_swapPanes_keeps_navigation_shell_stable
Navigation shell did not stay quiet during Swap Panes (left pane after swap);
pane=0, focusTarget=0, editMode=no, historyVisible=no, suggestVisible=no, popupVisible=no,
childWindows=0, currentPath='', historyCount=0, itemCount=1, nameFilterActive=no.
```

Latest failing evidence:

```text
C:\RSPerf\runs\20260707T175039Z-101860-05ee6a54a92b4ab8af658f176d7f21e5\
```

The failure shape is specific: after Swap Panes, the left pane navigation-shell snapshot is quiet in
terms of edit/history/suggest/popup state, but the pane path is empty and the item count is only 1.
Do not treat this as a generic flake. The next session should determine whether Swap Panes needs
deterministic path/item setup, a wait for post-swap pane stabilization, or a product/test invariant
fix around the left pane model after swapping.

### Next actions

- [ ] Inspect the latest run JSON and trace around
  `cmd_app_swapPanes_keeps_navigation_shell_stable`.
- [ ] Locate the Swap Panes test and product command path.
- [ ] Reproduce the failure shape with a focused case, then predecessor window, then prefix only if
  focused/predecessor are green.
- [ ] Patch the deterministic root cause. Do not make the case non-blocking and do not add anonymous
  quarantine.
- [ ] Add or adjust a source contract if the fix creates a new durable test-harness invariant.
- [ ] Rebuild if C++ changed.
- [ ] Rerun source contracts if source-contract guards changed.
- [ ] Rerun focused/predecessor/prefix proof for Swap Panes.
- [ ] Rerun broad Commands without `-FailFast` under `C:\RSPerf`.
- [ ] Run Suite Full under `C:\RSPerf` only after broad Commands is green.
- [ ] Archive final accepted evidence under `Specs/TestRuns/<current-head>/...` before cleaning
  volatile `C:\RSPerf` evidence.
- [ ] Merge durable requirements/lessons into authoritative specs or repo guidance.
- [ ] Move this plan to `Specs/Plans/Done/`.
- [ ] Commit only the intended stabilization files.

### Exact resume commands

```powershell
# Confirm no orphaned app/build process is holding the interactive desktop.
Get-Process RedSalamander,MSBuild,cl,link -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding

# Inspect the current first broad Commands failure.
$root='C:\RSPerf\runs\20260707T175039Z-101860-05ee6a54a92b4ab8af658f176d7f21e5\artifacts\selftest\last_run\commands'
$d=Get-Content "$root\results.json" -Raw | ConvertFrom-Json
$failed = for ($i=0; $i -lt $d.cases.Count; $i++) {
  if ($d.cases[$i].status -eq 'failed') {
    [pscustomobject]@{index=$i; name=$d.cases[$i].name; duration_ms=$d.cases[$i].duration_ms; reason=$d.cases[$i].reason}
  }
}
$failed | Format-List
$idx=$failed[0].index
($idx-8)..($idx+8) | Where-Object { $_ -ge 0 -and $_ -lt $d.cases.Count } |
  ForEach-Object { [pscustomobject]@{index=$_; name=$d.cases[$_].name; status=$d.cases[$_].status; duration_ms=$d.cases[$_].duration_ms; reason=$d.cases[$_].reason} } |
  Format-Table -AutoSize -Wrap
Select-String -LiteralPath "$root\trace.txt" -Pattern 'cmd_app_swapPanes_keeps_navigation_shell_stable|Swap Panes|swapPanes|Navigation shell' -Context 4,10 |
  ForEach-Object { $_.ToString() }

# Locate the test and command path.
rg -n "swapPanes_keeps_navigation_shell_stable|Swap Panes|swapPanes|cmd/app/swapPanes|IDM_APP_SWAP|SwapPanes" `
  RedSalamander\SelfTest\Commands RedSalamander -g "*.cpp" -g "*.h"

# Build a focused/predecessor proof set from the failed run once the case index is known.
$failedRun='C:\RSPerf\runs\20260707T175039Z-101860-05ee6a54a92b4ab8af658f176d7f21e5\artifacts\selftest\last_run\commands'
$d=Get-Content "$failedRun\results.json" -Raw | ConvertFrom-Json
$caseName='cmd_app_swapPanes_keeps_navigation_shell_stable'
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-20260707Tswap-panes-focused'
$args=@('--commands-selftest',"--selftest-case=$caseName",'--selftest-timeout-multiplier=2')
$p=Start-Process -FilePath (Resolve-Path '.\.build\x64\Debug\RedSalamander.exe') -ArgumentList $args -PassThru -Wait
Start-Sleep -Seconds 3
"processExit=$($p.ExitCode)"
$root="C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands"
$res=Get-Content "$root\results.json" -Raw | ConvertFrom-Json
$res | Select-Object suite,duration_ms,passed,failed,skipped,fail_fast,repeat_count | Format-List
$res.cases | Where-Object status -ne 'passed' | Format-Table -AutoSize

# Predecessor window around the broad failure.
$idx=-1
for ($i=0; $i -lt $d.cases.Count; $i++) {
  if ($d.cases[$i].name -eq $caseName) {
    $idx=$i
    break
  }
}
if ($idx -lt 0) { throw "Could not find $caseName in $failedRun" }
$start=[Math]::Max(0, $idx - 12)
$windowCases = (($start)..$idx | ForEach-Object { $d.cases[$_].name }) -join ','
$env:REDSALAMANDER_TEST_RUN_ID='manual-20260707Tswap-panes-predecessor-window'
$args=@('--commands-selftest',"--selftest-case=$windowCases",'--selftest-timeout-multiplier=2')
$p=Start-Process -FilePath (Resolve-Path '.\.build\x64\Debug\RedSalamander.exe') -ArgumentList $args -PassThru -Wait
Start-Sleep -Seconds 3
"processExit=$($p.ExitCode)"
$root="C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands"
$res=Get-Content "$root\results.json" -Raw | ConvertFrom-Json
$res | Select-Object suite,duration_ms,passed,failed,skipped,fail_fast,repeat_count | Format-List
$res.cases | Where-Object status -ne 'passed' | Format-Table -AutoSize

# Acceptance gates after the Swap Panes root cause is patched. Do not run GUI/native selftests in parallel.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2

$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

### Resume cautions

- [ ] Direct GUI selftests must be launched with `Start-Process -Wait`; direct `& exe` probes can
  return before the GUI process finishes and create weak or invalid evidence.
- [ ] After native GUI selftests, wait 3-6 seconds, then read settled
  `last_run\results.json` and `last_run\commands\results.json`.
- [ ] Shell exit code can be 0 even when a script prints `runnerExit=1`; inspect JSON and summaries.
- [ ] Pester can return control despite failures; inspect the pass/fail summary.
- [ ] In non-fail-fast GUI runs, treat later failures after the first failure as potentially
  contaminated until reproduced independently.
- [ ] Do not clean `C:\RSPerf` until useful manual and broad evidence has been archived.
- [ ] The current `C:\RSPerf` disk audit reports many `unexpected-test-run-dir` issues because
  useful manual runs have not yet been archived/cleaned. Archive useful evidence before removing old
  manual run directories.

### Worktree at save point

- [ ] Modified stabilization files visible in `git status`:
  `Common/Common/SettingsStore.cpp`,
  `Common/DxUi/DxUi.Menu.cpp`,
  `Common/DxUi/DxUi.h`,
  `RedSalamander/ManagePluginsDialog.cpp`,
  `RedSalamander/RedSalamander.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ThemesGeneralPanes.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`,
  `Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md`,
  `Specs/Plans/WIP/README.md`,
  `Tests/DxUiTests/DxUiTests.Menu.cpp`,
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`,
  and `Tools/Tests/WingetValidation.Tests.ps1`.
- [ ] Untracked review/spec material visible in `git status`:
  `Specs/Core/Core_FileSystemBridge.md`,
  `Specs/Plans/WIP/FileSystem_CrossFsBridgeImplementationAndPerformanceRedesign_2026-07-07.md`,
  `Specs/Plans/Done/Operation_Causeway_CrossFsBridgeSecurityAndPerformanceRemediation_2026-07-07.md`,
  `Specs/Plans/WIP/UI_FileOperationsPopupUxRefinementPlan_2026-07-07.md`,
  `Specs/Reviews/FileSystemBridge-2026-07-07-Findings.md`,
  and `Specs/Reviews/ThreeDayDiff-2026-07-06-Findings.md`. Do not stage/delete these unless they
  are intentionally part of the final scope.
- [ ] Broad Commands and Suite Full are still not accepted. Do not move this plan to Done yet.
- [ ] Do not stage or commit unrelated untracked review material during the final commit.

## Superseded resume checkpoint - 2026-07-07 18:07 +02:00

Historical checkpoint retained for audit only. Use the topmost current resume checklist when
resuming. At this point in history the operation was still not done: broad Commands and Suite Full
still needed green evidence under the final root `C:\RSPerf`, final evidence still needed archival,
durable requirements still needed to move out of WIP, this plan still needed to move to
`Specs/Plans/Done/`, and the intended stabilization files still needed to be committed. Flaky tests
were still required to be identified and fixed or replaced, not made non-blocking.

### At-a-glance checklist

- [x] Final proof root is specified and unchanged: use `REDSALAMANDER_TEST_ROOT=C:\RSPerf` for all
  accepted local proof runs.
- [x] Checkout at save point: `cb2478689bc1b37afc6066a6a5994d8f135a56cc` plus uncommitted
  stabilization edits.
- [x] No `RedSalamander.exe`, `MSBuild`, `cl`, or `link` process was listed at the 18:07 save
  check.
- [x] Preferences shared focus helper upgrade is saved and source-contract guarded.
- [x] Plugins header-rect wait guard is saved and source-contract guarded.
- [x] Keyboard deferred-search broad blocker was investigated and did not reproduce focused,
  in the retained-state predecessor pair, in the tail-25 window, or in the next full Commands
  fail-fast gate.
- [x] Broad Commands fail-fast is green under `C:\RSPerf`:
  `20260707T150904Z-93028-a711e3f4adc54622bb6671a8ab3b907c`, 786 passed / 0 failed / 2 skipped.
- [x] Source contracts pass after the latest retained guard:
  `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`,
  96 passed / 0 failed.
- [x] Latest build passed after the retained source edits:
  `.build\logs\msbuild-20260707_180227_176.log`, 0 warnings / 0 errors.
- [x] Broad Commands without `-FailFast` exposed the current unresolved gate:
  `20260707T152901Z-73760-075062df9a2942b78aada8ba8b1ad22e`, 784 passed / 2 failed / 2 skipped.
- [x] Find header-drag failure path now captures `DescribeFindSnapshotBrief(snapshot)`, requires
  `value.usesDxUiHost`, and is guarded by
  `keeps Find header-drag reorder failures diagnostic-rich` in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`.
- [x] The rejected menu rewrite attempts were reverted. Do not resurrect them blindly: synchronous
  click closed the temporary menu, mnemonic pre-open did not leave the popup visible, and the
  modal-worker mnemonic path opened a popup but did not exercise hover switching.
- [x] Focused/repeat proof after the menu revert and Find diagnostic patch is green:
  `manual-20260707Tresume-focused-find-and-menu-reverted-with-find-diag` passed 2 / 0, and
  `manual-20260707Tresume-repeat-find-menu-after-revert` passed 40 / 0 across 20 repeats of each
  failing case.
- [ ] Broad Commands without `-FailFast` has not passed after the latest retained state.
- [ ] Suite Full has not passed after the latest retained state.
- [ ] Final passing evidence has not been archived to `Specs/TestRuns/...`.
- [ ] This WIP plan has not been moved to `Specs/Plans/Done/`.
- [ ] Durable lessons still need to be merged into authoritative specs/guidance before closeout.
- [ ] Final intentional commit has not been made.

### Current evidence ledger

- [x] Keyboard old blocker:
  `manual-20260707Tresume-focused-keyboard-search-action` passed 1 / 0,
  `manual-20260707Tresume-keyboard-retained-then-deferred` passed 2 / 0, and
  `manual-20260707Tresume-tail25-to-keyboard-deferred` passed 25 / 0.
- [x] Fresh broad Commands fail-fast after Keyboard investigation:
  `C:\RSPerf\runs\20260707T150904Z-93028-a711e3f4adc54622bb6671a8ab3b907c\`,
  process exit 0, 786 passed / 0 failed / 2 skipped, elapsed 17m 33.6s.
- [x] Broad Commands non-fail-fast current red run:
  `C:\RSPerf\runs\20260707T152901Z-73760-075062df9a2942b78aada8ba8b1ad22e\`,
  process exit 1, 784 passed / 2 failed / 2 skipped, elapsed 17m 50.1s. Failures:
  `cmd_pane_find_dialog_header_drag_reorders_columns_without_sort` and
  `cmd_app_menuBar_hover_switches_top_level_popup`.
- [x] Focused proof for those two failures before/after retained edits:
  `manual-20260707Tresume-focused-find-header-drag` passed 1 / 0,
  `manual-20260707Tresume-focused-menubar-hover` passed 1 / 0,
  `manual-20260707Tresume-find-header-predecessor9` passed 9 / 0,
  `manual-20260707Tresume-menubar-hover-predecessor9` passed 9 / 0,
  `manual-20260707Tresume-focused-find-and-menu-reverted-with-find-diag` passed 2 / 0, and
  `manual-20260707Tresume-repeat-find-menu-after-revert` passed 40 / 0.
- [x] Latest retained code evidence:
  source contracts 96 / 0 and build `.build\logs\msbuild-20260707_180227_176.log`.

### Current blocker / next investigation

- [ ] Rerun broad Commands without `-FailFast` under `C:\RSPerf`; this is the next acceptance gate.
- [ ] If Find fails again, use the new diagnostic text from
  `DescribeFindSnapshotBrief(snapshot)` before changing behavior.
- [ ] If menu hover fails again, inspect whether the first click loses selection, whether the
  initial popup handle is null, whether hover delivery happened, and whether the replacement popup
  was observed. Do not repeat the three rejected rewrite approaches listed above.
- [ ] If broad Commands passes, run Suite Full under `C:\RSPerf`.
- [ ] If Suite Full still fails, investigate the previously seen `RedSalamanderMonitorEtwLatency`
  empty-output exit 1 and `ToolsPesterTests` empty-output exit 1 before closeout.
- [ ] Once all gates pass, archive the accepted evidence under `Specs/TestRuns/<current-head>/...`
  before cleaning volatile `C:\RSPerf` evidence.

### Exact resume commands

```powershell
# Confirm no orphaned app/build process is holding the interactive desktop.
Get-Process RedSalamander,MSBuild,cl,link -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding

# Confirm the suspended repeat proof, already observed green at the save point.
$repeat='C:\RSPerf\runs\manual-20260707Tresume-repeat-find-menu-after-revert\artifacts\selftest\last_run\commands'
Get-Content "$repeat\results.json" -Raw
Get-Content "$repeat\trace.txt" -Tail 80

# Baseline retained source state.
Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander

# Next acceptance gates. Do not run GUI/native selftests in parallel.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2

$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

### Resume cautions

- [ ] Direct GUI selftests must be launched with `Start-Process -Wait`; direct `& exe` probes can
  return before the GUI process finishes and create weak or invalid evidence.
- [ ] After native GUI selftests, wait 3-6 seconds, then read settled
  `last_run\results.json` and `last_run\commands\results.json`.
- [ ] Shell exit code can be 0 even when a script prints `runnerExit=1`; inspect JSON and summaries.
- [ ] Pester can return control despite failures; inspect the pass/fail summary.
- [ ] In non-fail-fast GUI runs, treat later failures after the first failure as potentially
  contaminated until reproduced independently.
- [ ] Do not clean `C:\RSPerf` until useful manual and broad evidence has been archived.

### Worktree at save point

- [ ] Modified stabilization files visible in `git status`:
  `Common/Common/SettingsStore.cpp`,
  `Common/DxUi/DxUi.Menu.cpp`,
  `Common/DxUi/DxUi.h`,
  `RedSalamander/ManagePluginsDialog.cpp`,
  `RedSalamander/RedSalamander.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`,
  `Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md`,
  `Specs/Plans/WIP/README.md`,
  `Tests/DxUiTests/DxUiTests.Menu.cpp`,
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`,
  and `Tools/Tests/WingetValidation.Tests.ps1`.
- [ ] Untracked review/spec material visible in `git status`:
  `Specs/Core/Core_FileSystemBridge.md`,
  `Specs/Plans/WIP/FileSystem_CrossFsBridgeImplementationAndPerformanceRedesign_2026-07-07.md`,
  `Specs/Plans/Done/Operation_Causeway_CrossFsBridgeSecurityAndPerformanceRemediation_2026-07-07.md`,
  `Specs/Reviews/FileSystemBridge-2026-07-07-Findings.md`,
  and `Specs/Reviews/ThreeDayDiff-2026-07-06-Findings.md`. Do not stage/delete these unless they
  are intentionally part of the final scope.
- [ ] Broad Commands and Suite Full are still not accepted. Do not move this plan to Done yet.
- [ ] Do not stage or commit unrelated untracked review material during the final commit.

## Superseded resume checkpoint - 2026-07-07 17:03 +02:00

Kept for history only. Use the current 18:07 checkpoint above when resuming. This older checkpoint
predates the Keyboard non-repro proof, the green Commands fail-fast gate, the broad non-fail-fast
Find/menu failures, and the retained Find diagnostic guard.

### At-a-glance checklist

- [x] Final proof root is specified and unchanged: use `REDSALAMANDER_TEST_ROOT=C:\RSPerf` for all
  accepted local proof runs.
- [x] Current checkout at save point: `cb2478689bc1b37afc6066a6a5994d8f135a56cc`.
- [x] No `RedSalamander.exe`, `MSBuild`, `cl`, or `link` process was running at this save point.
- [x] The shared Preferences focus helper has been reviewed against the stronger
  Search/ViewCommands foreground patterns and upgraded to use `AttachThreadInput(...)`,
  root activation, topmost nudge, `SetFocus(...)`, and pumped retry.
- [x] Plugins roundtrip focus failure has focused, predecessor, and exact-prefix green proof under
  `C:\RSPerf` after the focus-helper upgrade.
- [x] Plugins reordered/resized/sort/search header failure has focused, predecessor, and
  exact-prefix green proof under `C:\RSPerf` after the header-rect wait guard.
- [x] Source contracts passed after the latest guard additions: 95 passed / 0 failed.
- [x] Latest build passed after the latest source edits:
  `.build\logs\msbuild-20260707_164930_113.log`, 0 warnings / 0 errors.
- [x] Latest broad Commands fail-fast evidence is saved at
  `C:\RSPerf\runs\20260707T145553Z-71864-2cfa41388a354d26aff55a14b2f9eb5b\`:
  process exit 1, 330 passed / 1 failed / 457 skipped.
- [ ] 17:03 blocker to root-cause next at that older checkpoint:
  `cmd_preferences_dialog_keyboard_search_action_updates_dxui_surface`.
- [ ] Do not patch the Keyboard blocker from suspicion; inspect the saved trace/result JSON, then
  reproduce focused, local predecessor, and exact-prefix behavior.
- [ ] At 17:03, broad Commands fail-fast was still red.
- [ ] Broad Commands without `-FailFast` is still not accepted.
- [ ] Suite Full is still not accepted.
- [ ] Final evidence has not been archived to `Specs/TestRuns/...`.
- [ ] This WIP plan has not been moved to `Specs/Plans/Done/`.
- [ ] A final intentional commit has not been made.

### Done / saved

- [x] Earlier root-cause fixes remain saved in this WIP plan: SettingsStore test-root isolation,
  rapid Preferences fixture seeding, stale Preferences category-tree UIA selection wait,
  plugin-config logical-focus recovery, preview fallback atomic wait, and Quick Search pane
  sort/extension restoration.
- [x] Hot Paths browse remains diagnostic-only. It did not reproduce focused or in local predecessor
  proof: `manual-20260707T1530-focused-hotpaths-browse` passed 1 / 0 / 0,
  `manual-20260707T1531-hotpaths-browse-predecessor-cluster` passed 13 / 0 / 0, and
  `manual-20260707T1532-preferences-80-to-hotpaths-browse` passed 81 / 0 / 0.
- [x] Viewers theme-cycle did not reproduce focused or in its immediate Preferences predecessor
  window: `manual-20260707T1545-focused-viewers-theme-cycle` passed 1 / 0 / 0 and
  `manual-20260707T1546-pref-window-to-viewers-theme-cycle` passed 12 / 0 / 0.
- [x] Viewers roundtrip category navigation now uses condition-based
  `FocusWindowAndWait(...)` instead of two one-shot raw `SetFocus(...)` assertions.
- [x] Shared `FocusWindowAndWait(...)` in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.cpp` was upgraded from the
  provisional root activation retry to the proven foreground pattern: resolve root, accept child
  focus, attach the foreground thread with `AttachThreadInput(...)`, call `ShowWindow(...)`,
  `BringWindowToTop(...)`, `SetActiveWindow(...)`, `SetForegroundWindow(...)`, `SetFocus(...)`,
  topmost/not-topmost nudge, `UpdateWindow(...)`, and pump/retry.
- [x] Source-contract guard for Preferences focus acquisition now checks the stronger helper shape,
  including `targetHasFocus`, `IsChild`, `AttachThreadInput(TRUE/FALSE)`, `ShowWindow`,
  `BringWindowToTop`, `SetForegroundWindow`, topmost/not-topmost `SetWindowPos`, `SetFocus`, and
  retry-loop behavior.
- [x] Focus helper validation after the upgrade:
  `manual-20260707T1634-focused-plugins-roundtrip-attach-focus` passed 1 / 0 / 0,
  `manual-20260707T1634-plugins-roundtrip-predecessor-attach-focus` passed 9 / 0 / 0,
  `manual-20260707T1634-focused-themes-roundtrip-attach-focus` passed 1 / 0 / 0,
  `manual-20260707T1634-focused-viewers-roundtrip-attach-focus` passed 1 / 0 / 0, and
  `manual-20260707T1636-prefix-to-plugins-roundtrip-attach-focus` passed 230 / 0 / 0.
- [x] Broad Commands fail-fast after the focus-helper upgrade moved the first failure to Plugins
  header geometry:
  `20260707T143801Z-100144-8417d280f1854434b21362831bd406a6`,
  process exit 1, 253 passed / 1 failed / 534 skipped. First failure:
  `cmd_preferences_dialog_plugins_reordered_resized_columns_survive_sort_cycles_and_search_roundtrip`,
  reason `Failed to capture the visible Preferences Plugins Name header rect before
  reordered-resized-sort/search validation.`
- [x] Plugins header failure did not reproduce focused, in a 5-case predecessor window, or in the
  exact 254-case prefix before the patch:
  `manual-20260707T1643-focused-plugins-reordered-resized-sort-search` passed 1 / 0 / 0,
  `manual-20260707T1645-plugins-header-predecessor-to-sort-search` passed 5 / 0 / 0, and
  `manual-20260707T1646-prefix-to-plugins-reordered-resized-sort-search` passed 254 / 0 / 0.
- [x] Plugins reordered/resized/sort/search now waits for visible, ordered Name/Type header rects
  before column drag/sort/search interaction, and the category-tree focus path uses
  `FocusWindowAndWait(...)` with diagnostic-rich failure text.
- [x] Source-contract guard
  `waits for Plugins sort-search header rects before reorder-resize interactions` now locks the
  Plugins header wait shape.
- [x] Source contracts passed after the Plugins header guard:
  `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`,
  95 passed / 0 failed.
- [x] Build initially caught a real diagnostic bug in the Plugins header patch: `std::format` was
  handed `PrefCategory`. The diagnostic now casts enum values to `int`.
- [x] Rebuild after the diagnostic fix passed:
  `.build\logs\msbuild-20260707_164930_113.log`, 0 warnings / 0 errors.
- [x] Plugins header validation after the patch:
  `manual-20260707T1652-focused-plugins-sort-search-header-wait` passed 1 / 0 / 0,
  `manual-20260707T1652-plugins-header-predecessor-header-wait` passed 5 / 0 / 0, and
  `manual-20260707T1653-prefix-to-plugins-sort-search-header-wait` passed 254 / 0 / 0.
- [x] Latest broad Commands fail-fast after the Plugins header fix is saved at
  `C:\RSPerf\runs\20260707T145553Z-71864-2cfa41388a354d26aff55a14b2f9eb5b\`:
  process exit 1, 330 passed / 1 failed / 457 skipped. First failure:
  `cmd_preferences_dialog_keyboard_search_action_updates_dxui_surface`, reason:
  `Preferences Keyboard page did not settle before deferred-search validation.`
- [x] Direct GUI selftests must be launched with `Start-Process -Wait`; direct `& exe` probes can
  return before the GUI process finishes and create weak or invalid evidence.

### Current blocker / next investigation

- [ ] Inspect the latest broad Keyboard trace and result JSON:
  `C:\RSPerf\runs\20260707T145553Z-71864-2cfa41388a354d26aff55a14b2f9eb5b\artifacts\selftest\last_run\commands\`.
- [ ] Locate and read
  `cmd_preferences_dialog_keyboard_search_action_updates_dxui_surface`, its page-settle helper, and
  the deferred-search assertion path before editing.
- [ ] Run focused
  `cmd_preferences_dialog_keyboard_search_action_updates_dxui_surface` under `C:\RSPerf` with
  `Start-Process -Wait`.
- [ ] If focused Keyboard passes, derive the immediate predecessor window from the saved broad
  result JSON and replay that predecessor cluster under `C:\RSPerf`.
- [ ] If the predecessor cluster passes, replay the exact broad prefix ending at the Keyboard case
  and narrow the minimal predecessor window. Do not patch from suspicion.
- [ ] If focused Keyboard fails, inspect `commands\trace.txt`, result JSON, and Preferences debug
  snapshot text before changing code.
- [ ] If the root cause is a wait/diagnostic contract issue, add or update source-contract guards in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` alongside the C++ fix.

### Remaining to finish this operation

- [ ] Root-cause and fix or replace
  `cmd_preferences_dialog_keyboard_search_action_updates_dxui_surface`; do not classify it as flaky
  and move on.
- [ ] Re-run source contracts and require 0 failures.
- [ ] Rebuild `RedSalamander`.
- [ ] Re-run focused Keyboard search-action and its predecessor cluster under `C:\RSPerf`.
- [ ] Re-run the exact latest broad prefix under `C:\RSPerf`.
- [ ] Re-run broad Commands fail-fast under `C:\RSPerf`.
- [ ] If broad Commands finds another broad-only failure, identify it, root-cause it, and fix or
  replace the brittle path. Do not label it flaky and move on.
- [ ] If `cmd_connection_manager_window_rejects_blank_profile_name` reappears, use the new
  diagnostic snapshot traces before changing behavior.
- [ ] If `cmd_preferences_dialog_hot_paths_browse_live_dx_interaction` reappears, use the new
  diagnostic-rich snapshot text before changing behavior.
- [ ] Re-run broad Commands without `-FailFast` under `C:\RSPerf`.
- [ ] Re-run Suite Full under `C:\RSPerf`.
- [ ] If still failing, explain or fix `RedSalamanderMonitorEtwLatency`; a prior Suite Full saw exit
  1 with an empty output log.
- [ ] If still failing, explain or fix `ToolsPesterTests`; a prior Suite Full saw exit 1 with no
  captured output log.
- [ ] Archive final passing evidence under `Specs/TestRuns/<current-head>\...` or another clearly
  named continuation folder before cleaning volatile `C:\RSPerf` evidence.
- [ ] Clean or reset `C:\RSPerf` only after needed evidence is archived; current manual run folders
  make disk-audit checks dirty.
- [ ] Before closeout, move this plan to `Specs/Plans/Done/` and merge durable requirements into
  `Specs/Testing/*`, `Tests/README.md`, `README.md`, or `AGENTS.md` as appropriate.
- [ ] Commit only intentional stabilization files; do not stage unrelated review material.

### Exact resume commands

```powershell
# Confirm no orphaned app/build process is holding the interactive desktop.
Get-Process RedSalamander,MSBuild,cl,link -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding

# Inspect the latest broad first-failure trace and result order.
$run='C:\RSPerf\runs\20260707T145553Z-71864-2cfa41388a354d26aff55a14b2f9eb5b'
Get-Content "$run\artifacts\selftest\last_run\commands\trace.txt" -Tail 140
$d=Get-Content "$run\artifacts\selftest\last_run\commands\results.json" -Raw | ConvertFrom-Json
$executed=$d.cases | Where-Object { $_.status -ne 'skipped' }
$executed | Select-Object -Last 30 name,status,durationMs,reason | Format-Table -AutoSize

# Locate the Keyboard deferred-search page-settle code before editing.
rg -n "cmd_preferences_dialog_keyboard_search_action_updates_dxui_surface|Keyboard page did not settle|deferred-search" RedSalamander\SelfTest\Commands

# Focused Keyboard proof/repro. Use Start-Process -Wait for GUI selftests.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-resume-focused-keyboard-search-action'
$args=@(
  '--commands-selftest',
  '--selftest-case=cmd_preferences_dialog_keyboard_search_action_updates_dxui_surface',
  '--selftest-timeout-multiplier=2'
)
$p=Start-Process -FilePath (Resolve-Path '.\.build\x64\Debug\RedSalamander.exe') `
  -ArgumentList $args -PassThru -Wait
Start-Sleep -Seconds 3
Write-Host "processExit=$($p.ExitCode)"
Get-Content "C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands\results.json" -Raw
Get-Content "C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands\trace.txt" -Tail 160

# After the root-cause fix.
Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

### Resume cautions

- [ ] Do not run GUI/native selftests in parallel.
- [ ] Use `Start-Process -Wait` for direct native GUI selftests.
- [ ] After native GUI selftests, wait 3-6 seconds, then read settled
  `last_run\results.json` and `last_run\commands\results.json`.
- [ ] Shell exit code can be 0 even when a script prints `runnerExit=1`; inspect JSON and summaries.
- [ ] Pester can return control despite failures; inspect the pass/fail summary.
- [ ] In non-fail-fast GUI runs, treat later failures after the first failure as potentially
  contaminated until reproduced independently.
- [ ] Do not clean `C:\RSPerf` until useful manual and broad evidence has been archived.

### Worktree at save point

- [ ] Modified stabilization files visible in `git status`:
  `Common/Common/SettingsStore.cpp`,
  `Common/DxUi/DxUi.Menu.cpp`,
  `Common/DxUi/DxUi.h`,
  `RedSalamander/ManagePluginsDialog.cpp`,
  `RedSalamander/RedSalamander.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`,
  `Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md`,
  `Specs/Plans/WIP/README.md`,
  `Tests/DxUiTests/DxUiTests.Menu.cpp`,
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`,
  and `Tools/Tests/WingetValidation.Tests.ps1`.
- [ ] Untracked review/spec material visible in `git status`:
  `Specs/Core/Core_FileSystemBridge.md`,
  `Specs/Plans/WIP/FileSystem_CrossFsBridgeImplementationAndPerformanceRedesign_2026-07-07.md`,
  `Specs/Plans/Done/Operation_Causeway_CrossFsBridgeSecurityAndPerformanceRemediation_2026-07-07.md`,
  `Specs/Reviews/FileSystemBridge-2026-07-07-Findings.md`,
  and `Specs/Reviews/ThreeDayDiff-2026-07-06-Findings.md`. Do not stage/delete these unless they
  are intentionally part of the final scope.
- [ ] Local continuation evidence confirmed on disk includes all run ids named in this current
  checkpoint and the superseded checkpoints below.
- [ ] Broad Commands and Suite Full are still not accepted. Do not move this plan to Done yet.

## Superseded resume checkpoint - 2026-07-07 16:24 +02:00

Kept for history only. Use the current 17:03 checkpoint above when resuming. This older checkpoint
predates the `AttachThreadInput(...)` Preferences focus-helper upgrade, the Plugins header-rect
wait guard, and the latest Keyboard deferred-search page-settle blocker.

### Done / saved

- [x] Final proof root is explicit and unchanged: use `REDSALAMANDER_TEST_ROOT=C:\RSPerf` for all
  accepted local proof runs.
- [x] Current checkout at save point: `cb2478689bc1b37afc6066a6a5994d8f135a56cc`.
- [x] No `RedSalamander.exe`, `MSBuild`, `cl`, or `link` process was running at this save point.
- [x] Earlier root-cause fixes remain saved in this WIP plan: SettingsStore test-root isolation,
  rapid Preferences fixture seeding, stale Preferences category-tree UIA selection wait,
  plugin-config logical-focus recovery, preview fallback atomic wait, and Quick Search pane
  sort/extension restoration.
- [x] Hot Paths browse remains diagnostic-only. It did not reproduce focused or in local predecessor
  proof: `manual-20260707T1530-focused-hotpaths-browse` passed 1 / 0 / 0,
  `manual-20260707T1531-hotpaths-browse-predecessor-cluster` passed 13 / 0 / 0, and
  `manual-20260707T1532-preferences-80-to-hotpaths-browse` passed 81 / 0 / 0.
- [x] Viewers theme-cycle did not reproduce focused or in its immediate Preferences predecessor
  window: `manual-20260707T1545-focused-viewers-theme-cycle` passed 1 / 0 / 0 and
  `manual-20260707T1546-pref-window-to-viewers-theme-cycle` passed 12 / 0 / 0.
- [x] The old Viewers roundtrip first failure was captured in
  `manual-20260707T1548-prefix-to-viewers-theme-cycle-failfast`: process exit 1,
  231 passed / 1 failed / 1 skipped.
- [x] Before the Viewers focus patch, the focused Viewers roundtrip and local predecessor proof did
  not reproduce the failure:
  `manual-20260707T1552-focused-viewers-roundtrip` passed 1 / 0 / 0,
  `manual-20260707T1553-viewers-roundtrip-predecessor-cluster` passed 11 / 0 / 0, and
  `manual-20260707T1554-prefix-to-viewers-roundtrip-red-repro` passed 232 / 0 / 0.
- [x] Viewers roundtrip category navigation now uses condition-based
  `FocusWindowAndWait(...)` instead of two one-shot raw `SetFocus(...)` assertions, with diagnostic
  failure text in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`.
- [x] After the Viewers focus patch, source contracts passed 93 / 0 and build
  `.build\logs\msbuild-20260707_155826_875.log` had no MSBuild/MSVC warning or error diagnostics.
- [x] After the Viewers focus patch, focused/predecessor/prefix proof stayed green:
  `manual-20260707T1601-focused-viewers-roundtrip-focus-wait` passed 1 / 0 / 0,
  `manual-20260707T1601-viewers-roundtrip-predecessor-focus-wait` passed 11 / 0 / 0, and
  `manual-20260707T1602-prefix-to-viewers-roundtrip-focus-wait` passed 232 / 0 / 0.
- [x] Broad Commands fail-fast then moved the first failure to Themes:
  `20260707T140500Z-73276-4cba059777c244b99dbdd3b37d37070c`, process exit 1,
  264 passed / 1 failed / 523 skipped. First failure:
  `cmd_preferences_dialog_themes_roundtrip_restores_dxui_surface`, reason
  `Failed to focus the Preferences category host for Themes round-trip test.`
- [x] Themes did not reproduce focused or in a 25-case predecessor window before the shared helper
  change: `manual-20260707T1610-focused-themes-roundtrip` passed 1 / 0 / 0 and
  `manual-20260707T1611-themes-roundtrip-predecessor-25` passed 25 / 0 / 0.
- [x] Shared `FocusWindowAndWait(...)` in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.cpp` was changed to retry root
  activation/focus while pumping. This is provisional: the current version uses
  `BringWindowToTop(root)` and `SetForegroundWindow(root)`, and must be reviewed against the
  stronger `AttachThreadInput(...)` foreground patterns in Search/ViewCommands before it is trusted.
- [x] Themes roundtrip has diagnostic-rich failure text in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`.
- [x] Source contracts passed after the focus-helper/source-contract edits:
  `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`,
  94 passed / 0 failed.
- [x] Builds after the focus-helper and Themes diagnostic edits passed:
  `.build\logs\msbuild-20260707_161249_278.log` and
  `.build\logs\msbuild-20260707_161612_326.log`, with no MSBuild/MSVC warning or error diagnostics.
- [x] Focused Themes failed once after the shared helper change before the diagnostic rebuild:
  `manual-20260707T1615-focused-themes-roundtrip-focus-helper` failed 0 / 1 / 0 with the generic
  focus-host failure. Treat this as a warning sign, not as proof of a fix.
- [x] After the Themes diagnostic rebuild, focused and predecessor Themes proof passed:
  `manual-20260707T1618-focused-themes-roundtrip-focus-diagnostic` passed 1 / 0 / 0 and
  `manual-20260707T1619-themes-roundtrip-predecessor-25-focus-helper` passed 25 / 0 / 0.
- [x] Latest broad Commands fail-fast is saved at
  `C:\RSPerf\runs\20260707T142028Z-70416-fdc68581974845e58d44fe7bcea87278\`:
  process exit 1, 229 passed / 1 failed / 558 skipped. First failure:
  `cmd_preferences_dialog_plugins_roundtrip_restores_dxui_surface`, reason:
  `Failed to focus the Preferences category host before leaving Plugins for General; nativeFocus=0x0,
  categoryHost=0x96D1CA8, activePage=0x9DF15AC, shellHost=0x17E1CC6, currentCategory=7,
  pageTitle='Plugins'.`
- [x] Direct GUI selftests must be launched with `Start-Process -Wait`; direct `& exe` probes can
  return before the GUI process finishes and create weak or invalid evidence.

### Current blocker / next investigation

- [ ] Re-read `FocusWindowAndWait(...)` in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.cpp`.
- [ ] Re-read the stronger focus/foreground helpers in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp` and
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp`, especially their
  `AttachThreadInput(...)` handling.
- [ ] Decide whether the shared helper should be rolled back to retry-only
  `SetActiveWindow(root)` + `SetFocus(hwnd)`, or upgraded to the proven `AttachThreadInput(...)`
  foreground pattern. Do not leave the current helper unexamined.
- [ ] Inspect the latest broad trace:
  `C:\RSPerf\runs\20260707T142028Z-70416-fdc68581974845e58d44fe7bcea87278\artifacts\selftest\last_run\commands\trace.txt`.
- [ ] Run focused
  `cmd_preferences_dialog_plugins_roundtrip_restores_dxui_surface` under `C:\RSPerf` with
  `Start-Process -Wait`.
- [ ] If focused Plugins passes, run the local predecessor cluster ending at Plugins roundtrip.
- [ ] If the predecessor cluster passes, replay the exact 0..229 fail-fast prefix from the latest
  broad run and narrow the minimal predecessor window. Do not patch from suspicion.
- [ ] If Plugins fails, inspect `commands\trace.txt`, result JSON, and Preferences debug snapshot
  text before changing code.
- [ ] Add or update source-contract guards only after the root cause is known.

### Remaining to finish this operation

- [ ] Root-cause and fix or replace
  `cmd_preferences_dialog_plugins_roundtrip_restores_dxui_surface`; do not classify it as flaky and
  move on.
- [ ] Re-review the provisional shared `FocusWindowAndWait(...)` change and either keep it with a
  real root-cause justification or replace it with a safer helper.
- [ ] Re-run source contracts and require 0 failures.
- [ ] Rebuild `RedSalamander`.
- [ ] Re-run focused Plugins roundtrip and its predecessor cluster under `C:\RSPerf`.
- [ ] Re-run the exact latest broad prefix under `C:\RSPerf`.
- [ ] Re-run broad Commands fail-fast under `C:\RSPerf`.
- [ ] If broad Commands finds another broad-only failure, identify it, root-cause it, and fix or
  replace the brittle path. Do not label it flaky and move on.
- [ ] If `cmd_connection_manager_window_rejects_blank_profile_name` reappears, use the new
  diagnostic snapshot traces before changing behavior.
- [ ] If `cmd_preferences_dialog_hot_paths_browse_live_dx_interaction` reappears, use the new
  diagnostic-rich snapshot text before changing behavior.
- [ ] Re-run broad Commands without `-FailFast` under `C:\RSPerf`.
- [ ] Re-run Suite Full under `C:\RSPerf`.
- [ ] If still failing, explain or fix `RedSalamanderMonitorEtwLatency`; a prior Suite Full saw exit
  1 with an empty output log.
- [ ] If still failing, explain or fix `ToolsPesterTests`; a prior Suite Full saw exit 1 with no
  captured output log.
- [ ] Archive final passing evidence under `Specs/TestRuns/<current-head>\...` or another clearly
  named continuation folder before cleaning volatile `C:\RSPerf` evidence.
- [ ] Clean or reset `C:\RSPerf` only after needed evidence is archived; current manual run folders
  make disk-audit checks dirty.
- [ ] Before closeout, move this plan to `Specs/Plans/Done/` and merge durable requirements into
  `Specs/Testing/*`, `Tests/README.md`, `README.md`, or `AGENTS.md` as appropriate.
- [ ] Commit only intentional stabilization files; do not stage unrelated review material.

### Exact resume commands

```powershell
# Confirm no orphaned app/build process is holding the interactive desktop.
Get-Process RedSalamander,MSBuild,cl,link -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding

# Read the current focus helper and the stronger existing foreground patterns.
$lines=Get-Content 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.cpp'
$lines[65..120]
$lines=Get-Content 'RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp'
$lines[160..210]
$lines=Get-Content 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Search.cpp'
$lines[1150..1195]

# Inspect the latest broad first-failure trace.
Get-Content 'C:\RSPerf\runs\20260707T142028Z-70416-fdc68581974845e58d44fe7bcea87278\artifacts\selftest\last_run\commands\trace.txt' -Tail 180

# Focused Plugins roundtrip proof/repro. Use Start-Process -Wait for GUI selftests.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-resume-focused-plugins-roundtrip-focus-helper'
$args=@(
  '--commands-selftest',
  '--selftest-case=cmd_preferences_dialog_plugins_roundtrip_restores_dxui_surface',
  '--selftest-timeout-multiplier=2'
)
$p=Start-Process -FilePath (Resolve-Path '.\.build\x64\Debug\RedSalamander.exe') `
  -ArgumentList $args -PassThru -Wait
Start-Sleep -Seconds 3
Write-Host "processExit=$($p.ExitCode)"
Get-Content "C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands\results.json" -Raw
Get-Content "C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands\trace.txt" -Tail 160

# Local predecessor cluster through Plugins roundtrip if focused passes.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-resume-plugins-roundtrip-predecessor-cluster'
$cases=@(
  'cmd_preferences_dialog_category_tree_uses_dxui_host_without_visible_legacy_treeview',
  'cmd_preferences_dialog_opens_with_french_satellite_resources',
  'cmd_preferences_dialog_category_tree_exposes_live_uia_selection',
  'cmd_preferences_dialog_shell_uses_dxui_header_footer_without_visible_legacy_shell_controls',
  'cmd_preferences_dialog_escape_prompts_before_dirty_close',
  'cmd_preferences_dialog_page_host_uses_dxui_surface',
  'cmd_preferences_dialog_viewers_page_uses_dxui_combo_and_button_chrome',
  'cmd_preferences_dialog_plugins_page_uses_dxui_shell_chrome',
  'cmd_preferences_dialog_plugins_roundtrip_restores_dxui_surface'
) -join ','
$args=@('--commands-selftest', "--selftest-case=$cases", '--selftest-fail-fast', '--selftest-timeout-multiplier=2')
$p=Start-Process -FilePath (Resolve-Path '.\.build\x64\Debug\RedSalamander.exe') `
  -ArgumentList $args -PassThru -Wait
Start-Sleep -Seconds 3
Write-Host "processExit=$($p.ExitCode)"
Get-Content "C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands\results.json" -Raw
Get-Content "C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands\trace.txt" -Tail 180

# After the root-cause fix.
Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

### Resume cautions

- [ ] Do not run GUI/native selftests in parallel.
- [ ] Use `Start-Process -Wait` for direct native GUI selftests.
- [ ] After native GUI selftests, wait 3-6 seconds, then read settled
  `last_run\results.json` and `last_run\commands\results.json`.
- [ ] Shell exit code can be 0 even when a script prints `runnerExit=1`; inspect JSON and summaries.
- [ ] Pester can return control despite failures; inspect the pass/fail summary.
- [ ] In non-fail-fast GUI runs, treat later failures after the first failure as potentially
  contaminated until reproduced independently.
- [ ] Do not clean `C:\RSPerf` until useful manual and broad evidence has been archived.

### Worktree at save point

- [ ] Modified stabilization files visible in `git status`:
  `Common/Common/SettingsStore.cpp`,
  `Common/DxUi/DxUi.Menu.cpp`,
  `Common/DxUi/DxUi.h`,
  `RedSalamander/ManagePluginsDialog.cpp`,
  `RedSalamander/RedSalamander.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`,
  `Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md`,
  `Specs/Plans/WIP/README.md`,
  `Tests/DxUiTests/DxUiTests.Menu.cpp`,
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`,
  and `Tools/Tests/WingetValidation.Tests.ps1`.
- [ ] Untracked review/spec material visible in `git status`:
  `Specs/Core/Core_FileSystemBridge.md`,
  `Specs/Plans/Done/Operation_Causeway_CrossFsBridgeSecurityAndPerformanceRemediation_2026-07-07.md`,
  `Specs/Reviews/FileSystemBridge-2026-07-07-Findings.md`,
  and `Specs/Reviews/ThreeDayDiff-2026-07-06-Findings.md`. Do not stage/delete these unless they
  are intentionally part of the final scope.
- [ ] Local continuation evidence confirmed on disk includes all run ids named in this current
  checkpoint and the superseded checkpoints below.
- [ ] Broad Commands and Suite Full are still not accepted. Do not move this plan to Done yet.

## Superseded resume checkpoint - 2026-07-07 15:50 +02:00

Kept for history only. Use the current 17:03 checkpoint above when resuming. This older checkpoint
predates the Viewers focus patch, the provisional shared focus-helper change, the Themes diagnostic
checkpoint, and the latest Plugins roundtrip blocker.

### Done / saved

- [x] Final proof root is explicit and unchanged: use `REDSALAMANDER_TEST_ROOT=C:\RSPerf` for all
  accepted local proof runs.
- [x] Current checkout at save point: `cb2478689bc1b37afc6066a6a5994d8f135a56cc`.
- [x] No `RedSalamander.exe`, `MSBuild`, `cl`, or `link` process was running at this save point.
- [x] SettingsStore isolation is patched in `Common/Common/SettingsStore.cpp`: app settings route
  through `runs\<runId>\scratch\settings-store\RedSalamander\Settings` when
  `REDSALAMANDER_TEST_ROOT` and `REDSALAMANDER_TEST_RUN_ID` are set.
- [x] SettingsStore isolation has a source-contract guard in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`; user-profile protection was proven in
  `C:\RSPerf\runs\manual-20260707T1336Z-settings-isolation-compare-commands\`.
- [x] `cmd_preferences_dialog_rapid_switches_keep_page_specific_uia_subtrees` was root-caused as a
  clean-settings fixture bug and fixed by seeding deterministic editor actions in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp`.
- [x] Rapid-switch red/green evidence exists under `C:\RSPerf`:
  `manual-20260707T1408Z-red-rapid-preferences` failed 0 / 1 / 0 before the patch;
  `manual-20260707T1410Z-green-rapid-preferences` passed 1 / 0 / 0 after it.
- [x] Connection Manager double-click and blank-profile-name symptoms did not reproduce focused or
  in local predecessor windows. Do not patch them blindly; use the diagnostic traces in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp` if they recur.
- [x] A stale Preferences category-tree UIA selection was fixed by waiting for provider selection to
  settle on the expected tree item in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`.
- [x] Plugin configuration tab traversal logical-focus recovery is patched in
  `RedSalamander/ManagePluginsDialog.cpp`; the selftest helper no longer sends an extra Tab after a
  successful advance in `RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp`.
- [x] Preview fallback failure was root-caused as a test race and fixed in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp` with
  `WaitForPreviewPaneTextAndFocus(...)`.
- [x] Quick Search broad-only failure was root-caused as leaked pane sort/extension state and fixed
  in `RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp` by owning/restoring pane sort
  and extension visibility.
- [x] Hot Paths browse failure from broad run
  `20260707T131755Z-80540-0d84ab2c177f4f92bef55cba18ec84d6` now has diagnostic-rich failure text in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp`; this was
  intentionally diagnostic, not a behavioral fix.
- [x] Hot Paths diagnostic guard is present in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`.
- [x] Source contracts passed after the Hot Paths diagnostic guard:
  `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`,
  92 passed / 0 failed.
- [x] Build after the Hot Paths diagnostic guard passed:
  `.build\logs\msbuild-20260707_153449_079.log`, with no MSBuild/MSVC warning or error diagnostics.
- [x] Hot Paths browse did not reproduce focused or in local predecessor proof:
  `manual-20260707T1530-focused-hotpaths-browse` passed 1 / 0 / 0,
  `manual-20260707T1531-hotpaths-browse-predecessor-cluster` passed 13 / 0 / 0, and
  `manual-20260707T1532-preferences-80-to-hotpaths-browse` passed 81 / 0 / 0.
- [x] Exact 0..321 prefix after the diagnostic build is saved at
  `C:\RSPerf\runs\manual-20260707T1537-commands-prefix-to-hotpaths-browse-diagnostic\`:
  process exit 1, 312 passed / 10 failed / 0 skipped. Hot Paths browse passed in this run. Treat
  only the first failure as primary; later failures may be contaminated by the first failure.
- [x] Focused Viewers theme-cycle and its immediate Preferences predecessor window passed:
  `manual-20260707T1545-focused-viewers-theme-cycle` passed 1 / 0 / 0 and
  `manual-20260707T1546-pref-window-to-viewers-theme-cycle` passed 12 / 0 / 0.
- [x] Exact 0..232 prefix with `--selftest-fail-fast` is saved at
  `C:\RSPerf\runs\manual-20260707T1548-prefix-to-viewers-theme-cycle-failfast\`: process exit 1,
  231 passed / 1 failed / 1 skipped.
- [x] Current actionable first failure is
  `cmd_preferences_dialog_viewers_roundtrip_restores_dxui_surface`, reason:
  `Failed to focus the Preferences category host before leaving Viewers for General.`
- [x] Direct GUI selftests must be launched with `Start-Process -Wait`; direct `& exe` probes can
  return before the GUI process finishes and create weak or invalid evidence.

### Current blocker / next investigation

- [ ] Inspect `TestPreferencesDialogViewersRoundTripRestoresDxUiSurface(...)` and the exact failure
  path in `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`.
- [ ] Inspect `FocusWindowAndWait(...)` before changing the test. The next fix must explain whether
  the helper has a timing/focus contract bug, whether the Viewers roundtrip test is focusing the
  wrong host, or whether a prior prefix case leaks foreground/focus state.
- [ ] Run the focused current blocker under `C:\RSPerf` with `Start-Process -Wait`.
- [ ] If focused passes, run the local predecessor cluster for the Viewers roundtrip case.
- [ ] If the predecessor cluster passes, rerun the exact 0..232 fail-fast prefix and then narrow the
  minimal predecessor window from the prefix. Do not patch from suspicion.
- [ ] If the case fails, inspect `commands\trace.txt`, result JSON, and Preferences debug snapshot
  text before changing code.
- [ ] Add or update a source-contract guard only after the root cause is known.

### Remaining to finish this operation

- [ ] Root-cause and fix or replace
  `cmd_preferences_dialog_viewers_roundtrip_restores_dxui_surface`; do not classify it as flaky and
  move on.
- [ ] Re-run source contracts and require 0 failures.
- [ ] Rebuild `RedSalamander`.
- [ ] Re-run focused Viewers roundtrip and its predecessor cluster under `C:\RSPerf`.
- [ ] Re-run the exact 0..232 fail-fast prefix under `C:\RSPerf`.
- [ ] Re-run broad Commands fail-fast under `C:\RSPerf`.
- [ ] If broad Commands finds another broad-only failure, identify it, root-cause it, and fix or
  replace the brittle path. Do not label it flaky and move on.
- [ ] If `cmd_connection_manager_window_rejects_blank_profile_name` reappears, use the new diagnostic
  snapshot traces before changing behavior.
- [ ] If `cmd_preferences_dialog_hot_paths_browse_live_dx_interaction` reappears, use the new
  diagnostic-rich snapshot text before changing behavior.
- [ ] Re-run broad Commands without `-FailFast` under `C:\RSPerf`.
- [ ] Re-run Suite Full under `C:\RSPerf`.
- [ ] If still failing, explain or fix `RedSalamanderMonitorEtwLatency`; a prior Suite Full saw exit
  1 with an empty output log.
- [ ] If still failing, explain or fix `ToolsPesterTests`; a prior Suite Full saw exit 1 with no
  captured output log.
- [ ] Archive final passing evidence under `Specs/TestRuns/<current-head>\...` or another clearly
  named continuation folder before cleaning volatile `C:\RSPerf` evidence.
- [ ] Clean or reset `C:\RSPerf` only after needed evidence is archived; current manual run folders
  make disk-audit checks dirty.
- [ ] Before closeout, move this plan to `Specs/Plans/Done/` and merge durable requirements into
  `Specs/Testing/*`, `Tests/README.md`, `README.md`, or `AGENTS.md` as appropriate.
- [ ] Commit only intentional stabilization files; do not stage unrelated review material.

### Exact resume commands

```powershell
# Confirm no orphaned app/build process is holding the interactive desktop.
Get-Process RedSalamander,MSBuild,cl,link -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding

# Inspect the current blocker and helper.
rg -n "TestPreferencesDialogViewersRoundTripRestoresDxUiSurface|Failed to focus the Preferences category host before leaving Viewers|leaving Viewers for General|navigateToGeneralPage|navigateToViewersPage|FocusWindowAndWait" `
  RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.ChromeAndPlugins.cpp `
  RedSalamander\SelfTest\Commands `
  RedSalamander

# Focused Viewers roundtrip proof/repro. Use Start-Process -Wait for GUI selftests.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-resume-focused-viewers-roundtrip'
$args=@(
  '--commands-selftest',
  '--selftest-case=cmd_preferences_dialog_viewers_roundtrip_restores_dxui_surface',
  '--selftest-timeout-multiplier=2'
)
$p=Start-Process -FilePath (Resolve-Path '.\.build\x64\Debug\RedSalamander.exe') `
  -ArgumentList $args -PassThru -Wait
Start-Sleep -Seconds 3
Write-Host "processExit=$($p.ExitCode)"
Get-Content "C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands\results.json" -Raw
Get-Content "C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands\trace.txt" -Tail 160

# Local predecessor cluster through Viewers roundtrip if focused passes.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-resume-viewers-roundtrip-predecessor-cluster'
$cases=@(
  'cmd_preferences_dialog_category_tree_uses_dxui_host_without_visible_legacy_treeview',
  'cmd_preferences_dialog_opens_with_french_satellite_resources',
  'cmd_preferences_dialog_category_tree_exposes_live_uia_selection',
  'cmd_preferences_dialog_shell_uses_dxui_header_footer_without_visible_legacy_shell_controls',
  'cmd_preferences_dialog_escape_prompts_before_dirty_close',
  'cmd_preferences_dialog_page_host_uses_dxui_surface',
  'cmd_preferences_dialog_viewers_page_uses_dxui_combo_and_button_chrome',
  'cmd_preferences_dialog_plugins_page_uses_dxui_shell_chrome',
  'cmd_preferences_dialog_plugins_roundtrip_restores_dxui_surface',
  'cmd_preferences_dialog_plugins_theme_cycle_keeps_surface_legible',
  'cmd_preferences_dialog_viewers_roundtrip_restores_dxui_surface'
) -join ','
$args=@('--commands-selftest', "--selftest-case=$cases", '--selftest-fail-fast', '--selftest-timeout-multiplier=2')
$p=Start-Process -FilePath (Resolve-Path '.\.build\x64\Debug\RedSalamander.exe') `
  -ArgumentList $args -PassThru -Wait
Start-Sleep -Seconds 3
Write-Host "processExit=$($p.ExitCode)"
Get-Content "C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands\results.json" -Raw
Get-Content "C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands\trace.txt" -Tail 180

# After the root-cause fix.
Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

### Resume cautions

- [ ] Do not run GUI/native selftests in parallel.
- [ ] Use `Start-Process -Wait` for direct native GUI selftests.
- [ ] After native GUI selftests, wait 3-6 seconds, then read settled
  `last_run\results.json` and `last_run\commands\results.json`.
- [ ] Shell exit code can be 0 even when a script prints `runnerExit=1`; inspect JSON and summaries.
- [ ] Pester can return control despite failures; inspect the pass/fail summary.
- [ ] In non-fail-fast GUI runs, treat later failures after the first failure as potentially
  contaminated until reproduced independently.
- [ ] Do not clean `C:\RSPerf` until useful manual and broad evidence has been archived.

### Worktree at save point

- [ ] Modified stabilization files visible in `git status`:
  `Common/Common/SettingsStore.cpp`,
  `Common/DxUi/DxUi.Menu.cpp`,
  `Common/DxUi/DxUi.h`,
  `RedSalamander/ManagePluginsDialog.cpp`,
  `RedSalamander/RedSalamander.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`,
  `Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md`,
  `Specs/Plans/WIP/README.md`,
  `Tests/DxUiTests/DxUiTests.Menu.cpp`,
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`,
  and `Tools/Tests/WingetValidation.Tests.ps1`.
- [ ] Untracked review/spec material visible in `git status`:
  `Specs/Core/Core_FileSystemBridge.md`,
  `Specs/Plans/Done/Operation_Causeway_CrossFsBridgeSecurityAndPerformanceRemediation_2026-07-07.md`,
  `Specs/Reviews/FileSystemBridge-2026-07-07-Findings.md`,
  and `Specs/Reviews/ThreeDayDiff-2026-07-06-Findings.md`. Do not stage/delete these unless they
  are intentionally part of the final scope.
- [ ] Local continuation evidence confirmed on disk includes:
  `C:\RSPerf\runs\manual-20260707T1530-focused-hotpaths-browse\`,
  `C:\RSPerf\runs\manual-20260707T1531-hotpaths-browse-predecessor-cluster\`,
  `C:\RSPerf\runs\manual-20260707T1532-preferences-80-to-hotpaths-browse\`,
  `C:\RSPerf\runs\manual-20260707T1537-commands-prefix-to-hotpaths-browse-diagnostic\`,
  `C:\RSPerf\runs\manual-20260707T1545-focused-viewers-theme-cycle\`,
  `C:\RSPerf\runs\manual-20260707T1546-pref-window-to-viewers-theme-cycle\`,
  `C:\RSPerf\runs\manual-20260707T1548-prefix-to-viewers-theme-cycle-failfast\`,
  and earlier run ids listed in the superseded checkpoints below.
- [ ] Broad Commands and Suite Full are still not accepted. Do not move this plan to Done yet.

## Superseded resume checkpoint - 2026-07-07 15:25 +02:00

Kept for history only. Use the current 15:50 checkpoint above when resuming. This older checkpoint
predates the Hot Paths diagnostic guard/proof runs and the current Viewers roundtrip blocker.

### Done / saved

- [x] Final proof root is explicit and unchanged: use `REDSALAMANDER_TEST_ROOT=C:\RSPerf` for all
  accepted local proof runs.
- [x] Current checkout at save point: `cb2478689bc1b37afc6066a6a5994d8f135a56cc`.
- [x] No `RedSalamander.exe`, `MSBuild`, `cl`, or `link` process was running at this save point.
- [x] SettingsStore isolation is patched in `Common/Common/SettingsStore.cpp`: app settings route
  through `runs\<runId>\scratch\settings-store\RedSalamander\Settings` when
  `REDSALAMANDER_TEST_ROOT` and `REDSALAMANDER_TEST_RUN_ID` are set.
- [x] SettingsStore isolation has a source-contract guard in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`; user-profile protection was proven in
  `C:\RSPerf\runs\manual-20260707T1336Z-settings-isolation-compare-commands\`.
- [x] `cmd_preferences_dialog_rapid_switches_keep_page_specific_uia_subtrees` was root-caused as a
  clean-settings fixture bug and fixed by seeding deterministic editor actions in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp`.
- [x] Rapid-switch red/green evidence exists under `C:\RSPerf`:
  `manual-20260707T1408Z-red-rapid-preferences` failed 0 / 1 / 0 before the patch;
  `manual-20260707T1410Z-green-rapid-preferences` passed 1 / 0 / 0 after it.
- [x] The clean-root rename and FolderView first symptoms did not reproduce focused after the
  rapid-switch fix: `manual-20260707T1411Z-focused-rename-prompt` passed 1 / 0 / 0 and
  `manual-20260707T1413Z-focused-folder-empty` passed 1 / 0 / 0.
- [x] A Connection Manager double-click symptom did not reproduce focused, in an immediate pair, or
  in a local predecessor window:
  `manual-20260707T1420Z-focused-connection-doubleclick` passed 1 / 0 / 0,
  `manual-20260707T1422Z-connection-live-then-doubleclick` passed 2 / 0 / 0, and
  `manual-20260707T1424Z-connection-local-window` passed 13 / 0 / 0. Do not patch it blindly.
- [x] A stale Preferences category-tree UIA selection was fixed by waiting for provider selection to
  settle on the expected tree item in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`.
- [x] Preferences category-tree proof passed:
  `manual-20260707T1434Z-green-pref-category-tree`, 1 passed / 0 failed / 0 skipped.
- [x] Plugin configuration tab traversal logical-focus recovery is patched in
  `RedSalamander/ManagePluginsDialog.cpp`; the helper also no longer sends an extra Tab after a
  successful advance in `RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp`.
- [x] Plugin-config source-contract guards are present in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`.
- [x] Plugin-config focused/pair proof passed:
  `manual-20260707T1433Z-green-plugin-config-tab-indexed-recovery` passed 1 / 0 / 0,
  `manual-20260707T1436Z-green-plugin-live-then-tab-indexed-recovery` passed 2 / 0 / 0,
  `manual-20260707T1446-plugin-config-tab-no-extra-advance` passed 1 / 0 / 0, and
  `manual-20260707T1447-plugin-live-then-tab-no-extra-advance` passed 2 / 0 / 0.
- [x] `cmd_connection_manager_window_rejects_blank_profile_name` was exposed once by broad
  Commands, then did not reproduce focused or in its local predecessor cluster:
  `manual-20260707T1451-focused-blank-profile-name` passed 1 / 0 / 0 and
  `manual-20260707T1452-connection-cluster-to-blank-name` passed 6 / 0 / 0.
- [x] `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp` now has diagnostic-only
  snapshot tracing around `TestConnectionManagerWindowRejectsProfileName(...)` for the next broad
  recurrence. This is not a behavioral fix.
- [x] Preview fallback failure was root-caused as a test race: the fallback test waited for text and
  sampled source-pane focus separately.
- [x] `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp` now has
  `WaitForPreviewPaneTextAndFocus(...)`, and the no-preview file/folder fallback assertions wait for
  fallback text, forbidden text absence, and expected source-pane focus in one predicate.
- [x] Preview fallback source-contract guard is present in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`.
- [x] Source contracts passed after the preview patch:
  `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`,
  90 passed / 0 failed.
- [x] Build after the preview patch passed:
  `.build\logs\msbuild-20260707_150116_469.log`, with no MSBuild/MSVC warning or error diagnostics.
- [x] Preview fallback focused proof passed:
  `manual-20260707T1504-preview-fallback-focus-wait`, 1 passed / 0 failed / 0 skipped.
- [x] Preview predecessor cluster proof passed:
  `manual-20260707T1506-preview-cluster-actual-to-fallback-focus-wait`, 4 passed / 0 failed /
  0 skipped.
- [x] Quick Search broad-only failure was root-caused as leaked pane sort/extension state from the
  preceding search/find cluster.
- [x] `RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp` now makes the integrated and
  hidden-extension quick-search tests own and restore pane sort and extension visibility.
- [x] Quick Search source-contract guard is present in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`.
- [x] Source contracts passed after the Quick Search patch:
  `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`,
  91 passed / 0 failed.
- [x] Build after the Quick Search patch passed:
  `.build\logs\msbuild-20260707_151503_009.log`, with no MSBuild/MSVC warning or error diagnostics.
- [x] Quick Search focused proof passed:
  `manual-20260707T1518-focused-quicksearch-sort-owned`, 1 passed / 0 failed / 0 skipped.
- [x] Quick Search predecessor cluster proof passed:
  `manual-20260707T1519-search-cluster-to-quicksearch-sort-owned`, 9 passed / 0 failed /
  0 skipped.
- [x] Latest broad Commands fail-fast run is saved at
  `C:\RSPerf\runs\20260707T131755Z-80540-0d84ab2c177f4f92bef55cba18ec84d6\`: 321 passed /
  1 failed / 466 skipped.
- [x] Current failing case from that broad run is
  `cmd_preferences_dialog_hot_paths_browse_live_dx_interaction`, reason:
  `Preferences Hot Paths page did not settle to the active DX surface before browse interaction validation.`
- [x] Direct GUI selftests must be launched with `Start-Process -Wait`; earlier direct `& exe` probes
  returned before the GUI process finished and produced weak or invalid evidence.

### Current blocker / next investigation

- [ ] Inspect the failing Hot Paths browse test and helpers:
  `rg -n "hot_paths_browse|Hot Paths page did not settle|browse_live|WaitFor.*Hot Paths"`
  in `RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp` and
  nearby Preferences helper files.
- [ ] Run the focused current blocker under `C:\RSPerf` with `Start-Process -Wait`.
- [ ] Run the 13-case local predecessor cluster under `C:\RSPerf` with `Start-Process -Wait`.
- [ ] If focused or predecessor cluster fails, inspect `commands\trace.txt`, result JSON, and any
  debug snapshot text before changing code.
- [ ] If focused passes but predecessor cluster fails, treat it as leaked state from one of the prior
  Hot Paths / Panes / Shell Footer tests and make the browse case own that state.
- [ ] Current leading suspicion is a narrow wait predicate or stale Hot Paths state before browse
  validation. This is an inference, not yet proof. Do not patch blindly.
- [ ] Add or update a source-contract guard only after the root cause is known.

### Remaining to finish this operation

- [ ] Fix or replace the current Hot Paths browse failure with a deterministic test path.
- [ ] Re-run source contracts and require 0 failures.
- [ ] Rebuild `RedSalamander`.
- [ ] Re-run focused Hot Paths browse and its predecessor cluster under `C:\RSPerf`.
- [ ] Re-run broad Commands fail-fast under `C:\RSPerf`.
- [ ] If broad Commands finds another broad-only failure, identify it, root-cause it, and fix or
  replace the brittle path. Do not label it flaky and move on.
- [ ] If `cmd_connection_manager_window_rejects_blank_profile_name` reappears, use the new diagnostic
  snapshot traces before changing behavior.
- [ ] Re-run broad Commands without `-FailFast` under `C:\RSPerf`.
- [ ] Re-run Suite Full under `C:\RSPerf`.
- [ ] If still failing, explain or fix `RedSalamanderMonitorEtwLatency`; a prior Suite Full saw exit
  1 with an empty output log.
- [ ] If still failing, explain or fix `ToolsPesterTests`; a prior Suite Full saw exit 1 with no
  captured output log.
- [ ] Archive final passing evidence under `Specs/TestRuns/<current-head>\...` or another clearly
  named continuation folder before cleaning volatile `C:\RSPerf` evidence.
- [ ] Clean or reset `C:\RSPerf` only after needed evidence is archived; current manual run folders
  make disk-audit checks dirty.
- [ ] Before closeout, move this plan to `Specs/Plans/Done/` and merge durable requirements into
  `Specs/Testing/*`, `Tests/README.md`, `README.md`, or `AGENTS.md` as appropriate.
- [ ] Commit only intentional stabilization files; do not stage unrelated review material.

### Exact resume commands

```powershell
# Confirm no orphaned app/build process is holding the interactive desktop.
Get-Process RedSalamander,MSBuild,cl,link -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding

# Inspect the current blocker.
rg -n "hot_paths_browse|Hot Paths page did not settle|browse_live|WaitFor.*Hot Paths" `
  RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp `
  RedSalamander\SelfTest\Commands

# Focused Hot Paths browse proof/repro. Use Start-Process -Wait for GUI selftests.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-resume-focused-hotpaths-browse'
$args=@(
  '--commands-selftest',
  '--selftest-case=cmd_preferences_dialog_hot_paths_browse_live_dx_interaction',
  '--selftest-timeout-multiplier=2'
)
$p=Start-Process -FilePath (Resolve-Path '.\.build\x64\Debug\RedSalamander.exe') `
  -ArgumentList $args -PassThru -Wait
Start-Sleep -Seconds 3
Write-Host "processExit=$($p.ExitCode)"
Get-Content "C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands\results.json" -Raw
Get-Content "C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands\trace.txt" -Tail 120

# Local predecessor cluster from the last broad failure.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-resume-hotpaths-browse-predecessor-cluster'
$cases=@(
  'cmd_preferences_dialog_panes_page_uses_dxui_statics_and_toggles',
  'cmd_preferences_dialog_shell_footer_live_dx_interaction',
  'cmd_preferences_dialog_shell_footer_access_keys_route_expected_actions',
  'cmd_preferences_dialog_panes_roundtrip_restores_dxui_surface',
  'cmd_preferences_dialog_panes_live_dx_interaction',
  'cmd_preferences_dialog_panes_theme_cycle_keeps_surface_legible',
  'cmd_preferences_dialog_panes_tab_traversal_live_dx_interaction',
  'cmd_preferences_dialog_panes_history_size_live_dx_interaction',
  'cmd_preferences_dialog_panes_combo_then_toggle_live_dx_interaction',
  'cmd_preferences_dialog_hot_paths_page_uses_dxui_statics_and_toggles',
  'cmd_preferences_dialog_hot_paths_live_dx_interaction',
  'cmd_preferences_dialog_hot_paths_open_prefs_toggle_live_dx_interaction',
  'cmd_preferences_dialog_hot_paths_browse_live_dx_interaction'
) -join ','
$args=@('--commands-selftest', "--selftest-case=$cases", '--selftest-timeout-multiplier=2')
$p=Start-Process -FilePath (Resolve-Path '.\.build\x64\Debug\RedSalamander.exe') `
  -ArgumentList $args -PassThru -Wait
Start-Sleep -Seconds 3
Write-Host "processExit=$($p.ExitCode)"
Get-Content "C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands\results.json" -Raw
Get-Content "C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands\trace.txt" -Tail 160

# After the root-cause fix.
Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

### Resume cautions

- [ ] Do not run GUI/native selftests in parallel.
- [ ] Use `Start-Process -Wait` for direct native GUI selftests.
- [ ] After native GUI selftests, wait 3-6 seconds, then read settled
  `last_run\results.json` and `last_run\commands\results.json`.
- [ ] Shell exit code can be 0 even when a script prints `runnerExit=1`; inspect JSON and summaries.
- [ ] Pester can return control despite failures; inspect the pass/fail summary.
- [ ] Do not clean `C:\RSPerf` until useful manual and broad evidence has been archived.

### Worktree at save point

- [ ] Modified stabilization files visible in `git status`:
  `Common/Common/SettingsStore.cpp`,
  `Common/DxUi/DxUi.Menu.cpp`,
  `Common/DxUi/DxUi.h`,
  `RedSalamander/ManagePluginsDialog.cpp`,
  `RedSalamander/RedSalamander.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`,
  `Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md`,
  `Specs/Plans/WIP/README.md`,
  `Tests/DxUiTests/DxUiTests.Menu.cpp`,
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`,
  and `Tools/Tests/WingetValidation.Tests.ps1`.
- [ ] Untracked review/spec material visible in `git status`:
  `Specs/Core/Core_FileSystemBridge.md`,
  `Specs/Plans/Done/Operation_Causeway_CrossFsBridgeSecurityAndPerformanceRemediation_2026-07-07.md`,
  `Specs/Reviews/FileSystemBridge-2026-07-07-Findings.md`,
  and `Specs/Reviews/ThreeDayDiff-2026-07-06-Findings.md`. Do not stage/delete these unless they
  are intentionally part of the final scope.
- [ ] Local continuation evidence confirmed on disk includes:
  `C:\RSPerf\runs\manual-20260707T1504-preview-fallback-focus-wait\`,
  `C:\RSPerf\runs\manual-20260707T1506-preview-cluster-actual-to-fallback-focus-wait\`,
  `C:\RSPerf\runs\manual-20260707T1518-focused-quicksearch-sort-owned\`,
  `C:\RSPerf\runs\manual-20260707T1519-search-cluster-to-quicksearch-sort-owned\`,
  and `C:\RSPerf\runs\20260707T131755Z-80540-0d84ab2c177f4f92bef55cba18ec84d6\`.
- [ ] Earlier local continuation evidence may also still be on disk; use the superseded checkpoints
  below for exact older run ids if needed.
- [ ] Broad Commands and Suite Full are still not accepted. Do not move this plan to Done yet.

## Superseded resume checkpoint - 2026-07-07 14:55 +02:00

Kept for history only. Use the current 15:25 checkpoint above when resuming. This older checkpoint
predates the preview fallback fix, the Quick Search state-ownership fix, and the current Hot Paths
browse failure.

### Done / saved at 14:55

- [x] Final proof root is explicit and unchanged: use `REDSALAMANDER_TEST_ROOT=C:\RSPerf` for all
  final accepted runs.
- [x] Current checkout at save point: `cb2478689bc1b37afc6066a6a5994d8f135a56cc`.
- [x] No `RedSalamander.exe`, `MSBuild`, `cl`, or `link` process was running at this save point.
- [x] SettingsStore isolation is patched in `Common/Common/SettingsStore.cpp`: app settings route
  through `runs\<runId>\scratch\settings-store\RedSalamander\Settings` when
  `REDSALAMANDER_TEST_ROOT` and `REDSALAMANDER_TEST_RUN_ID` are set.
- [x] SettingsStore isolation has a source-contract guard in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`; it failed red before the patch and passed
  after it.
- [x] User-profile protection was proven in
  `C:\RSPerf\runs\manual-20260707T1336Z-settings-isolation-compare-commands\`: the real
  `%LOCALAPPDATA%\RedSalamander\Settings\RedSalamander-debug.settings.json` timestamp stayed
  unchanged while the isolated settings file was written under `C:\RSPerf\runs\...`.
- [x] `cmd_preferences_dialog_rapid_switches_keep_page_specific_uia_subtrees` was root-caused as a
  clean-settings fixture bug: the test expected editor file-action rows while
  `DefaultEditorFileActionsSettings()` is intentionally empty.
- [x] The rapid-switch case now seeds deterministic editor actions in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp` and
  restores `g_settings.fileActions.editors` with `wil::scope_exit`.
- [x] Rapid-switch red/green evidence under `C:\RSPerf`: `manual-20260707T1408Z-red-rapid-preferences`
  failed 0 / 1 / 0 before the fixture patch; `manual-20260707T1410Z-green-rapid-preferences`
  passed 1 / 0 / 0 after it.
- [x] The old clean-root rename and FolderView first symptoms did not reproduce focused after the
  rapid-switch fix: `manual-20260707T1411Z-focused-rename-prompt` passed 1 / 0 / 0 and
  `manual-20260707T1413Z-focused-folder-empty` passed 1 / 0 / 0.
- [x] A Connection Manager broad-only double-click symptom was investigated, but did not reproduce
  focused, in an immediate pair, or in the local predecessor window:
  `manual-20260707T1420Z-focused-connection-doubleclick` passed 1 / 0 / 0,
  `manual-20260707T1422Z-connection-live-then-doubleclick` passed 2 / 0 / 0, and
  `manual-20260707T1424Z-connection-local-window` passed 13 / 0 / 0. Do not patch it blindly;
  revisit only if it reappears.
- [x] A stale Preferences category-tree UIA selection was root-caused and fixed by waiting for the
  provider selection to settle on the expected tree item in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`.
- [x] Preferences category-tree verification passed after the settling patch:
  `manual-20260707T1434Z-green-pref-category-tree` passed 1 / 0 / 0.
- [x] Plugin configuration tab traversal first failed in broad Commands because logical focus was
  lost when native focus was temporarily missing; fallback then advanced from the wrong boundary.
- [x] `RedSalamander/ManagePluginsDialog.cpp` now records and recovers the logical debug focus index
  (`lastDebugFocusedHostIndex`) and fills the debug focus snapshot from the recovered host before
  falling back to the first visible interactive target.
- [x] A source-contract guard proves plugin-config debug tab traversal recovers from logical focus
  position before first-control fallback.
- [x] Focused/pair proof after the logical-focus recovery patch passed:
  `manual-20260707T1433Z-green-plugin-config-tab-indexed-recovery` passed 1 / 0 / 0 and
  `manual-20260707T1436Z-green-plugin-live-then-tab-indexed-recovery` passed 2 / 0 / 0.
- [x] Broad Commands then exposed a second root cause in the selftest helper: after a successful
  advance, the retry path could send another Tab instead of only waiting for the expected snapshot.
  Evidence is in
  `C:\RSPerf\runs\20260707T123429Z-93544-8a0409cbdde04f9480fbca3dc913a0ab\`.
- [x] `RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp` now separates
  `advanceOnce()` from `waitForExpectedSnapshot()`. After a successful advance, the retry only
  pumps and waits again; it does not send another Tab.
- [x] A source-contract guard proves plugin-config tab traversal retries cannot send an extra Tab
  after a successful advance.
- [x] Source contracts passed before the new preview-focus RED guard:
  `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`,
  89 passed / 0 failed.
- [x] Build after the retry-guard patch passed:
  `.build\logs\msbuild-20260707_143719_260.log`, with no MSBuild/MSVC warning or error diagnostics.
- [x] Focused plugin-config proof after the no-extra-advance retry patch passed:
  `manual-20260707T1446-plugin-config-tab-no-extra-advance`, 1 passed / 0 failed / 0 skipped.
- [x] Immediate plugin live->tab pair after the no-extra-advance retry patch passed:
  `manual-20260707T1447-plugin-live-then-tab-no-extra-advance`, 2 passed / 0 failed / 0 skipped
  after waiting for settled JSON.
- [x] Broad Commands fail-fast after the plugin fix exposed
  `cmd_connection_manager_window_rejects_blank_profile_name`: terminal-captured run
  `20260707T124652Z-96136-6d916f240bba41f8a345ed8a35e5c730`, 187 passed / 1 failed /
  600 skipped. That run directory is not currently present under `C:\RSPerf\runs\`.
- [x] The blank-profile failure did not reproduce focused or in its local predecessor cluster:
  `manual-20260707T1451-focused-blank-profile-name` passed 1 / 0 / 0 and
  `manual-20260707T1452-connection-cluster-to-blank-name` passed 6 / 0 / 0.
- [x] `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp` now has diagnostic-only
  snapshot tracing around `TestConnectionManagerWindowRejectsProfileName(...)` to make the next
  broad recurrence root-causeable instead of guessed.
- [x] Build after the Connection Manager diagnostic patch passed:
  `.build\logs\msbuild-20260707_145003_180.log`, with no MSBuild/MSVC warning or error diagnostics.
- [x] Broad Commands fail-fast after the diagnostic patch exposed the current actionable issue:
  `C:\RSPerf\runs\20260707T125215Z-95664-57c41f4f01cd40d684781d686c2dd3e5\`, 63 passed /
  1 failed / 724 skipped, failing
  `pane_view_options_preview_falls_back_to_item_properties_when_no_embedded_preview_matches`.
- [x] Current preview fallback broad failure was narrowed to a test race: the test waits for fallback
  preview text, then samples source-pane focus separately. Existing neighboring preview tests already
  put expected focus into the wait predicate.
- [x] Focused preview fallback proof passed:
  `manual-20260707T1455-focused-preview-fallback-focus`, 1 passed / 0 failed / 0 skipped.
- [x] Preview predecessor cluster proof passed:
  `manual-20260707T1456-preview-cluster-to-fallback`, 4 passed / 0 failed / 0 skipped.
- [x] A new RED source-contract guard has been added to
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` requiring the preview fallback test to wait for
  fallback text and source-pane focus together. Current Pester result is intentionally red:
  89 passed / 1 failed / 0 skipped.

### Blocker / next edit at 14:55

- [ ] Patch `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`.
- [ ] Add helper `WaitForPreviewPaneTextAndFocus(...)` near `WaitForPreviewPaneText(...)`.
- [ ] The helper must poll `DebugGetPreviewPaneSnapshot(...)`, expected preview text, forbidden
  preview text, and `g_folderWindow.GetFocusedFolderViewHwnd() == expectedFocus` in one condition.
- [ ] Replace both `WaitForPreviewPaneText(...)` calls in
  `TestPaneViewOptionsPreviewFallsBackToItemPropertiesWhenNoEmbeddedPreviewMatches(...)` with
  `WaitForPreviewPaneTextAndFocus(..., expectedFocus, ..., ...)`.
- [ ] Keep or improve the explicit post-wait focus assertions, but the acceptance requirement is that
  the wait predicate proves text and source-pane focus together.
- [ ] Re-run the source-contract guard and expect 90 passed / 0 failed.

### Remaining at 14:55

- [ ] Run source contracts after the preview-focus patch.
- [ ] Rebuild `RedSalamander`.
- [ ] Re-run focused preview fallback under `C:\RSPerf`.
- [ ] Re-run the preview predecessor cluster under `C:\RSPerf`.
- [ ] Re-run broad Commands fail-fast under `C:\RSPerf`.
- [ ] If broad Commands finds another broad-only failure, identify it, root-cause it, and fix or
  replace the brittle path. Do not label it flaky and move on.
- [ ] If `cmd_connection_manager_window_rejects_blank_profile_name` reappears, use the new diagnostic
  snapshot traces before changing behavior.
- [ ] Re-run broad Commands without `-FailFast` under `C:\RSPerf`.
- [ ] Re-run Suite Full under `C:\RSPerf`.
- [ ] If still failing, explain or fix `RedSalamanderMonitorEtwLatency`; a prior Suite Full saw exit
  1 with an empty output log.
- [ ] If still failing, explain or fix `ToolsPesterTests`; a prior Suite Full saw exit 1 with no
  captured output log.
- [ ] Archive final passing evidence under `Specs/TestRuns/<current-head>\...` or another clearly
  named continuation folder before cleaning volatile `C:\RSPerf` evidence.
- [ ] Clean or reset `C:\RSPerf` only after needed evidence is archived; current manual run folders
  make disk-audit checks dirty.
- [ ] Before closeout, move this plan to `Specs/Plans/Done/` and merge durable requirements into
  `Specs/Testing/*`, `Tests/README.md`, `README.md`, or `AGENTS.md` as appropriate.
- [ ] Commit only intentional stabilization files; do not stage unrelated review material.

### Exact resume commands

```powershell
# Confirm no orphaned app/build process is holding the interactive desktop.
Get-Process RedSalamander,MSBuild,cl,link -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding

# Inspect the exact patch site.
rg -n "WaitForPreviewPaneText|PreviewFallsBackToItemProperties" `
  RedSalamander\SelfTest\Commands\Commands.SelfTest.Settings.cpp `
  Tools\Tests\TestHarnessSourceContracts.Tests.ps1

# After patching Settings.cpp, source contract must go green.
Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru

# Rebuild before native selftests.
.\build.ps1 -ProjectName RedSalamander

# Focused preview fallback proof after the combined text+focus wait patch.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-resume-preview-fallback-focus-wait'
.\.build\x64\Debug\RedSalamander.exe `
  --commands-selftest `
  --selftest-case=pane_view_options_preview_falls_back_to_item_properties_when_no_embedded_preview_matches `
  --selftest-timeout-multiplier=2
Start-Sleep -Seconds 4
Get-Content "C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands\results.json" -Raw

# Preview predecessor cluster.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-resume-preview-cluster-to-fallback'
.\.build\x64\Debug\RedSalamander.exe `
  --commands-selftest `
  --selftest-case=pane_view_options_toggle_preview_pane_tabs_and_selection,pane_view_options_configured_preview_handler_uses_embedded_preview,pane_view_options_built_in_empty_assoc_uses_embedded_preview,pane_view_options_preview_falls_back_to_item_properties_when_no_embedded_preview_matches `
  --selftest-timeout-multiplier=2
Start-Sleep -Seconds 4
Get-Content "C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands\results.json" -Raw

# Broad gates. Use JSON results as truth; native process exit can be less informative.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

### Resume cautions

- [ ] Do not run GUI/native selftests in parallel.
- [ ] Direct native selftests can exit before JSON is fully flushed; wait 3-6 seconds, then read
  settled `last_run\results.json` and `last_run\commands\results.json`.
- [ ] Shell exit code can be 0 even when a script prints `runnerExit=1`; inspect the JSON and summary.
- [ ] Pester can return control despite failures; inspect the pass/fail summary.
- [ ] Do not clean `C:\RSPerf` until useful manual and broad evidence has been archived.

### Worktree at save point

- [ ] Intentional dirty stabilization files:
  `Common/Common/SettingsStore.cpp`,
  `Common/DxUi/DxUi.h`,
  `Common/DxUi/DxUi.Menu.cpp`,
  `RedSalamander/ManagePluginsDialog.cpp`,
  `RedSalamander/RedSalamander.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`,
  `Tests/DxUiTests/DxUiTests.Menu.cpp`,
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`,
  `Tools/Tests/WingetValidation.Tests.ps1`,
  and this plan.
- [ ] Untracked review/spec material present at save point:
  `Specs/Core/Core_FileSystemBridge.md`,
  `Specs/Reviews/FileSystemBridge-2026-07-07-Findings.md`,
  and `Specs/Reviews/ThreeDayDiff-2026-07-06-Findings.md`. Do not stage/delete these unless they
  are intentionally part of the final scope.
- [ ] Additional plan-folder changes visible in `git status` but not classified as part of this
  test-suite stabilization checkpoint:
  `Specs/Plans/WIP/README.md` and
  `Specs/Plans/Done/Operation_Causeway_CrossFsBridgeSecurityAndPerformanceRemediation_2026-07-07.md`.
  Do not stage/delete them unless they are intentionally part of the final scope.
- [ ] Local continuation evidence confirmed on disk includes:
  `C:\RSPerf\runs\manual-20260707T1446-plugin-config-tab-no-extra-advance\`,
  `C:\RSPerf\runs\manual-20260707T1447-plugin-live-then-tab-no-extra-advance\`,
  `C:\RSPerf\runs\manual-20260707T1451-focused-blank-profile-name\`,
  `C:\RSPerf\runs\manual-20260707T1452-connection-cluster-to-blank-name\`,
  `C:\RSPerf\runs\20260707T125215Z-95664-57c41f4f01cd40d684781d686c2dd3e5\`,
  `C:\RSPerf\runs\manual-20260707T1455-focused-preview-fallback-focus\`,
  and `C:\RSPerf\runs\manual-20260707T1456-preview-cluster-to-fallback\`.
- [ ] Earlier local continuation evidence may also still be on disk:
  `C:\RSPerf\runs\manual-20260707T1336Z-settings-isolation-compare-commands\`,
  `C:\RSPerf\runs\manual-20260707T1408Z-red-rapid-preferences\`,
  `C:\RSPerf\runs\manual-20260707T1410Z-green-rapid-preferences\`,
  `C:\RSPerf\runs\manual-20260707T1411Z-focused-rename-prompt\`,
  `C:\RSPerf\runs\manual-20260707T1413Z-focused-folder-empty\`,
  `C:\RSPerf\runs\manual-20260707T1420Z-focused-connection-doubleclick\`,
  `C:\RSPerf\runs\manual-20260707T1422Z-connection-live-then-doubleclick\`,
  `C:\RSPerf\runs\manual-20260707T1424Z-connection-local-window\`,
  `C:\RSPerf\runs\manual-20260707T1434Z-green-pref-category-tree\`,
  `C:\RSPerf\runs\manual-20260707T1433Z-green-plugin-config-tab-indexed-recovery\`,
  `C:\RSPerf\runs\manual-20260707T1436Z-green-plugin-live-then-tab-indexed-recovery\`,
  and `C:\RSPerf\runs\20260707T123429Z-93544-8a0409cbdde04f9480fbca3dc913a0ab\`.
- [ ] Terminal-captured-only broad fail-fast summaries whose run directories are not currently
  present under `C:\RSPerf\runs\`:
  `20260707T120651Z-96976-6900d35244d6496b8c95dd8924e16a12`,
  `20260707T121109Z-97900-1fbb21ad448a4192a1bf5afea1993a80`, and
  `20260707T124652Z-96136-6d916f240bba41f8a345ed8a35e5c730`.
- [ ] Broad Commands and Suite Full are still not accepted. Do not move this plan to Done yet.

## Superseded resume checkpoint - 2026-07-07 14:41 +02:00

Kept for history only. Use the current 14:55 checkpoint above when resuming. This older checkpoint
predates the post-retry-guard plugin proof, the Connection Manager diagnostic probe, and the preview
fallback focus RED guard.

### Done / saved at 14:41

- [x] Final proof root is explicit and unchanged: use `REDSALAMANDER_TEST_ROOT=C:\RSPerf` for all
  final accepted runs.
- [x] Current checkout at save point: `cb2478689bc1b37afc6066a6a5994d8f135a56cc`.
- [x] No `RedSalamander.exe`, `MSBuild`, `cl`, or `link` process was running at this save point.
- [x] `Common/Common/SettingsStore.cpp` routes app settings through the unified test root when
  `REDSALAMANDER_TEST_ROOT` and `REDSALAMANDER_TEST_RUN_ID` are present:
  `runs\<runId>\scratch\settings-store\RedSalamander\Settings`.
- [x] SettingsStore isolation has a source-contract guard in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`; it failed red before the patch and passed
  after it.
- [x] User-profile protection was proven in
  `C:\RSPerf\runs\manual-20260707T1336Z-settings-isolation-compare-commands\`: the real
  `%LOCALAPPDATA%\RedSalamander\Settings\RedSalamander-debug.settings.json` timestamp stayed
  unchanged while the isolated test settings file was written under `C:\RSPerf\runs\...`.
- [x] `cmd_preferences_dialog_rapid_switches_keep_page_specific_uia_subtrees` was root-caused as a
  clean-settings fixture bug: the test expected editor file-action rows while
  `DefaultEditorFileActionsSettings()` is intentionally empty.
- [x] The rapid-switch case now seeds deterministic editor actions in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp` and
  restores `g_settings.fileActions.editors` with `wil::scope_exit`.
- [x] Rapid-switch red/green evidence under `C:\RSPerf`: `manual-20260707T1408Z-red-rapid-preferences`
  failed 0 / 1 / 0 before the fixture patch; `manual-20260707T1410Z-green-rapid-preferences`
  passed 1 / 0 / 0 after it.
- [x] The old clean-root rename and FolderView first symptoms did not reproduce focused after the
  rapid-switch fix: `manual-20260707T1411Z-focused-rename-prompt` passed 1 / 0 / 0 and
  `manual-20260707T1413Z-focused-folder-empty` passed 1 / 0 / 0.
- [x] A Connection Manager broad-only double-click symptom was investigated, but did not reproduce
  focused, in an immediate pair, or in the local predecessor window:
  `manual-20260707T1420Z-focused-connection-doubleclick` passed 1 / 0 / 0,
  `manual-20260707T1422Z-connection-live-then-doubleclick` passed 2 / 0 / 0, and
  `manual-20260707T1424Z-connection-local-window` passed 13 / 0 / 0. Do not patch it blindly;
  revisit only if it reappears.
- [x] A stale Preferences category-tree UIA selection was root-caused and fixed by waiting for the
  provider selection to settle on the expected tree item in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`.
- [x] Preferences category-tree verification passed after the settling patch:
  `manual-20260707T1434Z-green-pref-category-tree` passed 1 / 0 / 0.
- [x] Plugin configuration tab traversal first failed in broad Commands because logical focus was
  lost when native focus was temporarily missing; fallback then advanced from the wrong boundary.
- [x] `RedSalamander/ManagePluginsDialog.cpp` now records and recovers the logical debug focus index
  (`lastDebugFocusedHostIndex`) and fills the debug focus snapshot from the recovered host before
  falling back to the first visible interactive target.
- [x] A source-contract guard proves plugin-config debug tab traversal recovers from logical focus
  position before first-control fallback.
- [x] Focused/pair proof after the logical-focus recovery patch passed:
  `manual-20260707T1433Z-green-plugin-config-tab-indexed-recovery` passed 1 / 0 / 0 and
  `manual-20260707T1436Z-green-plugin-live-then-tab-indexed-recovery` passed 2 / 0 / 0.
- [x] Broad Commands then exposed a second root cause in the selftest helper: after a successful
  advance, the retry path could send another Tab instead of only waiting for the expected snapshot.
  Evidence is in
  `C:\RSPerf\runs\20260707T123429Z-93544-8a0409cbdde04f9480fbca3dc913a0ab\`.
- [x] `RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp` now separates
  `advanceOnce()` from `waitForExpectedSnapshot()`. After a successful advance, the retry only
  pumps and waits again; it does not send another Tab.
- [x] A source-contract guard proves plugin-config tab traversal retries cannot send an extra Tab
  after a successful advance.
- [x] Latest source contracts passed after the retry-guard patch:
  `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`,
  89 passed / 0 failed.
- [x] Latest build after the retry-guard patch completed in
  `.build\logs\msbuild-20260707_143719_260.log`; the log contains no MSBuild/MSVC error or warning
  diagnostics and produced `.build\x64\Debug\RedSalamander.exe`.

### Remaining at 14:41

- [ ] Re-run focused plugin configuration tab traversal after the no-extra-advance retry patch.
- [ ] Re-run the immediate plugin live->tab pair after the no-extra-advance retry patch.
- [ ] Re-run broad Commands fail-fast under `C:\RSPerf`.
- [ ] If broad Commands finds another broad-only failure, identify it, root-cause it, and fix or
  replace the brittle test path. Do not label it flaky and move on.
- [ ] Re-run broad Commands without `-FailFast` under `C:\RSPerf`.
- [ ] Re-run Suite Full under `C:\RSPerf`.
- [ ] If still failing, explain or fix `RedSalamanderMonitorEtwLatency`; a prior Suite Full saw exit
  1 with an empty output log.
- [ ] If still failing, explain or fix `ToolsPesterTests`; a prior Suite Full saw exit 1 with no
  captured output log.
- [ ] Archive final passing evidence under `Specs/TestRuns/<current-head>\...` or another clearly
  named continuation folder before cleaning volatile `C:\RSPerf` evidence.
- [ ] Clean or reset `C:\RSPerf` only after needed evidence is archived; current manual run folders
  make disk-audit checks dirty.
- [ ] Before closeout, move this plan to `Specs/Plans/Done/` and merge durable requirements into
  `Specs/Testing/*`, `Tests/README.md`, `README.md`, or `AGENTS.md` as appropriate.
- [ ] Commit only intentional stabilization files; do not stage unrelated review material.

### Latest blocker evidence before the retry-guard patch

- [x] Broad fail-fast run:
  `C:\RSPerf\runs\20260707T123429Z-93544-8a0409cbdde04f9480fbca3dc913a0ab\`.
- [x] Result:
  `run-all-tests-results.json` = 174 passed / 1 failed / 613 skipped / 87918 ms.
- [x] Failed case:
  `cmd_plugin_configuration_dialog_tab_traversal_live_dx_interaction`.
- [x] Failure reason:
  `Plugin configuration dialog tab traversal did not complete; last focused label was 'reverse Use HTTPS toggle'.`
- [x] Important trace sequence:
  `currentIndex=4 -> nextIndex=3` succeeded for reverse Verify TLS, then
  `currentIndex=3 -> nextIndex=2` succeeded for reverse Use HTTPS, then the old retry sent another
  reverse Tab after the snapshot wait missed and ended on `Cancel`.
- [x] This evidence is now considered addressed by the `advanceOnce()` / `waitForExpectedSnapshot()`
  split, but it still needs focused/pair/broad verification.

### Exact resume commands

```powershell
# Confirm no orphaned app/build process is holding the interactive desktop.
Get-Process RedSalamander,MSBuild,cl,link -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding

# Confirm source-contract and build state after the latest patches.
Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander

# Focused plugin-config tab traversal after the retry-guard patch.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-resume-plugin-config-tab-no-extra-advance'
.\.build\x64\Debug\RedSalamander.exe `
  --commands-selftest `
  --selftest-case=cmd_plugin_configuration_dialog_tab_traversal_live_dx_interaction `
  --selftest-timeout-multiplier=2
Start-Sleep -Seconds 2
Get-Content "C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands\results.json" -Raw

# Immediate predecessor pair after the retry-guard patch.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-resume-plugin-live-then-tab-no-extra-advance'
.\.build\x64\Debug\RedSalamander.exe `
  --commands-selftest `
  --selftest-case=cmd_plugin_configuration_dialog_live_dx_interaction,cmd_plugin_configuration_dialog_tab_traversal_live_dx_interaction `
  --selftest-timeout-multiplier=2
Start-Sleep -Seconds 2
Get-Content "C:\RSPerf\runs\$env:REDSALAMANDER_TEST_RUN_ID\artifacts\selftest\last_run\commands\results.json" -Raw

# Broad gates. Use JSON results as truth; native process exit can be less informative.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

### Worktree at save point

- [ ] Intentional dirty stabilization files:
  `Common/Common/SettingsStore.cpp`,
  `Common/DxUi/DxUi.h`,
  `Common/DxUi/DxUi.Menu.cpp`,
  `RedSalamander/ManagePluginsDialog.cpp`,
  `RedSalamander/RedSalamander.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`,
  `Tests/DxUiTests/DxUiTests.Menu.cpp`,
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`,
  `Tools/Tests/WingetValidation.Tests.ps1`,
  and this plan.
- [ ] Untracked review/spec material present at save point:
  `Specs/Core/Core_FileSystemBridge.md`,
  `Specs/Reviews/FileSystemBridge-2026-07-07-Findings.md`,
  and `Specs/Reviews/ThreeDayDiff-2026-07-06-Findings.md`. Do not stage/delete these unless they
  are intentionally part of the final scope.
- [ ] Local continuation evidence still on disk includes:
  `C:\RSPerf\runs\manual-20260707T1336Z-settings-isolation-compare-commands\`,
  `C:\RSPerf\runs\manual-20260707T1408Z-red-rapid-preferences\`,
  `C:\RSPerf\runs\manual-20260707T1410Z-green-rapid-preferences\`,
  `C:\RSPerf\runs\manual-20260707T1411Z-focused-rename-prompt\`,
  `C:\RSPerf\runs\manual-20260707T1413Z-focused-folder-empty\`,
  `C:\RSPerf\runs\manual-20260707T1420Z-focused-connection-doubleclick\`,
  `C:\RSPerf\runs\manual-20260707T1422Z-connection-live-then-doubleclick\`,
  `C:\RSPerf\runs\manual-20260707T1424Z-connection-local-window\`,
  `C:\RSPerf\runs\manual-20260707T1434Z-green-pref-category-tree\`,
  `C:\RSPerf\runs\manual-20260707T1433Z-green-plugin-config-tab-indexed-recovery\`,
  `C:\RSPerf\runs\manual-20260707T1436Z-green-plugin-live-then-tab-indexed-recovery\`,
  and `C:\RSPerf\runs\20260707T123429Z-93544-8a0409cbdde04f9480fbca3dc913a0ab\`.
- [ ] Terminal-captured-only broad fail-fast summaries whose run directories are not currently
  present under `C:\RSPerf\runs\`:
  `20260707T120651Z-96976-6900d35244d6496b8c95dd8924e16a12` and
  `20260707T121109Z-97900-1fbb21ad448a4192a1bf5afea1993a80`.
- [ ] Broad Commands and Suite Full are still not accepted. Do not move this plan to Done yet.

## Superseded resume checkpoint - 2026-07-07 14:23 +02:00

Kept for history only; use the current resume checklist above when resuming. At this checkpoint the
plan was still not done because broad Commands and Suite Full were not green under the final root
`C:\RSPerf`.

### Done / saved

- [x] Final root requirement is explicit: all final proof runs use
  `REDSALAMANDER_TEST_ROOT=C:\RSPerf`.
- [x] Current checkout at save point: `cb2478689bc1`.
- [x] No `RedSalamander.exe` process was running at this save point.
- [x] `Common/Common/SettingsStore.cpp` now routes app settings through the unified test root when
  `REDSALAMANDER_TEST_ROOT` and `REDSALAMANDER_TEST_RUN_ID` are present:
  `runs\<runId>\scratch\settings-store\RedSalamander\Settings`.
- [x] `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` has a source-contract guard for that
  SettingsStore isolation. It failed red before the patch and passed after it.
- [x] User-profile protection was proven in
  `manual-20260707T1336Z-settings-isolation-compare-commands`: the real
  `%LOCALAPPDATA%\RedSalamander\Settings\RedSalamander-debug.settings.json` timestamp stayed
  unchanged, while the isolated settings file was written under `C:\RSPerf\runs\...`.
- [x] `cmd_preferences_dialog_rapid_switches_keep_page_specific_uia_subtrees` was root-caused as a
  clean-settings fixture bug: it expected editor file-action rows while
  `DefaultEditorFileActionsSettings()` is intentionally empty.
- [x] The rapid-switch case now seeds deterministic editor actions in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp` and
  restores `g_settings.fileActions.editors` with `wil::scope_exit`.
- [x] Rapid-switch red/green evidence under `C:\RSPerf`:
  `manual-20260707T1408Z-red-rapid-preferences` failed 0 / 1 / 0 before the fixture patch;
  `manual-20260707T1410Z-green-rapid-preferences` passed 1 / 0 / 0 after it.
- [x] The old clean-root rename and FolderView first symptoms did not reproduce focused after the
  rapid-switch fix:
  `manual-20260707T1411Z-focused-rename-prompt` passed 1 / 0 / 0 and
  `manual-20260707T1413Z-focused-folder-empty` passed 1 / 0 / 0.
- [x] A broad fail-fast run then exposed a Connection Manager broad-only double-click symptom. The
  run directory is no longer present, so keep this as terminal-captured evidence:
  `20260707T120651Z-96976-6900d35244d6496b8c95dd8924e16a12` passed 200 / failed 1 / skipped 587
  on `cmd_connection_manager_window_textfield_doubleclick_selects_word`.
- [x] The Connection Manager symptom did not reproduce focused, in an immediate pair, or in the
  local predecessor window:
  `manual-20260707T1420Z-focused-connection-doubleclick` passed 1 / 0 / 0,
  `manual-20260707T1422Z-connection-live-then-doubleclick` passed 2 / 0 / 0, and
  `manual-20260707T1424Z-connection-local-window` passed 13 / 0 / 0. Do not patch it blindly.
- [x] A later broad fail-fast run exposed a stale Preferences category-tree UIA selection. The run
  directory is no longer present, so keep this as terminal-captured evidence:
  `20260707T121109Z-97900-1fbb21ad448a4192a1bf5afea1993a80` passed 221 / failed 1 / skipped 566
  because the expected selected item was `General` but the observed selected item was `S3`.
- [x] `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp` now waits
  for the category-tree UIA selection to settle on the expected tree item instead of sampling a
  stale provider state once.
- [x] Preferences category-tree verification passed after the settling patch:
  `manual-20260707T1434Z-green-pref-category-tree` passed 1 / 0 / 0.
- [x] Latest source contracts passed after the current patches:
  `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`,
  87 passed / 0 failed.
- [x] Latest build passed after the current patches, 0 warnings / 0 errors:
  `.build\logs\msbuild-20260707_141556_603.log`.
- [x] Latest broad fail-fast run reached the current blocker and remains on disk:
  `C:\RSPerf\runs\20260707T121827Z-94508-93bad59133984b45b29959a3b747c6e0\artifacts\selftest\last_run\run-all-tests-results.json`,
  174 passed / 1 failed / 613 skipped.
- [x] Current blocker focused probes passed:
  `manual-20260707T1440Z-focused-plugin-config-tab` passed 1 / 0 / 0 and
  `manual-20260707T1442Z-plugin-live-then-tab` passed 2 / 0 / 0.

### Remaining gates

- [ ] Root-cause and fix/replace the current broad-only plugin configuration tab-traversal failure.
  It is not accepted as flaky, and it must not be made non-blocking.
- [ ] Re-run source contracts and rebuild after the plugin configuration fix.
- [ ] Re-run focused plugin configuration tab traversal, the immediate plugin live->tab pair, and a
  broad Commands fail-fast run under `C:\RSPerf`.
- [ ] Re-run broad Commands without `-FailFast` under `C:\RSPerf`.
- [ ] Re-run Suite Full under `C:\RSPerf`.
- [ ] If still failing, explain or fix `RedSalamanderMonitorEtwLatency`; a prior Suite Full saw exit
  1 with an empty output log.
- [ ] If still failing, explain or fix `ToolsPesterTests`; a prior Suite Full saw exit 1 with no
  captured output log.
- [ ] Archive final passing evidence under `Specs/TestRuns/<current-head>\...` or another clearly
  named continuation folder.
- [ ] Before closeout, move this plan to `Specs/Plans/Done/` and merge durable requirements into
  `Specs/Testing/*`, `Tests/README.md`, `README.md`, or `AGENTS.md` as appropriate.
- [ ] Commit only intentional stabilization files; do not stage unrelated review material.

### Current blocker

- [ ] `cmd_plugin_configuration_dialog_tab_traversal_live_dx_interaction` failed only in the latest
  broad fail-fast run:
  `C:\RSPerf\runs\20260707T121827Z-94508-93bad59133984b45b29959a3b747c6e0\`.
- [ ] Failure reason:
  `Plugin configuration dialog tab traversal did not complete; last focused label was 'reverse Verify TLS certificate toggle'.`
- [ ] Trace diagnostic to inspect:
  `plugin-config tab step miss: step='reverse Verify TLS certificate toggle' expectedKind=4 expectedLabel='Verify TLS certificate' labelMatches=0 actualFocusKind=6 actualFocusLabel='OK' usesButtons=1 usesSurface=1 usesStatics=1 usesInputs=1 dxButtons=2 dxStatics=21 dxInputs=8 legacyButtons=0 legacyControls=0 legacyStatics=0 scrollY=0`.
- [ ] Trace also showed:
  `plugin-config advance command: no focused HWND; attempting fallback focus recovery`.
- [ ] Working hypothesis to prove or disprove before patching:
  `DebugAdvancePluginConfigurationDialogTab(reverse)` loses the logical last-focused host when
  Win32 focus is temporarily missing, so fallback recovery can advance from the wrong boundary and
  leave focus on `OK` when the test expects `Verify TLS certificate`.
- [ ] Do not solve this by adding blind sleeps or by weakening the assertion. The fix should preserve
  deterministic logical focus/advance semantics or replace the brittle path with an equivalent
  deterministic test path.

### Next steps

- [ ] Inspect the plugin configuration tab helper and its fallback logic:
  `RedSalamander/ManagePluginsDialog.cpp` around `DebugAdvancePluginConfigurationDialogTab`,
  `plugin-config advance command`, `no focused HWND`, and fallback focus recovery.
- [ ] Inspect the failing test:
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp` around
  `TestPluginConfigurationDialogTabTraversalLiveDxInteraction`, `sendTab(...)`, and
  `plugin-config tab step miss`.
- [ ] Patch the smallest deterministic root cause. Likely acceptable outcomes:
  preserve/restore the last logical focused host across no-HWND fallback, or make the debug advance
  command advance from an explicit logical focus snapshot instead of an arbitrary recovered control.
- [ ] Verify with the focused case, immediate pair, source contracts, build, and broad Commands.
- [ ] Revisit the Connection Manager double-click terminal-captured symptom only if it reappears
  after the plugin configuration fix; the focused/pair/window probes already passed.

### Exact resume commands

```powershell
# Confirm no orphaned app process is holding the interactive desktop.
Get-Process RedSalamander -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding

# Inspect current broad blocker evidence.
$run='20260707T121827Z-94508-93bad59133984b45b29959a3b747c6e0'
$last="C:\RSPerf\runs\$run\artifacts\selftest\last_run"
Get-Content "$last\run-all-tests-results.json" -Raw | ConvertFrom-Json |
  Select-Object passed,failed,skipped,duration_ms | Format-List
Get-Content "$last\commands\trace.txt" |
  Select-String -Pattern 'plugin-config tab step miss|no focused HWND|fallback focus recovery' -Context 4,8

# Read the likely root-cause code.
$p='RedSalamander\ManagePluginsDialog.cpp'
$lines=Get-Content -LiteralPath $p
$lines[4215..4410]
$lines[4785..4845]
rg -n "DebugAdvancePluginConfigurationDialogTab|plugin-config advance command|no focused HWND|fallback focus recovery|reusing last focused|Advance.*PluginConfiguration|last focused" `
  RedSalamander\ManagePluginsDialog.cpp RedSalamander\ManagePluginsDialog.h

# Read the failing test.
$p='RedSalamander\SelfTest\Commands\Commands.SelfTest.PluginConfig.cpp'
$lines=Get-Content -LiteralPath $p
$lines[3940..4165]
$lines[5095..5115]

# Re-verify guard/build state before and after the fix.
Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander

# Focused current blocker and immediate predecessor pair.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-resume-focused-plugin-config-tab'
.\.build\x64\Debug\RedSalamander.exe `
  --commands-selftest `
  --selftest-case=cmd_plugin_configuration_dialog_tab_traversal_live_dx_interaction `
  --selftest-timeout-multiplier=2

$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-resume-plugin-live-then-tab'
.\.build\x64\Debug\RedSalamander.exe `
  --commands-selftest `
  --selftest-case=cmd_plugin_configuration_dialog_live_dx_interaction,cmd_plugin_configuration_dialog_tab_traversal_live_dx_interaction `
  --selftest-timeout-multiplier=2

# Final gates after clustered fixes.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

### Worktree at save point

- [ ] Intentional dirty files:
  `Common/Common/SettingsStore.cpp`,
  `Common/DxUi/DxUi.h`,
  `Common/DxUi/DxUi.Menu.cpp`,
  `RedSalamander/RedSalamander.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`,
  `Tests/DxUiTests/DxUiTests.Menu.cpp`,
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`,
  `Tools/Tests/WingetValidation.Tests.ps1`,
  and this plan.
- [ ] Local continuation evidence still on disk:
  `C:\RSPerf\runs\manual-20260707T1336Z-settings-isolation-compare-commands\`,
  `C:\RSPerf\runs\manual-20260707T1408Z-red-rapid-preferences\`,
  `C:\RSPerf\runs\manual-20260707T1410Z-green-rapid-preferences\`,
  `C:\RSPerf\runs\manual-20260707T1411Z-focused-rename-prompt\`,
  `C:\RSPerf\runs\manual-20260707T1413Z-focused-folder-empty\`,
  `C:\RSPerf\runs\manual-20260707T1420Z-focused-connection-doubleclick\`,
  `C:\RSPerf\runs\manual-20260707T1422Z-connection-live-then-doubleclick\`,
  `C:\RSPerf\runs\manual-20260707T1424Z-connection-local-window\`,
  `C:\RSPerf\runs\manual-20260707T1434Z-green-pref-category-tree\`,
  `C:\RSPerf\runs\20260707T121827Z-94508-93bad59133984b45b29959a3b747c6e0\`,
  `C:\RSPerf\runs\manual-20260707T1440Z-focused-plugin-config-tab\`,
  and `C:\RSPerf\runs\manual-20260707T1442Z-plugin-live-then-tab\`.
- [ ] Terminal-captured-only broad fail-fast summaries whose run directories are not currently
  present under `C:\RSPerf\runs\`:
  `20260707T120651Z-96976-6900d35244d6496b8c95dd8924e16a12` and
  `20260707T121109Z-97900-1fbb21ad448a4192a1bf5afea1993a80`.
- [ ] Untracked `Specs/Reviews/ThreeDayDiff-2026-07-06-Findings.md` is unrelated review material;
  do not stage/delete it unless the user explicitly asks.
- [ ] Broad Commands and Suite Full are still not accepted. Do not move this plan to Done yet.

## Superseded resume checkpoint - 2026-07-07 14:00 +02:00

Kept for history only; use the topmost current resume checklist above when resuming. At this
checkpoint the plan was still not done because broad Commands and Suite Full were not green under
the final root `C:\RSPerf`.

### Done / saved

- [x] Final root requirement is explicit: all final proof runs use
  `REDSALAMANDER_TEST_ROOT=C:\RSPerf`.
- [x] Current checkout at save point: `cb2478689`.
- [x] No `RedSalamander.exe` process was running at this save point.
- [x] Added a source-contract guard proving app settings must be routed through the unified
  `TestSandbox` root during test runs. New contract in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` requires `Common/Common/SettingsStore.cpp` to
  consume `REDSALAMANDER_TEST_ROOT` + `REDSALAMANDER_TEST_RUN_ID` and build
  `runs\<runId>\scratch\settings-store\RedSalamander\Settings`.
- [x] Verified the guard failed red before the code change: source contracts were 86 passed /
  1 failed because `SettingsStore.cpp` did not mention `REDSALAMANDER_TEST_ROOT`.
- [x] Patched `Common/Common/SettingsStore.cpp` so `GetSettingsDirectoryPath()` first uses
  `GetUnifiedTestSettingsDirectoryPathFromEnvironment()` when both test-root env vars are present,
  falling back to `%LOCALAPPDATA%\RedSalamander\Settings` only for normal app runs.
- [x] The test-root settings path rejects unsafe run ids and normalizes the configured root before
  composing the selftest path.
- [x] Source contracts passed after the `SettingsStore` patch:
  `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`,
  87 passed / 0 failed.
- [x] Build passed after the `SettingsStore` patch, 0 warnings / 0 errors:
  `.build\logs\msbuild-20260707_133322_749.log`.
- [x] Manual Compare-then-Commands repro was rerun under clean settings isolation:
  `manual-20260707T1336Z-settings-isolation-compare-commands`.
  Compare passed 219 / 0 / 30; Commands failed 750 / 36 / 2.
- [x] That manual repro proved user settings isolation works: real
  `%LOCALAPPDATA%\RedSalamander\Settings\RedSalamander-debug.settings.json` timestamp stayed at
  2026-07-07 11:20:43 UTC before and after the run, and the isolated file was created at
  `C:\RSPerf\runs\manual-20260707T1336Z-settings-isolation-compare-commands\scratch\settings-store\RedSalamander\Settings\RedSalamander-debug.settings.json`.
- [x] Focused clean-root repro of
  `cmd_preferences_dialog_rapid_switches_keep_page_specific_uia_subtrees` failed 0 / 1 / 0 under
  `manual-20260707T1412Z-focused-clean-rapid-preferences`, proving at least the first Preferences
  failure is a hidden ambient-settings assumption, not Compare timing.
- [x] Do not back out settings isolation. It is a real fix: tests were previously reading/writing
  the user's app settings path, which invalidated the `C:\RSPerf` root contract.

### Remaining gates

- [ ] Fix the hidden ambient-settings assumptions exposed by clean isolation. First known case:
  `cmd_preferences_dialog_rapid_switches_keep_page_specific_uia_subtrees` expects editor file-action
  rows even though `DefaultEditorFileActionsSettings()` is empty under clean settings.
- [ ] Root-cause and fix/replace the clean-root rename prompt failures:
  `cmd_pane_rename_prompt_uses_dxui_surface`,
  `cmd_pane_rename_prompt_live_dx_interaction`, and
  `cmd_pane_rename_prompt_long_run_open_close_stays_stable`.
- [ ] Root-cause and fix/replace the clean-root folder/pane layout failures. Many failures report
  zero-width or unwarmed FolderView state after clean settings, so look for deterministic startup
  layout/default-settings assumptions before chasing 20 separate symptoms.
- [ ] Broad Commands still needs a fresh archived green run under `C:\RSPerf`.
- [ ] Suite Full still needs a green run under `C:\RSPerf`.
- [ ] `RedSalamanderMonitorEtwLatency` must be explained or fixed if it still fails after Commands
  is deterministic; the prior Suite Full saw exit 1 with an empty output log.
- [ ] `ToolsPesterTests` must be explained or fixed if it still fails after Commands is
  deterministic; the prior Suite Full saw exit 1 with no captured output log.
- [ ] Before closeout, move this plan to `Specs/Plans/Done/` and merge durable requirements into
  `Specs/Testing/*`, `Tests/README.md`, `README.md`, or `AGENTS.md` as appropriate.
- [ ] Commit only intentional stabilization files; do not stage unrelated review material.

### Current blocker

- [ ] **Primary blocker: clean-root Commands now fails because tests/app startup relied on ambient
  user settings.** This is progress, not a regression to hide: the suite is no longer polluting the
  user's settings file, but several Commands cases now need deterministic setup/defaults.
- [ ] Current clean-root Commands failure population from
  `C:\RSPerf\runs\manual-20260707T1336Z-settings-isolation-compare-commands\artifacts\selftest\last_run\commands\results.json`:
  - Preferences: `cmd_preferences_dialog_rapid_switches_keep_page_specific_uia_subtrees`.
  - Rename prompt / item properties: `cmd_pane_rename_prompt_uses_dxui_surface`,
    `cmd_pane_rename_prompt_live_dx_interaction`,
    `cmd_pane_rename_prompt_long_run_open_close_stays_stable`,
    `cmd_pane_itemProperties_window_long_run_open_close_stays_stable`.
  - App/menu/pane shell: `cmd_app_openDriveMenus_keeps_navigation_shell_stable`,
    `cmd_app_menuBar_persistent_direct_hover_switches_top_level_popup`,
    `cmd_app_menuBar_mouse_opened_popup_processes_keyboard_before_mouse_move`,
    `cmd_app_swapPanes_keeps_navigation_shell_stable`,
    `cmd_pane_statusBar_uses_owned_window_and_sort_click_opens_menu`.
  - FolderView correctness/perf: `folderView_empty_folder_state`,
    `folderView_refresh_to_paint_metric_clears_after_failed_render`,
    `folderView_render_device_loss_recovers`, `folderView_draw_item_brush_reuse_guard`,
    `folderView_dpi_change_repaints_both_panes`, `folderView_column_widths_audit`,
    `folderView_visible_column_widths`, `folderView_thumbnail_aspect_ratio`,
    `folderView_thumbnail_scroll_stress`, `folderView_thumbnail_scroll_requeues_visible`,
    `folderView_thumbnail_resize_requeues_visible`,
    `folderView_thumbnail_size_change_regenerates_fallback_icons`,
    `folderView_thumbnail_return_to_normal_icon_size`,
    `folderView_thumbnail_sort_popup_slider`, `folderView_perf_large_folder_baseline`,
    `folderView_perf_sort_toggle_stress`, `folderView_perf_overlay_invalidation_stress`,
    `folderView_perf_scroll_render_stress`, `folderView_perf_huge_folder_scale`,
    `folderView_perf_cold_first_visit`, `folderView_perf_slow_virtual_provider`,
    `folderView_perf_relayout_churn_while_scrolled`,
    `folderView_perf_directory_change_storm`, `folderView_perf_refresh_preservation`,
    `folderView_perf_icon_pipeline_cold_slow`, and `folderView_perf_iconcache_contention`.
- [ ] Do not classify those as flaky/non-blocking. For each cluster, either seed/restore the
  required deterministic settings in the test fixture, make app startup deterministic under clean
  settings, or replace the case with a deterministic equivalent.

### Next steps

- [ ] Inspect the failing Preferences test first:
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp`
  around `TestPreferencesDialogRapidSwitchesKeepPageSpecificUiaSubtrees`.
- [ ] Inspect deterministic file-action construction patterns in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp` and
  `Common/SettingsStore.h` before patching.
- [ ] Patch the Preferences rapid-switch case to seed and restore deterministic editor actions if
  the product default is intentionally empty. Likely pattern: save `g_settings.fileActions.editors`,
  restore with `wil::scope_exit`, add at least one editor `FileActionDefinition` plus association
  row before opening Preferences.
- [ ] Re-run focused
  `cmd_preferences_dialog_rapid_switches_keep_page_specific_uia_subtrees` under a fresh clean
  `C:\RSPerf` run id.
- [ ] After the Preferences case is fixed, run focused clean-root repros for one rename prompt
  failure and `folderView_empty_folder_state` to decide whether the remaining 35 failures are a few
  clustered roots or many independent test bugs.
- [ ] Re-run source contracts and rebuild after source edits.
- [ ] Re-run manual Compare-then-Commands, broad Commands, then Suite Full under `C:\RSPerf`.
- [ ] Archive final passing evidence under `Specs/TestRuns/<current-head>\...` or another clearly
  named continuation folder.

### Exact resume commands

```powershell
# Confirm no orphaned app process is holding the interactive desktop.
Get-Process RedSalamander -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding

# Inspect the latest clean-isolation Commands failure population.
$run='manual-20260707T1336Z-settings-isolation-compare-commands'
$manual="C:\RSPerf\runs\$run\artifacts\selftest\last_run\commands\results.json"
$j=Get-Content $manual -Raw | ConvertFrom-Json
$j | Select-Object passed,failed,skipped,duration_ms | Format-List
$j.cases | Where-Object status -eq 'failed' |
  Select-Object name,reason,duration_ms | Format-List

# Inspect the first known focused failure and settings construction helpers.
$p='RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp'
$lines=Get-Content $p
$lines[8210..8370]

$p='Common\SettingsStore.h'
$lines=Get-Content $p
$lines[430..675]

$p='RedSalamander\SelfTest\Commands\Commands.SelfTest.Settings.cpp'
$lines=Get-Content $p
$lines[160..225]

# Re-verify current guard/build state before changing more code.
Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander

# Focused repro for the first clean-root Preferences blocker.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-resume-focused-rapid-preferences'
.\.build\x64\Debug\RedSalamander.exe `
  --commands-selftest `
  --selftest-case=cmd_preferences_dialog_rapid_switches_keep_page_specific_uia_subtrees `
  --selftest-timeout-multiplier=2

# Final gates after clustered fixes.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

### Worktree at save point

- [ ] Intentional dirty files:
  `Common/Common/SettingsStore.cpp`,
  `Common/DxUi/DxUi.h`,
  `Common/DxUi/DxUi.Menu.cpp`,
  `RedSalamander/RedSalamander.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`,
  `Tests/DxUiTests/DxUiTests.Menu.cpp`,
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`,
  `Tools/Tests/WingetValidation.Tests.ps1`,
  and this plan.
- [ ] Local continuation evidence exists under `C:\RSPerf\runs\manual-20260707T1336Z-settings-isolation-compare-commands\`.
  The older copied evidence under
  `Specs\TestRuns\4cb089111a23\Continuation\2026-07-07_manual_compare_then_commands_cross_suite_failure\`
  is still useful history but predates the `SettingsStore` isolation patch.
- [ ] Untracked `Specs/Reviews/ThreeDayDiff-2026-07-06-Findings.md` is unrelated review material;
  do not stage/delete it unless the user explicitly asks.
- [ ] Broad Commands and Suite Full are still not accepted. Do not move this plan to Done yet.

## Superseded resume checkpoint - 2026-07-07 13:26 +02:00

Historical detail retained for audit. Do not use this as the active resume checklist; use the
topmost current resume checklist above.

### Done / saved

- [x] Final root requirement is explicit: all final proof runs use
  `REDSALAMANDER_TEST_ROOT=C:\RSPerf`.
- [x] Current checkout at save point: `cb2478689`.
- [x] No `RedSalamander.exe` process was running at this save point.
- [x] Latest source-contract suite passed after the Compare diagnostic guard was added:
  `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`,
  86 passed / 0 failed.
- [x] Latest completed build passed after source edits, 0 warnings / 0 errors:
  `.build\logs\msbuild-20260707_115733_574.log`.
- [x] Focused pair passed after that build:
  `20260707T095934Z-6440-0f6fa5aa4a2d424983808720cda84955`,
  2 passed / 0 failed, disk audit clean.
- [x] Compare predecessor window passed after that build:
  `20260707T095152Z-32356-3fae3786f36c4596b22d50ec9dcfd141`,
  21 passed / 0 failed, disk audit clean.
- [x] Menu predecessor window passed after that build:
  `20260707T095554Z-44604-ef5e1ab532504990bb13cfb490037840`,
  12 passed / 0 failed, disk audit clean.
- [x] Broad Commands passed once after that build:
  `20260707T095945Z-67328-f05e8d2fd91d4b0b8deffbb8ba0a1e18`,
  786 passed / 0 failed / 2 skipped, disk audit clean. This is useful evidence only; a later run
  cleaned those artifacts, so this does not satisfy final archival proof.
- [x] Suite Full was attempted after that build:
  `20260707T101710Z-82808-998b5f930245410aa3575b1337f6ee1d`,
  1103 passed / 18 failed / 52 skipped, disk audit clean. Later cleanup removed the artifacts, but
  the failure list below was captured from the dashboard output.
- [x] Focused rerun of `cmd_compare_directories_options_live_dx_body_interaction` passed after the
  Suite Full failure:
  `20260707T105906Z-81448-36196891d528438f9331e4a0926c8f50`, proving that failure is
  predecessor-sensitive, not focused deterministic.
- [x] Manual Compare-then-Commands reproduction under `C:\RSPerf` reproduced a narrower
  cross-suite Commands failure:
  `manual-20260707T110215Z-cross-suite-commands`; Compare passed 219 / 0 / 30, then Commands
  failed 782 / 4 / 2.
- [x] Manual reproduction artifacts were copied for resume under
  `Specs\TestRuns\4cb089111a23\Continuation\2026-07-07_manual_compare_then_commands_cross_suite_failure\`.
  `Copy-Item` reported transient reparse/symlink copy misses under Compare work folders; trust the
  copied `results.json` and command artifacts, not the copied work-tree scratch directories.

### Remaining gates

- [ ] Broad Commands still needs a fresh archived green run under `C:\RSPerf`.
- [ ] Suite Full still needs a green run under `C:\RSPerf`.
- [ ] `RedSalamanderMonitorEtwLatency` must be explained or fixed; Suite Full saw exit 1 with an
  empty output log.
- [ ] `ToolsPesterTests` must be explained or fixed; Suite Full saw exit 1 with no captured output
  log.
- [ ] Before closeout, move this plan to `Specs/Plans/Done/` and merge durable requirements into
  `Specs/Testing/*`, `Tests/README.md`, `README.md`, or `AGENTS.md` as appropriate.
- [ ] Commit only intentional stabilization files; do not stage unrelated review material.

### Current blocker

- [ ] **Primary blocker: cross-suite Compare -> Commands contamination.** Manual reproduction
  `manual-20260707T110215Z-cross-suite-commands` ran Compare first and then Commands in the same
  `C:\RSPerf` run id. Compare passed; Commands failed these four Preferences cases:
  - `cmd_preferences_dialog_keyboard_swap_assign_live_dx_interaction`: filtered DX row did not
    expose the original selected shortcut before invoking Assign for swap validation.
  - `cmd_preferences_dialog_plugins_custom_paths_remove_live_dx_interaction`: Plugins page did not
    restore the selected custom path after shell Cancel discarded pending remove.
  - `cmd_preferences_dialog_editors_mouse_roundtrip_restore_dxui_notes`: Editors page did not
    repaint and restore the stabilized one-host DxUi note surface after returning from General.
  - `cmd_preferences_dialog_category_tree_handles_reverse_keyboard_navigation`: VK_UP did not move
    exactly one visible category upward; captured snapshot showed focus on the category tree, page
    title `Monitor`, selected visible index 35, and zero named focus targets.
- [ ] Treat `%LOCALAPPDATA%\RedSalamander\Settings` as the prime suspect until disproven: the
  settings store is not rooted under `C:\RSPerf`, so Compare/Full can plausibly leave ambient user
  settings that Commands then loads. Prove this with a failing guard or diagnostic before patching.
  Likely real fixes are selftest settings isolation under `TestSandbox`, or strict suite-level
  snapshot/restore/cleanup of the normal app settings path.
- [ ] Suite Full had 16 Commands failures, including Compare options live restore, navigation shell
  stability, About/prompt/fatal-error live UIA, Open Drive Menu focus return, path edit-mode tab
  handoff, persistent menu hover, and `folderView_empty_folder_state`. Do not chase all 16
  independently until the manual Compare-then-Commands contamination is explained.
- [ ] Keep `cmd_pane_navigationView_path_region_keyboard_activation_enters_edit_mode`,
  `cmd_compare_directories_options_scroll_to_lower_cards_stays_stable`, and
  `cmd_app_menuBar_persistent_view_to_plugins_hover_switches_popup` on the watch list until broad
  Commands and Suite Full are green with archived evidence.

### Next steps

- [ ] Inspect the manual reproduction results first:
  `C:\RSPerf\runs\manual-20260707T110215Z-cross-suite-commands\artifacts\selftest\last_run\commands\results.json`
  and the archived copy under
  `Specs\TestRuns\4cb089111a23\Continuation\2026-07-07_manual_compare_then_commands_cross_suite_failure\artifacts\selftest\last_run\commands\results.json`.
- [ ] Add a red/green guard or diagnostic proving whether Commands is loading ambient settings or
  other Compare-owned process/UI state. Do not patch from guesswork.
- [ ] Fix the proven root cause. Acceptable outcomes are deterministic isolation, deterministic
  reset/snapshot/restore, or a deterministic replacement test. Unacceptable outcomes are retry-only,
  non-blocking flaky labels, or anonymous quarantine.
- [ ] Re-run source contracts, rebuild if source changed, rerun manual Compare-then-Commands, rerun
  broad Commands, then rerun Suite Full under `C:\RSPerf`.
- [ ] Re-run `RedSalamanderMonitorEtwLatency` and `ToolsPesterTests` isolated if Suite Full still
  reports those standalone failures after Commands state bleed is fixed.
- [ ] Archive final passing evidence under `Specs/TestRuns/<current-head>\...` or another clearly
  named continuation folder.

### Exact resume commands

```powershell
# Confirm no orphaned app process is holding the interactive desktop.
Get-Process RedSalamander -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding

# Inspect the saved manual reproduction.
$manual='C:\RSPerf\runs\manual-20260707T110215Z-cross-suite-commands\artifacts\selftest\last_run\commands\results.json'
$j=Get-Content $manual -Raw | ConvertFrom-Json
$j | Select-Object passed,failed,skipped,duration_ms | Format-List
$j.cases | Where-Object status -eq 'failed' | Select-Object name,reason,duration_ms | Format-List

# Re-verify the latest source-contract guard and build before changing code.
Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander

# Reproduce the current primary blocker with explicit process waits.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
$env:REDSALAMANDER_TEST_RUN_ID='manual-resume-compare-then-commands'
$exe='Z:\src\RedSalamander\.build\x64\Debug\RedSalamander.exe'
foreach ($arg in @('--compare-selftest','--commands-selftest')) {
  $psi=[System.Diagnostics.ProcessStartInfo]::new()
  $psi.FileName=$exe
  $psi.WorkingDirectory='Z:\src\RedSalamander'
  $psi.UseShellExecute=$false
  [void]$psi.ArgumentList.Add($arg)
  [void]$psi.ArgumentList.Add('--selftest-timeout-multiplier=2')
  $p=[System.Diagnostics.Process]::Start($psi)
  $p.WaitForExit()
  "$arg exit=$($p.ExitCode)"
}

# Final gates after fixes.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

### Worktree at save point

- [ ] Intentional dirty files:
  `Common/DxUi/DxUi.h`,
  `Common/DxUi/DxUi.Menu.cpp`,
  `RedSalamander/RedSalamander.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`,
  `Tests/DxUiTests/DxUiTests.Menu.cpp`,
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`,
  `Tools/Tests/WingetValidation.Tests.ps1`,
  and this plan.
- [ ] Intentional saved evidence:
  `Specs\TestRuns\4cb089111a23\Continuation\2026-07-07_manual_compare_then_commands_cross_suite_failure\`
  (copied from `C:\RSPerf\runs\manual-20260707T110215Z-cross-suite-commands`; reparse/symlink
  work-folder copy warnings occurred, but result JSONs were preserved). This path is ignored by
  `.gitignore`, so it is local continuation evidence unless explicitly exported or force-added.
- [ ] Untracked `Specs/Reviews/ThreeDayDiff-2026-07-06-Findings.md` is unrelated review material;
  do not stage/delete it unless the user explicitly asks.
- [ ] Ignored scratch helper `.build\compare-prefix-filter.txt` exists with the 0..529 broad-order
  prefix from the archived Commands run; use it only if runner case-list support exists, otherwise
  regenerate smaller chunks.
- [ ] Broad Commands and Suite Full are still not accepted. Do not move this plan to Done yet.

## Superseded resume checkpoint - 2026-07-07 11:47 +02:00

Historical detail retained for audit. Do not use this as the active resume checklist; use the
topmost current resume checklist above.

### Status at a glance

- [x] Final root requirement is explicit: all final proof runs use
  `REDSALAMANDER_TEST_ROOT=C:\RSPerf`.
- [x] No `RedSalamander.exe` process was running at this save point.
- [x] Latest source-contract suite passed after the menu-suppression contract was added:
  `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`,
  85 passed / 0 failed.
- [x] Last completed build passed, 0 warnings / 0 errors:
  `.build\logs\msbuild-20260707_111213_387.log`.
- [ ] The latest 11:44 source edits in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp` and
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` have not been rebuilt yet; the 11:14 build is
  older than those files.
- [x] Focused current-pair verification passed before the latest 11:44 menu patch:
  `20260707T093756Z-89216-0a20335a5d68456198f48e806c8aa66a`, 2 passed / 0 failed.
- [x] Compare predecessor window passed before the latest 11:44 menu patch:
  `20260707T093816Z-77996-b31cfa9cdc57411799ba73588b2e80b8`,
  21 passed / 0 failed, disk audit clean.
- [x] Menu predecessor window passed before the latest 11:44 menu patch:
  `20260707T094206Z-440-ed9e0e8f8eb04a329cda37a7ce814cc1`,
  12 passed / 0 failed, disk audit clean.
- [x] Post-build navigation slice passed:
  `20260707T091442Z-76804-b559ace5b0594bfe97c57a5a2c784248`,
  41 passed / 0 failed, disk audit clean.
- [x] Post-build menu predecessor window passed:
  `20260707T091421Z-62572-c8464beb16834871a762b81e603a1769`,
  12 passed / 0 failed, disk audit clean.
- [x] Latest broad Commands run was archived in the repo:
  `Specs\TestRuns\4cb089111a23\Continuation\2026-07-07_menu_suppression_attempt_broad_commands_current_failures\`.
- [ ] Broad Commands is not green. Latest run:
  `20260707T091535Z-90924-7da4844a3fee4022bc4598a7ca5f7367`,
  784 passed / 2 failed / 2 skipped, disk audit clean.
- [ ] Suite Full has not been rerun after the latest broad Commands failure.
- [ ] Do not move this plan to `Specs/Plans/Done/` yet.

### Non-negotiable acceptance gates

- [ ] Do not make flaky tests non-blocking. A flaky or broad-only failure must be identified,
  root-caused, and fixed or replaced with a deterministic equivalent.
- [ ] No anonymous permanent quarantine. Any quarantine must be named, temporary, owned, and paired
  with a repair/removal plan.
- [ ] Final local evidence root is `C:\RSPerf`. Do not use
  `%LOCALAPPDATA%\Temp\RedSalamander-TestSandbox`, `C:\RST`, or any other shorter root as final
  proof.
- [ ] Broad Commands must pass under `REDSALAMANDER_TEST_ROOT=C:\RSPerf`.
- [ ] Suite Full must pass under `REDSALAMANDER_TEST_ROOT=C:\RSPerf`.
- [ ] Before closeout, move this plan to `Specs/Plans/Done/` and merge durable requirements into
  `Specs/Testing/*`, `Tests/README.md`, `README.md`, or `AGENTS.md` as appropriate.

### Done in the latest stabilization slice

- [x] The previous five broad-only failures from the 10:29 checkpoint now pass in broad order:
  `cmd_preferences_dialog_viewers_add_update_live_dx_interaction`,
  `cmd_preferences_dialog_plugins_custom_paths_page_exposes_live_uia_grid_selection`,
  `cmd_compare_directories_options_live_dx_body_interaction`,
  `cmd_app_menuBar_hover_switches_top_level_popup`, and
  `cmd_app_menuBar_submenu_placement_matches_spec`.
- [x] Root cause for the live UIA/value broad failures was identified: the Commands
  `RunUiaActionWithMessagePump(...)` helper timed out after 3 s while DxUi accessibility dispatch
  can legitimately wait up to 5 s, then detached the worker. That allowed late UI mutations to
  contaminate later broad-order cases.
- [x] `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp` now gives UIA dispatch a
  7000 ms scaled budget, joins the worker instead of detaching it, and reports timeout without
  leaving an orphaned mutation path.
- [x] Root cause for the raw-provider fallback was identified: `CreateWindowHostAccessibilityProvider`
  rejects off-window-thread callers, so raw provider value-setting through the worker path was
  inert for same-thread windows.
- [x] `SetWindowHostRawProviderValueByNameWithMessagePump(...)` now uses the direct raw-provider
  setter first when the target HWND belongs to the current UI thread, then falls back to the pumped
  worker helper only when needed.
- [x] TDD source-contract proof was added in `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
  for the deterministic UIA action helper and same-thread raw-provider path. It failed first on
  the old implementation, then passed after the patch.
- [x] A second TDD source-contract proof was added for
  `cmd_app_menuBar_persistent_view_to_plugins_hover_switches_popup`: the test now requires the
  non-mouseleave path to call `ResolveMenuBarHoverSuppressionPoint(...)` instead of hard-coding the
  stale View menu suppression point. It failed first on the old implementation, then passed after
  the patch.
- [x] A newer TDD source-contract proof was added for the same menu case: the non-mouseleave path
  must park the real cursor on the View menu point before calling
  `ResolveMenuBarHoverSuppressionPoint(...)`. It failed first with
  `Expected {9283} to be less than {5925}`, then passed after the patch.
- [x] `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp` now resolves and logs
  the actual menu-bar suppression point/hit index for the persistent View-to-Plugins hover test.
- [x] Current working menu root cause: `TestMainMenuPersistentViewToPluginsHoverSwitchesPopup`
  resolved the suppression point before parking the real cursor on View. In broad order this could
  sample a stale cursor from an earlier case, re-arm suppression to the wrong top-level item, and
  let synthetic Plugins hover briefly win before the stale live-hover point switched selection away.
- [x] `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp` now captures/restores
  the original cursor, parks the non-mouseleave cursor before resolving suppression, and initializes
  `hoverCursorMoved` from whether pre-positioning actually succeeded.
- [x] Source-contract verification passed:
  `Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru`,
  85 passed / 0 failed.
- [x] Build passed after the earlier patch, 0 warnings / 0 errors:
  `.build\logs\msbuild-20260707_111213_387.log`.
- [ ] Build has not been rerun after the 11:44 menu ordering patch.
- [x] Focused current-pair verification passed under `C:\RSPerf` before the 11:44 patch:
  `20260707T093756Z-89216-0a20335a5d68456198f48e806c8aa66a`, 2 passed / 0 failed.
- [x] Compare predecessor window passed under `C:\RSPerf` before the 11:44 patch:
  `20260707T093816Z-77996-b31cfa9cdc57411799ba73588b2e80b8`,
  21 passed / 0 failed, disk audit clean.
- [x] Menu predecessor window passed under `C:\RSPerf` before the 11:44 patch:
  `20260707T094206Z-440-ed9e0e8f8eb04a329cda37a7ce814cc1`,
  12 passed / 0 failed, disk audit clean.
- [x] Navigation predecessor windows passed under `C:\RSPerf`; one early pair of runs was polluted
  by parallel root sharing and must be treated as functional-only, but the clean post-build
  41-case slice passed:
  `20260707T091442Z-76804-b559ace5b0594bfe97c57a5a2c784248`.
- [x] Menu predecessor window passed under `C:\RSPerf` after the earlier suppression-point patch:
  `20260707T091421Z-62572-c8464beb16834871a762b81e603a1769`.
- [x] Latest broad Commands was rerun under `C:\RSPerf`:
  `20260707T091535Z-90924-7da4844a3fee4022bc4598a7ca5f7367`,
  784 passed / 2 failed / 2 skipped, disk audit clean.
- [x] Latest broad failure artifacts were copied for resume under
  `Specs\TestRuns\4cb089111a23\Continuation\2026-07-07_menu_suppression_attempt_broad_commands_current_failures\`.
- [x] Compare investigation so far: the focused pair and the 21-case predecessor window passed, so
  the failure is still broad-order or predecessor-sensitive. The current assertion message is too
  thin because it does not say whether the final snapshot lost focus, lost scroll offset, or both.
- [x] Scratch prefix filter generated at `.build\compare-prefix-filter.txt` with 0..529 broad-order
  case names from the archived run. It is about 32 KB and may be too long for `-CaseFilter`; first
  inspect whether the runner supports a case/order file, otherwise bisect with smaller cumulative
  chunks.
- [x] No `RedSalamander.exe` process was running at this save point.

### Current remaining failures from broad Commands

- [ ] `cmd_compare_directories_options_scroll_to_lower_cards_stays_stable`
  - Broad failure reason:
    `Compare Directories options lower toggle did not stay in view after focus moved into the
    scrolled body.`
  - This is a current broad-only blocker from
    `20260707T091535Z-90924-7da4844a3fee4022bc4598a7ca5f7367`. Root-cause it before changing the
    assertion; likely areas are dialog scroll anchoring, focus-into-body side effects, viewport
    settle timing, or broad-order state carried from earlier Compare cases.
- [ ] `cmd_app_menuBar_persistent_view_to_plugins_hover_switches_popup`
  - Broad failure reason:
    `Delivering a Plugins hover while View is open did not select the Plugins top-level menu
    (mode=mouse-move, point=(485,237), hitResolved=yes, hitIndex=4, selectedAfterWait=6,
    expectedPluginsIndex=4, pointerSwitches=5).`
  - The earlier selftest-side suppression-point patch was not sufficient in broad Commands. The
    current 11:44 working fix is not yet built or broad-verified. If it still fails after rebuild,
    investigate stale queued root-hover messages, live-cursor influence from the real desktop
    cursor, switch-loop exit semantics, or a product-side guarantee that the delivered synthetic
    hover remains authoritative long enough to assert the replacement root.

### Watch list, not current broad blockers

- [ ] `cmd_pane_navigationView_path_region_keyboard_activation_enters_edit_mode`
  - It failed once in the clean 41-case broad-order slice:
    `20260707T090451Z-90144-73d3f1d93d3f460d8edb2cd5c957721c`.
  - It then passed the focused pair, diagnostic/slower 41-case slice, clean rerun, post-build
    41-case slice, and latest broad Commands run. That is not a fix; it is only evidence that the
    failure is intermittent or predecessor-sensitive. Keep it on the watch list until broad Commands
    and Suite Full are green.

### Remaining work, in order

- [ ] Re-read the current broad failure source ranges: the Compare Directories lower-card scroll
  test and the menu persistent View-to-Plugins hover test in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`.
- [ ] Inspect exact helpers before patching:
  `ResolveMenuBarHoverSuppressionPoint`,
  `DebugSetMainMenuBarHoverSuppressionCursorOverride`,
  `DebugGetMainMenuBarSelectedIndex`,
  `WaitForMainMenuBarSelectedIndex`,
  `DebugHitTestMainMenuBarScreenPoint`,
  `WaitForRootSwitchedDxUiContextMenuWindow`,
  `DebugGetContextMenuPopupState`,
  Compare Directories option-dialog scroll/focus helpers,
  `CommandFocusAddressBar`,
  `DebugFocusNavigationViewRegion`,
  `DebugGetNavigationViewSnapshot`,
  `currentEditHostHwnd`,
  `currentEditInputHwnd`,
  `debugEnterEdit`, and composition/edit-entry state.
- [x] Reproduce the two current failures focused together under `C:\RSPerf`; latest focused run
  passed, so the failure is not simple focused-case deterministic.
- [x] Reproduce the Compare predecessor window under `C:\RSPerf`; latest 21-case predecessor run
  passed, so the Compare failure needs a broader prefix, chunked predecessor search, or better
  snapshot diagnostics.
- [x] Reproduce the menu predecessor window under `C:\RSPerf`; latest 12-case predecessor run passed
  before the 11:44 working fix.
- [ ] Rebuild after the 11:44 menu ordering patch.
- [ ] Rerun source contracts, focused current failures, Compare predecessor, and menu predecessor
  after the rebuild.
- [ ] For Compare, inspect whether the harness supports a case-list/order file. If not, bisect with
  smaller cumulative broad-order chunks; do not hand-wave the broad-only failure away.
- [ ] If Compare still will not reproduce outside broad Commands, add diagnostic-rich snapshot
  failure text with source-contract red/green proof, then run broad Commands again to capture
  focus/scroll state.
- [ ] Keep the navigation predecessor window available as a watch-list rerun if broad Commands or
  Suite Full exposes it again.
- [ ] Root-cause both broad-only failures as real setup/teardown, timing, input, or provider defects.
  Do not quarantine them and do not make them non-blocking.
- [ ] Patch deterministic setup/teardown or replace each brittle assertion with an equivalent
  deterministic assertion.
- [ ] Re-run source contracts if touched, rebuild, rerun focused current failures, rerun predecessor
  windows, rerun broad Commands without `-FailFast`, then rerun Suite Full under `C:\RSPerf`.
- [ ] Archive final evidence under `Specs/TestRuns/4cb089111a23/...`.
- [ ] Commit only intentional stabilization files unless the user explicitly asks to include
  unrelated review material.

### Exact resume commands

```powershell
# Confirm no orphaned app process is holding the interactive desktop.
Get-Process RedSalamander -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding

# Re-verify the latest source-contract guard, then rebuild because the latest source edits are
# newer than the last MSBuild log.
Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander

# Inspect the failing source areas.
rg -n "cmd_compare_directories_options_scroll_to_lower_cards_stays_stable|lower toggle|scroll.*body|Compare Directories" RedSalamander\SelfTest\Commands -g "*.cpp" -g "*.h"
rg -n "ResolveMenuBarHoverSuppressionPoint|DebugSetMainMenuBarHoverSuppressionCursorOverride|DebugGetMainMenuBarSelectedIndex|WaitForMainMenuBarSelectedIndex|DebugHitTestMainMenuBarScreenPoint|WaitForRootSwitchedDxUiContextMenuWindow|DebugGetContextMenuPopupState" RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp Common\DxUi\DxUi.Menu.cpp Common\DxUi\DxUi.h

# Inspect whether a long broad-order prefix can be supplied through a file; if not, chunk the
# archived broad-order case list into smaller cumulative filters.
rg -n "selftest-case|selftest-list|case-file|order|CaseFilter|SelfTestShuffleSeed|SelfTestRepeat" Tools\Run-AllTests.ps1 RedSalamander -g "*.ps1" -g "*.cpp" -g "*.h" |
  Select-Object -First 120

# Re-run the two current broad-only failures together.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter 'cmd_compare_directories_options_scroll_to_lower_cards_stays_stable,cmd_app_menuBar_persistent_view_to_plugins_hover_switches_popup' -TimeoutMultiplier 2

# Compare predecessor window from latest broad run indexes 514-534.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter 'cmd_app_shortcuts_restores_reordered_sorted_grid_layout,cmd_app_shortcuts_restores_combined_view_state,cmd_app_shortcuts_column_reorder_survives_search_roundtrip,cmd_app_shortcuts_reordered_resized_copy_follows_visible_columns_after_search_roundtrip,cmd_app_shortcuts_column_reorder_survives_sort_cycles,cmd_app_shortcuts_reordered_resized_columns_survive_sort_cycles,cmd_app_shortcuts_reordered_resized_copy_follows_visible_columns_after_sort_cycles,cmd_app_shortcuts_reordered_resized_columns_survive_sort_cycles_and_search_roundtrip,cmd_app_shortcuts_reordered_resized_copy_follows_visible_columns_after_sort_cycles_and_search_roundtrip,cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics,cmd_compare_directories_window_uses_dxui_menu_bar_and_banner_buttons,cmd_compare_directories_options_long_run_open_close_stays_stable,cmd_compare_directories_options_live_dx_body_interaction,cmd_compare_directories_options_pointer_click_toggles_live_dx_interaction,cmd_compare_directories_options_tab_traversal_live_dx_interaction,cmd_compare_directories_options_scroll_to_lower_cards_stays_stable,cmd_compare_directories_options_enter_and_escape_route_default_cancel,cmd_compare_directories_options_access_keys_focus_expected_controls,cmd_compare_directories_options_theme_cycle_keeps_surface_legible,cmd_compare_directories_non_file_plugin_path_form_selection_and_empty_state,cmd_compare_directories_leave_scope_prompt_defers_out_of_navigation_callback' -TimeoutMultiplier 2

# Menu predecessor window.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter 'cmd_app_menuBar_hidden_model_not_owner_draw,cmd_app_menuBar_popup_host_matches_flyout_contract,cmd_app_menuBar_arrow_switches_top_level_popup,cmd_app_menuBar_hover_switches_top_level_popup,cmd_app_menuBar_persistent_direct_hover_switches_top_level_popup,cmd_app_menuBar_persistent_view_to_plugins_hover_switches_popup,cmd_app_menuBar_persistent_view_to_files_hover_highlight_follows_pointer,cmd_app_menuBar_persistent_mouseleave_clears_hover_without_live_cursor_switch,cmd_app_menuBar_mouse_open_keeps_popup_selection_clear,cmd_app_menuBar_mouse_opened_popup_processes_keyboard_before_mouse_move,cmd_app_menuBar_top_level_mapping_matches_raw_menu,cmd_app_menuBar_submenu_placement_matches_spec' -TimeoutMultiplier 2

# Navigation watch-list window if it reappears.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter 'cmd_app_managePlugins_keeps_navigation_shell_stable,cmd_pane_focusAddressBar_tab_traversal,cmd_pane_navigationView_path_doubleClick_enters_edit_mode,cmd_pane_navigationView_path_doubleClick_ignores_stale_restore_focus,cmd_pane_navigationView_path_doubleClick_after_focusAddressBar_tab_traversal,cmd_pane_navigationView_path_region_keyboard_activation_enters_edit_mode,cmd_pane_navigationView_path_ancestor_click_navigates_to_ancestor,cmd_pane_navigationView_unfocused_pane_click_focuses_target_pane,cmd_pane_navigationView_full_path_popup_edit_route,cmd_pane_navigationView_full_path_popup_ancestor_click_navigates_to_ancestor,cmd_pane_navigationView_region_tab_traversal' -TimeoutMultiplier 2

# Final gates after fixes.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

### Worktree at save point

- [ ] Intentional dirty files:
  `Common/DxUi/DxUi.h`,
  `Common/DxUi/DxUi.Menu.cpp`,
  `RedSalamander/RedSalamander.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`,
  `Tests/DxUiTests/DxUiTests.Menu.cpp`,
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`,
  `Tools/Tests/WingetValidation.Tests.ps1`,
  and this plan.
- [ ] Untracked `Specs/Reviews/ThreeDayDiff-2026-07-06-Findings.md` is unrelated review material;
  do not stage/delete it unless the user explicitly asks.
- [ ] Ignored scratch helper `.build\compare-prefix-filter.txt` exists with the 0..529 broad-order
  prefix from the archived Commands run; use it only if runner case-list support exists, otherwise
  regenerate smaller chunks.
- [ ] Broad Commands and Suite Full are still red/not complete. Do not move this plan to Done yet.

## Superseded resume checkpoint - 2026-07-07 10:29 +02:00

Historical detail retained for audit. Do not use this as the active resume checklist; the five
failures listed here have since been fixed in broad order, and the current status is in the
topmost resume checklist above.

### Non-negotiable acceptance gates

- [ ] Do not make flaky tests non-blocking. A flaky or broad-only failure must be identified,
  root-caused, and fixed or replaced with a deterministic equivalent.
- [ ] No anonymous permanent quarantine. Any quarantine must be named, temporary, owned, and paired
  with a repair/removal plan.
- [ ] Final local evidence root is `C:\RSPerf`. Do not use
  `%LOCALAPPDATA%\Temp\RedSalamander-TestSandbox`, `C:\RST`, or any other shorter root as final
  proof.
- [ ] Broad Commands must pass under `REDSALAMANDER_TEST_ROOT=C:\RSPerf`.
- [ ] Suite Full must pass under `REDSALAMANDER_TEST_ROOT=C:\RSPerf`.
- [ ] Before closeout, move this plan to `Specs/Plans/Done/` and merge durable requirements into
  `Specs/Testing/*`, `Tests/README.md`, `README.md`, or `AGENTS.md` as appropriate.

### Done in this latest stabilization slice

- [x] `cmd_connection_manager_window_protocol_churn_keeps_form_and_uia_stable` was root-caused as
  a test-specific broad timing/open-detection weakness, not a general Connection Manager open
  failure. Neighboring Connection Manager cases opened and passed in broad order, while the failing
  test used a shorter bespoke open path instead of the shared helper.
- [x] `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp` now opens the
  protocol-churn dialog through `OpenConnectionManagerForSelfTest(...)`, adds trace around the open,
  and uses the shared 3000 ms close wait instead of the shorter bespoke wait.
- [x] Build after the protocol-churn patch passed, 0 warnings / 0 errors:
  `.build\logs\msbuild-20260707_094424_444.log`.
- [x] Protocol-churn focused verification passed under `C:\RSPerf`:
  `20260707T074638Z-39392-e647c7d234144badb117dba1d49707dc`, 1 passed / 0 failed.
- [x] Immediate Connection Manager cluster passed under `C:\RSPerf`:
  `20260707T074648Z-83492-18cc39f9f55548848e059def674b9905`, 7 passed / 0 failed.
- [x] Source-order 63-case predecessor window passed under `C:\RSPerf`:
  `20260707T074657Z-30916-f2156706efce4c7f883d13ceacc0a1e7`, 63 passed / 0 failed.
- [x] Post-patch broad Commands run under `C:\RSPerf` proved protocol-churn no longer fails:
  `20260707T074801Z-89868-eaa3885bf7b0426a8b680400fd0b2e54`, 781 passed / 5 failed / 2 skipped,
  disk audit clean. This volatile `C:\RSPerf` run has already been cleaned, so the failure summary
  below is the durable tracked record.
- [x] The five new broad failures passed together when filtered:
  `20260707T080528Z-44160-1dd79aedc90248629e358cdc904323c1`, 5 passed / 0 failed.
- [x] Local predecessor windows for the five new failures passed functionally under `C:\RSPerf`:
  `20260707T082255Z-85876-0bcaf3d11b9d433793ff3505f8193582`,
  `20260707T082328Z-85876-7397cfc5fa0d43a0a73836333f2ec0ce`,
  `20260707T082356Z-85876-966d21a0b4c34262b6088ddd3091e9b6`,
  `20260707T082706Z-85876-7b1ec40e053240319c2f7f528a3b460d`, and
  `20260707T082709Z-85876-f80c6b13dc6f4190bb1e839e43173837`.
- [x] Resume-safe artifact copies were saved under ignored local evidence roots:
  `Specs\TestRuns\4cb089111a23\Continuation\2026-07-07_protocol_churn_broad_failure\` and
  `Specs\TestRuns\4cb089111a23\Continuation\2026-07-07_protocol_churn_fixed_remaining_broad_failures\`.
- [x] No `RedSalamander.exe` process was running at this save point.

### Current remaining failures from broad Commands

- [ ] `cmd_preferences_dialog_viewers_add_update_live_dx_interaction` failed in broad order because
  the visible DX match-value edit did not settle to the edited value after live UIA mutation.
- [ ] `cmd_preferences_dialog_plugins_custom_paths_page_exposes_live_uia_grid_selection` failed in
  broad order because the Preferences Plugins page did not expose its custom-paths DX grid surface
  for UIA selection validation.
- [ ] `cmd_compare_directories_options_live_dx_body_interaction` failed in broad order because the
  `Ignore files` edit mutated from expected `selftest-ignore-pattern` to observed
  `selfntest-ignore-pattern`, indicating input/text mutation contamination or a non-atomic setter.
- [ ] `cmd_app_menuBar_hover_switches_top_level_popup` failed in broad order after `SC_KEYMENU`
  selected the temporary menu bar, but the synthetic click did not open the expected first top-level
  popup.
- [ ] `cmd_app_menuBar_submenu_placement_matches_spec` failed in broad order because a cascading
  submenu closed when the pointer returned to the parent item that owns it.

### Remaining work, in order

- [ ] Inspect shared UIA value/grid helpers and the failing call sites before patching:
  `SetVisibleDescendantValueByNameWithMessagePump`,
  `SetWindowHostRawProviderValueByNameWithMessagePump`, and
  `WaitForAnyVisibleGridSelectionState`.
- [ ] Search for deterministic debug/direct setters for Preferences viewer/plugin fields and
  Compare Directories option fields. Use direct deterministic hooks only when they preserve or
  replace the same behavioral contract; do not silently delete live UIA coverage.
- [ ] Root-cause the five broad-only failures as real setup/teardown, timing, input, or UIA-provider
  defects. Do not quarantine them and do not mark them non-blocking.
- [ ] Patch deterministic setup/teardown or replace each brittle assertion with an equivalent
  deterministic assertion.
- [ ] Rebuild, then rerun the five focused failures, their source-order/predecessor windows, broad
  Commands without `-FailFast`, and finally Suite Full under `C:\RSPerf`.
- [ ] Archive final evidence under `Specs/TestRuns/4cb089111a23/...`.
- [ ] Commit only intentional stabilization files unless the user explicitly asks to include
  unrelated review material.

### Exact resume commands

```powershell
# Confirm no orphaned app process is holding the interactive desktop.
Get-Process RedSalamander -ErrorAction SilentlyContinue |
  Select-Object Id,ProcessName,CPU,StartTime,Responding

# Inspect the shared helpers and candidate deterministic hooks.
rg -n "SetVisibleDescendantValueByNameWithMessagePump|SetWindowHostRawProviderValueByNameWithMessagePump|WaitForAnyVisibleGridSelectionState" RedSalamander\SelfTest\Commands -g "*.cpp" -g "*.h"
rg -n "Debug.*Preferences.*(Value|Text|CustomPaths|Viewers)|Debug.*CompareDirectoriesOptions.*(Value|Text|Ignore)|Set.*ValueForSelfTest|Set.*TextForSelfTest" RedSalamander -g "*.cpp" -g "*.h"

# Re-run the current five broad-only failures together.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter 'cmd_preferences_dialog_viewers_add_update_live_dx_interaction,cmd_preferences_dialog_plugins_custom_paths_page_exposes_live_uia_grid_selection,cmd_compare_directories_options_live_dx_body_interaction,cmd_app_menuBar_hover_switches_top_level_popup,cmd_app_menuBar_submenu_placement_matches_spec' -TimeoutMultiplier 2

# Final gates after fixes.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

### Worktree at save point

- [ ] Intentional dirty files:
  `Common/DxUi/DxUi.h`,
  `Common/DxUi/DxUi.Menu.cpp`,
  `RedSalamander/RedSalamander.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`,
  `Tests/DxUiTests/DxUiTests.Menu.cpp`,
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`,
  `Tools/Tests/WingetValidation.Tests.ps1`,
  and this plan.
- [ ] Untracked `Specs/Reviews/ThreeDayDiff-2026-07-06-Findings.md` is unrelated review material;
  do not stage/delete it unless the user explicitly asks.
- [ ] Earlier accidental unfiltered Commands run `20260707T080709Z-18324-db2b73207e4941e09d4f180eff4d639c`
  was killed after it was recognized as an unintended full run caused by a stale/empty filter path.
  Do not treat it as failure evidence.

## Superseded resume checkpoint - 2026-07-07 09:37 +02:00

Historical detail retained for audit. Do not use this as the active resume checklist; protocol-churn
has since been fixed and the current status is in the topmost resume checklist above.

### Non-negotiable acceptance gates

- [ ] Do not make flaky tests non-blocking. A flaky or broad-only test failure must be
  identified, root-caused, and fixed or replaced with a deterministic equivalent.
- [ ] No anonymous permanent quarantine. Any quarantine must be named, temporary, owned, and
  paired with a repair/removal plan.
- [ ] Final local evidence root is `C:\RSPerf`. Do not use
  `%LOCALAPPDATA%\Temp\RedSalamander-TestSandbox`, `C:\RST`, or any other shorter root as final
  proof.
- [ ] Broad Commands must pass under `REDSALAMANDER_TEST_ROOT=C:\RSPerf`.
- [ ] Suite Full must pass under `REDSALAMANDER_TEST_ROOT=C:\RSPerf`.
- [ ] Before closeout, move this plan to `Specs/Plans/Done/` and merge durable requirements into
  `Specs/Testing/*`, `Tests/README.md`, `README.md`, or `AGENTS.md` as appropriate.

### Done in the current dirty slice

- [x] Fatal-error dialog selftests were restored by making
  `DebugShowFatalErrorDialog(...)` bypass the generic selftest non-modal suppression path and
  directly show `FatalErrorDialogWindow`.
- [x] Several path-budget-sensitive Commands fixtures now use shorter run-local names so
  `C:\RSPerf` runs stay within classic Win32 path limits.
- [x] Menu-bar input isolation and stale queued hover-message handling were fixed:
  real-cursor residue is neutralized, posted hover root switches carry index + sequence, and stale
  sequence messages are ignored.
- [x] DxUi menu stale-message coverage was added in `Tests/DxUiTests/DxUiTests.Menu.cpp`.
- [x] Pester compatibility/path-budget fixes were added in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` and
  `Tools/Tests/WingetValidation.Tests.ps1`.
- [x] Connection Manager modeless-connect selftests now acquire a Commands `SelfTest::TestSandbox`,
  force the tested pane to `builtin/file-system`, suppress real connect navigation while asserting
  debug-posted navigation, and restore pane/settings state deterministically.
- [x] Preferences Keyboard broad-only failure was root-caused to non-deterministic shortcut state:
  `g_settings.shortcuts.emplace(ShortcutDefaults::CreateDefaultShortcuts())` was a no-op when
  broad-order state had already populated the optional. The affected Preferences Keyboard tests now
  assign `g_settings.shortcuts = ShortcutDefaults::CreateDefaultShortcuts()` explicitly in
  `Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp` and
  `Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`.
- [x] Temporary diagnostics for the Keyboard theme-cycle baseline wait were added so future failures
  report actual theme flags, category, row counts, visible grid state, capture/search/focus state,
  child-window counts, resize failures, list render/resize counts, page render count, and title.

### Verification captured and reusable

- [x] Build passed, 0 warnings / 0 errors:
  `.build\logs\msbuild-20260707_091513_784.log`.
- [x] DxUi tests passed after stale hover-message coverage:
  `.\.build\x64\Debug\DxUiTests.exe`.
- [x] Focused historical Commands failures passed under `C:\RSPerf`:
  `20260706T183750Z-77120-bd82e087232449ada2a64c33d3b22e44`, 13 passed / 0 failed.
- [x] Tools Pester passed under `C:\RSPerf`:
  `Invoke-Pester -Script .\Tools\Tests -ExcludeTag RequiresBuildToolchain -PassThru`,
  193 passed / 0 failed.
- [x] Menu-bar source-order/focused evidence passed after the sequence patch:
  `20260707T054029Z-73180-a59c1ffa47a54217af20081d5d9612c7` and
  `20260706T200741Z-18348-b4a2c223ceb04d7ea034831fddceb6f5`.
- [x] Connection Manager left/right focused and source-order cluster passed after deterministic
  cleanup:
  `20260707T062620Z-72648-5e5d5a9e6ab1419993028a44a5c7088f` and
  `20260707T062629Z-74824-89e230866d8c4cc7bbc31b2fe1acdbab`.
- [x] Broad Commands no longer hangs on the previous modeless-connect failure. Archived run
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-07_084410` completed with
  785 passed / 1 failed / 2 skipped; the single failure was the now-fixed Preferences Keyboard
  theme-cycle case.
- [x] After the shortcut defaulting fix, focused Keyboard theme-cycle passed under `C:\RSPerf`:
  `20260707T071733Z-22060-c26b8b41d06e4a5387197db7ad769b79`,
  1 passed / 0 failed / 0 skipped.
- [x] After the shortcut defaulting fix, the 31-case Preferences window passed under `C:\RSPerf`:
  `20260707T071753Z-77944-1426eb0066714c279bf298e013a6aa8f`,
  31 passed / 0 failed / 0 skipped.
- [x] After the shortcut defaulting fix, broad Commands with `-FailFast` passed under `C:\RSPerf`:
  `20260707T065642Z-9760-1a974568920c495aba20f6bf0bcbc25d`,
  786 passed / 0 failed / 2 skipped, disk audit clean.

### Broad no-failfast run after shortcut fix

- [x] Broad Commands without `-FailFast` completed under `REDSALAMANDER_TEST_ROOT=C:\RSPerf`:
  run `20260707T071835Z-68992-c70f505525594b40b3ca77cd0f4f9c47`,
  785 passed / 1 failed / 2 skipped, exit code 1, duration 17m 44s.
- [x] Disk audit was clean:
  `test_sandbox_audit.is_clean=true`, `issue_count=0`, `test_root=C:\RSPerf`.
- [x] Result JSON:
  `C:\RSPerf\runs\20260707T071835Z-68992-c70f505525594b40b3ca77cd0f4f9c47\artifacts\selftest\last_run\run-all-tests-results.json`.
- [x] Logs:
  `.build\logs\commands-broad-rsperf-20260707_091834.out.log`,
  `.build\logs\commands-broad-rsperf-20260707_091834.err.log`,
  `C:\RSPerf\runs\20260707T071835Z-68992-c70f505525594b40b3ca77cd0f4f9c47\artifacts\selftest\last_run\commands\trace.txt`.
- [x] Resume-safe artifact copy exists on disk. This path is ignored by `.gitignore`, so the tracked
  checklist above carries the durable summary:
  `Specs\TestRuns\4cb089111a23\Continuation\2026-07-07_protocol_churn_broad_failure\`.
- [x] The now-fixed broad-only Keyboard failure passed in this run:
  `cmd_preferences_dialog_keyboard_theme_cycle_keeps_surface_legible`.
- [ ] Remaining failure:
  `cmd_connection_manager_window_protocol_churn_keeps_form_and_uia_stable` failed with
  `Connection Manager window did not open for the focused protocol-churn test.`

### Remaining work, in order

- [ ] Root-cause `cmd_connection_manager_window_protocol_churn_keeps_form_and_uia_stable`. It is a
  broad-order failure, not a general Connection Manager open failure: neighboring Connection Manager
  cases before and after it opened and passed in broad order.
- [ ] Reproduce with the focused protocol-churn case under `C:\RSPerf`.
- [ ] Reproduce with the source-order cluster covering the four preceding Connection Manager cases,
  protocol churn, and the next two Connection Manager cases.
- [ ] Patch the deterministic setup/teardown or replace the test with an equivalent deterministic
  assertion. Do not quarantine it and do not make it non-blocking.
- [ ] Rebuild, then rerun focused protocol churn, the source-order cluster, broad Commands without
  `-FailFast`, and finally Suite Full under `C:\RSPerf`.
- [ ] Archive final evidence under `Specs/TestRuns/4cb089111a23/...`.
- [ ] Commit only the intentional stabilization files unless the user explicitly asks to include
  unrelated review material.

### Exact resume commands

```powershell
# Re-open the completed broad run result.
Get-Content 'C:\RSPerf\runs\20260707T071835Z-68992-c70f505525594b40b3ca77cd0f4f9c47\artifacts\selftest\last_run\commands\trace.txt' -Tail 120
Get-Content 'C:\RSPerf\runs\20260707T071835Z-68992-c70f505525594b40b3ca77cd0f4f9c47\artifacts\selftest\last_run\run-all-tests-results.json' -Raw

# Locate the next failing case in source.
rg -n "protocol_churn|ProtocolChurn|cmd_connection_manager_window_protocol_churn" RedSalamander\SelfTest\Commands

# Focused repro.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter 'cmd_connection_manager_window_protocol_churn_keeps_form_and_uia_stable' -TimeoutMultiplier 2

# Source-order cluster.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter 'cmd_connection_manager_window_uses_dxui_command_buttons,cmd_connection_manager_window_uses_dxui_form_inputs,cmd_connection_manager_window_layout_keeps_cards_and_fields_clean,cmd_connection_manager_window_uses_dxui_form_action_buttons,cmd_connection_manager_window_protocol_churn_keeps_form_and_uia_stable,cmd_connection_manager_window_mtp_picker_populates_profile,cmd_connection_manager_window_cloud_profiles_persist_oauth2' -TimeoutMultiplier 2

# Final gates after the fix.
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

### Worktree at save point

- [ ] Intentional dirty files:
  `Common/DxUi/DxUi.h`,
  `Common/DxUi/DxUi.Menu.cpp`,
  `RedSalamander/RedSalamander.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`,
  `Tests/DxUiTests/DxUiTests.Menu.cpp`,
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`,
  `Tools/Tests/WingetValidation.Tests.ps1`,
  and this plan.
- [ ] Ignored on-disk evidence bundle:
  `Specs/TestRuns/4cb089111a23/Continuation/2026-07-07_protocol_churn_broad_failure/`
  contains the run JSON, commands trace, stdout/stderr logs, and a short README.
- [ ] Untracked `Specs/Reviews/ThreeDayDiff-2026-07-06-Findings.md` is unrelated review material;
  do not stage/delete it unless the user explicitly asks.

## Previous continuation checklist - 2026-07-07 (superseded by resume checklist above)

Historical detail retained for audit. Use the topmost current resume checklist first when
continuing. The goal is still to implement this plan until it can move to `Specs/Plans/Done/`; do
not narrow success to the green slices below.

### Acceptance gate - still open

- [ ] Do not make flaky tests non-blocking. Every flaky/broad-only failure must be identified,
  root-caused, and fixed or replaced with a deterministic equivalent.
- [ ] No anonymous permanent quarantine. Any quarantine must be named, temporary, owned, and paired
  with a repair/removal plan.
- [ ] Final local evidence root is `C:\RSPerf`. Do not use `%LOCALAPPDATA%\Temp\RedSalamander-TestSandbox`,
  `C:\RST`, or another shorter root as the final proof root.
- [ ] Broad Commands must pass under `REDSALAMANDER_TEST_ROOT=C:\RSPerf`.
- [ ] Suite Full must pass under `REDSALAMANDER_TEST_ROOT=C:\RSPerf`.
- [ ] Before closeout, move this plan to `Specs/Plans/Done/` and merge durable requirements into
  `Specs/Testing/*`, `Tests/README.md`, `README.md`, or `AGENTS.md` as appropriate.

### Done in the current dirty slice

- [x] `DebugShowFatalErrorDialog(...)` in `RedSalamander/RedSalamander.cpp` now bypasses the generic
  selftest non-modal suppression path and directly shows `FatalErrorDialogWindow`, restoring the
  intentional fatal-error dialog selftests.
- [x] Path-budget-sensitive Commands fixtures in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp` now use shorter run-local
  names for navigation full-path popups, thumbnail scroll, column audit, visible column widths, and
  scroll/render stress.
- [x] `cmd_pane_navigationView_path_region_keyboard_activation_enters_edit_mode` now targets the
  native edit input HWND when available and verifies focus belongs to the active native edit surface
  before IME/key injection.
- [x] Menu-bar hover/root-switch selftests now restore directed real-cursor warps and suppress stale
  live-cursor residue without suppressing the intended synthetic hover target. Root cause of the
  reproduced failure was input isolation: `cmd_app_menuBar_persistent_view_to_plugins_hover_switches_popup`
  parked the real cursor on `View` and the next temporary hover test inherited that live cursor, so
  the popup switched to visual index 5 instead of the intended synthetic index 1.
- [x] DxUi menu-bar root-hover switching now carries the posted hover index and a monotonic sequence
  through `ContextMenuSessionCallbacks::switchRootFromMenuBarHover(...)`. The modal menu loop no
  longer consumes a mutable "latest pending hover" slot that can be changed by a newer hover before
  an older queued message is processed.
- [x] `MainWindow::PostPendingMenuBarHoverRootSwitch(...)` now posts the hover sequence in `LPARAM`,
  and the callback rejects stale queued messages unless both index and sequence match the pending
  root switch.
- [x] `Tests/DxUiTests/DxUiTests.Menu.cpp` now has explicit stale-message coverage: a stale
  `(hoverIndex=0, sequence=1)` message is ignored before the valid `(hoverIndex=1, sequence=2)`
  message switches roots.
- [x] `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` now has a Windows PowerShell-compatible
  relative-path helper instead of requiring `[System.IO.Path]::GetRelativePath`.
- [x] `Tools/Tests/WingetValidation.Tests.ps1` VC runtime fixtures now use shorter case names and
  `VS\VC\Redist\...` synthetic roots so `C:\RSPerf` runs stay inside classic Win32 path limits.
- [x] `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp` now makes the
  connection-manager modeless-connect selftests deterministic: each case acquires a Commands
  `SelfTest::TestSandbox`, forces the tested pane to `builtin/file-system` plus that sandbox root
  before opening Connection Manager, suppresses real connect navigation while asserting the posted
  debug navigation, and leaves cleanup on deterministic local pane state while restoring
  `g_settings.connections` and the active pane.
- [x] The previous broad-only connection-manager hang has a continuation evidence bundle under
  `Specs/TestRuns/4cb089111a23/Continuation/2026-07-07_connection_manager_modeless_hang/` with
  the source-order trace/result files and root-cause notes needed to audit the fix later.

### Verification already captured

- [x] Build passed, 0 warnings / 0 errors:
  `.build\logs\msbuild-20260706_220451_531.log`.
- [x] Focused historical Commands failures passed under
  `REDSALAMANDER_TEST_ROOT=C:\RSPerf`:
  run `20260706T183750Z-77120-bd82e087232449ada2a64c33d3b22e44`, 13 passed / 0 failed, disk audit 0.
- [x] Tools Pester passed under `REDSALAMANDER_TEST_ROOT=C:\RSPerf`:
  `Invoke-Pester -Script .\Tools\Tests -ExcludeTag RequiresBuildToolchain -PassThru`,
  193 passed / 0 failed.
- [x] Reproduced menu-bar order/input leak before the fix:
  run `20260706T194814Z-97120-cf119896b31a48749472050c72b34b09`,
  failed on repeat 2 with `selectedBefore=5`, `selectedAfter=5`, `expectedSecond=1`,
  `pointerSwitches=3`.
- [x] Menu-bar cluster passed after the fix:
  run `20260706T200655Z-20652-effa7d14296e476e8280ecafb20baf3c`,
  60 passed / 0 failed / 0 skipped, disk audit 0.
- [x] Focused menu-bar hover repeat passed after the fix:
  run `20260706T200741Z-18348-b4a2c223ceb04d7ea034831fddceb6f5`,
  25 passed / 0 failed / 0 skipped, disk audit 0.

### Verification captured on 2026-07-07

- [x] Build passed, 0 warnings / 0 errors:
  `.build\logs\msbuild-20260707_073644_814.log`.
- [x] `.\.build\x64\Debug\DxUiTests.exe` passed all DxUi tests, including the updated stale
  menu-root message coverage.
- [x] Source-order menu-bar cluster passed before the sequence patch, proving the earlier live-cursor
  hardening was not enough to expose the stale queued-message defect in focused order:
  run `20260707T053402Z-70276-43c422f1365a49e1aa74d7ac4017e4d6`,
  150 passed / 0 failed / 0 skipped, disk audit 0.
- [x] Same source-order menu-bar cluster passed after the sequence patch:
  run `20260707T054029Z-73180-a59c1ffa47a54217af20081d5d9612c7`,
  150 passed / 0 failed / 0 skipped, disk audit 0.
- [x] Post-patch broad Commands reached the connection-manager area without reproducing the repaired
  menu-bar failure. The run id was `20260707T054154Z-68236-0bb5ee889daa489aa45fa606101ba77f`;
  the run artifact was later removed by stale-run cleanup after the hung child process was killed,
  so keep these captured facts in this checklist:
  `cmd_plugin_configuration_dialog_theme_cycle_keeps_surface_legible` passed, the active case at
  timeout was `cmd_connection_manager_window_modeless_connect_posts_left_navigation`, `last_run\trace.txt`
  contained `ConnectionManager modeless-connect-left: complete`, and `commands\trace.txt` never
  logged `Case returned`.
- [x] Focused connection-manager left case passed under `C:\RSPerf`:
  run `20260707T061343Z-51340-496a77d0d4d4425bbc4eebfaae856643`,
  1 passed / 0 failed / 0 skipped, disk audit 0.
- [x] Source-order cluster from plugin configuration through connection-manager left/right passed:
  run `20260707T061352Z-69052-91122767bc944bbcb0432d087a8bb0cf`,
  19 passed / 0 failed / 0 skipped, disk audit 0.
- [x] Build passed after the deterministic connection-manager cleanup patch, 0 warnings / 0 errors:
  `.build\logs\msbuild-20260707_082356_956.log`.
- [x] Focused connection-manager left/right passed under `C:\RSPerf`:
  run `20260707T062620Z-72648-5e5d5a9e6ab1419993028a44a5c7088f`,
  2 passed / 0 failed / 0 skipped, disk audit 0.
- [x] Source-order cluster from plugin configuration through connection-manager left/right passed
  after the deterministic cleanup patch:
  run `20260707T062629Z-74824-89e230866d8c4cc7bbc31b2fe1acdbab`,
  19 passed / 0 failed / 0 skipped, disk audit 0.
- [x] Broad Commands no longer hangs under `REDSALAMANDER_TEST_ROOT=C:\RSPerf`. It completed in
  17m 8s with run id `20260707T062703Z-30480-aa30c6cf84b348e2953c8d6179f13d9e` and is archived at
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-07_084410`.
  Result: 785 passed / 1 failed / 2 skipped, disk audit 0. The remaining failure is
  `cmd_preferences_dialog_keyboard_theme_cycle_keeps_surface_legible`.
- [x] Focused `cmd_preferences_dialog_keyboard_theme_cycle_keeps_surface_legible` passed under
  `C:\RSPerf`: run `20260707T064506Z-8916-8a33339813ae4702bd23ef4f5050e81c`,
  1 passed / 0 failed / 0 skipped, disk audit 0.
- [x] The 13-case local predecessor window ending at
  `cmd_preferences_dialog_keyboard_theme_cycle_keeps_surface_legible` passed under `C:\RSPerf`:
  run `20260707T064601Z-71308-6a15c1210610499c96cf1f1f85776727`,
  13 passed / 0 failed / 0 skipped, disk audit 0.
- [x] The 31-case Preferences window from Themes through General, Panes, Hot Paths, and Keyboard
  theme-cycle passed under `C:\RSPerf`:
  run `20260707T064634Z-66964-b40f4e97e3e246ffaa3ee8100d5e287d`,
  31 passed / 0 failed / 0 skipped, disk audit 0.

### Current blocker - fix next

- [ ] Broad Commands completes but is still red on
  `cmd_preferences_dialog_keyboard_theme_cycle_keeps_surface_legible` with:
  `Preferences Keyboard page did not settle to the baseline dark theme-cycle state.`
- [ ] Failure evidence lives in
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-07_084410\commands_results.json`;
  the failing case is index 325 in that broad run.
- [ ] The failure is broad-order or broad-state dependent, not a simple focused failure: the focused
  Keyboard case, the 13-case immediate predecessor window, and the 31-case Preferences window all
  pass under `C:\RSPerf`.
- [ ] Next root-cause step: add a diagnostic formatter around the Keyboard theme-cycle baseline wait
  in `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp` so the
  failure reports actual category, theme flags, row counts, visible rows/columns/cells, capture
  state, search text, focus target, page-child count, and page resize failures.
- [ ] Inspect shortcut/settings mutations before patching behavior:
  `rg -n "g_settings\.shortcuts|shortcuts\s*=|shortcuts\.emplace|CreateDefaultShortcuts" RedSalamander\SelfTest\Commands -g "*.cpp"`.
  Current strong suspect is that the Keyboard theme-cycle case calls
  `g_settings.shortcuts.emplace(ShortcutDefaults::CreateDefaultShortcuts());`, which only initializes
  defaults when the optional is empty. In broad order, earlier shortcut/preference tests may leave a
  non-default shortcut model, so the case may not start from deterministic keyboard rows.
- [ ] After adding diagnostics or a confirmed fix, run this verification ladder under
  `REDSALAMANDER_TEST_ROOT=C:\RSPerf`: focused Keyboard theme-cycle, the 13-case predecessor window,
  the 31-case Preferences window, broad Commands with `-FailFast`, broad Commands without
  `-FailFast`, then Suite Full.

### Current worktree and handoff notes

- [ ] Dirty files intentionally related to this slice:
  `Common/DxUi/DxUi.h`,
  `Common/DxUi/DxUi.Menu.cpp`,
  `RedSalamander/RedSalamander.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.Connections.cpp`,
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`,
  `Tests/DxUiTests/DxUiTests.Menu.cpp`,
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`,
  `Tools/Tests/WingetValidation.Tests.ps1`,
  and this plan.
- [ ] Untracked `Specs/Reviews/ThreeDayDiff-2026-07-06-Findings.md` is unrelated review material;
  do not stage or delete it unless the user explicitly asks.
- [ ] Use `REDSALAMANDER_TEST_ROOT=C:\RSPerf` for final local path-sensitive evidence. Do not use
  `%LOCALAPPDATA%\Temp\RedSalamander-TestSandbox` or `C:\RST` as final evidence roots.
- [ ] Do not stage/commit the unrelated untracked review file unless explicitly requested. If the user
  asks to commit all plan/test-suite stabilization work, include the intentional dirty files above
  and leave unrelated review material out unless they confirm otherwise.
- [ ] Next best action: root-cause the broad-only Preferences Keyboard theme-cycle failure with
  diagnostics first, then patch the deterministic state setup or theme/page settle condition proven
  by the diagnostic output.
- [ ] Only update this status toward Done after broad Commands and Suite Full evidence are green under
  `C:\RSPerf`, with no anonymous permanent quarantine and no flaky-nonblocking policy.

Method: 7-axis parallel code investigation + adversarial verification (39 candidate root
causes → 5 confirmed at high/critical) + completeness critique (10 gaps closed). All
file:line citations were read, not inferred.

## Implementation progress

### 2026-07-06 — Phase 0 runner sandbox bridge

Completed the runner-side first slice of §8/R1-R2:

- `Tools/TestRunPlan.ps1` now defines the canonical `REDSALAMANDER_TEST_ROOT` sandbox base,
  defaulting to `<repoRoot>\.build\TestSandbox`.
- `Run-AllTests.ps1` now creates a per-run `runId`, sets `REDSALAMANDER_TEST_ROOT`, bridges the
  legacy native `REDSALAMANDER_SELFTEST_ROOT\last_run` writer under
  `REDSALAMANDER_TEST_ROOT\runs\<runId>\artifacts\selftest\last_run`, and records `test_root` +
  `run_id` in `run-all-tests-results.json`.
- `Tools\Tests\RunAllTestsPlan.Tests.ps1` covers default sandbox selection, explicit
  `REDSALAMANDER_TEST_ROOT`, conflicting legacy/unified roots, and aggregate artifact metadata.
- Specs updated: `Specs/Testing/Testing_SelfTests.md`, `Specs/Testing/Testing_TestCoverage.md`,
  and `Tests/README.md`.

Verification:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru"
```

Result: 10 passed / 0 failed.

Later Phase 0 checkpoints close retry/classification/quarantine, the legacy cleanup reaper, and direct
native consumption of `REDSALAMANDER_TEST_ROOT\runs\<runId>\...`.
Still open from this area: executing the full-run disk proof and fixing any non-FileOps scratch
migrations it exposes.

### 2026-07-06 — Phase 0 native selftest root consumption

Completed the direct native consumption slice from §8/R1-R2:

- `SelfTestCommon.cpp` now resolves normal self-test artifacts from
  `REDSALAMANDER_TEST_ROOT` + `REDSALAMANDER_TEST_RUN_ID` as
  `runs\<runId>\artifacts\selftest`, then appends the existing `last_run` suite layout.
- `REDSALAMANDER_SELFTEST_ROOT` remains an explicit compatibility override for deliberate
  launches, including the reviewed quarantine repair lane, but `Run-AllTests.ps1` now ignores it
  while creating the normal run context and clears inherited values before child-suite execution.
- `Run-AllTests.ps1` now sets `REDSALAMANDER_TEST_RUN_ID` for every normal invocation and no longer
  sets the legacy per-run bridge root.
- `TestHarnessSourceContracts.Tests.ps1` guards the native env names, layout constants, direct
  resolver, runner env forwarding, and legacy-root clearing.
- Specs/docs updated: `Specs/Testing/Testing_SelfTests.md`,
  `Specs/Testing/Testing_TestCoverage.md`, `README.md`, and `Tests/README.md`.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
$env:REDSALAMANDER_SELFTEST_ROOT='Z:\src\RedSalamander\.build\BogusLegacySelfTestRoot'; .\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter showIdentical -TimeoutMultiplier 0.1 -SkipLegacySandboxCleanup
```

Result: `TestHarnessSourceContracts` 62 / 0 / 0, `RunAllTestsPlan` 29 / 0 / 0,
`TestInventory` 5 / 0 / 0, PowerShell parser checks passed, Debug build log
`.build\logs\msbuild-20260706_125917_741.log` had 0 warnings / 0 errors, and the focused Compare
runtime proof passed with 1 passed / 0 failed while `REDSALAMANDER_SELFTEST_ROOT` was intentionally
set to an unrelated bogus path. The runner still selected
`Z:\src\RedSalamander\.build\TestSandbox` and wrote under
`.build\TestSandbox\runs\20260706T110346Z-87828-dab956516dda4f7ca2c88af9da3d1580\artifacts\selftest\last_run`.
The runner summary recorded `test_root=Z:\src\RedSalamander\.build\TestSandbox`,
`run_id=20260706T110346Z-87828-dab956516dda4f7ca2c88af9da3d1580`, and
`compare\results.json` recorded `showIdentical` as `passed`.

### 2026-07-06 — Phase 1 Compare dummy filesystem TestSandbox scratch

Completed another §8/R3-R6 scratch-root migration:

- The Compare dummy filesystem cases `dummy_content`,
  `normalized_name_collision_preserves_same_side_entries`, `deep_tree`, and `invalidate` no longer
  allocate under hard-coded `Y:\CompareSelfTest_*`, `Z:\CompareSelfTest_*`, or
  `W:\CompareSelfTest_*` roots.
- Each case now acquires a native Compare scratch root through
  `SelfTest::AcquireTestSandbox(SelfTest::SelfTestSuite::CompareDirectories, <case>)` and fails the
  case cleanly if sandbox acquisition is unavailable.
- `TestHarnessSourceContracts.Tests.ps1` guards the four acquisitions and rejects the old
  `CompareSelfTest_`/drive-root pattern.
- Durable contract updates: `Specs/Testing/Testing_SelfTests.md`,
  `Specs/Testing/Testing_TestCoverage.md`, and `Tests/README.md`.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter 'dummy_content,normalized_name_collision_preserves_same_side_entries,deep_tree,invalidate' -TimeoutMultiplier 2 -SkipLegacySandboxCleanup
rg -n 'CompareSelfTest_|std::filesystem::path\(L"[A-Z]:\\\\"\)|RedSalamanderCrossVolumeSelfTest_|std::filesystem::temp_directory_path|GetTempPathW|GetTempFileNameW' RedSalamander\SelfTest Tests Tools\Tests -g '*.cpp' -g '*.h' -g '*.hpp' -g '*.ps1'
```

Result: the new source-contract guard first failed at 81 / 1 / 0 against the still-hard-coded
`dummy_content` root, then passed at 82 / 0 / 0 after migration. `TestInventory` first failed because
the Tools Pester count moved from 189 to 190, then passed at 5 / 0 / 0 after count/doc updates.
Debug `RedSalamander` build log `.build\logs\msbuild-20260706_162932_943.log` had 0 warnings /
0 errors. Focused Compare proof passed 4 / 0 / 0 with run id
`20260706T143152Z-74460-563d6bd4c17b4a509b9d606ccb14c2da`; the shared trace records the four
scratch roots under
`.build\TestSandbox\runs\20260706T143152Z-74460-563d6bd4c17b4a509b9d606ccb14c2da\scratch\compare\<case>`.
The broad source scan now reports only guard/cleanup literals in `Tools\Tests`.

Note: the first focused runner attempt without `-SkipLegacySandboxCleanup` did not reach the cases
because a stale historical dump under `%LOCALAPPDATA%\RedSalamander\SelfTest\last_run\commands\hang_dumps`
returned access denied during legacy cleanup. The next checkpoint hardens that cleanup path.

### 2026-07-06 — Phase 1 legacy cleanup locked-target hardening

Completed the cleanup robustness follow-up exposed by the Compare dummy filesystem proof:

- `Tools\Clean-TestSandbox.ps1` now treats each legacy cleanup target independently.
- Successful removals are returned with `Status=Removed`; skipped `ShouldProcess` targets report
  `Status=Skipped`; locked or access-denied targets warn and report `Status=Failed` plus the removal
  error.
- A stale locked historical artifact can no longer abort `Run-AllTests.ps1` before build/test
  execution starts. The failure remains visible in the warning stream and result rows.
- `RunAllTestsPlan.Tests.ps1` now exercises a locked legacy `%LOCALAPPDATA%\RedSalamander\SelfTest`
  tree inside a repo-local proof root and verifies the script reports failure without throwing.
- `TestHarnessSourceContracts.Tests.ps1` guards the `-ErrorAction Stop`/`catch`/`Write-Warning`/
  failed-status behavior.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter 'dummy_content,normalized_name_collision_preserves_same_side_entries,deep_tree,invalidate' -TimeoutMultiplier 2
```

Result: `RunAllTestsPlan` 30 / 0 / 0 and `TestHarnessSourceContracts` 82 / 0 / 0. The focused
runner proof with legacy cleanup enabled emitted a warning for the pre-existing locked dump
`RedSalamander-commands-find-setup-hang-p37204.dmp`, then continued and passed Compare 4 / 0 / 0
with run id `20260706T143655Z-85708-3594b0f9798e45989d6c1bdbffb097cf`.

### 2026-07-06 — Phase 1 broad raw-root source-guard lint

Completed the §8/R6 source-guard slice for raw temp/profile/legacy test-root acquisition:

- `TestHarnessSourceContracts.Tests.ps1` now enumerates primary gated test sources under
  `RedSalamander\SelfTest`, `Tests`, and `Tools\Tests`, excluding meta source-contract literals.
- The guard fails on `std::filesystem::temp_directory_path(...)`, `GetTempPath*`,
  `GetTempFileName*`, PowerShell `[System.IO.Path]::GetTempPath()`, direct
  `GetEnvironmentVariableW(L"LOCALAPPDATA", ...)`, direct C getenv use of `LOCALAPPDATA`/`TEMP`/
  `TMP`, and the known legacy root families `CompareSelfTest_*`,
  `RedSalamanderCrossVolumeSelfTest_*`, `C:\BatchRename*SelfTest`,
  `Specs\TestRuns\DxUiGallery`, and `Specs\TestRuns\local_scratch`.
- The test includes a synthetic violation block so the patterns prove they would catch the banned
  acquisition forms before scanning the actual tree.
- The cleanup-plan Pester test remains allowed to mention legacy cross-volume cleanup targets; those
  strings are cleanup evidence, not new allocation.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
```

RED proof: the first run failed because the draft guard used Pester 5 `Should -Be` syntax, then
because the draft legacy-name regex caught harmless BatchRename function names and cleanup-plan
assertion strings. After tightening the assertion and path-acquisition patterns, GREEN
`TestHarnessSourceContracts` passed 83 / 0 / 0.

### 2026-07-06 — Phase 1 TestSandbox disk-audit evidence plumbing

Completed the first half of the full-run disk-audit proof: structured audit collection and
runner-summary preservation.

- `Tools\TestRunPlan.ps1` now exposes `Get-RSTestSandboxDiskAudit(...)`, returning schema
  `red-salamander.test-sandbox-disk-audit.v1`.
- The audit reports unexpected children directly under the TestSandbox root, stale sibling run
  directories under `runs\`, and resolved legacy cleanup targets from the existing cleanup plan:
  `%LOCALAPPDATA%\RedSalamander\SelfTest`, historical `%TEMP%` standalone/perf roots, and
  fixed-drive `RedSalamanderCrossVolumeSelfTest_*` roots.
- The current runner-owned `runId` is allowed, so the audit can be collected before the current
  run's artifacts are archived or uploaded.
- `Run-AllTests.ps1` now stores the audit object in `run-all-tests-results.json` as
  `test_sandbox_audit` and prints the issue count in the artifacts section.
- This is non-blocking evidence for now. The remaining closeout work is to execute the full-run
  disk proof, fix any issues it reports, and only then decide whether the audit becomes a hard gate.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter showIdentical -TimeoutMultiplier 0.1 -SkipLegacySandboxCleanup
```

Result: `RunAllTestsPlan` passed 31 / 0 / 0. The new Pester case proves clean current-run state,
stale sibling run detection, unexpected TestSandbox child detection, and legacy
selftest/temp/cross-volume root detection. The focused runner proof passed Compare 1 / 0 / 0 with
run id `20260706T145229Z-74348-9f42fae656424187a176718f128249f9`; its
`run-all-tests-results.json` preserved `test_sandbox_audit.schema=red-salamander.test-sandbox-disk-audit.v1`
and `issue_count=60` (`unexpected-test-run-dir=59`, `legacy-selftest-root=1`). Those issues are
non-blocking evidence for this checkpoint and confirm the remaining full-run cleanup proof is still
real work.

### 2026-07-06 — Phase 1 stale TestSandbox run cleanup

Completed the §8/R5 pre-run dead-PID sibling cleanup slice:

- `Tools\TestRunPlan.ps1` now parses runner-generated run ids of the form
  `<UTC timestamp>-<PID>-<GUID>` and treats unparseable or overflowed names as manual/unknown
  directories, not cleanup targets.
- `Resolve-RSTestSandboxStaleRunTargets(...)` selects only sibling `runs\<runId>` directories with
  a parseable owner PID that is not currently live, while preserving the current run id, explicitly
  allowed run ids, live-PID run dirs, and unparseable/manual directories.
- `Remove-RSTestSandboxStaleRunDirectories(...)` removes those dead-PID siblings before child tests
  start, verifies each target remains under `REDSALAMANDER_TEST_ROOT\runs\`, and reports
  `Status=Removed` or `Status=Failed` rows. Removal failures warn and remain visible instead of
  becoming silent green evidence.
- `Run-AllTests.ps1` invokes the dead-PID sweep after creating the current run context and before the
  legacy cleanup reaper. This is current TestSandbox hygiene, so it still runs when
  `-SkipLegacySandboxCleanup` is supplied for diagnosing historical `%TEMP%`/profile roots.
- `RunAllTestsPlan.Tests.ps1` covers current/live/allowed/manual preservation and dead-PID removal.
- Specs/docs updated: `Specs/Testing/Testing_SelfTests.md`,
  `Specs/Testing/Testing_TestCoverage.md`, and `Tests/README.md`.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru
```

Result: `RunAllTestsPlan` passed 32 / 0 / 0. The first focused runner proof
`.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter showIdentical -TimeoutMultiplier 0.1 -SkipLegacySandboxCleanup`
passed Compare 1 / 0 / 0 with run id
`20260706T145934Z-78632-fe9b8b010cce4dbba8cef15d81d31009`. The runner removed 48 stale
PID/GUID-scoped sibling run directories before execution, with 0 cleanup failures. The aggregate
`test_sandbox_audit.issue_count` dropped from the prior focused proof's 60 issues to 13:
12 `unexpected-test-run-dir` entries for unparseable historical proof directories
(`crashproof_*`, `redconfigure-sandbox-proof-*`, `viewer-sqlite-sandbox-proof-*`,
`viewer-pe-sandbox-proof-*`, `dxui-artifact-proof-*`, `performance-tests2-sandbox-proof-*`, and
`crash-handling-sandbox-proof-*`) plus the known locked `legacy-selftest-root`. Those remaining
issues were local historical proof debris, not source-generated names.

After removing those local historical proof dirs under `.build\TestSandbox\runs\` and repairing the
ACL on legacy hang-dump artifacts under `%LOCALAPPDATA%\RedSalamander\SelfTest`, the focused proof
with legacy cleanup enabled:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter showIdentical -TimeoutMultiplier 0.1
```

passed Compare 1 / 0 / 0 with run id
`20260706T151836Z-38328-966786696a054e21872134aeee3860a6`. The runner removed the prior stale
PID/GUID-scoped run directory, reported 0 cleanup failures, and the post-run aggregate audit reported
`Disk audit issues: 0`. This proves the focused TestSandbox cleanup path is clean on this machine;
the remaining §8 exit proof is the actual `-Suite Full` run plus injected crash audit.

### 2026-07-06 — Suite Full disk proof first pass and Pester adapter fix

Executed the first actual `-Suite Full` disk-proof pass after the focused cleanup was clean:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

Result: the run exited 1 with run id
`20260706T151921Z-74032-457d5f5270ae4489954b239518e2d079`, but its aggregate
`test_sandbox_audit` was clean: `is_clean=true`, `issue_count=0`, and no legacy cleanup targets.
This proves the full runner's disk-audit plumbing can stay clean after the stale-run and legacy-root
cleanup work; the remaining blockers are functional suite failures, not TestSandbox sprawl.

Functional failure population from this pass:

- CompareDirectories: 3 failures. Two are default `Z:` sandbox filesystem-capability failures
  (`local_index_core_snapshot_reload` requiring NTFS and
  `local_search_native_unicode_long_path_matches_host_fallback` failing deep Unicode path creation);
  one is `search_service_sqlite_status_reports_maintenance_history` returning `0x80070002`.
- Commands: 19 failures. Several are default `Z:` sandbox filesystem-capability failures (alternate
  stream creation, indexed local-provider enumeration, long/large FolderView fixture writes, item
  properties/rename/navigation fixture roots), plus seven fatal-error dialog cases that did not open
  the dialog under selftest mode.
- DxUiTests: one executable failure at
  `TestMenuRainbowHoverUsesSeededHighlightContrast` (`rainbow hover highlight paint state reports the
  row as hovered`).
- RedSalamanderMonitorEtwLatency: executable exited 1. This pass used `-SkipBuild`, so monitor
  selftest hooks may not match the `Suite Full` build contract until rerun after a normal Full build.
- ToolsPesterTests: runner bug. `Run-AllTests.ps1` splatted raw
  `Invoke-Pester -Path ... -PassThru` strings, but this machine has Pester 3.4, whose path parameter
  is `-Script`; PowerShell misbound the raw strings into `EnableExit`.

Fixed in this checkpoint:

- `Tools\TestRunPlan.ps1` now exposes `New-RSPesterInvokeParameters(...)`, binding `-Script` when
  the installed `Invoke-Pester` exposes the Pester 3 parameter set and `-Path` when newer Pester
  exposes it. The helper explicitly forwards only reviewed `-Tag` and `-ExcludeTag` arguments.
- `Run-AllTests.ps1` uses the helper instead of raw string splatting.
- `RunAllTestsPlan.Tests.ps1` covers Pester 3 and newer parameter sets.
- Inventory docs updated for the new Tools Pester case.

Verification:

```powershell
Invoke-Pester -Path @('.\Tools\Tests\RunAllTestsPlan.Tests.ps1', '.\Tools\Tests\TestInventory.Tests.ps1') -PassThru
Invoke-Pester -Script .\Tools\Tests -ExcludeTag RequiresBuildToolchain -PassThru
```

Result: focused helper/inventory tests passed 38 / 0 / 0, and artifact-only Tools Pester passed
193 / 0 / 0 with `ExcludeTag=RequiresBuildToolchain`. Remaining work is to rerun the Full suite with
the adapter fix and correct filesystem/build environment, then attack the surviving failures
worst-first.

### 2026-07-06 — Phase 1 short NTFS root and Compare long-path fixture repair

Locked the local capability root and fixed the next real Compare blocker without quarantine:

- New local short NTFS root for capability-sensitive diagnostic and perf evidence:
  `REDSALAMANDER_TEST_ROOT=C:\RSPerf`.
- This root is used when the default workspace root is not NTFS, is redirected/network-backed, or is
  too long for path-sensitive Compare/search fixtures. It must keep the normal
  `runs\<runId>\scratch\...` and `runs\<runId>\artifacts\...` layout.
- `%LOCALAPPDATA%\Temp\RedSalamander-TestSandbox` is no longer acceptable for final path-sensitive
  evidence because that profile path is long enough to invalidate the Unicode long-path fixture.
- `C:\RST` remains legacy cleanup debt, not a blessed root.
- `local_search_native_unicode_long_path_matches_host_fallback` now creates its deep fixture
  directories and file through selftest-only extended Win32 paths (`\\?\...`). The product path under
  test is still passed as a normal root, so this fixes fixture setup instead of weakening the search
  behavior being verified.

RED proof:

```powershell
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter 'local_search_name_windows_filesystem_case_parity,local_search_native_unicode_long_path_matches_host_fallback,search_service_sqlite_query_failure_falls_back_live_scan,search_service_sqlite_midquery_failure_restarts_live_scan_without_duplicates,search_service_sqlite_prefilter_roundtrip' -TimeoutMultiplier 2
```

Result before the fixture repair: Compare 1 / 1 / 3 with run id
`20260706T163913Z-45728-09fd050947c446d8b1cc9dfd237a5598`; the failing case was
`local_search_native_unicode_long_path_matches_host_fallback` with `Failed to create a deep Unicode
search path.`

GREEN proof:

```powershell
.\build.ps1 -ProjectName RedSalamander
$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter 'local_search_name_windows_filesystem_case_parity,local_search_native_unicode_long_path_matches_host_fallback,search_service_sqlite_query_failure_falls_back_live_scan,search_service_sqlite_midquery_failure_restarts_live_scan_without_duplicates,search_service_sqlite_prefilter_roundtrip' -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter 'local_search_service_disconnect_falls_back_local_index,search_service_multi_client_and_rebuild_control' -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter 'search_service_foreground_rejects_second_instance,search_service_foreground_logs_request_status,search_service_status_and_query_roundtrip,search_service_query_reports_live_progress,search_service_filters_cached_descendants_denied_to_client,search_service_transient_authorization_failure_is_incomplete_not_cached,search_service_candidate_impersonation_failure_is_incomplete_warning,local_search_service_indexed_name_latency_and_parity,local_search_service_matches_host_fallback,local_search_service_single_request_uses_query,local_search_service_content_early_stop_stats,local_search_service_protocol_mismatch_falls_back_local_index,local_search_service_disconnect_falls_back_local_index,search_service_multi_client_and_rebuild_control' -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter 'search_service_multi_client_and_rebuild_control' -SelfTestRepeat 20 -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -TimeoutMultiplier 2
```

Result after the fixture repair: Debug `RedSalamander` build log
`.build\logs\msbuild-20260706_184141_548.log` had 0 warnings / 0 errors. Focused Compare passed
2 / 0 / 3 with run id `20260706T164345Z-66432-c101d5812182430e84fb1ac9940dfa45`; the skipped
SQLite direct-currentness cases still require a live journal cursor and are environmental skips, not
quarantine. The adjacent disconnect/multi-client probe passed 2 / 0 / 0 with run id
`20260706T165206Z-67140-68347f2c5b7f45689b04cf1f0809ba15`. The wider service cluster passed
14 / 0 / 0 with run id `20260706T165441Z-83624-65d6da5091864fd281e045de92ab2764`. The
multi-client case passed 20 / 0 / 0 under repeat with run id
`20260706T165802Z-80364-cd17c6bf6a1b45b99558f12010945dc2`. Fresh full Compare passed 219 / 0 / 30
with run id `20260706T165814Z-93164-30befa2eb8024a4ca9743760b9a43e86` and `Disk audit issues: 0`.

One earlier full Compare attempt after the long-path repair failed once in
`search_service_multi_client_and_rebuild_control` with `ERROR_BROKEN_PIPE`; this was not quarantined
or made non-blocking. It did not reproduce in adjacent order, the wider service cluster, 20x repeat,
or the fresh full Compare pass above. If it appears again, it remains a blocking service-lifecycle
repair item and must be fixed or replaced with service-side diagnostics, not suppressed.

### 2026-07-06 — Phase 3 local index snapshot reload advisory timing

Completed the first timing/perf-axis slice:

- `local_index_core_snapshot_reload` no longer fails correctness on the arbitrary
  `warmElapsedMs < 1000u` ceiling.
- The case still asserts the causal correctness signals: snapshot reload state, candidate set, and
  query result shape.
- The removed ceiling was replaced by the advisory metric
  `compare.selftest.local_index.snapshot_reload_us`, emitted with `Debug::Perf::EmitDurationUs(...)`.
  `value0` records the warm candidate count and `value1` records snapshot bytes.
- `TestHarnessSourceContracts.Tests.ps1` guards both sides of the contract: the metric must remain,
  and the old `warmElapsedMs < 1000u` / "Warm indexed query took too long" gate must not return.

Perf Measurement Record:

- Scenario: Local index core snapshot reload after cache eviction and small-tree warmup.
- Subsystem: Compare/search local index selftest.
- Change type: stabilization.
- User-visible risk protected: a correct warm indexed query must not be reported as a regression
  solely because machine load, Defender, or filesystem state pushed an arbitrary elapsed ceiling over
  1000 ms.
- Metric keys: `compare.selftest.local_index.snapshot_reload_us`.
- Metric units and sample grain: microseconds per warm reload, one row per case run; `value0` is
  candidate count, `value1` is snapshot file bytes.
- Existing instrumentation reused or new instrumentation added: new advisory
  `Debug::Perf::EmitDurationUs(...)` row in `local_index_core_snapshot_reload`.
- Deterministic validation: `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter local_index_core_snapshot_reload -TimeoutMultiplier 0.1 -SkipLegacySandboxCleanup`.
- Build flavor: Debug diagnostic proof; Release baseline/candidate evidence remains required before
  making a throughput, latency, p95/p99, or budget claim.
- Baseline run: [blocked] not captured before this stabilization change.
- Candidate run: `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-06_131232`.
- Same-machine and same-suite comparison: [blocked] baseline missing; candidate-only result is
  directional.
- Analyzer command: `.\Tools\Show-PerfRuns.ps1 -Run .\Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-06_131232 -Metric compare.selftest.local_index.snapshot_reload_us -ShowBuildFlavor`.
- Sample-quality result: one advisory sample at 6687 us; no p95/p99 claim because sample count is 1.
- Environment matrix: historical Debug diagnostic proof used
  `REDSALAMANDER_TEST_ROOT=C:\Users\eric\AppData\Local\Temp\RedSalamander-TestSandbox` because the
  default `Z:\src\RedSalamander\.build\TestSandbox` root does not satisfy this case's NTFS
  requirement. That long profile path is not final evidence after the short-root checkpoint above;
  final path-sensitive evidence must use `REDSALAMANDER_TEST_ROOT=C:\RSPerf`.
- Archived evidence: `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-06_131232\perf\perf_metrics.jsonl`.
- Before/after delta: [blocked] no same-machine baseline archive exists.
- Caveats or blockers: Debug-only, one-sample diagnostic evidence; final perf claim needs
  test-enabled Release evidence and a same-machine baseline.
- Authoritative spec updates required: `Operation_PerfMeasurementContract_2026-07-06.md` records the
  exact metric key, sample fields, root override rule, and candidate evidence.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
$env:REDSALAMANDER_TEST_ROOT='C:\Users\eric\AppData\Local\Temp\RedSalamander-TestSandbox'; try { .\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter local_index_core_snapshot_reload -TimeoutMultiplier 0.1 -SkipLegacySandboxCleanup } finally { Remove-Item Env:REDSALAMANDER_TEST_ROOT -ErrorAction SilentlyContinue }
.\Tools\Show-PerfRuns.ps1 -Run .\Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-06_131232 -Metric compare.selftest.local_index.snapshot_reload_us -ShowBuildFlavor
```

Result: `TestHarnessSourceContracts` 63 / 0 / 0, `TestInventory` 5 / 0 / 0, Debug build log
`.build\logs\msbuild-20260706_130836_929.log` had 0 warnings / 0 errors, focused Compare runtime
proof passed 1 / 0 / 0, and the analyzer showed the advisory metric row in Debug with count=1 and
6687 us. The same focused run against the default `Z:` sandbox failed before executing the metric
because the case correctly requires NTFS; that is now recorded as an environment/root requirement
rather than a flaky-test exception.

### 2026-07-06 — Phase 3 FolderView overlay advisory sampling

Completed the overlay sample-count stabilization slice:

- `folderView_perf_overlay_invalidation_stress` no longer uses the fixed `Scale(4200ms)` collection
  window.
- The case now collects toward the frame sample target with a pass-count safety cap and records
  `overlayFrameCollectionPassCount`.
- `overlaySamplesEnoughForP95` is no longer a correctness `Require`; it remains in `metricQuality`
  and in the case artifact as advisory sample-quality evidence.
- The overlay correctness gate remains: at least 8 natural overlay animation frames must still be
  observed.
- Strict perf-budget runs still enforce `FolderViewPerfBudgets.json5` `minimumSamples` through
  `CheckFolderViewPerfBudgets(...)`; this change only removes the normal selftest's flaky
  animation-window hard failure.
- Durable rule merged into `Specs\Testing\Testing_PerformanceValidation.md`; source guard added to
  `Tools\Tests\TestHarnessSourceContracts.Tests.ps1`; coverage counts updated in
  `Specs\Testing\Testing_TestCoverage.md` and `Tests\README.md`.

Perf Measurement Record:

- Scenario: FolderView incremental-search overlay plus busy/cancel overlay animation under
  deterministic rendering load.
- Subsystem: FolderView Commands selftest.
- Change type: stabilization.
- User-visible risk protected: overlay animation and repaint load must not turn CI load into a false
  frame-sample failure while still preserving analyzer-ready frame metrics.
- Metric keys: `folder.frame.total_us`, `folder.frame.present_us`,
  `folder.frame.overlay_animation_count`, `folder.frame.overlay_dirty_rect_area_px`,
  `render.incremental_search_effect_updates`, and artifact `metricQuality`.
- Metric units and sample grain: microseconds per frame for `folder.frame.*_us`; one artifact
  `metricQuality` summary per case run.
- Existing instrumentation reused or new instrumentation added: reused existing FolderView frame and
  overlay metrics; added `overlayFrameCollectionPassCount` and
  `overlaySamplesEnoughForP95Advisory` to the focused artifact.
- Deterministic validation: `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_perf_overlay_invalidation_stress -TimeoutMultiplier 0.2 -SkipLegacySandboxCleanup`.
- Build flavor: Debug diagnostic proof; test-enabled Release evidence remains required before any
  throughput, latency, p95/p99, or budget claim.
- Baseline run: [blocked] not captured before this stabilization change.
- Candidate run: `Specs\TestRuns\4cb089111a23\Commands\2026-07-06_135251`.
- Same-machine and same-suite comparison: [blocked] baseline missing; candidate-only result is
  directional.
- Analyzer command: `.\Tools\Show-PerfRuns.ps1 -Run .\Specs\TestRuns\4cb089111a23\Commands\2026-07-06_135251 -FolderViewPreset -FailOnQuality`.
- Sample-quality result: analyzer showed `folder.frame.total_us` count=1192 / P95Quality=pass /
  P99Quality=pass and `folder.frame.present_us` count=1192 / P95Quality=pass / P99Quality=pass;
  artifact `metricQuality` recorded `folderFrameTotal.count=1143`,
  `folderFramePresent.count=1143`, `samplesEnoughForP95=true`, and `samplesEnoughForP99=true`.
- Environment matrix: diagnostic proof used default
  `REDSALAMANDER_TEST_ROOT=Z:\src\RedSalamander\.build\TestSandbox`; run id
  `20260706T115223Z-80136-987d3d9b853548e6a7113195405ec23e`.
- Archived evidence: `Specs\TestRuns\4cb089111a23\Commands\2026-07-06_135251`.
- Before/after delta: [blocked] no same-machine baseline archive exists.
- Caveats or blockers: Debug-only diagnostic evidence; final perf claim needs test-enabled Release
  evidence and a same-machine baseline.
- Authoritative spec updates required: `Specs\Testing\Testing_PerformanceValidation.md` documents
  the overlay-specific advisory sample-quality rule.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter folderView_perf_overlay_invalidation_stress -TimeoutMultiplier 0.2 -SkipLegacySandboxCleanup
.\Tools\Show-PerfRuns.ps1 -Run .\Specs\TestRuns\4cb089111a23\Commands\2026-07-06_135251 -FolderViewPreset -FailOnQuality
```

Result: `TestHarnessSourceContracts` 67 / 0 / 0, `TestInventory` 5 / 0 / 0, Debug build log
`.build\logs\msbuild-20260706_135015_373.log` had 0 warnings / 0 errors, focused Commands runtime
proof passed 1 / 0 / 0, and the analyzer quality gate passed for the frame p95 rows.

### 2026-07-06 — Phase 3 named raw-wait scaling

Completed the named raw-wait slice from §1/Phase 3:

- `Commands.SelfTest.Search.cpp` now scales `WaitForFlag(...)` internally through
  `SelfTest::ScaleTimeout(timeoutMs)`, so callers keep the scenario base budget while the runner
  multiplier controls the actual deadline.
- `Commands.SelfTest.BatchRename.cpp` now routes the gated provider `WaitForSingleObject(...)`
  through `SelfTest::ScaleTimeout(30'000u)`.
- `FolderWindow.FileOperations.SelfTest.cpp` replaced `kMaxAttempts=120` with a scaled 6-second
  deadline and 50 ms retry slices for temp-root recreation under AV/indexer churn.
- `TestHarnessSourceContracts.Tests.ps1` guards all three named offenders so the raw literals cannot
  silently return.
- Durable rule merged into `Specs\Testing\Testing_SelfTests.md`; coverage counts updated in
  `Specs\Testing\Testing_TestCoverage.md` and `Tests\README.md`.

Perf Measurement Record:

- Scenario: Search directory-watch callback entry, BatchRename gated provider execution, and
  FileOperations setup temp-root recreation under slow runner or filesystem churn.
- Subsystem: Commands selftest and FileOperations selftest.
- Change type: stabilization.
- User-visible risk protected: correct asynchronous progress must not be reported as a timeout
  solely because a loaded runner needs longer than an unscaled fixed wait.
- Metric keys: native result `duration_ms` fields in `commands_results.json`,
  `fileops_results.json`, and aggregate `selftest_run_results.json`; no subsystem latency metric was
  added because this slice does not claim a product performance improvement.
- Metric units and sample grain: milliseconds per focused selftest case or phase.
- Existing instrumentation reused or new instrumentation added: reused native selftest result
  durations and source-contract guard; no new `Debug::Perf` row because the protected behavior is
  timeout scaling, not a latency budget.
- Deterministic validation:
  `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter filesystem_local_watch_unwatch_drains_inflight_callback -TimeoutMultiplier 1.0 -SkipLegacySandboxCleanup`,
  `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_pane_batchRename_execute_while_busy_returns_busy -TimeoutMultiplier 1.0 -SkipLegacySandboxCleanup`, and
  `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Phase8_InvalidDestinationRejected -TimeoutMultiplier 1.0 -SkipLegacySandboxCleanup`.
- Build flavor: Debug focused proof after a fresh Debug build; no Release throughput, latency, p95,
  p99, or budget claim is made.
- Baseline run: [blocked] not captured before this stabilization change.
- Candidate run: `Specs\TestRuns\4cb089111a23\Commands\2026-07-06_140144`,
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-06_140144_001`, and
  `Specs\TestRuns\4cb089111a23\FileOps\2026-07-06_140134`.
- Same-machine and same-suite comparison: [blocked] baseline missing; candidate-only evidence proves
  focused correctness after scaling, not a before/after performance delta.
- Analyzer command: not applicable for this correctness-scaling slice; no percentile or subsystem
  perf-metric claim is made.
- Sample-quality result: one focused sample per protected path; p95/p99 quality is not claimed.
- Environment matrix: archived `env.txt` files in the candidate folders record machine/environment
  context; summaries record `timeout_scale=1.0`.
- Archived evidence: candidate folders listed above include result JSON and trace artifacts.
- Before/after delta: [blocked] no same-machine baseline archive exists.
- Caveats or blockers: Release evidence remains required before making any product latency or
  throughput statement; this slice only proves the named correctness waits honor the multiplier.
- Authoritative spec updates required: `Specs\Testing\Testing_SelfTests.md` now requires correctness
  waits to establish scaled deadlines and allows fixed literals only as short polling slices.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter filesystem_local_watch_unwatch_drains_inflight_callback -TimeoutMultiplier 1.0 -SkipLegacySandboxCleanup
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_pane_batchRename_execute_while_busy_returns_busy -TimeoutMultiplier 1.0 -SkipLegacySandboxCleanup
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Phase8_InvalidDestinationRejected -TimeoutMultiplier 1.0 -SkipLegacySandboxCleanup
```

Result: RED source-contract Pester first failed at 67 / 1 / 0 against the unscaled waits; GREEN
`TestHarnessSourceContracts` 68 / 0 / 0, `TestInventory` 5 / 0 / 0, Debug build log
`.build\logs\msbuild-20260706_135919_196.log` had 0 warnings / 0 errors, Search focused proof
passed 1 / 0 / 0, BatchRename focused proof passed 1 / 0 / 0, and FileOps focused proof passed
3 / 0 / 0.

### 2026-07-06 — Phase 1 native TestSandbox scratch acquisition helper

Completed the first §8/R3 migration slice:

- `SelfTestCommon` now exposes `SelfTest::TestSandbox` and
  `SelfTest::AcquireTestSandbox(suite, caseName)`.
- Under normal runner execution the helper acquires scratch paths under
  `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\<suite>\<case>\`, separate from
  `artifacts\selftest\last_run`.
- Deliberate ad hoc/legacy runs fall back under `SelfTestRoot()\last_run\<suite>\scratch\<case>` so
  existing direct launches keep working while the migration proceeds.
- Compare foreground search-service stdout capture now uses `AcquireTestSandbox(...)` and no longer
  calls `GetTempPathW` / `GetTempFileNameW`.
- `TestHarnessSourceContracts.Tests.ps1` now guards the helper shape and this first direct temp-root
  migration.
- Durable rule merged into `Specs\Testing\Testing_SelfTests.md`; coverage counts updated in
  `Specs\Testing\Testing_TestCoverage.md` and `Tests\README.md`.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
$env:REDSALAMANDER_TEST_ROOT='C:\Users\eric\AppData\Local\Temp\RedSalamander-TestSandbox'; try { .\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_service_foreground_logs_request_status -TimeoutMultiplier 0.2 -SkipLegacySandboxCleanup } finally { Remove-Item Env:REDSALAMANDER_TEST_ROOT -ErrorAction SilentlyContinue }
```

Result: `TestHarnessSourceContracts` 65 / 0 / 0, `TestInventory` 5 / 0 / 0, Debug build log
`.build\logs\msbuild-20260706_132313_125.log` had 0 warnings / 0 errors, and focused Compare runtime
proof passed 1 / 0 / 0 with `run_id=20260706T112546Z-628-4e4d443598ea4851bd5ebbb5cc65f270`.
The trace recorded
`TestSandbox: suite=compare case=foreground_search_service_stdout root='C:\Users\eric\AppData\Local\Temp\RedSalamander-TestSandbox\runs\20260706T112546Z-628-4e4d443598ea4851bd5ebbb5cc65f270\scratch\compare\foreground_search_service_stdout'`,
and the archive path was
`Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-06_132548`.

### 2026-07-06 — Phase 1 FileOps alternate-volume TestSandbox migration

Completed the §8/R4 migration for the real cross-volume FileOps case:

- `SelfTestCommon` now exposes `SelfTest::AcquireTestSandboxOnVolume(...)` for the one sanctioned
  alternate-volume exception.
- FileOperations real cross-volume move fallback now allocates under
  `<AltDrive>:\RedSalamanderTestSandbox\runs\<runId>\scratch\fileops\real_cross_volume_move`
  instead of creating `RedSalamanderCrossVolumeSelfTest_*` at the drive root.
- The case removes the alternate-volume case root and prunes empty `fileops`, `scratch`, and
  `<runId>` parents after cleanup; the stable `RedSalamanderTestSandbox\runs` root may remain.
- `TestHarnessSourceContracts.Tests.ps1` guards the new helper, sanctioned alternate root,
  FileOps call site, cleanup pruning helper, and absence of the legacy drive-root prefix.
- Durable rule merged into `Specs\Testing\Testing_SelfTests.md`; coverage counts updated in
  `Specs\Testing\Testing_TestCoverage.md` and `Tests\README.md`.

Perf Measurement Record:

- Scenario: FileOperations recursive-copy matrix real cross-volume move fallback.
- Subsystem: FileOperations selftest.
- Change type: stabilization.
- User-visible risk protected: the cross-volume fallback must continue to exercise a real alternate
  fixed volume while no longer poisoning future runs with drive-root scratch directories.
- Metric keys: `FileOps.SelfTest.CopyRecursiveMatrixRealCrossVolumeMove`.
- Metric units and sample grain: microseconds per focused case run; `value0` is callback count and
  `value1` is matching stream count.
- Existing instrumentation reused or new instrumentation added: reused existing FileOps recursive
  matrix metric row and added source-contract coverage for the new root.
- Deterministic validation:
  `.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Phase7_CopyRecursiveParallelismMatrix -TimeoutMultiplier 1.0 -SkipLegacySandboxCleanup`.
- Build flavor: Debug focused proof after a fresh Debug build; no Release throughput, latency, p95,
  p99, or budget claim is made.
- Baseline run: [blocked] not captured before this root migration.
- Candidate run: `Specs\TestRuns\4cb089111a23\FileOps\2026-07-06_142222`.
- Same-machine and same-suite comparison: [blocked] baseline missing; candidate-only evidence proves
  focused correctness and root cleanup, not a before/after performance delta.
- Analyzer command: not applicable for this correctness/root migration slice; no percentile claim is
  made.
- Sample-quality result: one focused metric sample; p95/p99 quality is not claimed.
- Environment matrix: archived `env.txt` in the candidate folder records machine/environment context.
- Archived evidence: `Specs\TestRuns\4cb089111a23\FileOps\2026-07-06_142222` contains result JSON,
  trace artifacts, and `perf\perf_metrics.jsonl`.
- Before/after delta: [blocked] no same-machine baseline archive exists.
- Caveats or blockers: final §8 closeout still requires the full-run disk audit and any remaining
  non-FileOps scratch migrations that audit exposes.
- Authoritative spec updates required: `Specs\Testing\Testing_SelfTests.md` now documents the
  alternate-volume exception and legacy `RedSalamanderCrossVolumeSelfTest_*` cleanup-only status.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter Phase7_CopyRecursiveParallelismMatrix -TimeoutMultiplier 1.0 -SkipLegacySandboxCleanup
```

Result: RED source-contract Pester first failed at 68 / 1 / 0 against the missing
`PruneEmptyAlternateVolumeSandboxParents` helper; GREEN `TestHarnessSourceContracts` 69 / 0 / 0,
`TestInventory` 5 / 0 / 0, Debug build log `.build\logs\msbuild-20260706_142008_343.log` had
0 warnings / 0 errors, and focused FileOps runtime proof passed 3 / 0 / 0. The trace recorded
`TestSandbox alternate-volume: suite=fileops case=real_cross_volume_move root='C:\RedSalamanderTestSandbox\runs\20260706T122220Z-80748-1db6d56794f4467296704966c31214b9\scratch\fileops\real_cross_volume_move'`.
The metric row recorded `FileOps.SelfTest.CopyRecursiveMatrixRealCrossVolumeMove` with
`detail="shape=real-cross-volume-move-fallback"`, `durationUs=31000`, `value0=3`, `value1=3`,
`hr=0`, build `Debug`, machine hash `4cb089111a23`, and run id `2026-07-06_122220`.
Post-run cleanup check proved the alternate-volume `<runId>`, `scratch`, `fileops`, and
`real_cross_volume_move` paths were absent, and a `C:\RedSalamanderCrossVolumeSelfTest_*` scan
returned no directories.

### 2026-07-06 — Phase 1 MTP fake-journal LOCALAPPDATA TestSandbox migration

Completed the §8/R3/R6 migration for the fake MTP overwrite-journal selftests:

- `CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp` now redirects `LOCALAPPDATA` through
  `SelfTest::AcquireTestSandbox(SelfTest::SelfTestSuite::CompareDirectories, ...)` for the six
  fake MTP overwrite-journal cases instead of reading the user's ambient `%LOCALAPPDATA%`.
- The actual runner-visible case names are unchanged; only the scratch-root segments are shortened
  (`mtp_journal_orphan_cleanup`, `mtp_journal_temp_retry`, `mtp_journal_no_temp_puid`,
  `mtp_journal_rename_reject`, `mtp_overwrite_safety_matrix`, and
  `mtp_journal_completed_swap`) so the per-run sandbox root plus
  `RedSalamander\PluginState\FileSystemMtp\<hash>\overwrite-journal.json` stays under classic
  Win32 path limits used by the local test journal writer.
- `wil::scope_exit` restores the previous `LOCALAPPDATA` after the MTP plugin instances and fake
  backends have torn down, so per-case state cannot leak into the next case or the user's profile.
- `TestHarnessSourceContracts.Tests.ps1` guards the helper, the TestSandbox acquisition, the
  `LOCALAPPDATA` redirect, the shortened scratch segments, and the absence of the old
  `GetEnvironmentVariableW(L"LOCALAPPDATA")` / `getLocalAppDataPath` helpers.
- Coverage counts updated in `Specs\Testing\Testing_TestCoverage.md` and `Tests\README.md`.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
git diff --check
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter 'mtp_overwrite_journal_replay_removes_temp_when_final_exists,mtp_overwrite_journal_replay_temp_cleanup_delete_failure_retries,mtp_overwrite_journal_recovers_committed_temp_without_tempPuid,mtp_overwrite_journal_replay_rename_rejection_is_bounded,mtp_overwrite_never_duplicates_or_halfwrites,mtp_overwrite_journal_clears_completed_swap_without_temp' -TimeoutMultiplier 1.0 -SkipLegacySandboxCleanup
```

Result: RED source-contract Pester first failed at 79 / 1 / 0 against the legacy `LOCALAPPDATA`
reader. The first runtime attempt after redirecting `LOCALAPPDATA` failed 4 / 2 / 0 because long
scratch segments produced journal paths up to 261 characters. After shortening the scratch
segments, GREEN `TestHarnessSourceContracts` 80 / 0 / 0, `TestInventory` 5 / 0 / 0, `git diff
--check` had only line-ending warnings, Debug build log `.build\logs\msbuild-20260706_161107_992.log`
had 0 warnings / 0 errors, and focused Compare runtime proof passed 6 / 0 / 0 with run id
`20260706T141313Z-74896-721c0782d2494b1f9456647a1ff0dc51`. The trace recorded `LOCALAPPDATA`
redirected to `.build\TestSandbox\runs\20260706T141313Z-74896-721c0782d2494b1f9456647a1ff0dc51\scratch\compare\<short-mtp-segment>`
for all six cases.

### 2026-07-06 — Phase 1 Tools Pester TestSandbox migration

Completed the §8/R3/R6 tooling-side scratch migration for the remaining Pester tests that were
still creating roots under process temp:

- `Tools\TestRunPlan.ps1` now exposes `New-RSTestSandboxScratchDirectory(...)`, which validates
  harness/case path segments, resolves the canonical `REDSALAMANDER_TEST_ROOT` +
  `REDSALAMANDER_TEST_RUN_ID` context, creates
  `runs\<runId>\scratch\<harness>\<case>`, and returns a full path.
- `WingetValidation.Tests.ps1`, `VcpkgInstallSafety.Tests.ps1`, and `ShowPerfRuns.Tests.ps1`
  now route fake manifest roots, VC runtime redist fixtures, vcpkg merge fixtures, and synthetic
  perf-run trees under `runs\<runId>\scratch\tools-pester\<case>` instead of
  `[System.IO.Path]::GetTempPath()`.
- `Get-RSTestSandboxLegacyCleanupPlan(...)` now includes historical Pester temp prefixes:
  `RSWingetValidationTest_*`, `RSVcRuntimeTest_*`, `rs-vcpkg-install-root-test*`,
  `rs-vcpkg-single-file-merge-*`, and `rs-show-perfruns-tests-*`.
- `TestHarnessSourceContracts.Tests.ps1` guards the migrated Pester scripts and helper shape;
  `RunAllTestsPlan.Tests.ps1` covers helper behavior and path-segment rejection.
- Durable rules were merged into `Specs\Testing\Testing_SelfTests.md`; inventory counts were
  updated in `Specs\Testing\Testing_TestCoverage.md`, `Tests\README.md`, and
  `TestInventory.Tests.ps1`.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\WingetValidation.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\VcpkgInstallSafety.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\ShowPerfRuns.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
rg -n "\[System\.IO\.Path\]::GetTempPath\(\)" Tools\Tests -g "*.ps1"
rg -n "GetTempPath|GetTempFileName|temp_directory_path|RedSalamanderCrossVolumeSelfTest_|\[System\.IO\.Path\]::GetTempPath\(\)" Tools\Tests Tests RedSalamander\SelfTest -g "*.ps1" -g "*.cpp" -g "*.h" -g "*.hpp"
git diff --check
```

Result: RED source-contract Pester first failed at 80 / 1 / 0 against the missing Pester
TestSandbox helper and migrated scripts. GREEN `RunAllTestsPlan` 30 / 0 / 0,
`TestHarnessSourceContracts` 81 / 0 / 0, `WingetValidation` 17 / 0 / 0,
`VcpkgInstallSafety` 5 / 0 / 0, `ShowPerfRuns` 5 / 0 / 0, and `TestInventory` 5 / 0 / 0.
The raw `[System.IO.Path]::GetTempPath()` scan under `Tools\Tests` returned no matches; the broader
temp-pattern scan returned only guard/cleanup literals in `TestHarnessSourceContracts` and
`RunAllTestsPlan`. `git diff --check` reported only LF/CRLF normalization warnings.

### 2026-07-06 — Phase 1 RedConfigureTests TestSandbox migration

Completed the first standalone harness migration from §8/R1-R3:

- `RedConfigureTests.cpp` now resolves scratch paths from `REDSALAMANDER_TEST_ROOT` and
  `REDSALAMANDER_TEST_RUN_ID`.
- Direct launches without the runner use the same layout under `.build\TestSandbox`, choosing the
  sibling `.build\TestSandbox` when launched from `.build\<platform>\<configuration>`.
- All 14 formerly fixed `%TEMP%\RedConfigure*` roots now live under
  `<REDSALAMANDER_TEST_ROOT>\runs\<runId>\scratch\redconfigure\<case>`.
- The workspace-root negative test now uses a pure no-marker relative path so the repo-local
  `TestSandbox` parent cannot invalidate the assertion by design.
- `TestHarnessSourceContracts.Tests.ps1` guards the helper shape, the unified env names, the
  `runs\...\scratch\redconfigure` layout, and the absence of
  `std::filesystem::temp_directory_path`.
- Durable rule merged into `Specs\Testing\Testing_SelfTests.md`; coverage counts updated in
  `Specs\Testing\Testing_TestCoverage.md` and `Tests\README.md`.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedConfigureTests -Configuration Debug
$env:REDSALAMANDER_TEST_ROOT=(Resolve-Path '.\.build\TestSandbox').Path
$env:REDSALAMANDER_TEST_RUN_ID='redconfigure-sandbox-proof-20260706-1431'
.\.build\x64\Debug\RedConfigureTests.exe
```

Result: RED source-contract Pester first failed at 69 / 1 / 0 against the missing
`AcquireRedConfigureTestSandbox` helper; GREEN `TestHarnessSourceContracts` 70 / 0 / 0, Debug
`RedConfigureTests` build log `.build\logs\msbuild-20260706_143123_906.log` had 0 warnings /
0 errors, and focused standalone runtime proof passed with all 23 RedConfigure tests green. The run
created `Z:\src\RedSalamander\.build\TestSandbox\runs\redconfigure-sandbox-proof-20260706-1431\scratch\redconfigure`;
after per-case cleanup the `redconfigure` root existed with no child fixtures.

### 2026-07-06 — Phase 1 ViewerSqliteTests TestSandbox migration

Completed the next standalone harness migration from §8/R1-R3:

- `ViewerSqliteTests.cpp` now resolves its temporary database root from
  `REDSALAMANDER_TEST_ROOT` and `REDSALAMANDER_TEST_RUN_ID`.
- Direct launches without the runner use the same layout under `.build\TestSandbox`, choosing the
  sibling `.build\TestSandbox` when launched from `.build\<platform>\<configuration>`.
- The temporary SQLite database now lives under
  `<REDSALAMANDER_TEST_ROOT>\runs\<runId>\scratch\viewer-sqlite\database\viewer-sqlite-tests-*.sqlite`.
- `TempDatabase` removes the database file and the per-case `database` scratch root during cleanup.
- `TestHarnessSourceContracts.Tests.ps1` guards the helper shape, the unified env names, the
  `runs\...\scratch\viewer-sqlite` layout, and the absence of
  `std::filesystem::temp_directory_path`.
- Durable rule merged into `Specs\Testing\Testing_SelfTests.md`; coverage counts updated in
  `Specs\Testing\Testing_TestCoverage.md` and `Tests\README.md`.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
.\build.ps1 -ProjectName ViewerSqliteTests -Configuration Debug
$env:REDSALAMANDER_TEST_ROOT=(Resolve-Path '.\.build\TestSandbox').Path
$env:REDSALAMANDER_TEST_RUN_ID='viewer-sqlite-sandbox-proof-20260706-1439'
.\.build\x64\Debug\ViewerSqliteTests.exe
```

Result: RED source-contract Pester first failed at 70 / 1 / 0 against the missing
`AcquireViewerSqliteTestSandbox` helper; GREEN `TestHarnessSourceContracts` 71 / 0 / 0, Debug
`ViewerSqliteTests` build log `.build\logs\msbuild-20260706_143835_385.log` had 0 warnings /
0 errors, and focused standalone runtime proof passed. The run created
`Z:\src\RedSalamander\.build\TestSandbox\runs\viewer-sqlite-sandbox-proof-20260706-1439\scratch\viewer-sqlite`;
after database cleanup the `viewer-sqlite` root existed with no child fixtures.

### 2026-07-06 — Phase 1 CrashHandlingTests TestSandbox migration

Completed the next standalone harness migration from §8/R1-R3:

- `CrashHandlingTests.cpp` now resolves its crash-marker test root from
  `REDSALAMANDER_TEST_ROOT` and `REDSALAMANDER_TEST_RUN_ID`.
- Direct launches without the runner use the same layout under `.build\TestSandbox`, choosing the
  sibling `.build\TestSandbox` when launched from `.build\<platform>\<configuration>`.
- The crash-handler marker helper tests now write under
  `<REDSALAMANDER_TEST_ROOT>\runs\<runId>\scratch\crash-handling\marker-files`.
- `TestHarnessSourceContracts.Tests.ps1` guards the helper shape, the unified env names, the
  `runs\...\scratch\crash-handling` layout, and the absence of
  `std::filesystem::temp_directory_path`.
- Durable rule merged into `Specs\Testing\Testing_SelfTests.md`; coverage counts updated in
  `Specs\Testing\Testing_TestCoverage.md` and `Tests\README.md`.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
.\build.ps1 -ProjectName CrashHandlingTests -Configuration Debug
$env:REDSALAMANDER_TEST_ROOT=(Resolve-Path '.\.build\TestSandbox').Path
$env:REDSALAMANDER_TEST_RUN_ID='crash-handling-sandbox-proof-20260706-1441'
.\.build\x64\Debug\CrashHandlingTests.exe
```

Result: RED source-contract Pester first failed at 71 / 1 / 0 against the missing
`AcquireCrashHandlingTestSandbox` helper; GREEN `TestHarnessSourceContracts` 72 / 0 / 0, Debug
`CrashHandlingTests` build log `.build\logs\msbuild-20260706_144107_084.log` had 0 warnings /
0 errors, and focused standalone runtime proof passed. The run created
`Z:\src\RedSalamander\.build\TestSandbox\runs\crash-handling-sandbox-proof-20260706-1441\scratch\crash-handling`;
after marker-file cleanup the `crash-handling` root existed with no child fixtures.

### 2026-07-06 — Phase 1 Commands plugin-config TestSandbox migration

Completed the next native Commands fixture migration from §8/R1-R3:

- `Commands.SelfTest.PluginConfig.cpp` now routes the ViewerText hex-byte perf fixture,
  ViewerText diff perf fixture, ViewerImgRaw close-roundtrip fixture, and ViewerWeb close-roundtrip
  fixture through `SelfTest::AcquireTestSandbox(SelfTest::SelfTestSuite::Commands, ...)`.
- The migrated case roots are
  `<REDSALAMANDER_TEST_ROOT>\runs\<runId>\scratch\commands\viewer_text_hex_byte_color_perf`,
  `...\viewer_text_diff_perf`, `...\viewer_imgraw_close_roundtrip`, and
  `...\viewer_web_close_roundtrip`.
- The old `RedSalamander.ViewerTextPerf.*`, `RedSalamander.ViewerTextDiffPerf.*`,
  `RedSalamander.ViewerImgRawClose.*`, and `RedSalamander.ViewerWebClose.*` process-temp roots are
  gone from the source.
- `TestHarnessSourceContracts.Tests.ps1` guards all four native `AcquireTestSandbox(...)` calls and
  forbids `std::filesystem::temp_directory_path` in the plugin-config source.
- Durable rule merged into `Specs\Testing\Testing_SelfTests.md`; coverage counts updated in
  `Specs\Testing\Testing_TestCoverage.md` and `Tests\README.md`.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter 'viewer_text_hex_byte_color_perf,viewer_text_diff_perf,viewer_imgraw_close_roundtrip,viewer_web_close_roundtrip' -TimeoutMultiplier 1.0 -SkipLegacySandboxCleanup
```

Result: RED source-contract Pester first failed at 72 / 1 / 0 against the four process-temp roots;
GREEN `TestHarnessSourceContracts` 73 / 0 / 0, Debug `RedSalamander` build log
`.build\logs\msbuild-20260706_144915_221.log` had 0 warnings / 0 errors, and focused Commands
runtime proof passed 4 / 0 / 0 with run id
`20260706T125122Z-78724-35dc0b8689764f4eaf08d9357f6d1939`. The trace recorded all four
`TestSandbox: suite=commands` case roots under
`Z:\src\RedSalamander\.build\TestSandbox\runs\20260706T125122Z-78724-35dc0b8689764f4eaf08d9357f6d1939\scratch\commands\...`,
the archive landed at `Specs\TestRuns\4cb089111a23\Commands\2026-07-06_145125`, and
`scratch\commands` had no remaining child fixtures after cleanup.

### 2026-07-06 — Phase 1 ShellCommands shortcut-save TestSandbox migration

Completed the in-product ShellCommands Win32 temp-file migration from §8/R1-R3:

- `SaveShellLinkForShellCommandTest(...)` no longer calls `GetTempPathW` or `GetTempFileNameW` for
  the long-path `.lnk` save fallback.
- The fallback now creates a short unique `.lnk` file under
  `SelfTest::AcquireTestSandbox(SelfTest::SelfTestSuite::Commands, L"shell_shortcut_save_temp")`,
  then keeps the existing `IPersistFile::Save(temp)` plus `MoveFileExW(temp, longLinkPath, ...)`
  behavior.
- `TestHarnessSourceContracts.Tests.ps1` guards the new native sandbox acquisition and rejects
  `GetTempPathW` / `GetTempFileNameW` in `Commands.SelfTest.ShellCommands.cpp`.
- Durable rule merged into `Specs\Testing\Testing_SelfTests.md`; coverage counts updated in
  `Specs\Testing\Testing_TestCoverage.md` and `Tests\README.md`.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter 'cmd_pane_goToShortcutOrLinkTarget_lnk_navigates_to_file_target,cmd_pane_goToShortcutOrLinkTarget_lnk_navigates_to_directory_target,cmd_pane_goToShortcutOrLinkTarget_broken_lnk_reports_alert,cmd_pane_itemProperties_show_shortcut_and_reparse_targets' -TimeoutMultiplier 1.0 -SkipLegacySandboxCleanup
```

Result: RED source-contract Pester first failed at 73 / 1 / 0 against the Win32 process-temp fallback;
GREEN `TestHarnessSourceContracts` 74 / 0 / 0, Debug `RedSalamander` build log
`.build\logs\msbuild-20260706_145453_462.log` had 0 warnings / 0 errors, and focused ShellCommands
shortcut-helper runtime proof passed 4 / 0 / 0 with run id
`20260706T125718Z-73056-35658bb24a43400da7d6788684b726bb`. The archive landed at
`Specs\TestRuns\4cb089111a23\Commands\2026-07-06_145720`.

### 2026-07-06 — Phase 1 PerformanceTests2 TestSandbox migration

Completed the standalone CppUnitTest perf-fixture migration from §8/R1-R3:

- `PerformanceTests2::AcquirePerformanceTestSandbox(...)` now resolves direct standalone launches
  from `REDSALAMANDER_TEST_ROOT` + `REDSALAMANDER_TEST_RUN_ID`, falling back to
  `<cwd>\.build\TestSandbox` or the sibling `.build\TestSandbox` when launched from
  `.build\<platform>\<configuration>`.
- `FolderIconEnumerationPerfTest`, `FolderIconEnumerationDuplicatePathPerfTest`, and
  `FolderViewRefreshDuplicatePathPerfTest` now create fixture roots under
  `runs\<runId>\scratch\performance-tests2\<case>` and no longer call
  `std::filesystem::temp_directory_path`.
- `TestHarnessSourceContracts.Tests.ps1` guards the helper shape, unified env names, the
  `runs\...\scratch\performance-tests2` layout, and the absence of raw temp-root calls in the
  migrated PerformanceTests2 sources.
- Durable rule merged into `Specs\Testing\Testing_SelfTests.md`; coverage counts updated in
  `Specs\Testing\Testing_TestCoverage.md` and `Tests\README.md`.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
.\build.ps1 -ProjectName PerformanceTests2 -Configuration Debug
$env:REDSALAMANDER_TEST_ROOT = (Resolve-Path '.\.build\TestSandbox').Path
$env:REDSALAMANDER_TEST_RUN_ID = 'performance-tests2-sandbox-proof-20260706-1508'
& 'C:\Program Files\Microsoft Visual Studio\18\Insiders\Common7\IDE\Extensions\TestPlatform\vstest.console.exe' '.\PerformanceTests2.dll' '/Platform:x64'
rg -n "std::filesystem::temp_directory_path|GetTempPathW|GetTempFileNameW" Tests RedSalamander\SelfTest
```

Result: RED source-contract Pester first failed at 74 / 1 / 0 against the missing standalone helper;
GREEN `TestHarnessSourceContracts` 75 / 0 / 0, Debug `PerformanceTests2` build log
`.build\logs\msbuild-20260706_150744_273.log` had 0 warnings / 0 errors, and direct VSTest runtime
proof passed 12 / 0 / 0 with run id `performance-tests2-sandbox-proof-20260706-1508`. The
`scratch\performance-tests2` parent existed with no child fixtures after cleanup. The raw temp-root
source scan for `std::filesystem::temp_directory_path`, `GetTempPathW`, and `GetTempFileNameW` under
`Tests` and `RedSalamander\SelfTest` returned no remaining matches.

### 2026-07-06 — Phase 1 BatchRename fixed-root TestSandbox migration

Completed another §8/R1-R3 non-raw-root migration slice:

- `Commands.SelfTest.BatchRename.cpp` now routes fixed BatchRename window fixture roots through
  `AcquireBatchRenameCommandsSandboxRoot(...)`, which calls
  `SelfTest::AcquireTestSandbox(SelfTest::SelfTestSuite::Commands, ...)`.
- The migrated cases no longer use fixed `C:\BatchRename*SelfTest` roots for pane/window contexts.
- The source guard rejects both the old literal pattern and direct `const std::filesystem::path root =
  L"C:\BatchRename..."` reintroduction in BatchRename window fixture tests.
- Durable rule merged into `Specs\Testing\Testing_SelfTests.md`; coverage counts updated in
  `Specs\Testing\Testing_TestCoverage.md` and `Tests\README.md`.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter 'cmd_pane_batchRename_opens_from_active_pane,cmd_pane_batchRename_window_rules_recompute_preview,cmd_pane_batchRename_window_preview_context_menu_copies_rows,cmd_pane_batchRename_preview_clipboard_honors_display_order,cmd_pane_batchRename_stale_generation_payloads_are_ignored,cmd_pane_batchRename_theme_accessibility_snapshot,cmd_pane_batchRename_window_rule_controls_drive_preview,cmd_pane_batchRename_window_debounces_text_preview,cmd_pane_batchRename_window_uses_and_persists_settings,cmd_pane_batchRename_window_manual_mode_controls_drive_preview,cmd_pane_batchRename_window_helper_buttons_insert_into_rule_fields,cmd_pane_batchRename_duplicate_source_target_refresh_consumes_distinct_rows,cmd_pane_batchRename_manual_sort_like_preview_with_hide_unchanged,cmd_pane_batchRename_path_sort_uses_displayed_path_and_stable_ties' -TimeoutMultiplier 1.0 -SkipLegacySandboxCleanup
```

Result: RED source-contract Pester first failed at 76 / 1 / 0 against the hardcoded
`C:\BatchRename*SelfTest` roots; GREEN `TestHarnessSourceContracts` 77 / 0 / 0 and
`TestInventory` 5 / 0 / 0, Debug `RedSalamander` build log
`.build\logs\msbuild-20260706_153105_584.log` had 0 warnings / 0 errors, and focused Commands
runtime proof passed 14 / 0 / 0 with run id
`20260706T133314Z-83440-414bdc55817245afa50b7b5d94b1fc91`, archived at
`Specs\TestRuns\4cb089111a23\Commands\2026-07-06_153316`.

### 2026-07-06 — Phase 1 ViewerPETests fixture TestSandbox migration

Completed another §8/R1-R3 standalone fixture-root migration slice:

- `ViewerPETests.cpp` now resolves `REDSALAMANDER_TEST_ROOT` and `REDSALAMANDER_TEST_RUN_ID` directly,
  falling back to `.build\TestSandbox` for direct launches.
- Runtime fixture files now live under
  `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\viewer-pe\<case>` instead of `ViewerWebTests`,
  `ViewerImgRawTests`, `ViewerImgRawPngTests`, `ViewerTextTests`, `ViewerSpaceTests`, or
  `ViewerVLCTests` folders in the build output directory.
- The source guard rejects reintroduction of those build-dir fixture roots and proves the standalone
  helper/walled-garden layout.
- Durable rule merged into `Specs\Testing\Testing_SelfTests.md`; coverage counts updated in
  `Specs\Testing\Testing_TestCoverage.md` and `Tests\README.md`.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
.\build.ps1 -ProjectName ViewerPETests -Configuration Debug
$env:REDSALAMANDER_TEST_ROOT=(Resolve-Path '.\.build\TestSandbox').Path; $env:REDSALAMANDER_TEST_RUN_ID='viewer-pe-sandbox-proof-20260706-1543'; try { .\.build\x64\Debug\ViewerPETests.exe TestViewerImgRawDecodesPngThroughWicWithoutErrorAlert } finally { Remove-Item Env:REDSALAMANDER_TEST_ROOT -ErrorAction SilentlyContinue; Remove-Item Env:REDSALAMANDER_TEST_RUN_ID -ErrorAction SilentlyContinue }
```

Result: RED source-contract Pester first failed at 77 / 1 / 0 against the missing
`AcquireViewerPETestSandbox(...)` helper and build-dir fixture roots; GREEN
`TestHarnessSourceContracts` passed 78 / 0 / 0 and `TestInventory` passed 5 / 0 / 0. Debug
`ViewerPETests` build log `.build\logs\msbuild-20260706_154301_566.log` had 0 warnings / 0 errors,
and focused runtime proof passed `TestViewerImgRawDecodesPngThroughWicWithoutErrorAlert` with
`REDSALAMANDER_TEST_RUN_ID=viewer-pe-sandbox-proof-20260706-1543`. The runtime case root
`.build\TestSandbox\runs\viewer-pe-sandbox-proof-20260706-1543\scratch\viewer-pe\viewerimgraw_png_decode`
existed with `child_count=0` after cleanup.

### 2026-07-06 — Phase 1 DxUiTests generated artifact default TestSandbox migration

Completed another §8/R1-R3 standalone generated-artifact migration slice:

- `DxUiTestHelpers.h` now resolves `REDSALAMANDER_TEST_ROOT` and
  `REDSALAMANDER_TEST_RUN_ID` directly for generated DxUi artifact defaults, falling back to
  `.build\TestSandbox` for direct launches.
- Default generated artifacts now live under
  `REDSALAMANDER_TEST_ROOT\runs\<runId>\artifacts\dxui` instead of `Specs\TestRuns\DxUiGallery`
  or `Specs\TestRuns\local_scratch`.
- Explicit caller-supplied paths remain caller-owned: `--gallery-output`,
  `--gallery-output-directory`, `--button-audit-output`, and `--perf-jsonl` still use the provided
  path.
- The source guard proves the ButtonContrast/control-gallery defaults and the Animation/WindowHost
  local perf JSONL defaults route through `GetDxUiTestArtifactPath(...)` and reject the old
  repo-local `Specs\TestRuns` defaults.
- Durable rule merged into `Specs\Testing\Testing_SelfTests.md`; coverage counts updated in
  `Specs\Testing\Testing_TestCoverage.md` and `Tests\README.md`.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
$env:REDSALAMANDER_TEST_ROOT=(Resolve-Path '.\.build\TestSandbox').Path; $env:REDSALAMANDER_TEST_RUN_ID='dxui-artifact-proof-20260706-1557'; try { .\.build\x64\Debug\DxUiTests.exe --suite=ButtonContrast; .\.build\x64\Debug\DxUiTests.exe --suite=Animation; .\.build\x64\Debug\DxUiTests.exe --suite=WindowHost } finally { Remove-Item Env:REDSALAMANDER_TEST_ROOT -ErrorAction SilentlyContinue; Remove-Item Env:REDSALAMANDER_TEST_RUN_ID -ErrorAction SilentlyContinue }
```

Result: RED source-contract Pester first failed at 78 / 1 / 0 against the missing
`GetDxUiTestArtifactPath(...)` helper and old `Specs\TestRuns` defaults; GREEN
`TestHarnessSourceContracts` passed 79 / 0 / 0. Debug `DxUiTests` build log
`.build\logs\msbuild-20260706_155547_780.log` had 0 warnings / 0 errors. Focused runtime proof
passed `DxUiTests.exe --suite=ButtonContrast`, `--suite=Animation`, and `--suite=WindowHost` with
`REDSALAMANDER_TEST_RUN_ID=dxui-artifact-proof-20260706-1557`; artifact root
`.build\TestSandbox\runs\dxui-artifact-proof-20260706-1557\artifacts\dxui` contained
`DxUiButtonContrast.png` (380558 bytes), `dxui_animation_scheduler_testlocal.jsonl` (85902 bytes),
and `dxui_windowhost_stage_metrics_testlocal.jsonl` (6288 bytes).

### 2026-07-06 — Phase 1 foreground search-service JobObject isolation

Completed the orphan-process part of §8/R7 and checklist item 14 for foreground Compare selftests:

- `ForegroundSearchServiceProcess` now owns a `wil::unique_handle _job`.
- `CreateKillOnCloseJob(...)` creates a JobObject and sets
  `JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE`.
- The foreground `RedSalamanderSearchService.exe --run-foreground` child starts suspended with
  `CREATE_SUSPENDED`, is assigned with `AssignProcessToJobObject`, and resumes only after assignment
  succeeds.
- Failure before resume terminates/waits the child, releases the JobObject, and restores the
  temporary client retry override.
- The focused trace records `Foreground search service JobObject assigned.` so runtime artifacts prove
  the isolation path executed.
- Durable rule merged into `Specs\Testing\Testing_SelfTests.md`; source guard added to
  `Tools\Tests\TestHarnessSourceContracts.Tests.ps1`; coverage counts updated in
  `Specs\Testing\Testing_TestCoverage.md` and `Tests\README.md`.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
$env:REDSALAMANDER_TEST_ROOT='C:\Users\eric\AppData\Local\Temp\RedSalamander-TestSandbox'; try { .\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter search_service_foreground_logs_request_status -TimeoutMultiplier 0.2 -SkipLegacySandboxCleanup } finally { Remove-Item Env:REDSALAMANDER_TEST_ROOT -ErrorAction SilentlyContinue }
```

Result: `TestHarnessSourceContracts` 66 / 0 / 0, `TestInventory` 5 / 0 / 0, Debug build log
`.build\logs\msbuild-20260706_133722_724.log` had 0 warnings / 0 errors, and focused Compare runtime
proof passed 1 / 0 / 0 with `run_id=20260706T113931Z-59180-53d9685f252d40f3b4f39defbd1c2f6c`.
The trace recorded
`Foreground search service JobObject assigned.` The archive path was
`Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-06_133931`.

### 2026-07-06 — Phase 2 DxUi focus/foreground interactive-desktop gating

Completed the real-Win32 focus/foreground slice from the Phase 2 audit:

- `DxUiTestHelpers.h` now exposes `SkipDxUiTest`, `WaitForDxUiThreadFocus`,
  `TryFocusDxUiTestWindow`, and `TryActivateDxUiTestWindow`.
- `DxUiTests.NativeTextInput.cpp` uses the same-thread focus probe before the
  host-focus and system-caret assertions that call `GetFocus()` and
  `GetGUIThreadInfo(...)`.
- `DxUiTests.Menu.cpp` uses the foreground activation probe before
  foreground-owned popup/menu focus assertions and removed direct discarded
  `SetForegroundWindow(ownerWindow.Hwnd())` calls from test bodies.
- Foreground-only paths still assert the real product behavior when the desktop
  grants focus/foreground capability; when that environmental precondition is
  absent, the case emits an explicit `SKIPPED:` reason instead of reporting a
  false regression.
- Durable rule merged into `Specs\Testing\Testing_SelfTests.md`; source guard
  added to `Tools\Tests\TestHarnessSourceContracts.Tests.ps1`; coverage counts
  updated in `Specs\Testing\Testing_TestCoverage.md` and `Tests\README.md`.

Verification:

```powershell
Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru
.\build.ps1 -ProjectName DxUiTests -Configuration Debug
.\.build\x64\Debug\DxUiTests.exe --suite=NativeTextInput
.\.build\x64\Debug\DxUiTests.exe --suite=Menu
```

Result: RED source-contract Pester first failed at 75 / 1 / 0 against the
missing focused helper split; GREEN `TestHarnessSourceContracts` 76 / 0 / 0 and
`TestInventory` 5 / 0 / 0, Debug `DxUiTests` build log
`.build\logs\msbuild-20260706_152402_723.log` had 0 warnings / 0 errors,
`NativeTextInput` passed with no focus-gate skips in this session, and `Menu`
passed while foreground-only cases emitted explicit `SKIPPED:` reasons in this
background desktop session.

### 2026-07-06 — Phase 0 CI runner routing

Completed the first CI-routing slice from §4:

- `Tools/TestRunPlan.ps1` now defines `-Suite CI` as the GitHub Actions PR gate lane, separate from
  the broader `-Suite Full` closeout lane.
- `-Suite CI` runs the three in-product self-test suites as separate processes, splits DxUiTests by
  suite to stop Menu/NativeTextInput focus state sharing one process, preserves the explicit
  ViewerPE find/goto prompt invocations, runs PerformanceTests2 through the existing VSTest resolver,
  keeps artifact-only Pester behavior with `-ExcludeTag RequiresBuildToolchain`, and adds the
  deterministic FileSystemCurlTests + RedConfigureTests gates.
- `.github/workflows/ci.yml` now sets
  `REDSALAMANDER_TEST_ROOT=${{ github.workspace }}\.build\TestSandbox`, invokes
  `.\Tools\Run-AllTests.ps1 -Suite CI -SkipBuild -TimeoutMultiplier 2.0`, and uploads the sandbox
  tree instead of `%LOCALAPPDATA%\RedSalamander\SelfTest`.
- The sandbox resolver rejects unrelated `REDSALAMANDER_SELFTEST_ROOT` conflicts but accepts the
  runner-owned per-run legacy bridge under `REDSALAMANDER_TEST_ROOT\runs\<runId>\...`, so the
  runner can safely execute Pester tests inside its own paired environment variables.
- PluginContractTests, SettingsSchemaTests, CrashHandlingTests, and RedSalamanderMonitorEtwLatency
  are deliberately kept in `-Suite Full` only for this slice, with the workflow comment requiring a
  future CI adapter/quarantine policy before they enter the PR gate.
- Specs/docs updated: `Specs/Testing/Testing_SelfTests.md`,
  `Specs/Testing/Testing_TestCoverage.md`, `README.md`, and `Tests/README.md`.

Verification:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru"
```

Result: 11 passed / 0 failed.

Also verified the same focused Pester suite with both `REDSALAMANDER_TEST_ROOT` and the runner-owned
per-run `REDSALAMANDER_SELFTEST_ROOT` bridge set; result remained 11 passed / 0 failed.

Still open from Track A: partial/crash result preservation, shuffle/order triage, and the full
native `TestSandbox` migration/reaper/lint work from §8.

### 2026-07-06 — Phase 0 blocking failure classifier

Completed the first §2c runner classifier slice:

- `Tools/TestRunPlan.ps1` now exposes `Add-RSTestResultClassification`, preserving
  `classification`, `classification_reason`, and retry evidence in `run-all-tests-results.json`.
- `Run-AllTests.ps1 -Suite CI` enables failure classification automatically; non-CI suites can opt
  in with `-ClassifyFailures`.
- Failed standalone entries rerun once at entry level. Pass-on-rerun is labeled `FLAKY`, fail-again
  is labeled `REGRESSION`; both keep the aggregate exit code non-zero.
- Failed broad in-product self-test cases rerun via `--selftest-case=<failedCase>` when possible.
  If they pass in isolation, the runner labels the suite `ISOLATION_SUSPECT` instead of falsely
  calling it flaky; the later shuffle-triage checkpoint below adds the final classification step.
- Runner summaries include top-level classification counts for `flaky`, `regression`,
  `isolation_suspect`, and `unclassified_failure`.
- Specs/docs updated: `Specs/Testing/Testing_SelfTests.md`,
  `Specs/Testing/Testing_TestCoverage.md`, `README.md`, and `Tests/README.md`.

Verification:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru"
```

Result: 14 passed / 0 failed.

Also parsed `Tools\Run-AllTests.ps1` and `Tools\TestRunPlan.ps1` with
`System.Management.Automation.Language.Parser`; both reported zero parse errors.

The later shuffle-triage classifier, injected runtime proof, and per-case history/dashboard
checkpoints below close the first runner-side consumption and artifact slices.

### 2026-07-06 — Phase 0 reviewed quarantine metadata gate

Completed the first §2d quarantine slice:

- Added `Tools/test-quarantine.jsonl` as the only runner-recognized flaky-test quarantine ledger.
  Blank lines mean no active quarantine.
- `Tools/TestRunPlan.ps1` now validates quarantine entries with required `harness`, `name`,
  `owner`, `opened`, `expires`, `issue`, `root_cause_hypothesis`, and `fix_or_replace_plan`
  fields. Dates must use `YYYY-MM-DD`, entries cannot be expired, and the maximum quarantine
  window is 30 days so "temporary" cannot become permanent by omission.
- `Run-AllTests.ps1` reads the quarantine file by default, records active/invalid entries in
  `run-all-tests-results.json`, prints quarantine blockers, and keeps the aggregate run red while
  any active or invalid entry exists.
- Specs/docs updated: `Specs/Testing/Testing_SelfTests.md`,
  `Specs/Testing/Testing_TestCoverage.md`, `README.md`, and `Tests/README.md`.

Verification:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru"
```

Result: 16 passed / 0 failed.

### 2026-07-06 — Phase 0 quarantine repair lane

Completed the §2d repair-lane execution slice:

- `Tools/TestRunPlan.ps1` now builds a quarantine repair plan from active
  `Tools/test-quarantine.jsonl` entries and the actual `Run-AllTests.ps1` test plan. Entries whose
  `harness` does not match a runner adapter, or whose self-test `name` is absent from the adapter's
  case list, are invalid and remain blocking.
- Matching in-product self-test entries run as single-case repair attempts by preserving the suite
  flag and adding `--selftest-case=<name>`; matching standalone harnesses rerun their adapter entry.
- `Run-AllTests.ps1` executes the repair lane after the main gate, continues collecting evidence for
  every reviewed quarantine entry, and still keeps the aggregate run red while entries exist or are
  invalid.
- `run-all-tests-results.json` now preserves `quarantine.repair_attempts`,
  `quarantine.repair_attempt_count`, and `quarantine.repair_reproduced_count`.
- Specs/docs updated: `Specs/Testing/Testing_SelfTests.md`, `README.md`, `Tests/README.md`, and
  `Specs/Testing/Testing_TestCoverage.md`.

Verification:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru"
```

Result: 19 passed / 0 failed.

### 2026-07-06 — Phase 0 GitHub summary publication

Completed the §2d/CI summary publication slice:

- `Tools/TestRunPlan.ps1` now formats a GitHub Markdown step summary from
  `run-all-tests-results.json` data, including suite status, blocking classification counts, active
  quarantine owner/expiry/issue rows, and quarantine repair-lane pass/fail evidence.
- `Run-AllTests.ps1` appends that summary to `GITHUB_STEP_SUMMARY` when GitHub Actions provides the
  summary file, without masking the underlying test result if summary writing fails.
- Specs/docs updated: `Specs/Testing/Testing_SelfTests.md`,
  `Specs/Testing/Testing_TestCoverage.md`, `README.md`, and `Tests/README.md`.

Verification:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru"
```

Result: 20 passed / 0 failed.

Still open from §2d: require quarantine removal plus fix-or-replace evidence before closeout, and add
true shuffle/repeat exit evidence before declaring a previously flaky case fixed.

### 2026-07-06 — Phase 0 Commands repeat/shuffle detector

Completed the first §2b repeat/shuffle slice:

- Native self-test options now include `--selftest-repeat=N` and `--selftest-shuffle=SEED`.
  Repeat is shared through `SelfTest::RunCase` and emits one case result per attempt with
  `repeat_index`; suite and aggregate JSON record `repeat_count` and `shuffle_seed`.
- Commands now uses `SelfTest::BuildSelfTestCaseExecutionOrder(...)` for explicit seeded ordering,
  then replays the existing case registrations one exact case at a time. This gives real seed-order
  evidence without rewriting the Commands case declarations.
- `Tools\Run-AllTests.ps1` exposes `-SelfTestRepeat` and `-SelfTestShuffleSeed`. Repeat forwards to
  all in-product self-test suites: Commands and CompareDirectories repeat through
  `SelfTest::RunCase`, and FileOperations expands its phase-state run plan per repeat attempt so
  repeated results are real native `repeat_index` rows. Shuffle is forwarded only to Commands until
  CompareDirectories and FileOperations migrate to explicit ordering. Native non-Commands shuffle
  requests fail fast so a fixed-order run cannot masquerade as shuffled evidence.
- Runner result-coverage validation now accepts repeated actual case names only when each expected
  case appears exactly the requested repeat count.
- Runner-native `--selftest-list-cases` now pushes `writeJsonSummary=false` into the shared self-test
  options before building inventory JSON, so post-run case-list validation cannot overwrite the real
  suite `results.json`.
- `Tools\TestRunPlan.ps1` now formats fractional timeout multipliers with invariant `.` decimals;
  the first smoke attempt exposed the prior culture-sensitive `0,1` argument and native rejected it
  with exit code 2.
- Specs/docs updated: `Specs/Testing/Testing_SelfTests.md`,
  `Specs/Testing/Testing_TestCoverage.md`, `README.md`, and `Tests/README.md`.

Verification:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru"
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru"
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter modeless_window_ownership -SelfTestRepeat 2 -SelfTestShuffleSeed 123 -TimeoutMultiplier 0.1
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter FileOps_ProviderCapabilityMatrix -SelfTestRepeat 2 -TimeoutMultiplier 0.1
```

Results:

- GREEN `RunAllTestsPlan.Tests.ps1`: 23 passed / 0 failed.
- GREEN `TestHarnessSourceContracts.Tests.ps1`: 56 passed / 0 failed.
- GREEN Debug build: `.build\logs\msbuild-20260706_113753_405.log`, 0 warnings / 0 errors.
- GREEN Commands runtime smoke:
  `.build\TestSandbox\runs\20260706T094019Z-77120-cb4f59ca0800460c869afadf964b779e\artifacts\selftest\last_run`
  reports `modeless_window_ownership` twice, `repeat_index` values `1,2`, `repeat_count=2`,
  `shuffle_seed=123`, runner exit code `0`.
- GREEN FileOperations runtime smoke:
  `.build\TestSandbox\runs\20260706T094031Z-75960-afbfda54b95e4049a18c1a586671c147\artifacts\selftest\last_run`
  reports `Setup`, `FileOps_ProviderCapabilityMatrix`, and `Cleanup_RestorePluginConfig` for
  `repeat_index` values `1,1,1,2,2,2`, `repeat_count=2`, runner exit code `0`.

At this point FileOperations still needed explicit seeded order; the later FileOperations checkpoint
below closes that migration. The later shuffle-triage and nightly workflow checkpoints close the
remaining §2b/§2c runner consumption and scheduling work.

### 2026-07-06 — Phase 0 CompareDirectories repeat/shuffle detector

Completed the CompareDirectories §2b explicit-order slice:

- `Tools\TestRunPlan.ps1` now forwards `--selftest-shuffle=SEED` to both Commands and
  CompareDirectories. FileOperations forwarding was completed in the following checkpoint.
- Native argument validation now accepts `--selftest-shuffle` for `--compare-selftest`.
- CompareDirectories explicit order is orchestrated in `RedSalamander.cpp`, outside the fragile
  include-fragment case file chain: the runner builds `SelfTest::BuildSelfTestCaseExecutionOrder`,
  invokes `CompareDirectoriesSelfTest::Run(...)` once per exact case with `repeatCount=1`, suppresses
  per-case JSON writes, and emits one aggregate `compare\results.json` preserving `repeat_index`.
- The shared repeated-suite aggregator now preserves `(case name, repeat_index)` for both
  FileOperations phase-state repeats and CompareDirectories explicit-order repeats.

Verification:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru"
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru"
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter showIdentical -SelfTestRepeat 2 -SelfTestShuffleSeed 456 -TimeoutMultiplier 0.1
```

Results:

- GREEN `RunAllTestsPlan.Tests.ps1`: 23 passed / 0 failed.
- GREEN `TestHarnessSourceContracts.Tests.ps1`: 57 passed / 0 failed.
- GREEN Debug build: `.build\logs\msbuild-20260706_114750_564.log`, 0 warnings / 0 errors.
- GREEN CompareDirectories runtime smoke:
  `.build\TestSandbox\runs\20260706T094952Z-52548-5efbb28842ce47708551b52945b161da\artifacts\selftest\last_run`
  reports `showIdentical` twice, `repeat_index` values `1,2`, `repeat_count=2`, and
  `shuffle_seed=456` in both suite and aggregate JSON.

### 2026-07-06 — Phase 0 FileOperations repeat/shuffle detector

Completed the FileOperations §2b explicit-order slice:

- `Tools\TestRunPlan.ps1` now forwards `--selftest-shuffle=SEED` to FileOperations as well as
  Commands and CompareDirectories.
- Native argument validation now accepts `--selftest-shuffle` for `--fileops-selftest`.
- FileOperations keeps its existing family/phase-state repeat behavior when no seed is supplied. When
  a seed is supplied, `RedSalamander.cpp` expands the selected FileOperations phase set into
  individual phase filters, shuffles those phase names through `SelfTest::BuildSelfTestCaseExecutionOrder`,
  runs each phase with `repeatCount=1`, and emits one aggregate `fileops\results.json` preserving
  `Setup`, `Cleanup_RestorePluginConfig`, `repeat_index`, `repeat_count`, and `shuffle_seed`.

Verification:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru"
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru"
pwsh -NoProfile -ExecutionPolicy Bypass -Command "& { $parseErrors = $null; [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path '.\Tools\Run-AllTests.ps1'), [ref]$null, [ref]$parseErrors) > $null; if ($parseErrors.Count -gt 0) { $parseErrors | ForEach-Object { $_.ToString() }; exit 1 }; [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path '.\Tools\TestRunPlan.ps1'), [ref]$null, [ref]$parseErrors) > $null; if ($parseErrors.Count -gt 0) { $parseErrors | ForEach-Object { $_.ToString() }; exit 1 }; 'PowerShell parse OK' }"
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter FileOps_ProviderCapabilityMatrix -SelfTestRepeat 2 -SelfTestShuffleSeed 789 -TimeoutMultiplier 0.1
```

Results:

- GREEN `RunAllTestsPlan.Tests.ps1`: 23 passed / 0 failed.
- GREEN `TestHarnessSourceContracts.Tests.ps1`: 57 passed / 0 failed.
- GREEN PowerShell parser check for `Tools\Run-AllTests.ps1` and `Tools\TestRunPlan.ps1`.
- GREEN Debug build: `.build\logs\msbuild-20260706_120101_517.log`, 0 warnings / 0 errors.
- GREEN FileOperations runtime smoke:
  `.build\TestSandbox\runs\20260706T100312Z-74672-1393b09f8812458ba12f7890e2e352ca\artifacts\selftest\last_run`
  reports `Setup`, `FileOps_ProviderCapabilityMatrix`, and `Cleanup_RestorePluginConfig` for
  `repeat_index` values `1,1,1,2,2,2`, with `repeat_count=2`, `shuffle_seed=789`, and trace line
  `FileOpsSelfTest: explicit execution order count=2 repeat=2 shuffleSeed=789`.

Still open from §2b at this checkpoint: add the nightly shuffle+repeat lane. The later nightly
workflow checkpoint below closes it.

### 2026-07-06 — Phase 0 shuffle-triage classifier consumption

Completed the first §2c shuffle-consumption slice:

- `Add-RSTestResultClassification` now treats `shuffle-triage` retry attempts as first-class
  classification evidence. Broad self-test failures whose failed cases pass in isolation stay
  `ISOLATION_SUSPECT` until shuffle triage exists; pass-all-shuffle evidence becomes blocking
  `FLAKY`; fail-any-shuffle evidence becomes blocking `REGRESSION` with an isolation/order reason.
- `Run-AllTests.ps1` now follows successful failed-case reruns with three broad-suite shuffled
  probes. It preserves any original `--selftest-shuffle` seed and adds fresh seeds until it has
  three triage attempts.
- `run-all-tests-results.json` now preserves `retry_attempts[].shuffle_seed` so the exact
  reproducing seed is available from aggregate evidence.
- Specs/docs updated: `Specs/Testing/Testing_SelfTests.md`,
  `Specs/Testing/Testing_TestCoverage.md`, `README.md`, and `Tests/README.md`.

Verification:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru"
```

Results:

- RED `RunAllTestsPlan.Tests.ps1`: 23 passed / 2 failed against missing shuffle-specific
  classification reasons and missing `shuffle_seed` summary preservation.
- GREEN `RunAllTestsPlan.Tests.ps1`: 25 passed / 0 failed.

The runtime injected flaky/order-dependent classifier proof checkpoint below closes the remaining
§2c proof gap. The per-case history/dashboard checkpoint below closes the first §2f artifact slice.

### 2026-07-06 — Phase 0 runtime injected classifier proof

Completed the remaining §2c runtime-proof slice:

- `Run-AllTests.ps1` now forwards explicit debug proof hooks,
  `-SelfTestFlakyProofCase` and `-SelfTestOrderProofCase`, to in-product self-test suites.
- Native debug selftests expose `--selftest-flaky-proof-case=NAME` and
  `--selftest-order-proof-case=NAME`. The flaky hook fails the named case in suite context only;
  exact-case and shuffle-triage retries pass so the runner must label the aggregate result
  blocking `FLAKY`. The order hook fails in suite/shuffle context but passes exact-case isolation,
  so shuffle triage must label it blocking `REGRESSION` instead of `FLAKY`.
- Focused `-CaseFilter` classification no longer falls back to whole-entry retry. It runs the exact
  failed case first, then preserves the original focused subset for three shuffle-triage attempts.
- Explicit-order native dispatchers now carry classifier proof suite/shuffle context through
  Commands, CompareDirectories, and FileOperations, so a shuffled per-case replay is not mistaken
  for an isolated exact-case retry.
- Specs/docs updated: `Specs/Testing/Testing_SelfTests.md`,
  `Specs/Testing/Testing_TestCoverage.md`, `README.md`, and `Tests/README.md`.

Verification:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru"
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru"
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\TestInventory.Tests.ps1 -PassThru"
pwsh -NoProfile -ExecutionPolicy Bypass -Command '& { $parseErrors = $null; [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path ''.\Tools\Run-AllTests.ps1''), [ref]$null, [ref]$parseErrors) > $null; if ($parseErrors.Count -gt 0) { $parseErrors | ForEach-Object { $_.ToString() }; exit 1 }; [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path ''.\Tools\TestRunPlan.ps1''), [ref]$null, [ref]$parseErrors) > $null; if ($parseErrors.Count -gt 0) { $parseErrors | ForEach-Object { $_.ToString() }; exit 1 }; ''PowerShell parse OK'' }'
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter 'unique,showIdentical' -ClassifyFailures -SelfTestFlakyProofCase showIdentical -TimeoutMultiplier 0.1
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter 'unique,showIdentical' -ClassifyFailures -SelfTestOrderProofCase showIdentical -TimeoutMultiplier 0.1
```

Results:

- RED `TestHarnessSourceContracts.Tests.ps1`: 58 passed / 1 failed against missing
  classifier proof suite/shuffle context in native explicit-order dispatch.
- GREEN `TestHarnessSourceContracts.Tests.ps1`: 59 passed / 0 failed.
- GREEN `RunAllTestsPlan.Tests.ps1`: 28 passed / 0 failed.
- GREEN `TestInventory.Tests.ps1`: 5 passed / 0 failed.
- GREEN PowerShell parser check for `Tools\Run-AllTests.ps1` and `Tools\TestRunPlan.ps1`.
- GREEN Debug build: `.build\logs\msbuild-20260706_123922_046.log`, 0 warnings / 0 errors.
- GREEN flaky proof: runner exited `1` as expected and classified
  `20260706T104131Z-85784-c110c5635f9b4b9886e8f2432b8d4f1b` as blocking `FLAKY`
  with retry modes `failed-case,shuffle-triage,shuffle-triage,shuffle-triage`,
  shuffle seeds `885711464,370821081,1015852695`, and classification counts
  `flaky=1 regression=0 isolation=0`.
- GREEN order-dependent proof: runner exited `1` as expected and classified
  `20260706T104138Z-85864-cecd27054aed4e3387c8d36c284b1f50` as blocking
  `REGRESSION` with retry modes `failed-case,shuffle-triage,shuffle-triage,shuffle-triage`,
  shuffle seeds `1404560201,1737196888,1805700883`, shuffle exit codes `1,1,1`,
  and classification counts `flaky=0 regression=1 isolation=0`.

### 2026-07-06 — Phase 0 per-case history/dashboard artifact

Completed the first §2f history slice:

- `New-RSTestRunSummary` now preserves suite `shuffle_seed`, `repeat_count`, and per-case `attempt`
  data when available.
- `Tools\TestRunPlan.ps1` now derives `run-all-tests-case-history.jsonl` rows from aggregate
  results, including `harness`, `case`, `duration_ms`, `status`, `reason`, `classification`, `seed`,
  `attempt`, and `source` (`case` or `retry`).
- `Run-AllTests.ps1` writes both `run-all-tests-case-history.jsonl` and
  `run-all-tests-dashboard.md` beside `run-all-tests-results.json`. The dashboard lists the slowest
  rows plus failure/retry evidence, including shuffle-triage seeds.
- Specs/docs updated: `Specs/Testing/Testing_SelfTests.md`,
  `Specs/Testing/Testing_TestCoverage.md`, `README.md`, and `Tests/README.md`.

Verification:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru"
```

Results:

- RED `RunAllTestsPlan.Tests.ps1`: 25 passed / 1 failed against missing history/dashboard helpers.
- GREEN `RunAllTestsPlan.Tests.ps1`: 26 passed / 0 failed.
- GREEN FileOperations runtime smoke:
  `.build\TestSandbox\runs\20260706T102157Z-86268-9edd29d4ec87435f88502caf93daf214\artifacts\selftest\last_run`
  produced `run-all-tests-case-history.jsonl` with three `red-salamander.case-history.v1` rows
  including `reason`, and `run-all-tests-dashboard.md` with the slow-case table.

### 2026-07-06 — Phase 0 nightly shuffle+repeat workflow

Completed the §2b scheduled-lane slice:

- Added `.github/workflows/nightly-flake.yml` as a separate scheduled/manual workflow so the
  expensive flake detector does not inflate the PR gate.
- The workflow reuses the Debug self-test build artifact, stages themes, computes a run seed from
  `GITHUB_RUN_ID` (or UTC seconds for local/manual fallback), then runs
  `.\Tools\Run-AllTests.ps1 -Suite All -SkipBuild -TimeoutMultiplier 2.0 -SelfTestRepeat 5
  -SelfTestShuffleSeed $seed -ClassifyFailures`.
- The nightly lane uploads the runner-owned `REDSALAMANDER_TEST_ROOT` tree as
  `selftest-artifacts-nightly-shuffle`.
- `Tools\Tests\TestHarnessSourceContracts.Tests.ps1` now guards that the nightly lane carries
  repeat/shuffle/classification flags and that `.github/workflows/ci.yml` does not quietly absorb
  the expensive `-SelfTestRepeat 5` lane.
- Specs/docs updated: `Specs/Testing/Testing_SelfTests.md`,
  `Specs/Testing/Testing_TestCoverage.md`, `README.md`, and `Tests/README.md`.

Verification:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru"
```

Results:

- RED `TestHarnessSourceContracts.Tests.ps1`: 57 passed / 1 failed against missing
  `.github\workflows\nightly-flake.yml`.
- GREEN `TestHarnessSourceContracts.Tests.ps1`: 58 passed / 0 failed.

### 2026-07-06 — Phase 0 fatal-error modal self-test bypass

Completed the fatal-modal CI-hang prevention slice:

- `ShowFatalErrorDialog(...)` now checks `IsRunningAnySelfTest()` under `ENABLE_TESTS`, appends a
  self-test trace row, writes the caption/message to debug output, and returns before constructing
  the modal dialog. The SEH caller still returns non-zero, but headless CI no longer waits on a
  `ShowModal()` message loop after a self-test fatal path.
- `Tools\Tests\TestHarnessSourceContracts.Tests.ps1` now has a source contract proving the self-test
  guard exists, traces the suppression, and appears before `dialog.ShowModal()`.
- Specs/docs updated: `Specs/Testing/Testing_SelfTests.md`,
  `Specs/Testing/Testing_TestCoverage.md`, and `Tests/README.md`.

Verification:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru"
```

Result: 53 passed / 0 failed.

### 2026-07-06 — Phase 0 partial crash-result preservation

Completed the native/runner implementation slice for the crash-signal item:

- `SelfTest::RunCase` now records the current in-flight suite/case while a case body and its result
  emission are active.
- `SelfTest::AppendCaseResult` now flushes the suite `results.json` after every recorded case, so a
  later timeout/crash preserves prior case status evidence.
- `SelfTestCaseResult::Status::crashed` is a first-class JSON status. Crashed cases count against the
  suite failure total, and FileOps aggregation treats `crashed` as stronger than ordinary `failed`.
- `wWinMain`'s SEH path stays C++-object-free and calls `RecordSelfTestUnhandledExceptionCrash(...)`,
  which marks the in-flight case `crashed`, writes partial `last_run\results.json`, archives the
  partial run best-effort, releases the self-test mutex, and still returns non-zero through the fatal
  path.
- `--selftest-crash-case=NAME` now raises `EXCEPTION_ACCESS_VIOLATION` as soon as the exact matching
  case starts, giving the plan a deterministic proof hook for crash-signal preservation.
- `Tools\Run-AllTests.ps1` now colors `crashed` red and includes it in failed-case evidence.
- Specs/docs updated: `Specs/Testing/Testing_SelfTests.md`,
  `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`, and `README.md`.

Verification:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru"
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru"
pwsh -NoProfile -ExecutionPolicy Bypass -Command '& { $parseErrors = $null; [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path ''.\Tools\Run-AllTests.ps1''), [ref]$null, [ref]$parseErrors) > $null; if ($parseErrors.Count -gt 0) { $parseErrors | ForEach-Object { $_.ToString() }; exit 1 }; ''Tools\Run-AllTests.ps1 parse OK'' }'
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
```

Results:

- RED Pester first failed at 53 passed / 1 failed against the missing in-flight/crash hooks.
- GREEN `TestHarnessSourceContracts.Tests.ps1`: 54 passed / 0 failed.
- GREEN `RunAllTestsPlan.Tests.ps1`: 20 passed / 0 failed.
- GREEN parser check: `Tools\Run-AllTests.ps1 parse OK`.
- GREEN Debug build: `.build\logs\msbuild-20260706_105446_011.log`, 0 warnings / 0 errors.
- GREEN injected AV proof: `RedSalamander.exe --commands-selftest
  --selftest-case=modeless_window_ownership --selftest-crash-case=modeless_window_ownership
  --selftest-timeout-multiplier=0.1` exited `-1` in 915 ms, wrote partial
  `.build\TestSandbox\runs\crashproof_20260706_110550\artifacts\selftest\last_run\results.json`,
  and archived `Specs\TestRuns\4cb089111a23\Commands\2026-07-06_110551\selftest_run_results.json`
  with `modeless_window_ownership` status `crashed` and reason
  `crashed: unhandled exception code=0xC0000005 name=Access Violation`.

### 2026-07-06 — Phase 0 legacy TestSandbox cleanup reaper

Completed the §8/R5 legacy reaper slice:

- Added `Tools\Clean-TestSandbox.ps1` as the reviewed cleanup entrypoint. Direct invocation is
  dry-run by default; removal requires `-Apply`, still flows through `ShouldProcess`, and uses
  `Remove-Item -LiteralPath` only after wildcard plans resolve to concrete directories.
- `Tools\TestRunPlan.ps1` now exposes `Get-RSTestSandboxLegacyCleanupPlan`,
  `Resolve-RSTestSandboxCleanupTargets`, and `Get-RSFixedDriveRoots`. The cleanup plan covers
  `%LOCALAPPDATA%\RedSalamander\SelfTest`, known `%TEMP%` standalone/perf roots
  (`RedConfigure*`, `CrashHandlingTests-*`, `ViewerSqliteTests*`, ViewerText/ViewerImgRaw/ViewerWeb
  plugin temp roots, and PerformanceTests2 temp roots), and
  `<FixedDrive>:\RedSalamanderCrossVolumeSelfTest_*`.
- `Run-AllTests.ps1` now invokes `Tools\Clean-TestSandbox.ps1 -Apply -Confirm:$false` before the
  build step and before any child tests launch. `-SkipLegacySandboxCleanup` exists only for
  diagnostic runs.
- Specs/docs updated: `Specs\Testing\Testing_SelfTests.md`,
  `Specs\Testing\Testing_TestCoverage.md`, `README.md`, and `Tests\README.md`.

Verification:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\RunAllTestsPlan.Tests.ps1 -PassThru"
pwsh -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Pester -Path .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru"
.\Tools\Clean-TestSandbox.ps1 -LocalAppDataRoot 'Z:\nonexistent-localappdata-for-plan-test' -TempRoot 'Z:\nonexistent-temp-for-plan-test' -DriveRoots @('Z:\')
pwsh -NoProfile -ExecutionPolicy Bypass -Command '& { $parseErrors = $null; foreach ($path in @(''.\Tools\Run-AllTests.ps1'', ''.\Tools\TestRunPlan.ps1'', ''.\Tools\Clean-TestSandbox.ps1'')) { [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path $path), [ref]$null, [ref]$parseErrors) > $null; if ($parseErrors.Count -gt 0) { $parseErrors | ForEach-Object { $_.ToString() }; exit 1 } }; ''PowerShell parse OK'' }'
```

Results:

- RED `RunAllTestsPlan.Tests.ps1`: 28 passed / 1 failed against missing
  `Get-RSTestSandboxLegacyCleanupPlan`.
- RED `TestHarnessSourceContracts.Tests.ps1`: 59 passed / 1 failed against missing
  `Tools\Clean-TestSandbox.ps1`, then 60 passed / 1 failed against missing runner invocation.
- GREEN `RunAllTestsPlan.Tests.ps1`: 29 passed / 0 failed.
- GREEN `TestHarnessSourceContracts.Tests.ps1`: 61 passed / 0 failed.
- GREEN script dry-run: printed the dry-run message and removed nothing for nonexistent controlled
  roots.
- GREEN PowerShell parser check for `Tools\Run-AllTests.ps1`, `Tools\TestRunPlan.ps1`, and
  `Tools\Clean-TestSandbox.ps1`.

---

## What is the problem?

The suite never converges because **nondeterminism is generated by a handful of independent
mechanisms, the runner cannot tell a flaky failure from a real regression, and the flake
tooling that would fix that only covers a fraction of the gate.** Every red run looks like a
regression, gets "fixed," and the next run reds a *different* case — a fix → re-run →
new-failure loop that burns days of 40/50-minute CI cycles with no data to break it.

**Case order is NOT a nondeterminism source** — `RunCase` (`SelfTestCommon.h:205-280`) calls
cases sequentially with no shuffle/sort/random. Cross-run variation originates in four true
generators, which the missing isolation then *amplifies*:

1. **Focus/foreground asserted in a headless environment (the largest CI-side generator).**
   Dozens of `state.Require(GetFocus()==folderView,...)`
   (`Commands.SelfTest.Navigation.cpp:2694,3055,3153,3541,4081,4260,4402`), DxUiTests Menu
   `SetForegroundWindow`-with-discarded-return + capture-based popups (`DxUiTests.Menu.cpp:1174`),
   and NativeTextInput caret / `GetGUIThreadInfo().hwndCaret==host` checks on an off-screen
   never-activated window (`DxUiTests.NativeTextInput.cpp:677,735-741`; caret gated on
   `GetFocus()==_hwnd` at `DxUi.NativeTextInput.cpp:1643`) only hold when the process owns the
   interactive foreground desktop. On GitHub `windows-latest`, `SetFocus`/`SetForegroundWindow`
   silently no-op; *which* focus assert trips first depends on ambient foreground timing — the
   classic shifting failure set. A foreground steal during any `PumpMessages` injects
   `WM_KILLFOCUS`/`WM_ACTIVATEAPP` that tears down state between setup and assertion.

2. **Async/teardown races and crashes that destroy signal.** Live concurrency counters read
   across the worker/UI boundary under the wrong mutex or none (`_perItemMaxConcurrency` written
   unlocked, `State.cpp:6413`; readers take `_progressMutex`, `Phases07_09.cpp:3725`); a
   process-global 16-thread delta assertion (`Phases07_09.cpp:3765`) sensitive to unrelated
   thread churn; and Phase14 tearing down `FileOperationState` (destroying popup HWNDs) from a
   threadpool thread while the UI pump still dispatches completion messages
   (`FolderWindow.FileOperations.SelfTest.Phases14_16.cpp:146`), which can crash the whole process.

3. **A small set of genuinely fragile wall-clock ceilings.** The scaled deadline machinery
   (`Scale()`/`ScaleTimeout()`) is well-built, but ~13 unscaled upper-bound `elapsed < X ms`
   asserts bypass it — and the population is **small and mostly robust**. Exactly one is the
   real environment-dependent flake: `local_index_core_snapshot_reload`
   (`SearchAndIndex.cpp:1385`, `warmElapsedMs < 1000u`) times a real query over a 2-file tree
   with no latency to hide behind. The ten MTP `<1000ms` checks are ~20x-margin watchdog smoke
   tests (a trip signals a real watchdog regression). Separately, the overlay perf case
   (`ViewCommands.cpp:19689`) demands ≥200 frames inside a fixed `Scale(4200ms)` window and
   fails on *sample count* on a loaded runner.

4. **External-dep / environment poisoning across runs.** A fixed shared self-test root with no
   PID/GUID (`SelfTestCommon.cpp:1053-1081`), rotated destructively at startup
   (`RedSalamander.cpp:7333`) while an orphan-able out-of-process search service with no
   JobObject (`CompareDirectoriesEngine.SelfTest.cpp:4051`) holds handles under it; and remote
   skip-gates that check credential *presence*, not reachability, then run blocking network I/O
   on the tick thread (`Phases14_16.cpp:436,595,622`).

**The cascade multiplier.** `RunCase` snapshots only a fresh `CaseState{}` and never restores
process globals — `g_folderWindow`, `g_settings` (710 refs in `Commands.SelfTest.Settings.cpp`
alone), plugin manager, theme, active pane, Win32 focus. Isolation is 100% delegated to each
case's `wil::scope_exit`. Those restores are registered *before* mutation (so they fire on
early-return), but they call **live async UI ops** (`SetFolderPath` → async `NavigationRequest`,
theme `SendMessageW`, `PumpPendingMessages`) that partial-complete when a failing case left a
window disabled or a modal open. Shared helpers `FocusFolderViewPane` (`Settings.cpp:5979`) and
`ForceRefreshPaneForCommandSelfTest` (`Settings.cpp:6236`) mutate the active pane with **no
restore**. So a leak from any of generators 1-4 shifts the starting state of every later case,
and with `failFast=false` in CI (`SelfTestCommon.h:51`; `ci.yml:113,119`) the cascade is
un-firewalled.

**Why "days" specifically — the runner amplifies everything.** There is **no retry,
rerun-to-classify, or quarantine** (`ci.yml:110-120`; `Run-AllTests.ps1` has zero flake
matches). One 1-in-N flake is indistinguishable from a regression, so the only recovery is a
full 40/50-min re-run that surfaces a *different* flake. Multiple suites share one process
(`--compare-selftest --fileops-selftest`); `results.json` is written only at run end; on an AV
crash the SEH `__except` (`RedSalamander.cpp:7899-7907`) unconditionally calls
`ShowFatalErrorDialog → ShowModal()` — a blocking modal with **no CI gate** — turning a fast
crash into a 50-minute hang with zero per-case data.

**Two corrections that are load-bearing for the plan:**
- **The flake tooling covers only a third of the gate.** Repeat/shuffle/retry-classify all hook
  `RunCase` + `--selftest-case`, which exist **only** for the three in-process selftests. The
  standalone EXEs — DxUiTests (630, incl. Menu/NativeTextInput), ViewerPETests, ViewerSqliteTests,
  MonitorTest, PerformanceTests2 (VSTest DLL), Pester — have **no** `--selftest-case`. Retry-classify
  must be per-harness or it leaves DxUiTests, the worst offender, unclassified.
- **The real focus contention is intra-process, not inter-job.** CI already has one serial
  `selftest` job. "Enforce serial execution" changes nothing. DxUiTests runs all 630 tests in
  **one process** — Menu's `SetForegroundWindow` and NativeTextInput's `GetFocus` interfere
  *within* the exe (README:127). The fix is intra-exe ordering/isolation or splitting DxUiTests
  by focus-sensitivity, not job serialization.

---

## Confirmed root-cause table

| Cause | Category | Severity | Evidence | Failure scenario |
|---|---|---|---|---|
| Commands hard-asserts real Win32 `GetFocus()==window`; only true when process owns foreground desktop | Focus / input | Critical (primary generator) | `Navigation.cpp:2694,3055,3153,3541,4081,4260,4402`; `Preferences.cpp:82-99`; `ci.yml:119` | On headless CI `SetFocus` no-ops, `GetFocus()` returns NULL, assert fails; which focus assert trips first depends on ambient foreground timing → shifting failure set |
| NativeTextInput caret tests require real system caret / `GetGUIThreadInfo().hwndCaret==host` on an off-screen never-activated window | Focus / input | High (verified) | `DxUiTestHelpers.h:673-674` (window at -32000,-32000, never shown); `DxUiTests.NativeTextInput.cpp:677,735-741`; caret gated on `GetFocus()==_hwnd` (`DxUi.NativeTextInput.cpp:1643`) | Non-foreground session: `SetFocus` no-op → focus asserts fail while DIP-rect asserts pass — exactly the shifting subset |
| DxUiTests Menu calls `SetForegroundWindow` (return discarded) then relies on capture-based popups that self-dismiss on activation loss | Focus / input | Critical | `DxUiTests.Menu.cpp:1174-1176,1268-1271`; `DxUi.Menu.cpp:1999-2007` | Foreground refused/stolen → popup deactivates → `Dismiss()` → hover deadline times out → `Require` fails |
| All 630 DxUiTests run in one process; Menu + NativeTextInput focus state interfere intra-exe | Focus / input | High (corrects "serial job" framing) | `DxUiTests.cpp:132`; README:127; `ci.yml:135` | Menu's foreground manipulation leaves focus state NativeTextInput's `GetFocus` asserts inherit; a serial CI *job* does not fix this — only intra-exe split/ordering does |
| Phase14 tears down `FileOperationState` + popup HWNDs from a threadpool thread while UI pump dispatches completion msgs | Threads / async | High | `Phases14_16.cpp:146`; `State.Runtime.Part.cpp:797`; `FileOperations.cpp:1299` | Completion message dispatched into half-destroyed popup → SEH crash → whole suite aborts, different case set "fails" next run |
| Live concurrency counters read across worker/UI boundary under wrong/no mutex | Threads / async | High | `_perItemMaxConcurrency` unlocked write (`State.cpp:6413`); readers take `_progressMutex` (`Phases07_09.cpp:3725`) or nothing (`Phases10_13.cpp:2444`) | Reader samples torn/stale count mid-increment; `Phase7` fails "expected both tasks in-flight"; interleaving varies run to run |
| Process-global 16-thread delta assertion sensitive to unrelated thread churn | Threads / async | High | `GetProcessThreadCount` (`SelfTest.cpp:4156`); baseline `Phases07_09.cpp:3644`; fail if delta>16 (`:3765`) | A COM/plugin thread starts between snapshots → delta>16 → fail on busy runs only |
| Cross-case Dialog UIA worker detached (not stopped) on 3000ms timeout | Threads / async | Medium | `Dialogs.cpp:43-55` | Detached focus-mutating worker completes into next case's setup, stealing focus; next case fails nondeterministically |
| One genuinely fragile unscaled `elapsed < X ms` ceiling on a real operation (rest are ~20x-margin smoke tests) | Timing / perf | High (1 case) / Low (rest) | `SearchAndIndex.cpp:1385` (`warmElapsedMs<1000u`, 2-file tree, ~220ms idle); MTP `<1000` = watchdog smoke (`Mtp.cpp:1158,1344,1572`) | Cache eviction + Defender rescan pushes a correct 220ms query to 1150ms; multiplier=2 cannot relax a ceiling; stays red on the loaded box |
| Overlay perf case requires ≥200 frames inside a fixed `Scale(4200ms)` window; unconditional `Require` on sample count | Timing / perf | High | `ViewCommands.cpp:19689-19721`; `kFolderViewPerfMinimumSamplesForP95=200` (`:16942`); runs in `--commands-selftest` | Loaded runner produces ~150 frames; `Require(overlaySamplesEnoughForP95)` fails on count while every latency was fine; idle re-run passes |
| MITIGATED 2026-07-06 for the named offenders: raw unscaled waits bypass the multiplier | Timing / perf | Medium | `Search.cpp` now scales `WaitForFlag(...)` internally; `BatchRename.cpp` gates `_renameItemsGate` with `SelfTest::ScaleTimeout(30'000u)`; `FolderWindow.FileOperations.SelfTest.cpp` uses a scaled 6-second deadline for temp-root recreation; guarded by `TestHarnessSourceContracts.Tests.ps1` | A slow box needs longer to settle; these named correctness waits now honor `--selftest-timeout-multiplier` instead of expiring on raw wall-clock literals |
| `RunCase` resets no global between cases; isolation is 100% per-case `scope_exit` restoring *live* globals via async UI ops | State isolation | Critical (cascade multiplier) | `SelfTestCommon.h:205-280`; `Settings.cpp` (63 scope_exits, 710 refs); unrestored `FocusFolderViewPane` (`:5979`), `ForceRefreshPaneForCommandSelfTest` (`:6236`) | Restore issues async `SetFolderPath`/theme `SendMessageW` that partial-completes because the case left a window disabled; a code change makes A pass so a *different* leaker fails first next run |
| Commands suite has no fail-fast in CI; runs all later cases on poisoned globals | State isolation | High (amplifier) | `failFast=false` (`SelfTestCommon.h:51`); guard `:223`; `ci.yml:113,119` | First leak never firewalled; cascade membership depends on timing, producing a fluctuating failure tail |
| Fixed shared self-test root, no PID/GUID, rotated destructively at startup | External deps / env | Critical | `SelfTestCommon.cpp:1053-1081,1135-1164`; `RedSalamander.cpp:7333` | Timeout-kill leaves partial files; next step's `RotateSelfTestRuns` rename fails on sharing violation; Commands runs on half-cleaned tree |
| PARTIAL 2026-07-06: test scratch/artifact paths still need full-run disk proof, but FileOps real cross-volume, RedConfigureTests, ViewerSqliteTests, CrashHandlingTests, Commands plugin-config viewer fixtures, ShellCommands shortcut-save temp files, PerformanceTests2 fixtures, BatchRename fixed roots, ViewerPETests runtime fixtures, and DxUiTests generated artifact defaults no longer create fixed out-of-sandbox roots | External deps / env / scope | High (maintainer directive §8) | Current FileOps cross-volume source uses `AcquireTestSandboxOnVolume(...)` and `PruneEmptyAlternateVolumeSandboxParents(...)`; legacy cleanup still reaps historical `RedSalamanderCrossVolumeSelfTest_*`; RedConfigureTests, ViewerSqliteTests, CrashHandlingTests, Commands plugin-config viewer perf/close fixtures, ShellCommands long-path shortcut-save temp files, PerformanceTests2 fixtures, BatchRename window fixture roots, and ViewerPETests runtime fixtures now use the `runs\<runId>\scratch\<harness-or-suite>\<case>` layout; DxUiTests generated defaults now use `runs\<runId>\artifacts\dxui`; the broad raw-temp/profile/legacy-root source guard scans primary gated test sources | FileOps cross-volume orphan risk is mitigated by `<AltDrive>:\RedSalamanderTestSandbox\runs\<runId>\scratch\fileops\real_cross_volume_move` plus parent pruning, and the former fixed `%TEMP%\RedConfigure*`, ViewerSqlite database, CrashHandling marker roots, Commands plugin-config ViewerText/ViewerImgRaw/ViewerWeb roots, ShellCommands shortcut-save temp files, PerformanceTests2 roots, `C:\BatchRename*SelfTest` roots, ViewerPETests build-dir fixture folders, and DxUiTests `Specs\TestRuns\DxUiGallery` / `Specs\TestRuns\local_scratch` defaults now route through `REDSALAMANDER_TEST_ROOT`; source drift is now guarded, but final §8 proof still needs the full-run on-disk audit and any migrations that audit exposes |
| MITIGATED 2026-07-06 for foreground selftest launches: out-of-process search service previously launched with no JobObject; orphans could survive parent kill | External deps / env | Critical | `CompareDirectoriesEngine.SelfTest.cpp` now creates a kill-on-close JobObject, starts the foreground service suspended, assigns it before resume, and traces assignment | Parent killed mid-case can no longer leave the foreground `RedSalamanderSearchService.exe` child holding `index-v2.sqlite3-wal`; still open only if final audit requires rotation-time stray-process cleanup for older orphaned services |
| Remote skip-gates check credential presence, not reachability; run blocking network I/O on the tick thread | External deps / env | High | `CompareDirectoriesEngine.SelfTest.cpp:574-658`; `Phases14_16.cpp:436,595,622-634` | Credentialed-but-stalled S3 blocks inside `stepState==0`; `HasTimedOut` never fires; run consumes full 40-min step timeout |
| No retry / rerun-to-classify / quarantine in CI or runner | Runner / CI | Critical (amplifier) | `ci.yml:110-120`; `Run-AllTests.ps1` (zero flake matches) | One 1-in-N flake is indistinguishable from a regression; only recovery is a full re-run that surfaces a *different* flake |
| Flake/retry tooling is self-harness-only; standalone EXEs (DxUiTests/Viewer/Monitor/Perf/Pester) cannot be classified | Runner / CI | Critical (Track-A gap) | `--selftest-case`/`RunCase` exist only for the 3 in-process selftests (`SelfTestCommon.h:207`); DxUiTests filters by *suite* only (`DxUiTests.cpp:132`) | Retry-classify "re-run failed names via `--selftest-case`" is impossible for DxUiTests; a DxUiTests focus flake still reds the gate as a regression |
| Multiple suites share one process; `results.json` written only at run end; crash loses all per-case results | Runner / CI | Critical | `ci.yml:113`; `RedSalamander.cpp:1559,8996,9061`; `__except` returns -1 without `FinalizeSelfTestRun` (`:7872-7908`) | FileOps AV crash → no `results.json` → CI shows only "exit -1", no idea which of ~260 cases crashed |
| Crash path pops blocking modal with no CI gate; startup rotation discards a history level | Runner / CI | High | `RedSalamander.cpp:7899-7907`, `1391-1396`; `SelfTestCommon.cpp:1135-1164` | SEH AV reaches `__except` → modal `GetMessageW` blocks headless runner forever → full 50-min budget burned |
| Five suites (CrashHandling/PluginContract/SettingsSchema/FileSystemCurl/RedConfigure) exist but are never invoked by CI | Runner / CI / scope | High (scope error) | `Tests/` dirs present; verified NOT in `ci.yml` for all five | The gate cannot flake on suites it never runs; FileSystemCurl/RedConfigure make **no** network calls (README:185,201), so a "reachability probe" fix is misdirected |
| No seed/order capture; CI bypasses `Run-AllTests.ps1` coverage guard | Runner / CI | Medium | `Run-AllTests.ps1:561-595`; no `Run-AllTests` reference in `.github/workflows`; no shuffle/seed in harness | Order-dependent leak can't be reproduced/bisected; suite exits after 700/733 cases and CI never flags the 33 missing |

---

## Remediation plan

Two tracks move in parallel: **(A)** give the runner the ability to distinguish flake from
regression *immediately* — for **all** harnesses — so every red run is actionable instead of
ambiguous; **(B)** systematically drive the underlying nondeterminism to zero. Track A buys
signal and reproducibility. It does **not** make flakes acceptable: a classified flake is still
a failing repair item until the test or product defect is fixed or the test is replaced with
deterministic equivalent coverage.

**Scope guardrail (every phase).** All work targets the primary tree at `Z:\src\RedSalamander`
only. The repo contains 9 agent worktrees under `.claude/worktrees/*/` that hold full second
copies of `RedSalamander/SelfTest/...`. Every grep/audit/edit MUST exclude
`.claude/worktrees/**`, `.build/**`, and `*/x64/Debug/**` — otherwise the audit double-counts
files and effort estimates inflate ~2x.

### 0. Gate reconciliation — decide what is in scope

The gate cannot flake on suites it never runs. Verified NOT in CI:
`CrashHandlingTests`, `PluginContractTests`, `SettingsSchemaTests`, `FileSystemCurlTests`,
`RedConfigureTests`.
- **FileSystemCurlTests / RedConfigureTests** are deterministic with no live remote (verified:
  zero `curl_easy`/`http`/`localhost`/`socket` in `Tests/FileSystemCurlTests/*.cpp` — it tests
  `BuildImapMessageLeafName` string logic and a local `RunPerfProbe`). **Add them to the gate**;
  do **not** give them reachability probes.
- **SettingsSchemaTests / PluginContractTests / CrashHandlingTests:** add if cheap, else declare
  out of scope in a `ci.yml` comment with a reason.
- **MonitorTest is stable-by-design, not a timing flake:** `MonitorTest.exe` returns 0 even when
  `stats.etwFailed > 0` (prints a warning, `Tests/MonitorTest/*.cpp:441-452`); returns 2 only on
  ETW *registration* failure. ETW loss does not red the job. Monitor its duration in §2f; do not
  spend Phase-2 effort on it.
- **LocalizationTests:** static-string/resource contracts, no focus/network/timing surface —
  treat as stable, verify once, move on.

### 1. The Test-Stability CONTRACT (every case must satisfy)

1. **Isolation.** After a case returns, `g_settings`, `g_folderWindow` (active pane, both pane
   paths, focus target, theme), the plugin-manager registry, and the process temp root are back
   to the harness baseline.
2. **Deadline-based waits only.** No fixed `Sleep`/`WaitForSingleObject`/`WaitForFlag` literal
   gates correctness; every wait is `poll-until-condition` bounded by `Scale()/ScaleTimeout()`.
   No unscaled upper-bound `elapsed < X ms` correctness assert — latency ceilings become
   advisory metrics.
3. **Focus-independence.** No case asserts on process-global `GetFocus()`/`GetForegroundWindow()`/
   `GetGUIThreadInfo().hwndCaret` unless it first passes an interactive-desktop probe; else
   `Skip`s with a recorded reason. Logical focus is asserted through the app's own model.
4. **External-dep gating.** A remote/hardware-dependent case passes a bounded reachability probe
   (not credential presence) or `Skip`s. No unbounded synchronous network/pipe I/O on the harness
   tick thread. (Genuine remote surface only — not FileSystemCurl.)
5. **Deterministic order + declared state.** A case declares the exact knobs it needs at entry,
   provably order-independent (verified by shuffle in §2).
6. **Cross-thread reads are synchronized.** Any live counter read from the UI thread is read
   through the same mutex the writer holds, or via an atomic snapshot accessor.
7. **Single walled-garden root.** Every file/folder a case creates is acquired through `TestSandbox`
   under the one `REDSALAMANDER_TEST_ROOT` (§8). No case calls `temp_directory_path`/`GetTempPath`/
   `%LOCALAPPDATA%`/a drive root directly; every path is `<runId>`-scoped and torn down.
8. **No flaky-pass policy.** `FLAKY` is a blocking classification, not an excuse. A gate with
   `flaky-count > 0` is red. Any known flaky entry must have an owner, expiry, linked repair
   issue/spec note, and a fix-or-replace plan; anonymous or permanent quarantine is forbidden.

### 2. Tooling to BUILD

**2a. Harness per-case reset/teardown (highest-leverage single change).** In `SelfTest::RunCase`
(`SelfTestCommon.h:205-280`) bracket the functor call with a `CaseFixture`:
- *Before:* snapshot a **fixed default baseline** (`g_settings`, active pane + both paths, theme,
  focus target); reset globals to it; `IconCache::GetInstance().Clear()` + reset failure TTL
  (`IconCache.cpp:40,1169-1180`); close leftover modal/popup; restore capture.
- *After:* unconditionally restore baseline and **wait for pane navigation to settle** (bounded
  scaled deadline — `SetFolderPath` is async via `NavigationRequest`, `FolderWindow.cpp:1907`);
  then **assert globals returned to baseline** — on drift, mark the case a new `poisoned` status
  naming the drifted field. Converts silent leakage into a loud, attributable failure on the
  *leaking* case.
- **Migration hazard fix:** the 63 existing `scope_exit` restores in `Settings.cpp` restore to
  the pre-case live value, but many cases legitimately declare non-default state at entry. If the
  fixture forces a fixed default while a case still restores-to-prior, the two mechanisms race
  during migration. Prevent with: (i) a per-case `fixtureManaged` opt-in; the baseline-drift
  assert is armed only for cases that dropped their manual restores; (ii) a `RestoreAndSettlePane(path)`
  helper all restores migrate to (replaces fire-and-forget `SetFolderPath`); (iii) a CI lint that
  fails any case that is both `fixtureManaged` **and** self-restoring (half-migrated).

**2b. Flake detector (`--selftest-repeat=N`, `--selftest-shuffle=<seed>`).** Core runner/native
controls are now present for Commands, CompareDirectories, and FileOperations; classifier
consumption, nightly scheduling, and runtime injected classifier proofs are landed. Remaining
work is to use the evidence to fix or replace the actual flaky cases and finish the native sandbox
migration/reaper/lint work.
`--selftest-repeat=N` runs each matched case N times in-process; `--selftest-shuffle=<seed>`
shuffles order with a printed seed; the seed prints on every run and every failure. Wire into
`Tools/TestRunPlan.ps1` (`Get-RSSelfTestArguments`). Pass N-of-N fixed but fail-under-shuffle =
**isolation violation**; k-of-N fixed = **timing/env flake**.

**2c. Retry-and-classify — per-harness adapter + isolation-regression routing.** Retry-classify
must NOT be `--selftest-case`-only. Build a `HarnessAdapter` table in `TestRunPlan.ps1`:

| Harness | Rerun-failed mechanism | Granularity |
|---|---|---|
| Commands/Compare/FileOps selftests | `--selftest-case=<name>` (`SelfTestCommon.h:207`) | per case |
| DxUiTests | positional **suite** filter (`.\DxUiTests.exe NativeTextInput`; `DxUiTests.cpp:132`) | per suite |
| ViewerPETests | positional test-name arg (already used at `ci.yml:127-128`) | per test |
| PerformanceTests2 (VSTest) | `vstest.console.exe /Tests:<FQN>` | per test |
| Pester | `Invoke-Pester -FullNameFilter <name>` | per test |

Isolation-regression routing: for the in-process selftests, a case that **fails in-suite but
passes under `--selftest-case`** is NOT labeled FLAKY immediately (that would mask cascade bugs) —
it is re-run under `--selftest-shuffle` across the recorded seed + 2 fresh seeds. **Passes-all-shuffles
=> FLAKY (blocking repair item); fails-any-shuffle => ISOLATION REGRESSION (blocking).** Standalone
EXEs without a shuffle mode fall back to fixed-order rerun: FLAKY only after 2 consecutive
pass-on-reruns. In all cases, FLAKY means "nondeterministic failure reproduced by the classifier";
the runner exits non-zero and writes a repair record.

**2d. Temporary repair quarantine (per-harness, blocking aggregate).** Replace anonymous
`Tools/quarantine.txt` with reviewed `Tools/test-quarantine.jsonl`, one JSON object per entry:
`{ "harness": "...", "name": "...", "owner": "...", "opened": "YYYY-MM-DD",
"expires": "YYYY-MM-DD", "issue": "...", "rootCauseHypothesis": "...",
"repairPlan": "...", "replacementPlan": "...", "lastSeenRun": "..." }`.
Runner executes quarantined entries in a separate repair lane so one unstable case does not hide
the rest of the suite; that lane may use `continue-on-error` internally to collect all evidence,
but the aggregate `Run-AllTests.ps1` result is **failed** while any unexpired quarantine entry
exists or any quarantined case still reproduces. Entry: labeled FLAKY twice by §2c, or approved by
the maintainer for a crash/timeout risk with the same metadata. Invalid entry: missing owner,
expiry, issue, or repair plan; expired entry; or no matching case in the harness adapter. Invalid
entries fail CI.

**2e. Flaky-test fix-or-replace playbook.** Every FLAKY entry is triaged into one of four outcomes:
- **Product defect:** fix production code and keep the test; exit requires the original test to pass
  N-of-N under repeat/shuffle plus the focused regression command.
- **Harness defect:** remove nondeterminism from the test by replacing wall-clock ceilings, global
  focus/input, shared filesystem roots, or unordered async observation with deterministic hooks,
  scaled poll-until waits, `TestSandbox`, and synchronized snapshots.
- **Invalid test contract:** replace the test in the same PR with deterministic coverage of the
  real behavior, then delete the invalid case. Deleting a flaky test without replacement is a
  coverage regression unless the authoritative spec explicitly says the behavior is obsolete.
- **External dependency:** convert to a bounded precondition probe with `skipped` reason or a fake
  provider/harness stub. Retrying live remote I/O is not a fix.

Exit from quarantine requires: the code/test repair merged, the quarantine entry removed in the
same PR, fixed-seed reproduction passing, repeat/shuffle evidence passing for in-process tests
(or 10-of-10 fixed-order for standalone tests without shuffle), and the repair noted in the WIP
plan or authoritative testing spec.

**2f. Per-case timing/flake dashboard + history store.** Persist `{harness, case, durationMs,
status, classification, seed, attempt}` per run into an append-only CI artifact keyed by case name
(must live outside the rotated run tree). Surface: cases approaching their unscaled deadline;
flaky-rate per case; active quarantine age/owner/expiry; MonitorTest duration + `etwFailed` as
advisory drift signals.

**2g. Seed/order reproduction.** `--selftest-shuffle=<seed>` + printed seed makes any
order-dependent failure replayable for bisection.

### 3. Phased rollout

**Phase 0 — Stop evidence loss and classify (Track A, ~3-4 days). No individual case rewrites.**
- Ship §2c retry-and-classify (with per-harness adapter) and §2d temporary repair quarantine in
  `Run-AllTests.ps1`/`TestRunPlan.ps1`.
- **Route CI through `Run-AllTests.ps1`** (`ci.yml:110-120`) so §2c/§2d + the coverage guard
  (`Run-AllTests.ps1:561-595`) run in CI.
- **Reconcile the gate (§0)**: add FileSystemCurl/RedConfigure (+ others if cheap) or annotate out.
- **Split the combined process:** `--compare-selftest`, `--fileops-selftest`, `--commands-selftest`
  as three separate exe invocations (caps crash blast radius).
- **Gate the fatal-error modal:** under `IsRunningAnySelfTest()`, `ShowFatalErrorDialog`
  (`RedSalamander.cpp:1391-1396,7899-7907`) writes a dump and `exit`s non-zero — never `ShowModal()`.
- **DONE 2026-07-06:** write per-case results incrementally (flush after each case in `RunCase`);
  emit a partial `results.json` on the `__except` path marking the in-flight case `crashed`; prove it
  with an archived injected AV run that exits non-zero within the five-minute gate budget.
- **Per-run unique artifact root + walled garden (§8 R1/R2/R5):** stand up the one
  `REDSALAMANDER_TEST_ROOT=<repoRoot>\.build\TestSandbox` with a
  `runs\<UTC timestamp>-<PID>-<GUID>\` run dir shared by every harness (folds in
  `SelfTestCommon.cpp:1053-1081`), and ship `Tools/Clean-TestSandbox.ps1` to reap the historical
  sprawl — including every `<Drive>:\RedSalamanderCrossVolumeSelfTest_*` (the `C:\R…`/`C:\RST` dirs,
  `FolderWindow.FileOperations.SelfTest.cpp:4556-4585`) — in `Run-AllTests.ps1` and a CI `always()`
  step. Make `RotateSelfTestRuns` retry-with-backoff / move-aside while migrating it to the new
  `runs\<runId>\artifacts\<harness>\` layout.
- **DONE 2026-07-06 for native scratch helper / first migrations:** `SelfTest::AcquireTestSandbox`
  now routes native scratch to `runs\<runId>\scratch\<suite>\<case>`, Compare foreground
  search-service stdout capture no longer uses process temp, and FileOps real cross-volume scratch
  uses `AcquireTestSandboxOnVolume(...)` under
  `<AltDrive>:\RedSalamanderTestSandbox\runs\<runId>\scratch\fileops\real_cross_volume_move` with
  empty parent pruning. The broad raw-root source-guard lint now scans primary gated test sources.
  Still open: execute the full-run disk proof and fix any remaining non-FileOps scratch migrations it exposes.
- **DONE 2026-07-06 for foreground selftest launches:** `ForegroundSearchServiceProcess` wraps
  `RedSalamanderSearchService.exe --run-foreground` in a kill-on-close JobObject, starts the child
  suspended, assigns it before resume, and traces assignment. Still open only if final audit requires
  rotation-time stray-process cleanup for older orphaned services.

**Phase 1 — Harness isolation (the cascade multiplier).**
- Implement §2a `CaseFixture` + `poisoned` status + `fixtureManaged` opt-in + half-migration lint
  in `SelfTestCommon.h`. Migrate `FocusFolderViewPane` (`Settings.cpp:5979`) and
  `ForceRefreshPaneForCommandSelfTest` (`:6236`) to restore, or have the fixture cover them.
- Ship §2b flake-detector flags; wire into runner + a nightly shuffle+repeat lane.

**Phase 2 — Focus/foreground axis (largest CI-side generator).**
- `SelfTest::IsInteractiveForeground()` (probe `GetForegroundWindow()==our root` / `OpenInputDesktop`).
  Gate every real-Win32 focus assert in `Navigation.cpp` (`:2694,3055,3153,3541,4081,4260,4402`)
  and `Preferences.cpp:82-99` behind it; assert the app's own focus model by default; else `Skip`.
- **DxUiTests intra-process split:** focus-sensitive lane (`Menu`, `NativeTextInput`) run one
  suite per invocation; focus-neutral lane for the rest (existing suite filter `DxUiTests.cpp:132`;
  `ci.yml:135`). This removes the intra-exe interference README:127 documents; a serial *job* does not.
- DxUiTests Menu (`DxUiTests.Menu.cpp:1174,1268-1271`): check `SetForegroundWindow` return /
  `GetForegroundWindow()==owner`; `Skip` if it can't foreground, or suppress the activation-loss
  auto-dismiss (`DxUi.Menu.cpp:1999-2007`) under test.
- NativeTextInput (`DxUiTests.NativeTextInput.cpp:677,735-741`): gate the `GetFocus`/`hwndCaret`
  asserts behind the interactive probe (DIP-rect asserts are already focus-independent); filter
  `PumpMessages` (`DxUiTestHelpers.h:699-714`) to drop foreign `WM_ACTIVATEAPP`/`WM_KILLFOCUS`, or
  drive focus via `SendMessage(WM_SETFOCUS/WM_KILLFOCUS)` as the deterministic test at `:867-881` does.

**Phase 3 — Timing/perf axis (small, targeted — population ~13, mostly robust).**
Every item in this phase must carry a Perf Measurement Record per
`Operation_PerfMeasurementContract_2026-07-06.md`: scenario, metric keys, deterministic validation,
build flavor, archived evidence, baseline/candidate comparison, analyzer command, sample-quality
result, and environment matrix.
- **DONE 2026-07-06 — `local_index_core_snapshot_reload` advisory timing** (`SearchAndIndex.cpp:1385`):
  deleted `warmElapsedMs < 1000u`, kept the causal correctness assertions, and emitted
  `compare.selftest.local_index.snapshot_reload_us` as an advisory drift metric. Debug diagnostic
  evidence is archived; Release baseline/candidate evidence is still required before any perf claim.
- **DONE 2026-07-06 — overlay perf advisory sampling** (`ViewCommands.cpp`): replaced the fixed
  `Scale(4200ms)` collection window with sample-targeted collection, kept the overlay-animation
  correctness gate, and moved `overlaySamplesEnoughForP95` to advisory artifact/spec evidence.
- **Leave the ten MTP `<1000ms` checks + the ViewCommands `<500/<50` hooks as-is** (~20x-margin
  watchdog smoke tests) — only reclassify a trip as "watchdog regression," not flake.
- **DONE 2026-07-06 for the named raw waits:** `WaitForFlag` (`Search.cpp:120-134`) scales
  internally; `BatchRename.cpp:439`'s gated `30000u` wait goes through `ScaleTimeout`; and the
  FileOperations temp-root recreation budget (`FolderWindow.FileOperations.SelfTest.cpp:3175`) is a
  scaled 6-second deadline with 50 ms retry slices. Remaining raw-wait audit work stays under the
  broad Phase 1/§5 source-guard sweep.

**Phase 4 — Threads/async/teardown races.**
- **Concurrency counters:** atomic snapshot accessors on `FileOperationState`/`Task`; readers use
  `_perItemInFlightCallsMutex`; make `_perItemMaxConcurrency` atomic (`State.cpp:6413`); fix readers
  in `Phases07_09.cpp:3725`, `Phases10_13.cpp:2444`.
- **Thread-delta assertion** (`Phases07_09.cpp:3644,3765`): drop the process-global count; assert the
  scheduler's own reported active-worker count.
- **Phase14 teardown** (`Phases14_16.cpp:146`): marshal `FileOperationState::Shutdown()` onto the
  **UI thread** (post a message) so window destruction can't overlap `OnFileOperationCompleted`
  dispatch, or fully drain completion + stop the pump before Shutdown. **The crash source —
  highest-value async fix.**
- **Dialog UIA worker** (`Dialogs.cpp:43-55`): never `detach`; join with a scaled deadline and
  `Fail` on timeout, or poll the stop_token.

**Phase 5 — External deps (genuine remote surface only).**
- Bounded reachability probes in `CheckRemoteConnectionSecret`
  (`CompareDirectoriesEngine.SelfTest.cpp:574-658`); run remote ops on a worker with a hard cancel
  deadline so `HasTimedOut` can fire (`Phases14_16.cpp:436,595,622`). **Do NOT touch
  FileSystemCurlTests — it has no network.**
- Per-run-unique ReFS scratch path + cleanup (`SearchAndIndex.cpp:1482`).

### 4. CI changes (`.github/workflows/ci.yml`)
- Invoke `Tools/Run-AllTests.ps1` (not the raw exe) so retry-classify, quarantine, and coverage
  guard run in CI.
- One exe invocation **per suite** (three selftest steps, not the combined pair).
- **Split DxUiTests** into focus-sensitive (Menu, NativeTextInput — one suite per invocation) and
  focus-neutral lanes.
- Add FileSystemCurl/RedConfigure steps (and others if kept); annotate deliberate exclusions.
- Retry-failed-only second pass via the §2c adapter; **FLAKY, REGRESSION, and ISOLATION
  REGRESSION all block**; publish `flaky-count`/`regression-count`/`isolation-regression-count` and
  active quarantine owner/expiry to the job summary.
- Quarantine repair lane as a separate evidence-collection step keyed by the §2d JSONL entries; it
  may continue after individual failures, but the aggregate CI job remains red while entries exist
  or reproduce.
- Set `REDSALAMANDER_TEST_ROOT=${{ github.workspace }}\.build\TestSandbox` for the job and pass it
  to every harness; upload `results.json` + timing/flake history (`if: always()`). Native selftests
  consume `REDSALAMANDER_TEST_ROOT` plus `REDSALAMANDER_TEST_RUN_ID` directly. Treat
  `REDSALAMANDER_SELFTEST_ROOT` as a deprecated compatibility override only; normal runner execution
  ignores inherited values while choosing the run context and clears them before child processes,
  while the reviewed repair lane may set it
  temporarily for a single repair attempt.
- Keep `--selftest-timeout-multiplier=2.0`; add a **nightly** shuffle+repeat lane
  (`--selftest-shuffle`, `--selftest-repeat=5`) — NOT on every gate (§6 budget).
- Do **not** add "serial job"/matrix-removal changes — CI is already single-job serial; the fix is
  intra-exe (DxUiTests split).

### 5. Audit checklist (primary tree only)
For each gated test source file (exclude `.claude/worktrees/**`, `.build/**`), verify against §1:
1. Grep `Sleep(`, `WaitForSingleObject(`, `sleep_for`, `WaitForFlag(` with **literal** args → scaled poll-until.
2. Grep `elapsed`/`< NNNms`/`< NNNu` upper-bound asserts → delete or advisory (most survivors are wide-margin MTP smoke tests — reclassify, don't rewrite).
3. Grep `GetFocus`, `GetForegroundWindow`, `SetForegroundWindow`, `GetGUIThreadInfo`, `GetCursorPos`, `GetAsyncKeyState` → gate behind interactive probe or move to app focus model.
4. Grep `SetEnvironmentVariableW`, `GetInstance`, `FreeLibrary` → confirm restore + fixture coverage.
5. Confirm every `g_settings`/`g_folderWindow` mutation has a fixture-covered restore and declares needed state at entry; confirm no case is both `fixtureManaged` and self-restoring.
6. Run each file's cases under `--selftest-repeat=10` + `--selftest-shuffle` across several seeds; any k-of-N or shuffle-only failure is logged, classified, and fixed or replaced through §2e.
7. Confirm the genuine remote cases (MTP/SearchAndIndex/CheckRemoteConnectionSecret) probe reachability — and that FileSystemCurl/RedConfigure are NOT given probes.
8. Identify which of the 58 network/remote grep hits are real synchronous tick-thread I/O vs. compile-time strings; only the former get Phase-5 treatment.

**Priority files (criticals first):** `Commands.SelfTest.Navigation.cpp`, `Commands.SelfTest.Settings.cpp`,
`Commands.SelfTest.ViewCommands.cpp`, `FolderWindow.FileOperations.SelfTest.Phases14_16.cpp`,
`CompareDirectoriesEngine.SelfTest.Cases.{Mtp,SearchAndIndex}.cpp`, `DxUiTests.{Menu,NativeTextInput}.cpp`.
**PerformanceTests2 is NOT on this list** — its asserts are structural/correctness (`!s_items.empty()`,
checksums, column-width ordering; zero `elapsed<Xms` gates), and it is a separate VSTest DLL with its own
`PluginManager.TestStubs.cpp` the `SelfTestCommon.h` fixture does not reach.

### 6. Baseline-first exit criteria (no 15-hour gate)

**Step 0 — measure the current flaky-rate before committing effort.** Run an instrumented baseline:
each gated suite `--selftest-repeat=5` fixed + one shuffled pass, once, on a dedicated nightly job.
Output: the initial blocking repair population + per-case flaky-rate. Sizes Phase 2-5 effort empirically.

**Gate exit criteria (cheap, per-gate):**
- **Reproducibility:** a fixed seed reproduces the same pass/fail set (verified 3x in the nightly lane).
- **Runner classification working:** an injected 1-in-3 flaky case is auto-labeled FLAKY and fails
  the aggregate run; an injected order-dependent case is labeled ISOLATION REGRESSION (not FLAKY).
- **DONE 2026-07-06 - Crash signal preserved:** injected AV proof exited `-1` in 915 ms and archived
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-06_110551\selftest_run_results.json`, naming
  `modeless_window_ownership` with status `crashed` and `0xC0000005` Access Violation.
- **Gate green:** `flaky-count=0, regression-count=0, isolation-regression-count=0`, active
  quarantine entries = 0, and repair-lane reproductions = 0 for 10 consecutive gate runs.

**Nightly/sampled exit criteria (expensive — NOT on every gate):**
- **Isolation proven:** every in-process case passes 20-of-20 under `--selftest-shuffle` across ≥5 seeds;
  zero `poisoned`. Run **sharded across nights** (one suite/night or a rotating 1/5 shard) — a full 20x of
  ~1600 in-process + ~630 DxUiTests at 40-50 min/suite is ~15-17 h and must never sit on the gate.
- **Timing robustness:** sampled 20-of-20 under `--selftest-timeout-multiplier=1.0` on a CPU-loaded runner, sharded the same way.
- **Quarantine empties:** the Step-0 repair population trends to zero as Phase 2-5 land; no entry
  is allowed to renew without maintainer review and updated fix-or-replace evidence.

### 7. Sequencing & rough effort
- **Phase 0 (Track A + gate reconciliation + per-harness adapter):** ~3-4 days. Stops "days to converge"
  immediately — flakes stop being misdiagnosed as unrelated regressions; the gate remains red until
  each flake is fixed or replaced, and crashes stop destroying signal.
- **Baseline measurement (§6 Step 0):** ~1 day of nightly compute, in parallel with Phase 0.
- **Phase 1 (fixture + migration lint + flake detector):** ~4-5 days.
- **Phases 2-5:** driven by the §2f dashboard, §6 baseline, and §5 audit, worked worst-first out of
  the blocking repair population. Blast-radius order: Focus (2) > Async/teardown (4, Phase14 Shutdown marshalling is the hardest)
  > Timing (3, mostly the one SearchAndIndex ceiling + overlay rewrite) > External (5, genuine remote only).

**Discipline that keeps it converged:** nothing exits quarantine without a fix-or-replace PR and
passing flake-detector evidence; no new in-process case merges while both `fixtureManaged` and
self-restoring; no new standalone-suite case merges without a §2c adapter entry so it is classifiable.

---

## 8. Single walled-garden test root — MANDATORY (all harnesses, one location)

**Requirement (maintainer directive, 2026-07-04):** every file and folder any test creates MUST live under
**one** well-defined sandbox root, identical across **all** harnesses. No test may write anywhere else, and
the sandbox MUST be cleaned up. Today the working/scratch/artifact paths are scattered across at least five
unrelated locations with inconsistent uniqueness and no reliable cleanup:

| Location today | Owner | Uniqueness | Cleanup | Evidence |
|---|---|---|---|---|
| `%LOCALAPPDATA%\RedSalamander\SelfTest\last_run\` (+ `previous_run\`) | 3 in-process selftests + runner | none (fixed) | destructive startup rotation only | `SelfTestCommon.cpp:51-62,1133-1164`; `TestRunPlan.ps1:50-61` |
| **LEGACY:** `<Drive>:\RedSalamanderCrossVolumeSelfTest_<GUID>` at fixed drive roots (A:–Z:) — **the historical `C:\R…`/`C:\RST` dirs** | Historical FileOps cross-volume test roots | GUID | Cleanup-only legacy debt; new allocation uses `RedSalamanderTestSandbox\runs\<runId>\scratch\fileops\real_cross_volume_move` | Current source guard forbids the legacy prefix in FileOps allocation; `Tools\Clean-TestSandbox.ps1` still reaps historical roots |
| **LEGACY:** `%LOCALAPPDATA%\RedSalamander\PluginState\FileSystemMtp\<deviceHash>\overwrite-journal.json` | Historical fake MTP overwrite-journal selftests | ambient user profile + device hash | none; could leak profile state between runs | Current fake MTP journal tests redirect `LOCALAPPDATA` to `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\compare\<short-mtp-segment>` through `SelfTest::AcquireTestSandbox(...)`; source guard forbids the old `GetEnvironmentVariableW(L"LOCALAPPDATA")` helper |
| Per-case FileOps working trees (~30 `create_directories`) | FileOps phases | per-case name | per-case, best-effort | `FolderWindow.FileOperations.SelfTest.cpp:3182,3395,3457,3663,3956,5264-5295` |
| `PrepareSearchCaseRoot(root,"<case>",caseRoot)` + ReFS scratch | Compare/Search | per-case name | per-case | `CompareDirectoriesEngine.SelfTest.cpp:3815`; `SearchAndIndex.cpp` (52 `PrepareSearchCaseRoot` sites); ReFS scratch `SearchAndIndex.cpp:1482` |
| `std::filesystem::temp_directory_path()` / Win32 temp APIs / Pester `[System.IO.Path]::GetTempPath()` (= `%TEMP%`) named dirs and build-dir fixture folders | every standalone EXE plus historical perf/viewer fixtures and Tools Pester fixtures | mostly **none** | mostly none | The current raw-temp API scan under `Tests` and `RedSalamander\SelfTest` is empty, and the migrated `Tools\Tests` Pester scratch scripts no longer call `[System.IO.Path]::GetTempPath()`. Former `RedConfigureTests.cpp` fixed-name roots, `ViewerSqliteTests.cpp` database root, `CrashHandlingTests.cpp` marker root, PerformanceTests2 icon/refresh fixture roots, and ViewerPETests runtime fixture roots were migrated on 2026-07-06 to local standalone TestSandbox helpers under `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\<harness>\<case>`. Tools Pester fake winget/VC-runtime roots, vcpkg merge roots, and synthetic perf-run roots now use `New-RSTestSandboxScratchDirectory(...)` under `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\tools-pester\<case>`. DxUiTests generated artifact defaults now route under `REDSALAMANDER_TEST_ROOT\runs\<runId>\artifacts\dxui` instead of `Specs\TestRuns\DxUiGallery` / `Specs\TestRuns\local_scratch`; explicit caller output arguments remain caller-owned. Commands plugin-config ViewerText perf and ViewerImgRaw/ViewerWeb close-roundtrip fixtures, ShellCommands long-path shortcut-save temp files, and BatchRename window fixture roots were migrated to native `SelfTest::AcquireTestSandbox(...)` under `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\commands\<case>`. |

**82 temp/working-dir acquisition sites across ~30 files** each choose their own location independently. The
drive-root cross-volume dirs and the former fixed-name `%TEMP%\RedConfigure*` dirs had **no PID/GUID and
no guaranteed cleanup**, so orphans accumulated and (per §Diagnosis generator 4) poisoned later runs — a
direct contributor to the shifting failure set. FileOps real cross-volume, RedConfigureTests,
ViewerSqliteTests, CrashHandlingTests, Commands plugin-config viewer fixtures, ShellCommands
shortcut-save temp files, PerformanceTests2 icon/refresh fixtures, BatchRename fixed roots,
ViewerPETests runtime fixtures, DxUiTests generated artifact defaults, MTP fake-journal
`LOCALAPPDATA` state, and Tools Pester fake winget/VC-runtime/vcpkg/perf-run scratch roots have now
been migrated. The current direct raw-temp API scan under `Tests` and `RedSalamander\SelfTest` is
empty, and the `[System.IO.Path]::GetTempPath()` scan under `Tools\Tests` returns no migrated
scratch-script matches. `TestHarnessSourceContracts.Tests.ps1` now enforces the broad raw
temp/profile/legacy-root source ban across primary gated test sources. `Run-AllTests.ps1` now
preserves `test_sandbox_audit` in its aggregate summary. Remaining root-isolation work is to execute
the full-run on-disk proof and fix any non-pattern migrations that audit exposes.

**Target design — `TestSandbox` (one root, one owner, enforced):**

- **R1 — one env-anchored base for ALL tests.** Introduce `REDSALAMANDER_TEST_ROOT` with the canonical
  value `<repoRoot>\.build\TestSandbox` (for this workspace:
  `Z:\src\RedSalamander\.build\TestSandbox`; in GitHub Actions:
  `${{ github.workspace }}\.build\TestSandbox`). `Run-AllTests.ps1` MUST set this absolute path for
  every child process and MUST also set `REDSALAMANDER_TEST_RUN_ID`. Native in-process selftests consume
  those two variables directly and write under
  `<REDSALAMANDER_TEST_ROOT>\runs\<runId>\artifacts\selftest\last_run`. **Every** harness — the 3 in-process selftests **and** DxUiTests, ViewerPETests,
  ViewerSqliteTests, PerformanceTests2, RedConfigureTests, CrashHandlingTests, FileSystemCurlTests,
  MonitorTest, LocalizationTests, **and** the Pester tests — resolves *all* scratch, working, and
  artifact paths from this one base. `REDSALAMANDER_SELFTEST_ROOT` becomes a deprecated compatibility
  override for deliberate launches only; the normal runner ignores inherited values while choosing the
  run context, clears them before child tests start, and the quarantine repair lane scopes the
  override to one repair attempt. No test-created path
  may fall back to `%LOCALAPPDATA%`, `%TEMP%`, or a drive root.
- **R2 — deterministic layout:** `<REDSALAMANDER_TEST_ROOT>\runs\<runId>\scratch\<harness>\<case>\`
  for scratch and `<REDSALAMANDER_TEST_ROOT>\runs\<runId>\artifacts\<harness>\` for results, where
  `runId = <UTC timestamp>-<PID>-<GUID>` (plus the shuffle seed when set). This is the same
  run-unique root Phase 0 already requires — unify them.
- **R3 — single acquisition helper is the ONLY way to get a path.** A shared `SelfTest::TestSandbox` (C++, in
  `SelfTest/Common`, with the standalone EXEs linking the same helper) and a mirrored `Get-RSTestSandbox`
  (`TestRunPlan.ps1`) for Pester. It creates the directory, hands it back, registers it for teardown, and
  **fails the run** (does not silently fall back) if the root cannot be created or is not writable. Direct
  calls to `temp_directory_path()`, `GetTempPath`, `GetTempFileName`, `%LOCALAPPDATA%`, or any drive-root
  literal are banned in test code.
- **R4 — cross-volume is the one sanctioned exception, made explicit.** The cross-volume FileOps test
  (`:4556-4585`) genuinely needs a *second* physical volume, so it cannot sit under a single drive. It MUST
  still allocate through `TestSandbox` on the alternate volume as
  `<AltDrive>:\RedSalamanderTestSandbox\runs\<runId>\scratch\<harness>\<case>\...`, register it for
  teardown, and is the **only** documented escape from the single-drive root — PID/GUID-scoped and
  reaped like everything else. No other test may touch a drive root.
- **R5 — cleanup MUST be done (both directions).**
  - *Pre-run:* `TestSandbox` creates a fresh `<runId>` dir and sweeps sibling `<runId>` dirs whose PID is dead.
  - *Post-run:* unconditional RAII/`scope_exit` teardown of the run dir, plus a finalizer that also runs on the
    `__except` crash path (best-effort) so a crash no longer leaves the garden dirty.
  - *Legacy reaper:* a new `Tools/Clean-TestSandbox.ps1` (invoked by `Run-AllTests.ps1` and a CI `always()`
    step, and offered to the maintainer with `-WhatIf`) that deletes the historical sprawl:
    `%LOCALAPPDATA%\RedSalamander\SelfTest`, `%TEMP%\RedConfigure*`, `%TEMP%\CrashHandlingTests-*`, the Viewer/
    Sqlite `%TEMP%` dirs, **and every `<Drive>:\RedSalamanderCrossVolumeSelfTest_*` on every fixed drive** — i.e.
    the `C:\R…`/`C:\RST` dirs the maintainer sees today.
- **R6 — enforce the wall with a source-guard lint.** `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`
  now greps primary gated test sources for raw `temp_directory_path`/`GetTempPath`/`GetTempFileName`,
  PowerShell `[System.IO.Path]::GetTempPath()`, direct `LOCALAPPDATA`/`TEMP`/`TMP` env acquisition,
  and the known legacy root families; it fails if any path is acquired outside `TestSandbox`. This keeps
  the garden walled as new tests land — the same discipline gate as §7.
- **R7 — concurrency/orphan safety.** Because every path is `<runId>`-scoped, concurrent checkouts and the nightly
  shuffle lane never collide; combined with the search-service JobObject (Phase 0 / external-deps finding) no
  orphaned process can hold a handle under the garden and block cleanup.

**Where this lands in the rollout:** R1/R2/R5-legacy-reaper are **Phase 0** (bleeding-stop — they kill the
orphan/stale-file poisoning immediately and give the maintainer a clean box). R3/R4/R6 now have the native
  `TestSandbox` helper, FileOps cross-volume migration, broad source-guard lint, and disk-audit summary plumbing landed; remaining Phase 1 proof
  is executing the full-run disk audit and fixing any migrations it exposes. Exit criterion: the §6 gate adds **"after a full run + an injected crash,
  `REDSALAMANDER_TEST_ROOT\runs\` is the only test-created location on disk and has no live run dirs after
  archival/cleanup; zero `RedSalamanderCrossVolumeSelfTest_*` at any drive root; the source-guard lint passes."**

---

## Quick wins (highest leverage first)

1. **Ship retry-and-classify with a per-harness adapter and route CI through `Run-AllTests.ps1`.**
   Re-run only failed items via the right mechanism per harness. Label pass-on-rerun FLAKY (blocking
   repair item) vs fail-both REGRESSION. The single biggest lever on "days" — but only if it reaches
   DxUiTests and still fails the gate until fixed.
2. **Route in-suite-fail / isolated-pass to the shuffle detector, not a blanket FLAKY label** — else Track A
   masks the cascade bugs.
3. **No anonymous permanent quarantine.** Add `Tools/test-quarantine.jsonl` with owner, expiry,
   issue, root-cause hypothesis, and fix-or-replace plan; invalid or expired entries fail CI.
4. **Reconcile the gate first.** Add FileSystemCurl/RedConfigure (no network) or annotate the five ungated
   suites out of scope. Do NOT give FileSystemCurl a reachability probe.
5. **DONE 2026-07-06:** Gate the fatal-error modal under `IsRunningAnySelfTest()` — trace + return
   before `ShowModal()` so the caller can exit non-zero without hanging CI.
6. **Split the combined `--compare-selftest --fileops-selftest`** into separate invocations and **flush
   per-case results incrementally.**
7. **Add the `CaseFixture` to `RunCase`** with a `fixtureManaged` opt-in and half-migration lint; mark the
   *leaking* case `poisoned` on baseline drift.
8. **DONE 2026-07-06:** Gate every real-Win32 `GetFocus()`/`hwndCaret`/`GetForegroundWindow()`
   assert behind an interactive-desktop probe. `NativeTextInput` now uses
   `TryFocusDxUiTestWindow`; Menu foreground popup/focus paths use
   `TryActivateDxUiTestWindow`; foreground-only paths emit explicit `SKIPPED:`
   reasons when the desktop session cannot provide the required capability.
9. **Split DxUiTests into focus-sensitive (Menu, NativeTextInput, one suite per invocation) and focus-neutral
   lanes** — the real fix for intra-exe focus interference.
10. **DONE 2026-07-06 for in-product suite controls and runner classifier consumption:** Add
    `--selftest-shuffle=<seed>` + `--selftest-repeat=N` with seed/repeat metadata in artifacts, and
    make the classifier consume shuffle evidence before relabeling isolated-pass failures. Still
    open: run an instrumented baseline to size the actual flaky set.
11. **DONE 2026-07-06:** Fix the one genuinely fragile ceiling — delete `warmElapsedMs < 1000u`
    (`SearchAndIndex.cpp:1385`), preserve causal correctness assertions, and emit
    `compare.selftest.local_index.snapshot_reload_us`; leave the MTP watchdog smoke tests.
12. **Scope every grep/edit to the primary tree** — exclude `.claude/worktrees/**` and `.build/**`.
13. **DONE 2026-07-06:** Convert the overlay perf case to fixed-sample-count collection
    (`ViewCommands.cpp`) and make `samplesEnoughForP95` advisory.
14. **DONE 2026-07-06 for per-run root plus foreground JobObject:** the runner now gives each
    invocation a PID+GUID-scoped run id, native selftests consume it, and foreground Compare
    search-service launches are protected by `JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE`. Still open only if
    final audit requires rotation-time stray-process cleanup for older orphaned services.
15. **PARTIAL 2026-07-06:** Unify ALL test file/folder usage under one walled-garden root and ship a cleanup reaper (§8).
    Native `SelfTest::AcquireTestSandbox` exists, one Compare temp-file user is migrated, FileOps
    real cross-volume scratch now uses the sanctioned alternate-volume `RedSalamanderTestSandbox`
    root with empty parent pruning, RedConfigureTests now routes its former fixed `%TEMP%`
    roots through `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\redconfigure\<case>`, and
    ViewerSqliteTests now creates its database under
    `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\viewer-sqlite\database`, CrashHandlingTests
    now writes marker-file fixtures under
    `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\crash-handling\marker-files`, and Commands
    plugin-config ViewerText/ViewerImgRaw/ViewerWeb fixtures now acquire
    `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\commands\<case>` through native
    `SelfTest::AcquireTestSandbox(...)`; ShellCommands long-path shortcut-save temp files now use
    `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\commands\shell_shortcut_save_temp`, and
    PerformanceTests2 icon/refresh fixtures now use
    `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\performance-tests2\<case>`. BatchRename window
    fixture roots now use `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\commands\<case>` instead of
    fixed `C:\BatchRename*SelfTest` paths, and ViewerPETests runtime fixture files now use
    `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\viewer-pe\<case>` instead of build-dir
    `Viewer*Tests` folders. DxUiTests generated default artifacts now use
    `REDSALAMANDER_TEST_ROOT\runs\<runId>\artifacts\dxui` instead of `Specs\TestRuns\DxUiGallery`
    or `Specs\TestRuns\local_scratch`, while explicit output arguments remain caller-owned. MTP
    fake overwrite-journal cases now redirect `LOCALAPPDATA` to
    `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\compare\<short-mtp-segment>` for the plugin
    state path, with short scratch segments preserving Win32 path budget while keeping runner case
    names unchanged. Compare dummy filesystem scratch cases now use
    `REDSALAMANDER_TEST_ROOT\runs\<runId>\scratch\compare\<case>` through native
    `SelfTest::AcquireTestSandbox(...)` instead of fixed `Y:\CompareSelfTest_*`,
    `Z:\CompareSelfTest_*`, or `W:\CompareSelfTest_*` roots. A raw temp API scan
    for `std::filesystem::temp_directory_path`, `GetTempPathW`, and `GetTempFileNameW` under
    `Tests` and `RedSalamander\SelfTest` is empty. Still route every remaining harness through
    `REDSALAMANDER_TEST_ROOT=<repoRoot>\.build\TestSandbox` and `TestSandbox`; ban direct
    direct `temp_directory_path`/`GetTempPath`/`%LOCALAPPDATA`/drive-root use with the broad
    source-guard lint; and keep `Tools/Clean-TestSandbox.ps1` deleting the historical sprawl — `%LOCALAPPDATA%\RedSalamander\SelfTest`,
    `%TEMP%\RedConfigure*`, `%TEMP%\CrashHandlingTests-*`, and every `<Drive>:\RedSalamanderCrossVolumeSelfTest_*`
    (the `C:\R…`/`C:\RST` dirs, `FolderWindow.FileOperations.SelfTest.cpp:4556-4585`). Kills the orphan/stale-file
    cross-run poisoning at its source.
