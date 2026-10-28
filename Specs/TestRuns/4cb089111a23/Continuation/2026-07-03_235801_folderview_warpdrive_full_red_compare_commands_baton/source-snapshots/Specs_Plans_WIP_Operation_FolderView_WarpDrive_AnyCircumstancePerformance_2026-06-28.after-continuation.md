# FolderView Warp Drive Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make FolderView performance defensible under realistic fast, cold, huge, slow, remote, device-loss, and Release-build circumstances instead of only under small warm Debug selftests.

**Architecture:** Bootstrap the evidence plumbing first: correct build labels, add the shared deterministic latency hooks, and make perf artifacts analyzable before taking baselines. Then fix the confirmed UI-thread stall and rendering-resilience defects, expand representative scale/cold/slow fixtures, and gate future dirty-region, scroll-region, virtualization, and broad DxUi/Monitor ideas behind fresh evidence.

**Tech Stack:** C++23, Win32, WIL RAII, Direct2D/DirectWrite, DXGI, shell thumbnail/icon APIs, `Debug::Perf`, Commands selftests, `Specs/TestRuns/` archived evidence, vcpkg/MSBuild via `build.ps1`.

---

## Master Implementation Checklist

This is the single source of truth for *what to do and how to know it is done*. Each
task below has its own per-step boxes; this list is the roll-up. Work top to bottom.
Do not start a later phase until the gates of the earlier phase are green.

### Definition of Done (applies to every box)

A box may only be checked when **all** of these hold:

1. The change is committed in its **own task-scoped commit** (do not batch tasks).
2. Every PowerShell verification command listed in that task **exits `0`**.
3. The produced evidence (selftest trace path, perf artifact path, or
   `Specs/TestRuns/<MachineHash>/.../` archive path) is **pasted into the
   Implementation Findings Log** below.
4. The box is checked with the **commit short-hash** in the line, e.g.
   `- [x] Task 2 — Brush reuse (`a1b2c3d`).`
5. No previously-green selftest or perf-quality gate regressed.

### Pre-flight (do before any code change)

- [x] **P0 — Confirm clean build.** Done when `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` exits `0` on `045a1f773` (or current HEAD, recorded in the log). (`6fc7851d1`)
- [x] **P1 — Bootstrap evidence plumbing.** Done when Task 0 is green: selftest perf artifacts report the real build flavor, `Tools/Show-PerfRuns.ps1` can fail on insufficient sample quality, and shared test-only latency hooks exist for shell thumbnail lookup, icon extraction, and paste-shortcut save work. (`6fc7851d1`)
- [x] **P2 — Capture FRESH same-machine baselines after Task 0.** Done when Debug *and* test-enabled Release runs of `folderView_perf_scroll_render_stress`, `folderView_perf_overlay_invalidation_stress`, `folderView_perf_large_folder_baseline`, `folderView_perf_iconcache_contention`, and `folderView_perf_directory_change_storm` are archived under `Specs/TestRuns/<MachineHash>/...` and their paths are in the log. These are the *before* evidence. **Do not reuse pre-sink-fix (pre-2026-06-19) archives** — see Finding C4. (`4dd103b25`)
- [x] **P3 — Record the machine profile.** Done when GPU, driver, display refresh, DPI/scale, WARP availability, and local-console/RDP status are written into the log (needed for "any circumstance" attribution). (`4dd103b25`)

### Phase R0 — Readiness bootstrap (must land before R1-R4)

- [x] **Task 0 — Evidence labels, analyzer gates, and shared latency hooks.** Done when the build-label bug is fixed, the analyzer reports sample quality with non-zero failure on hard gates, and all later no-freeze tests can use one bounded deterministic latency mechanism instead of inventing separate delay hooks. (`6fc7851d1`)

### Phase R1 — No-freeze / resilience fixes (correctness-style, ship first)

- [x] **Task 1 — Recover from device loss.** Done when `folderView_render_device_loss_recovers` passes for both injected `EndDraw` and injected `Present` failures, proves the pane repaints after `DXGI_ERROR_DEVICE_REMOVED`, and `folderView_perf_scroll_render_stress` still passes. **Gate: verify `DiscardDeviceResources()` actually nulls the device pointers — see Finding C3.** (`afbbf465d`)
- [x] **Task 2 — Remove per-item brush creation from the draw loop.** Done when a wide-selection + hover guard proves the test-only draw-loop transient brush-create counter is `0`, and scroll + overlay cases still pass. (`9506d9104`)
- [x] **Task 3 — Thumbnail fast-first + close-safe.** Done when BOTH gates pass: (a) visible-path gate — `folderView_thumbnail_cached_only_no_close_stall` proves the visible thumbnail path issues zero provider-allowed shell lookups, asserted via `PaneViewOptionsDebugSnapshot.thumbnailShellProviderAllowedCount == 0` (Step 3 must add this snapshot plumbing); and (b) teardown-safety gate — close/navigate while thumbnail work is pending returns within a bounded deadline and reaches `thumbnailPendingCount == 0` without the UI thread blocking on any joined shell-provider worker. Scroll + valid-image cases still pass. NOTE: `SIIGBF_INCACHEONLY` removes the normal visible-path provider call; it does **not** make a synchronous `GetImage` interruptible if a future path reintroduces one. **Gates: use the Task 0 shared thumbnail latency hook, snapshot plumbing for the provider-allowed counter is required, and any abandonable shell work must capture only values/HWND/generation, never `this`.** (`6c26367b7`)
- [x] **Task 4 — Paste-shortcut off the UI thread.** Done when the new ShellCommands case proves the command returns before delayed shortcut creation completes, close/navigation does not wait for delayed `IPersistFile::Save`, completion posts via `PostMessagePayload`, stale generations are ignored, and existing paste-shortcut cases stay green. **Gate: add the new case to `Commands.SelfTest.ShellCommands.cpp`; the existing cases in Finding C1 are load-bearing regression guards.** (`0dc592dcf`)

### Phase R2 — Measurement truth (no perf claim is valid until this is in place)

- [x] **Task 5 — Statistically useful harness.** Done when each frame metric family reaches ≥200 samples (or the case fails its `metricQuality.samplesEnoughForP95` check), Release evidence is produced and labeled by the Task 0 build flavor, and the analyzer emits count/p50/p95/p99/max plus sample-quality verdicts. (`7a511bfeb`)
- [x] **Task 6 — Representative scale/cold/slow/config fixtures.** Done when `folderView_perf_huge_folder_scale` (≥10k, 50k behind a flag), `folderView_perf_cold_first_visit`, `folderView_perf_slow_virtual_provider`, `folderView_perf_relayout_churn_while_scrolled`, and the WARP/RDP/high-DPI matrix manifest archive artifacts with scale/cold/slow/config markers and sample-quality fields. (`b6195d15`, `075e98d4`, `c51d1830`, `4512fdcb`)
- [x] **Task 7 — Automated regression gates.** Done when `FolderViewPerfBudgets.json5` exists with thresholds derived from the *post-Task-6 fresh* baselines (each threshold cites its source archive), hard budgets are machine-keyed, a known-good budget passes, and an intentionally-tiny scratch budget fails (scratch budget NOT committed).

### Phase R3 — Evidence-gated FolderView backlog

- [x] **Task 8 — Measure & close refresh preservation.** Done when `folder.refresh.*` rows appear in archives, the refresh-preservation case asserts selection/focus/incremental-search survival, and the task closes as measured no-op *or* lands the smallest measured extension.
- [x] **Task 9 — Measure icon pipeline before adding concurrency.** Done when the queue/extract/convert metrics are separated and summarized, an icon-heavy cold+slow fixture proves one slow request cannot stall all visible icons, and concurrency is either accepted with same-machine evidence or closed as measured no-op.

### Phase R4 — Larger architecture, only if still justified by R2/R3 evidence

- [ ] **Task 10 — Dirty regions / scroll rects / virtualization.** Done when each sub-item is either implemented behind its passing gate *or* explicitly recorded in the log as "gate not met — not implemented," with the deciding metric cited.

### Phase R5 — Closeout

- [ ] **Task 11 — Closeout & spec migration.** Done when durable contracts are merged into `Specs/UI/UI_FolderView.md` + `Specs/Testing/*`, `.\Tools\Run-AllTests.ps1 -Suite Full` passes, the full perf matrix (Debug + test-enabled Release) is archived and cited, and this plan is moved to `Specs/Plans/Done/`.

> The detailed `- [ ]` boxes inside each Task section remain authoritative for the
> sub-steps. Keep both in sync: when a Task's sub-steps are all checked, check its
> roll-up box here and its line in the final **Acceptance Checklist**.

---

## Implementation Findings Log

Append every discovery here as implementation proceeds — surprises, corrections to
this plan, archive paths, accepted/rejected decisions, and measured numbers. Newest
entries at the bottom of each subsection. **This section is mandatory: a step is not
Done until its evidence path or finding is recorded here (see Definition of Done).**

Entry format: `- [<DATE> · Task N] <what was found / decided / measured>. Evidence: <path or commit>.`

### 2026-06-29 closeout blocker / continuation state

- [2026-06-29 · Task 11] Final closeout is still blocked by broad-suite DxUi/UIA reliability, not by the original FolderView perf gates. The S3 plugin configuration modal seen during the broad Commands run was not a live hang; the run moved past it and later exposed separate broad-suite failures. Current continuation evidence and next steps are archived in `Specs/Plans/WIP/DxUi_Uia_ContinuationBaton_2026-06-29.md`. Key archives: `Specs/TestRuns/4cb089111a23/Commands/2026-06-29_171144_commands_textinputframework_crash` (`textinputframework.dll` crash), `Specs/TestRuns/4cb089111a23/Commands/2026-06-29_172815` (order-sensitive Compare Directories failure with stale snapshot formatting), and `Specs/TestRuns/4cb089111a23/Commands/2026-06-29_173947` (order-sensitive Preferences category-host focus failure). Focused/cluster repros passed at `2026-06-29_171526`, `2026-06-29_171605`, `2026-06-29_172931`, `2026-06-29_173620`, `2026-06-29_174054`, and `2026-06-29_174214`. Do not move this plan to Done until full-suite and final perf evidence are green.
- [2026-06-29 · Task 11] Broad Commands verification is now green on the current dirty worktree after the Preferences/plugin-config/Find selftest hardening. Evidence: build `Z:\src\RedSalamander\.build\logs\msbuild-20260629_194323_230.log` (0 warnings / 0 errors); focused Find shortcut `Specs/TestRuns/4cb089111a23/Commands/2026-06-29_194555`; paired Find predecessor+shortcut `Specs/TestRuns/4cb089111a23/Commands/2026-06-29_194621`; full Commands `Specs/TestRuns/4cb089111a23/Commands/2026-06-29_200556` (774 passed / 0 failed / 2 opt-in ViewerSpace skips). The plan remains WIP because `.\Tools\Run-AllTests.ps1 -Suite Full` and the final Debug/test-enabled Release FolderView perf matrix are still not run/cited.
- [2026-06-29 · Task 11] Continuation sanity checks are also green: DxUiTests build `Z:\src\RedSalamander\.build\logs\msbuild-20260629_200837_714.log` (0 warnings / 0 errors), `.\.build\x64\Debug\DxUiTests.exe` passed with `All DxUi tests passed.`, and `git diff --check` passed with line-ending warnings only.
- [2026-06-29 · Task 10/11] Final FolderView perf matrix cleanup removed file-backed selftest trace writes from the measured `Present1` quick-search path and `OnBatchIconUpdate` apply path, then hardened quick-search activation, warm quiescence, and relayout phase attribution. Test-enabled Release build `Z:\src\RedSalamander\.build\logs\msbuild-20260629_210405_389.log` passed with 0 errors and the existing File Operations C4883 warning. Release budget matrix `Specs/TestRuns/4cb089111a23/Commands/2026-06-29_210835` passed 12/12 with `folder.frame.total_us` count 1709, p95 7956us, p99 8328us, max 20107us, and `folder.frame.present_us` p95 7000us (`Tools\Show-PerfRuns.ps1 -FailOnQuality -ShowBuildFlavor`). Debug build `Z:\src\RedSalamander\.build\logs\msbuild-20260629_210918_769.log` passed with 0 warnings/errors; Debug matrix `Specs/TestRuns/4cb089111a23/Commands/2026-06-29_211316` passed 12/12 with `folder.frame.total_us` count 1655, p95 23179us, p99 28212us, max 47462us. Relayout phase evidence is in `Specs/TestRuns/4cb089111a23/Commands/2026-06-29_210027`. Durable rule added to `Specs/Testing/Testing_PerformanceValidation.md`: do not put `SelfTest::AppendSelfTestTrace(...)` file writes inside measured render/draw/icon/present hot paths.
- [2026-06-29 · Task 11] `Tools\Tests\RedSalamanderPluginDeployment.Tests.ps1` now forces `RSBuildEnableTests=true` only around its targeted RedSalamander Debug child build and restores the previous process environment in `finally`; this preserves the Full-suite monitor selftest-hook contract after the Pester deployment check. Evidence: test-enabled Debug build `Z:\src\RedSalamander\.build\logs\msbuild-20260629_215806_212.log` passed with 0 warnings/errors; `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` (`Z:\src\RedSalamander\.build\logs\runall-full-skipbuild-20260629_220021_691.out.log`) reached `RedSalamanderMonitorEtwLatency` and it passed with exit code 0. The plan remains WIP because the same Full run is still red from non-FolderView blockers: Commands/FileOperations access violations already present in earlier 21:29/21:49 runs (`textinputframework.dll` offset `0x6a419` and `ntdll.dll` offset `0x1272b3`; focused Commands repro dump `C:\Users\eric\AppData\Local\CrashDumps\RedSalamander.exe.97872.dmp`, WER report `6ce2cd72-d03b-4eef-90c7-4fa7899dd888`), `ViewerPETests`/ViewerVLC HUD focus, and ToolsPester inventory drift (Full surfaces 17 vs 15, Commands 674 vs 660, NativeTextInput 117 vs 114, FileOperations steps 120 vs 114). Do not move this plan to Done until these broad-suite blockers are handled or explicitly waived.
- [2026-06-29 · Task 11] Pause archive for later continuation is now written to `Specs/Plans/WIP/FolderView_WarpDrive_ContinuationBaton_2026-06-29.md`. The volatile Full-suite output was copied to `Specs/TestRuns/4cb089111a23/Full/2026-06-29_220021_full_suite_blocked/` (`run-all-tests-results.json`, stdout/stderr logs, ViewerPETests output, crash dump manifest, WER events, focused Pester log, git status, and diff stat). The Pester inventory drift from the previous entry has been addressed in the dirty worktree: Full suite expected plan now includes `SettingsSchemaTests` and `CrashHandlingTests`; source-derived counts now match Commands 674, CompareDirectories 184, FileOperations active phases 120, NativeTextInput 117, and Tools Pester 108. Focused verification passed with 13 passed / 0 failed in `Specs/TestRuns/4cb089111a23/Full/2026-06-29_220021_full_suite_blocked/pester-inventory-focused-20260629_224800.log`. Remaining blockers at pause time: Commands/FileOperations access violations, ViewerPETests/ViewerVLC HUD focus, and a full Tools Pester rerun after the crash blockers are handled.
- [2026-06-30 · Task 11] Continuation archive refreshed after a broad Commands rerun: the earlier Connection Manager/TSF crash did not reproduce, and `cmd_connection_manager_window_uses_dxui_command_buttons` passed focused at `Specs/TestRuns/4cb089111a23/Commands/2026-06-30_080925`. The new broad Commands blocker is deterministic enough to resume from: `cmd_preferences_dialog_themes_search_roundtrip_preserves_retained_state` failed after 284 passed / 1 failed / 491 skipped with reason `Failed to refocus the Preferences category host before leaving Themes for General during retained-search round-trip validation.` Evidence and git snapshots are archived at `Specs/TestRuns/4cb089111a23/Commands/2026-06-30_081705` (`commands_results.json`, `commands_trace.txt`, runner stdout/stderr, `git-status-short.txt`, `git-diff-stat.txt`). The continuation baton `Specs/Plans/WIP/FolderView_WarpDrive_ContinuationBaton_2026-06-29.md` was updated; next work starts by fixing this Preferences focus failure, then rerunning FileOperations to see whether the older TSF/native-crash lead still applies.
- [2026-06-30 · Task 11] Pause archive refreshed again at `Specs/TestRuns/4cb089111a23/Continuation/2026-06-30_083849_folderview_warpdrive_pause/` with current `git status`, diff stat, full tracked dirty patch, untracked list, WIP spec copies, and latest Commands stdout/results. The prior Themes retained-search blocker no longer reproduces in focused/pair/prefix runs (`Specs/TestRuns/4cb089111a23/Commands/2026-06-30_082128`, `2026-06-30_082234`, `2026-06-30_082834`). Broad Commands now fails later at `cmd_preferences_dialog_panes_history_size_live_dx_interaction`: `Specs/TestRuns/4cb089111a23/Commands/2026-06-30_083421` (311 passed / 1 failed / 464 skipped) with reason `Preferences Panes page did not settle to the active DX surface before history-size validation.` Focused Panes runs passed at `2026-06-30_083612` and `2026-06-30_083633`. Resume by hardening the Panes `navigateToPanesPage` helper in `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ThemesGeneralPanes.cpp` from relative `VK_DOWN` navigation to deterministic `DebugSelectPreferencesCategory(kPrefCategoryPanes)`, then rerun broad Commands before returning to FileOperations/ViewerPE/ToolsPester/Full.
- [2026-06-30 · Task 11] Latest pause archive is `Specs/TestRuns/4cb089111a23/Continuation/2026-06-30_101733_folderview_warpdrive_pause/CONTINUATION.md`. DxUi stationary-pointer rerun is green (`.build/codex-runs/dxui_stationary_pointer_repro_20260630_resume/dxui-tests.log`). Focused Find is green after a test-harness cursor-settle wait in `Commands.SelfTest.Search.cpp` (`Specs/TestRuns/4cb089111a23/Commands/2026-06-30_101418`). The minimized `400..429 + 460..643` slice now gets past Find and fails at `cmd_pane_clipboardPaste_ignores_unfocused_navigation_edit` (`Specs/TestRuns/4cb089111a23/Commands/2026-06-30_101514`), and the focused case fails the same way (`Specs/TestRuns/4cb089111a23/Commands/2026-06-30_101544`): `Failed to refocus left folder view after opening the navigation edit field.` Resume by proving whether this is selftest focus setup, navigation-edit focus reclamation, or pane focus bookkeeping in `WaitForFolderViewPaneFocus` / `CommandChangeDirectory` / `TryHandleNavigationEditClipboardCommand`, then rerun focused, slice, broad Commands, FileOperations, ViewerPE, Tools Pester, Full, and final perf/spec closeout.
- [2026-06-30 · Task 11] Pause archive refreshed at `Specs/TestRuns/4cb089111a23/Continuation/2026-06-30_105105_folderview_warpdrive_pause/CONTINUATION.md`. The clipboard navigation-edit blocker no longer reproduces focused, and the exact 214-case slice is green after a selftest-only UIA polling hardening in `Commands.SelfTest.Settings.cpp` (`.build/codex-runs/commands_ranges_400_429_460_643_exactfilter_after_uia_poll_20260630/run.log`, 214 passed / 0 failed / 0 skipped). The next blocker is a broad Commands access violation (`-1073741819`) after 14m13s at `cmd_pane_filter_prompt_live_dx_interaction`; volatile `last_run` artifacts were copied into the continuation folder. Resume by reproducing that focused/prefix crash before returning to FileOperations, ViewerPE, Tools Pester, Full, and final perf/spec closeout.
- [2026-06-30 · Task 11] Pause archive refreshed at `Specs/TestRuns/4cb089111a23/Continuation/2026-06-30_110920_folderview_warpdrive_pause/CONTINUATION.md`. Focused `cmd_pane_filter_prompt_live_dx_interaction` passed (1/0/0), and the last80-plus-target prefix passed (81/0/0), so the previous filter-prompt AV remains order-dependent. The current minimized red point is now the last160-plus-target prefix failing earlier at `cmd_pane_clipboardPaste_ignores_unfocused_navigation_edit` (11 passed / 1 failed / 149 skipped) with diagnostics proving focus remains on `RedSalamander.NavigationView.DxHost`, navigation edit mode stays active, and no edit-mode exit occurs. Resume by proving/fixing that selftest setup path first, then rerun last160, broad Commands, FileOperations, ViewerPE, Tools Pester, Full, and final perf/spec closeout.
- [2026-06-30 · Task 11] Pause archive refreshed at `Specs/TestRuns/4cb089111a23/Continuation/2026-06-30_112600_folderview_warpdrive_pause/CONTINUATION.md`. The stale last160 clipboard/navigation-edit red no longer reproduces (`.build/codex-runs/commands_filter_prompt_prefix_last160_freshred_20260630_resume2`, 161 passed / 0 failed / 0 skipped). Broad Commands then failed at credential-prompt long-run UIA stats, but focused and immediate-cluster reruns passed (`commands_credential_prompt_longrun_focused_20260630_resume2`, `commands_credential_prompt_cluster_20260630_resume2`). The current broad-suite red point is `cmd_connection_manager_window_tab_traversal_live_dx_interaction` in `.build/codex-runs/commands_broad_credential_recheck_20260630_resume2` (203 passed / 1 failed / 572 skipped): the Protocol combo step starts from `preKind='Edit' preLabel='Name' preModifiers=0x4`, routes Tab, and lands on `CommandButton` `Remove` with `postModifiers=0x0`. Resume by focused-running that case, then proving whether modifier residue, native text focus, reverse traversal, or tab-order bookkeeping is the root cause before patching. Do not move this plan to Done until broad Commands, FileOperations, ViewerPE, Tools Pester, Full, and final perf/spec closeout are green.
- [2026-06-30 · Task 11] Pause archive refreshed at `Specs/TestRuns/4cb089111a23/Continuation/2026-06-30_120509_folderview_warpdrive_pause/CONTINUATION.md`. The 11:26 Connection Manager tab-traversal blocker no longer reproduces in focused, paired, block, plugin-config-plus-connection-manager, or prefix runs (`commands_connection_manager_tab_traversal_focused_20260630_resume3`, `commands_connection_manager_tab_traversal_pair_20260630_resume3`, `commands_connection_manager_block_to_tab_20260630_resume3`, `commands_plugin_config_plus_connection_manager_to_tab_20260630_resume3`, `commands_prefix_to_connection_tab_20260630_resume3`). Broad Commands now fails later at `cmd_compare_directories_options_live_dx_body_interaction` in `.build/codex-runs/commands_broad_after_connection_prefix_green_20260630_resume3` (519 passed / 1 failed / 256 skipped): `Compare Directories options DX edit 'Ignore files' did not restore its original value.` Focused Compare Options and the immediate Compare Options cluster pass, while the best repro is the 7-case Preferences Compare Directories plus Compare Options sequence `.build/codex-runs/commands_prefs_compare_then_compare_options_20260630_resume3` (6 passed / 1 failed / 0 skipped). Resume by adding narrow diagnostics around the failing `Ignore files` edit restore path, then prove whether the cause is Preferences Compare Directories deferred focus restoration, retained settings state, or wrong/stale UIA edit matching before patching. Do not run multiple `Run-AllTests.ps1` invocations in parallel; they share the volatile `last_run` directory.
- [2026-06-30 · Task 11] Pause archive refreshed at `Specs/TestRuns/4cb089111a23/Continuation/2026-06-30_123644_folderview_warpdrive_pause/CONTINUATION.md`. Debug build passed after diagnostic-only Compare Options edits (`Z:\src\RedSalamander\.build\logs\msbuild-20260630_121407_554.log`, 0 warnings / 0 errors). The previous Compare Options `Ignore files` blocker did not reproduce with diagnostics (`.build/codex-runs/commands_prefs_compare_then_compare_options_diag_20260630`, 7 passed / 0 failed / 0 skipped). Broad Commands now fails later at `cmd_pane_selection_select_unselect_dialogs_keep_navigation_shell_stable` in `.build/codex-runs/commands_broad_after_compare_diag_20260630` (565 passed / 1 failed / 210 skipped): `currentPath=''`, `refreshCount=21`, selected count and focused item otherwise match. Focused rerun and immediate 13-case selection cluster are green (`commands_selection_unselect_navshell_focused_20260630`, `commands_selection_navshell_cluster_553_565_20260630`). Resume by adding diagnostic context to `TestPaneSelectionSelectUnselectDialogsKeepNavigationShellStable` and inspecting `WaitForNavigationViewSnapshot` / `DebugGetNavigationViewSnapshot` before changing product behavior. Do not move this plan to Done until broad Commands, FileOperations, ViewerPE, Tools Pester, Full, and final perf/spec closeout are green.
- [2026-06-30 · Task 11] Pause archive refreshed at `Specs/TestRuns/4cb089111a23/Continuation/2026-06-30_130129_folderview_warpdrive_pause/CONTINUATION.md`. Debug build passed after selection/unselect diagnostic context was added (`Z:\src\RedSalamander\.build\logs\msbuild-20260630_124148_516.log`, 0 warnings / 0 errors). The prior selection/unselect blocker did not reproduce in focused, immediate 13-case, 35-case, or 66-case predecessor runs (`commands_selection_unselect_navshell_focused_diag_20260630`, `commands_selection_navshell_cluster_553_565_diag_20260630`, `commands_fileops_issues_to_selection_diag_20260630`, `commands_shortcuts_compare_fileops_selection_diag_20260630`). Broad Commands now fails earlier at `cmd_preferences_dialog_category_tree_handles_reverse_keyboard_navigation` in `.build/codex-runs/commands_broad_after_selection_diag_20260630` (381 passed / 1 failed / 394 skipped): `Preferences category host lost keyboard focus after the second reverse-navigation VK_UP.` Focused, 12-case, and 33-case repro attempts are green (`commands_prefs_category_reverse_focused_20260630`, `commands_prefs_compare_to_reverse_cluster_20260630`, `commands_prefs_350_to_reverse_cluster_20260630`). Resume by adding diagnostic context to `TestPreferencesDialogCategoryTreeReverseKeyboardNavigation` before changing product behavior. Do not move this plan to Done until broad Commands, FileOperations, ViewerPE, Tools Pester, Full, and final perf/spec closeout are green.

- [2026-06-30 · Task 11] Pause archive refreshed at `Specs/TestRuns/4cb089111a23/Continuation/2026-06-30_143026_folderview_warpdrive_pause/CONTINUATION.md`. The 14:05 Compare Options theme-cycle blocker is now green after `CollectVisibleDescendantValuePatternState(...)` was fixed to scan all visible edit candidates for `UIA_ValuePatternId` instead of failing on the first visible non-ValuePattern edit; focused and 10-case cluster evidence is in `.build/codex-runs/commands_compare_options_theme_focused_valuepattern_fix_20260630` and `.build/codex-runs/commands_compare_options_cluster_valuepattern_fix_20260630`. Broad Commands now gets past Compare Options and crashes later at `cmd_pane_pack_prompt_uses_dxui_unique_archive_and_packer_extensions` with `textinputframework.dll` `0xc0000005` at offset `0x000000000006a419`; focused pack prompt and the immediate 3-case cluster are green, but the explicit 44-case suffix `.build/codex-runs/commands_pack_prompt_suffix_600_643_names_20260630` reproduces the crash. Resume by inspecting DxUi native text input / TextStore / WindowHost teardown and the ArchivePackPrompt modal lifecycle, then add narrow lifecycle tracing before patching. Do not move this plan to Done until broad Commands, FileOperations, ViewerPE, Tools Pester, Full, and final perf/spec closeout are green.
- [2026-06-30 · Task 11] Pause archive refreshed at `Specs/TestRuns/4cb089111a23/Continuation/2026-06-30_145522_folderview_warpdrive_pause/CONTINUATION.md`. Diagnostic-only lifecycle tracing was added around `RunArchivePackPromptModalCycle(...)` and Archive Pack prompt create/show/close/destroy paths; Debug build `Z:\src\RedSalamander\.build\logs\msbuild-20260630_144735_069.log` passed. Focused Pack prompt, `622..643`, and `600..643` lifecycle-trace reruns were green, but `628..643` still crashed with `textinputframework.dll` `0xc0000005` at offset `0x000000000006a419` (`C:\Users\eric\AppData\Local\CrashDumps\RedSalamander.exe.65276.dmp`, WER report `1331c903-1ea9-47c2-817b-ad14a7ac654b`). The crashing trace reaches the second Pack prompt worker function end, then stops before `send IDM_PANE_PACK returned`; no product-side `archive_pack_prompt:` trace lines appeared, so resume by proving trace routing/compilation first, then add narrower debug API and cleanup-scope traces before changing TSF/window-host teardown. Do not move this plan to Done until broad Commands, FileOperations, ViewerPE, Tools Pester, Full, and final perf/spec closeout are green.
- [2026-06-30 · Task 11] Pause archive refreshed at `Specs/TestRuns/4cb089111a23/Continuation/2026-06-30_154019_folderview_warpdrive_pause/CONTINUATION.md`. Debug build `Z:\src\RedSalamander\.build\logs\msbuild-20260630_150318_443.log` passed with 0 warnings / 0 errors after the Pack prompt lifecycle trace layer. The explicit Pack-prompt suffix is now green at `.build/codex-runs/commands_pack_prompt_suffix_628_643_lifecycle_trace2_20260630` (16 passed / 0 failed / 0 skipped), and broad Commands now gets past Pack prompt but fails later at `.build/codex-runs/commands_broad_after_pack_lifecycle_trace2_20260630` (551 passed / 1 failed / 224 skipped): `cmd_pane_fileops_speedLimit_prompt_long_run_open_close_stays_stable`, reason `Custom speed-limit prompt ValuePattern should start with '64,0 KB' during cycle 2.` Resume by inspecting the speed-limit prompt long-run test, `CollectVisibleDescendantValuePatternState(...)`, and `FileOperationsSpeedLimitPromptWindow` snapshot/debug plumbing, then capture observed ValuePattern candidates before changing behavior. Do not move this plan to Done until broad Commands, FileOperations, ViewerPE, Tools Pester, Full, and final perf/spec closeout are green.
- [2026-06-30 · Task 11] Pause archive refreshed at `Specs/TestRuns/4cb089111a23/Continuation/2026-06-30_155822_folderview_warpdrive_pause/CONTINUATION.md`. Debug build `Z:\src\RedSalamander\.build\logs\msbuild-20260630_154734_089.log` passed with 0 warnings / 0 errors after diagnostic-only speed-limit ValuePattern context was added. The previous speed-limit blocker is green in focused, three-case, 21-case local fileops, and diagnostic focused reruns (`commands_speedlimit_longrun_focused_20260630_resume`, `commands_speedlimit_three_case_cluster_20260630_resume`, `commands_fileops_issues_to_speedlimit_cluster_20260630_resume`, `commands_speedlimit_longrun_focused_diag_20260630_resume`). Broad Commands now fails earlier at `.build/codex-runs/commands_broad_speedlimit_diag_20260630_resume` (215 passed / 1 failed / 560 skipped): `cmd_connection_credential_prompt_pointer_click_toggles_secret_visibility`. The broad trace proves the second click restored masked state (`visible=0`, `checked=0`) but focus dropped to `None` (`focus=0`); focused and 21-case connection/credential reruns are green (`commands_credential_secret_toggle_focused_20260630_resume`, `commands_connection_to_credential_toggle_cluster_20260630_resume`). Resume by adding second-click diagnostics in `Commands.SelfTest.Connections.cpp` before changing product behavior, and split the test failure wording between "mask restored" and "focus stayed on toggle." Do not move this plan to Done until broad Commands, FileOperations, ViewerPE, Tools Pester, Full, and final perf/spec closeout are green.
- [2026-06-30 · Task 11] Pause archive refreshed at `Specs/TestRuns/4cb089111a23/Continuation/2026-06-30_221502_folderview_warpdrive_pause/CONTINUATION.md`. Broad Commands is now green after the TSF teardown-order fix (`.build/codex-runs/commands_broad_after_tsf_teardown_order_20260630`, 774 passed / 0 failed / 2 skipped; build `Z:\src\RedSalamander\.build\logs\msbuild-20260630_211155_814.log`). The FileOps startup crash is fixed by replacing the newly-created popup path's synchronous `RedrawWindow(... RDW_UPDATENOW)` with asynchronous `InvalidateRect(...)`; focused crash repro `.build/codex-runs/fileops_phase1_async_popup_invalidate_20260630` passed 3/0/0 after dump evidence showed stack exhaustion in `FileOperationState::StartOperation -> EnsurePopupVisible -> FileOperationsPopupState::WndProc/OnNcPaint/Render/EnsureTarget`. The follow-on focused Floodgate selftest hook failure is fixed by normalizing slash separators before the selftest-only mutation-hook path comparison (`.build/codex-runs/fileops_floodgate_move_cleanup_corruption_normalized_hook_20260630`, 3/0/0). Full FileOps is now green at `.build/codex-runs/fileops_full_after_floodgate_hook_normalization_20260630` (102 passed / 0 failed / 20 skipped; build `Z:\src\RedSalamander\.build\logs\msbuild-20260630_220028_025.log`). Remaining before Done: remove temporary icon diagnostics in `IconCache.cpp` and `FolderView.Icons.cpp`, rebuild, rerun Commands/FileOps, run ViewerPETests/ViewerVLC HUD, Tools Pester, Full, final Debug/test-enabled Release FolderView perf matrix, then migrate final durable notes and move this plan to `Specs/Plans/Done/`.
- [2026-07-01 · Task 11] Temporary icon hot-path diagnostics were removed from `IconCache.cpp` and `FolderView.Icons.cpp`, and `rg -n "IconCache::ConvertIconToBitmap:|FolderView::OnCreateIconBitmap:" RedSalamander\IconCache.cpp RedSalamander\FolderView.Icons.cpp` returned no matches. Debug rebuilds passed at `Z:\src\RedSalamander\.build\logs\msbuild-20260701_080744_110.log` and, after adding diagnostic context to the non-reproduced General Preferences settle assertion, `Z:\src\RedSalamander\.build\logs\msbuild-20260701_090928_050.log`. The one-off broad Commands failure at `cmd_preferences_dialog_general_page_uses_dxui_toggle_cards` did not reproduce in focused, Themes-suffix, 81-case Preferences slice, or exact `0..297` prefix reruns; the helper now includes the missing snapshot fields if it recurs. Current green evidence: focused General `.build/codex-runs/commands_general_toggle_focused_diag_20260701` (1/0/0), prefix `.build/codex-runs/commands_prefix_0_to_general_toggle_diag_20260701` (298/0/0), broad Commands `.build/codex-runs/commands_broad_after_general_diag_20260701` (774 passed / 0 failed / 2 skipped), and full FileOps `.build/codex-runs/fileops_full_after_general_diag_20260701` (102 passed / 0 failed / 20 skipped). Remaining before Done: run `ViewerPETests.exe`, full Tools Pester, `.\Tools\Run-AllTests.ps1 -Suite Full`, final Debug/test-enabled Release FolderView perf matrix, cite the archives, then move this plan to `Specs\Plans\Done\`.
- [2026-07-01 · Task 11] Pause archive refreshed at `Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_134139_folderview_warpdrive_pause\CONTINUATION.md`. The prior NavigationView broad failures are now green after the disk-info menu path was aligned with existing dropdown focus restoration in `NavigationView.Menus.cpp`; evidence is `.build\codex-runs\commands_navigationview_family_after_disk_focus_restore_20260701` (16 passed / 0 failed) plus build `Z:\src\RedSalamander\.build\logs\msbuild-20260701_133358_346.log` (0 warnings / 0 errors). Broad Commands now fails earlier at `.build\codex-runs\commands_broad_after_disk_focus_restore_20260701` (227 passed / 1 failed / 548 skipped): `cmd_preferences_dialog_viewers_theme_cycle_keeps_surface_legible`, reason `Preferences Viewers page did not settle before theme-cycle validation.` Resume by inspecting/reproducing that Preferences Viewers theme-cycle settle path first, adding diagnostic context if needed before changing product behavior. Do not move this plan to Done until broad Commands, Full skip-build, and the final Debug plus test-enabled Release FolderView perf matrices are green and cited.
- [2026-07-01 · Task 11] Pause archive refreshed at `Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_140931_folderview_warpdrive_pause\CONTINUATION.md`. The prior Preferences Viewers blocker is now green in focused/predecessor-cluster evidence (`Specs\TestRuns\4cb089111a23\Commands\2026-07-01_134530`, `2026-07-01_134615`). The next broad Commands red at `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_135019` (204 passed / 1 failed / 571 skipped) was `cmd_connection_manager_window_enter_from_dx_input_routes_default_connect`; focused, predecessor pair, and full Connection Manager prefix reruns passed (`2026-07-01_135101`, `2026-07-01_135143`, `2026-07-01_135317`). A diagnostic-only Connection Manager selftest patch was built successfully in `Z:\src\RedSalamander\.build\logs\msbuild-20260701_135517_041.log` (0 warnings / 0 errors). Broad Commands now fails later at `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_140737` (445 passed / 1 failed / 330 skipped): `cmd_pane_find_dialog_restored_combined_view_state_copy_follows_visible_columns`, reason `Find logical Name sort or widened reordered layout did not settle before restored combined-view-state copy persistence.` Resume there first, adding narrow Find sort/layout/copy diagnostics before changing product behavior if focused or predecessor-cluster runs do not expose root cause. Do not move this plan to Done until broad Commands, Full skip-build, and the final Debug plus test-enabled Release FolderView perf matrices are green and cited.
- [2026-07-01 · Task 11] Pause archive refreshed at `Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_142558_folderview_warpdrive_pause\CONTINUATION.md`. The broad Find red at `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_140737` remains the active blocker, but focused and narrowed repros are green: focused (`2026-07-01_141539`), predecessor pair (`2026-07-01_141609`), 26-case Find block (`2026-07-01_141731`), diagnostic focused (`2026-07-01_142342`), and diagnostic 26-case block (`2026-07-01_142436`). Debug build with expanded Find sort/layout diagnostics passed at `Z:\src\RedSalamander\.build\logs\msbuild-20260701_142132_723.log` (0 warnings / 0 errors). Resume by rerunning broad Commands with fail-fast and redirected output to capture the expanded assertion. Do not patch product behavior until that evidence identifies the failing live/persisted sort, column order, width, row order, or resize predicate.

- [2026-07-01 · Task 11] Pause archive refreshed at `Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_145912_folderview_warpdrive_pause\CONTINUATION.md`. The broad Find diagnostic rerun crashed before writing a Commands archive (`.build\logs\commands-broad-find-diag-20260701_142923.out.log`, `textinputframework.dll`, `0xC0000005`, offset `0x000000000006a419`, dump `RedSalamander.exe.90048.dmp`). Narrowing found a 107-case FileOps+Navigation+Dialogs crash archive at `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_144819_fileops_navigation_dialogs_tsf_crash`; the first observable selftest failure before the crash was `cmd_pane_navigation_context_menu_keeps_navigation_shell_stable`, then the trace stopped in `cmd_pane_navigation_create_directory_prompt_keeps_navigation_shell_stable` with the same TSF fault (`RedSalamander.exe.74560.dmp`). A diagnostic-only Navigation context-menu assertion patch was built successfully at `Z:\src\RedSalamander\.build\logs\msbuild-20260701_145355_735.log` (0 warnings / 0 errors), and the exact 107-case filter then passed with fail-fast at `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_145653` (107 passed / 0 failed / 0 skipped). Resume by rerunning the 107-case filter once more to classify intermittency, then rerun broad Commands with redirected output; copy crash `last_run`/WER/dumps before any subsequent run. Do not patch product behavior until the expanded predicate or crash stack identifies a root cause, and do not move this plan to Done until broad Commands, Full skip-build, and the final Debug plus test-enabled Release FolderView perf matrices are green and cited.
- [2026-07-02 · Task 11] Pause archive refreshed at `Specs\TestRuns\4cb089111a23\Continuation\2026-07-02_101000_folderview_warpdrive_commands_av_baton\CONTINUATION.md`. `DxUiTests` is green after the MenuBar whitespace-insensitive source guard fix (`.build\codex-runs\dxuitests_after_menubar_guard_warningfix_20260702.log`), and the exact 96-case Commands slice from the previous prompt/focus baton passed again (`.build\codex-runs\commands_slice_532_627_repeat_after_dxuitests_20260702.log`). The first broad Commands rerun stopped near `cmd_pane_find_dialog_restored_combined_view_state_action_buttons_activate_selection`, but focused, suffix-5, suffix-15, suffix-45, and exact prefix-456 reruns all passed. The current blocker is the second broad Commands rerun: wrapper exit `1`, selftest child exit `-1073741819` (`0xC0000005`), WER faulting module `textinputframework.dll`, offset `0x000000000006a419`, and trace stopping at `cmd_pane_selection_invert` after 775 cases. The latest local dump pointer is `C:\Users\eric\AppData\Local\CrashDumps\RedSalamander.exe.73392.dmp`; it was intentionally not copied into the repo because it is about 120 MB. Resume by running focused `cmd_pane_selection_invert`, then the archived last-10 and last-20 filters, then a prefix ending at `cmd_pane_selection_invert` if needed. Do not move this plan to Done until broad Commands, Full, and final Debug plus test-enabled Release FolderView perf matrices are green and cited.
- [2026-07-02 · Task 11] Pause archive refreshed at `Specs\TestRuns\4cb089111a23\Continuation\2026-07-02_102745_folderview_warpdrive_prefs_commands_baton\CONTINUATION.md`. The focused `cmd_pane_selection_invert`, last-10, and last-20 slices all passed, so the prior selection-invert AV remains order-dependent and did not reduce locally; the huge prefix-to-invert command route was discarded as invalid evidence because it did not produce a trustworthy exit/result artifact. A broad Commands rerun then failed normally at `cmd_connection_manager_window_protocol_churn_keeps_form_and_uia_stable`, but focused, immediate 5-case, plugin-config-plus-Connection, and last-35 suffix reruns passed. The next broad rerun got past Connection Manager and failed at `cmd_preferences_dialog_plugins_page_exposes_live_uia_grid_selection` with `title='S3'`, `pluginItemSelected=true`, and `pluginsDetailsActive=true`; focused and 14-case predecessor reruns passed. A Preferences-family fail-fast reproduced a later Plugins/custom-path failure at `cmd_preferences_dialog_plugins_custom_paths_long_run_list_scrolling_stays_bounded` (`active category changed unexpectedly`), while the focused case passed. Resume by inspecting the Preferences Plugins root/detail contract around `DebugSelectPreferencesCategory(kPrefCategoryPlugins)`, `pluginsDetailsActive`, plugin/custom-path selected state, and `DebugScrollPreferencesPluginsCustomPathsListByWheelDetents(...)` before patching. Do not move this plan to Done until broad Commands, Full, and final Debug plus test-enabled Release FolderView perf matrices are green and cited.
- [2026-07-02 · Task 11] Pause archive refreshed at `Specs\TestRuns\4cb089111a23\Continuation\2026-07-02_103700_folderview_warpdrive_prefs_root_baton\CONTINUATION.md`. The fresh four-case Plugins root/detail probe passed, but the Preferences-family fail-fast now reproduces an earlier `cmd_preferences_dialog_viewers_live_search_dx_interaction` silent failure (14 passed / 1 failed / 158 skipped, bare `failed`). Code inspection found the likely Plugins product-contract leak: `PreferencesDialog::SelectCategory(...)` clears `pluginsSelectedPlugin`, `pluginsSelectedPluginId`, and `pluginsDetailsActive`, but does not clear `pluginsRetainedSelectedPluginId`, while the real tree category-root delegate does clear it and the Plugins pane can restore stale selection from that retained ID. The current resume slice is: clear retained plugin ID in `SelectCategory(...)`, harden the Viewers live-search setup to use deterministic `DebugSelectPreferencesCategory(kPrefCategoryViewers)` or at minimum emit diagnostic `state.Require(...)` output, rebuild Debug, rerun the focused Viewers case, the four-case Plugins probe, then `cmd_preferences_dialog_` fail-fast before broad Commands. Do not run multiple `Run-AllTests.ps1` invocations in parallel because they share `SelfTest\last_run`.
- [2026-07-02 · Task 11] Pause archive refreshed at `Specs\TestRuns\4cb089111a23\Continuation\2026-07-02_110950_folderview_warpdrive_compare_tab_baton\CONTINUATION.md`. The retained Plugins root-state fix and Viewers live-search setup hardening are in the dirty worktree; Debug build, focused Viewers live-search, Plugins four-case probe, and full `cmd_preferences_dialog_` family all passed. Broad Commands first failed later at `cmd_connection_credential_prompt_theme_cycle_keeps_surface_legible`, but focused and immediate-predecessor reruns passed. A second broad Commands rerun failed later at `cmd_compare_directories_options_tab_traversal_live_dx_interaction` after 522 passed / 1 failed / 255 skipped: the `Compare attributes` Tab step routed to `RedSalamander.CompareOptions.DxHost`, scrolled from `0/978` to `93/978`, then reported `actualFocus=0`. Resume by focused-running that Compare Options case, then its immediate predecessor cluster, then widening with the archived case-window file before deciding whether this is a selftest focus/scroll race or a product tab-routing bug. Do not move this plan to Done until broad Commands, Full, and final Debug plus test-enabled Release FolderView perf matrices are green and cited.
- [2026-07-02 · Task 11] Pause archive refreshed at `Specs\TestRuns\4cb089111a23\Continuation\2026-07-02_114847_folderview_warpdrive_editnew_baton\CONTINUATION.md`. The Compare Options broad red passed focused, in its immediate predecessor cluster, and in the archived `506..522` window; the next Plugin Configuration broad red passed focused and in its 7-case predecessor cluster; the next Credential Prompt broad red passed focused, in its 5-case credential cluster, in the exact `178..213` window, and in the full `0..213` prefix. Broad Commands then failed later at `cmd_pane_editNew_prompt_rejects_invalid_and_existing_names` after 641 passed / 1 failed / 136 skipped, and the focused case now fails the same way: `Edit New should show a validation message for every invalid file name.` Resume by inspecting `Commands.SelfTest.Dialogs.cpp` around the `missingValidationForNames` probe or adding diagnostic output naming the missing invalid/existing-name inputs before changing product behavior. Do not move this plan to Done until broad Commands, Full, and final Debug plus test-enabled Release FolderView perf matrices are green and cited.
- [2026-07-02 · Task 11] Pause archive refreshed at `Specs\TestRuns\4cb089111a23\Continuation\2026-07-02_124500_folderview_warpdrive_tracecleanup_tsf_crash_baton\CONTINUATION.md`. The Edit New validation red was fixed as a selftest race against the posted prompt-confirm command; focused and paired Edit New runs are green. Temporary Archive Pack trace diagnostics were removed, the source search for `archive_pack_*` trace markers and old icon diagnostics returns no matches, Debug rebuild is green, and focused Pack prompt is green without traces. The follow-up broad Commands run without Pack traces crashed at `cmd_pane_pack_prompt_uses_dxui_unique_archive_and_packer_extensions` with child exit `-1073741819` (`0xC0000005`), WER faulting module `textinputframework.dll`, offset `0x000000000006a419`, latest dump pointer `C:\Users\eric\AppData\Local\CrashDumps\RedSalamander.exe.73980.dmp`; `commands/results.json` was not written. Resume by inspecting the Archive Pack prompt debug close lifecycle against the posted-close prompt pattern, then run the archived four-case Pack window before another full broad Commands rerun. Do not move this plan to Done until broad Commands, Full, and final Debug plus test-enabled Release FolderView perf matrices are green and cited.
- [2026-07-02 · Task 11] Pause archive refreshed at `Specs\TestRuns\4cb089111a23\Continuation\2026-07-02_135700_folderview_warpdrive_perfscope_crash_baton\CONTINUATION.md`. The user pointed to the canonical RedSalamander crash-handler folder (`C:\Users\eric\AppData\Local\RedSalamander\Crashes`); the paired sidecar stacks for `p73832` and `p33336` prove both the pasted crash and the later column-audit/broad crash share the same root: `Debug::Perf::Scope::~Scope` writing JSONL from `FolderView::ProcessThumbnailLoadQueue` after `perf.SetDetail(request.fullPath.parent_path().native())` left `_detail` pointing at a dead temporary. `Debug::Perf::Scope` now owns metric/detail strings when ETW or JSONL is enabled, preserving the disabled fast path. A repo-local reusable skill was added at `.agents\skills\red-salamander-crash-forensics\SKILL.md` and now checks `LocalAppData\RedSalamander\Crashes` first for paired `.dmp`/`.txt` artifacts. Verification: Debug build `Z:\src\RedSalamander\.build\logs\msbuild-20260702_134221_182.log` passed with 0 warnings/errors; focused `folderView_column_widths_audit` passed (1/0/0) in `.build\codex-runs\folder_column_widths_audit_after_perfscope_20260702.log`; `folderView_` prefix passed (33/0/0) in `.build\codex-runs\folderview_prefix_after_perfscope_20260702.log`; broad Commands no longer crashed but failed normally at `cmd_preferences_dialog_category_tree_handles_reverse_keyboard_navigation` after 382 passed / 1 failed / 395 skipped (`.build\codex-runs\commands_broad_after_perfscope_20260702.log`), and that Preferences case passed focused immediately afterward (`.build\codex-runs\preferences_reverse_keyboard_focused_after_perfscope_20260702.log`). Resume with another broad Commands rerun; if the Preferences focus loss repeats, diagnose it as the current non-FolderView blocker. Do not move this plan to Done until broad Commands, Full, and final Debug plus test-enabled Release FolderView perf matrices are green and cited.
- [2026-07-03 · Task 11] Pause archive refreshed at `Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_214957_folderview_warpdrive_full_red_dxui_viewerpe_baton\CONTINUATION.md`. Since the 20:54 archive, `ViewerPETests` was rebuilt with extra VLC HUD child-window diagnostics and focused VLC, focused shell-combo, and direct broad ViewerPE all passed. Test inventory drift was corrected from 229 to 230 CompareDirectories registrations in `Tools\Tests\TestInventory.Tests.ps1` and `Specs\Testing\Testing_TestCoverage.md`; focused inventory Pester passed and Tools Pester passed with the explicit `$result.FailedCount` wrapper (`Passed=115 Failed=0 Skipped=0 Total=115`). The proper Full run is still red: `.\Tools\Run-AllTests.ps1 -Suite Full -TimeoutMultiplier 2` exited `1` with 1098 passed / 2 failed / 52 skipped / 1152 total. CompareDirectories, Commands, FileOperations, and ToolsPester were green; the remaining red gates are `DxUiTests` (`TestMenuBarHoverMessageSwitchesRootWhilePopupOverlapsMenuBar`: `overlapping View popup exposes debug state`) and `ViewerPETests` (`TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly fresh harness run passes`). Resume by reproducing/root-causing DxUi `--suite=Menu` and the ViewerPE fresh-harness wrapper before patching. Do not move this plan to Done until these two Full gates, final Debug and test-enabled Release FolderView perf matrices, durable spec migration, and plan move are complete.
- [2026-07-02 · Task 11] Commands convergence archive refreshed at `Specs\TestRuns\4cb089111a23\Continuation\2026-07-02_215000_folderview_warpdrive_commands_green_baton\CONTINUATION.md`. Debug build is clean with 0 warnings/errors (`.build\codex-runs\msbuild_after_nav_ime_failure_diagnostics_20260702.log`). Focused/adjacent gates for hot reload, Preferences reverse navigation, Compare Options tab traversal, and NavigationView path activation are green. The first full broad Commands run completed without crashing but had one order-sensitive NavigationView IME-start diagnostic miss (775 passed / 1 failed / 2 skipped); focused, local cluster, and app-shell slice repros all passed. The final broad Commands rerun is green: `.build\codex-runs\commands_broad_after_nav_ime_diag_cleanbuild_20260702.log`, 776 passed / 0 failed / 2 opt-in ViewerSpace skips, exit 0, about 14m42s. No new RedSalamander crash-handler dump appeared after the final broad run. Do not move this plan to Done yet: remaining gates are Full (or accepted Full skip-build), final Debug plus test-enabled Release FolderView perf matrix confirmation if current changes remain, durable spec migration, and Done move.
- [2026-07-02 · Task 11] Full-suite red baton archived at `Specs\TestRuns\4cb089111a23\Continuation\2026-07-02_223436_folderview_warpdrive_full_red_baton\CONTINUATION.md`. The post-Commands Full skip-build gate ran for about 42m and failed with `1091 passed / 8 failed / 52 skipped`; `Commands` stayed green inside Full (`776 passed / 0 failed / 2 skipped`) and `FileOperations` stayed green (`102 passed / 0 failed / 20 skipped`). Remaining red gates are outside the FolderView perf slice: 4 CompareDirectories search/service cases, ViewerVLC HUD focus inside `ViewerPETests`, missing `.build\x64\Debug\RedConfigureTests.exe`, `RedSalamanderMonitorEtwLatency` exit `1`, and `ToolsPesterTests` exit `1`. Resume from that baton by running those focused gates directly before another Full rerun. Do not move this plan to Done until Full is green and the final Debug plus test-enabled Release FolderView perf matrices are rerun and cited from the current HEAD.
- [2026-07-03 · Task 11] Compare red-gate continuation archive refreshed at `Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_083500_folderview_warpdrive_compare_green_viewerpe_red_baton\CONTINUATION.md`. Branch `codex/folderview-warpdrive` remains `25` ahead / `0` behind `origin/master` at HEAD `275c04034`. The four Full-red CompareDirectories search/service cases are now focused-green after selftest-source guard hardening and isolated foreground search-service storage/request budgeting: `compare_red_gates_green_after_live_progress_20260703.log` reports `4 passed / 0 failed / 0 skipped`. Debug builds after each slice passed. Direct `ViewerPETests.exe` still exits `1`, but the old ViewerVLC HUD/focus suspicion did not reproduce; the direct log shows all visible fresh child runs pass, including `TestViewerTextDiffModesAndPlaceholders`, then the parent prints `ViewerPETests failed.` Resume by inspecting the ViewerPE parent fresh-process aggregation before returning to missing `RedConfigureTests.exe`, `RedSalamanderMonitorEtwLatency`, `ToolsPesterTests`, Full, and final Debug plus test-enabled Release FolderView perf matrices. Do not move this plan to Done until Full is green and final perf evidence is rerun and cited from the current HEAD.
- [2026-07-03 · Task 11] Broad Commands is green again on the current worktree after selftest wait hardening for order-sensitive Preferences category-tree boundary navigation and status-bar focused-item updates. Build `Z:\src\RedSalamander\.build\codex-runs\msbuild_redsalamander_after_nav_pref_waits_fix_20260703.log` passed with `0 warning(s), 0 error(s)`. Focused/window gates passed: `runall_commands_pref_boundary_focused_after_nav_pref_waits_20260703.log` (1/0), `runall_commands_nav_status_focused_after_nav_pref_waits_20260703.log` (1/0), `runall_commands_nav_window11_after_nav_pref_waits_20260703.log` (11/0), and `runall_commands_pref_boundary_window_after_nav_pref_waits_20260703.log` (11/0). Final broad evidence: `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_123938` and `.build\codex-runs\commands_broad_after_nav_pref_waits_20260703.log` report `776 passed / 0 failed / 2 skipped`. The plan remains WIP: remaining closeout gates are ViewerPE direct, Monitor ETW direct, Full skip-build, final Debug plus test-enabled Release FolderView perf matrices, and then Done move.
- [2026-07-03 · Task 11] Continuation archive refreshed at `Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_142650_folderview_warpdrive_compare_cache_crash_baton\CONTINUATION.md`. ViewerPE direct and Monitor ETW direct are now green on the current worktree; proper Full with build got past those gates but stayed red on one CompareDirectories case and one Commands case (`.build\codex-runs\full_with_build_after_compare_monitor_green_20260703.log`, 1097 passed / 2 failed / 52 skipped). Focused fixes for those two reds now pass: `decision_cache_eviction_budget_pending_wide_tree` via the quiet pending-update drain helper and `cmd_pane_clipboardPaste_ignores_unfocused_navigation_edit` via a navigation-edit host/input-open predicate. The follow-up full CompareDirectories rerun crashed before writing `compare\results.json`: `.build\codex-runs\compare_full_after_quiet_drain_20260703.log` reports child exit `-1073741819`, and the paired crash sidecar `C:\Users\eric\AppData\Local\RedSalamander\Crashes\RedSalamander-20260703-142049-p98100.txt` shows an access violation in `CompareDirectoriesSession::ComputeDecisionForFolder` while probing the content-decision cache (`std::_Hash<...ContentCompareKey...>::find`, `CompareDirectoriesEngine.cpp:3551`). Resume by investigating that content-decision-cache concurrency/lifetime crash first, then rerun full CompareDirectories, the Commands clipboard/navigation-edit window and broad Commands, proper Full, and finally the Debug plus test-enabled Release FolderView perf matrices. Do not move this plan to Done until Full and final perf evidence are green and cited from this HEAD.
- [2026-07-03 · Task 11] Continuation archive refreshed at `Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_150535_folderview_warpdrive_worker_join_commands_status_baton\CONTINUATION.md`. The content-decision-cache crash root cause is now patched in the dirty worktree: `CompareDirectoriesSession` explicitly wakes and joins `_scanWorkers` before `_contentCompareWorkers`, because implicit reverse member destruction would otherwise destroy content-compare state before joining scan workers. Verification is green for the Debug `RedSalamander` rebuild (`.build\codex-runs\msbuild_redsalamander_after_compare_worker_join_20260703_rerun.log`, 0 warnings / 0 errors), the new shutdown guard (`compare_worker_shutdown_join_guard_20260703.log`, 1/0/0), the prior pending-wide-tree red (`compare_pending_wide_tree_after_worker_join_20260703.log`, 1/0/0), full CompareDirectories (`Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-03_144451`, 208 passed / 0 failed / 30 skipped), and the Commands clipboard/navigation-edit window (`Specs\TestRuns\4cb089111a23\Commands\2026-07-03_144529`, 5/0/0). Broad Commands now fails normally, without a new dump, at `cmd_pane_navigation_status_bar_keeps_navigation_shell_stable` (`Specs\TestRuns\4cb089111a23\Commands\2026-07-03_150248`, 775 passed / 1 failed / 2 skipped): the failure snapshot has `currentPath=''` and no status-bar snapshot while folder item state is otherwise intact. Resume by running that focused case, then the navigation/status predecessor window, before another broad Commands, proper Full, and final Debug plus test-enabled Release FolderView perf matrices. At archive time this branch was `25` ahead / `4` behind `origin/master`, so recheck/rebase before final gates.
- [2026-07-03 · Task 11] Continuation archive refreshed at `Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_162914_folderview_warpdrive_speedlimit_focus_baton\CONTINUATION.md`. The 16:15 broad crash pointed at the speed-limit prompt live-Dx interaction, but focused speed-limit, two-case speed-limit, and tail `531..552` repros all passed (`Specs\TestRuns\4cb089111a23\Commands\2026-07-03_162111`, `2026-07-03_162136`, `2026-07-03_162221`). The wider `517..552` slice failed earlier at `cmd_pane_fileops_issues_pane_hide_restores_folder_focus` (`Specs\TestRuns\4cb089111a23\Commands\2026-07-03_162631`, 35 passed / 1 failed) with `focusHwnd=0x0` while both speed-limit cases later passed. No new crash-handler dump appeared after the focused/tail probes; the latest paired dump remains the 16:14 `p81732` sidecar already archived by the 16:15 baton. Resume by inspecting `FileOperationsIssuesPaneState::SelfTestFocusGrid` / hide-close focus restoration and reproducing focused plus `517..546` / `531..546` predecessor slices before patching. Do not weaken the final behavior assertion that hiding a focused issues pane restores folder focus.

### Pre-implementation verification (done 2026-06-28, against HEAD `045a1f773`)

These were confirmed by reading the current code so the plan's instructions are
safe to follow. Treat them as ground truth until a commit changes them.

**Verified accurate — the plan can rely on these:**

- The `_forceFullRenderOnNextPaint` member **already exists** (`FolderView.h:961`; the success present paths reset it at `FolderView.Rendering.cpp:1919` and `:1943`). Task 1 Step 2 reuses it correctly — do **not** add a duplicate flag.
- The device-loss failure branches to patch are `FolderView.Rendering.cpp:1876-1881` (EndDraw), `:1912-1918` (`Present1`), and `:1936-1942` (legacy `Present`). Each currently does only `ReleaseSwapChain(); EnsureSwapChain();` — exactly as the plan describes.
- `CreateShellShortcut(...)` exists at `FolderView.FileOps.cpp:272`; `PasteShortcutFromClipboard()` at `:816`; the UI-thread cost is the per-source `CreateShellShortcut` loop at `:857` (which runs `GenerateShortcutPath` + `IPersistFile::Save`). Task 4's worker calling `CreateShellShortcut` is correct.
- `kFolderViewPasteShortcutComplete = WM_APP + 0x30A` is genuinely the next free slot. Used offsets in `Common/WindowMessages.h` are `0x300`–`0x309` (note `0x305` = `kNetworkConnectivityChanged`, out of numeric order). `0x30A` is unused.
- `DirectoryInfoCache::NotifyFolderContentsChanged(...)` exists (`DirectoryInfoCache.h:139`). Task 4 Step 4 uses it correctly.
- `DebugThumbnailProviderMode` plumbing exists (`FolderView.h:385`, atomic at `:1533`, consumed at `FolderView.Icons.cpp:975`).
- `folderView_thumbnail_valid_images_shell_fail` (Task 3 verify) exists at `Commands.SelfTest.ViewCommands.cpp:19886`. `folderView_thumbnail_scroll_stress` exists at `:19894`.
- `Specs/Testing/Testing_PerformanceValidation.md`, `Specs/Testing/Testing_TestCoverage.md`, and `Tools/Run-AllTests.ps1` all exist.

**Corrections / gaps — fix while implementing:**

- **C1 [Task 4 Step 6 — corrected]:** The original grep was mis-scoped to `Commands.SelfTest.ViewCommands.cpp`. The clipboard-paste selftests EXIST in `RedSalamander/SelfTest/Commands/Commands.SelfTest.ShellCommands.cpp`: `cmd_pane_clipboardPaste_uses_preferred_move_effect` (def `TestPaneClipboardPasteUsesPreferredMoveEffect` ShellCommands.cpp:2046, registered :2997), `cmd_pane_clipboardPasteShortcut_creates_unique_links` (def :2776, registered :3011), and `cmd_pane_clipboardPasteShortcut_rejects_missing_clipboard_paths` (def :2893, registered :3013). Task 4 adds `cmd_pane_clipboardPasteShortcut_returns_before_worker_complete` and `cmd_pane_clipboardPasteShortcut_close_does_not_wait_for_worker`. `cmd_pane_clipboardPasteShortcut_creates_unique_links` is a LOAD-BEARING regression gate for Task 4: it seeds CF_HDROP, sends IDM_PANE_CLIPBOARD_PASTE_SHORTCUT, waits up to 3000ms for the pane to refresh with the .lnk files, and asserts focus on the last created link. Moving creation to a worker changes this from synchronous to async, so the refactor MUST keep both shortcut cases green (refresh + last-link focus must now occur via `OnPasteShortcutComplete`). The verify line that runs `cmd_pane_clipboardPaste_uses_preferred_move_effect` is valid as-written — keep it. Add the new cases to ShellCommands.cpp, NOT ViewCommands.cpp.
- **C2 [Task 3 Step 4]:** None of the existing `DebugThumbnailProviderMode` values (`Shell`, `ForceFallback`, `ForceSyntheticSuccess`, `ForceShellFailureAllowWic`, `ForceSyntheticWideSuccess`) can simulate the two distinct conditions Step 4 needs. TWO mechanisms are required, not one: (a) a counter proving the visible path issued zero provider-allowed lookups — surface `thumbnailShellProviderAllowedCount` via `PaneViewOptionsDebugSnapshot` (Step 3) and assert `== 0`; and (b) a BLOCKING latency-injection hook on the provider-allowed branch, because today there is no delay/block hook anywhere in the thumbnail path: `ExtractShellThumbnailBitmap` (FolderView.Icons.cpp:80-102) takes no mode and cannot block, and the `ENABLE_TESTS` branch (FolderView.Icons.cpp:974-1001) only swaps outcomes. A mode that merely asserts "provider not called" proves the path skipped the provider but proves nothing about teardown timing under provider latency. Task 0 now creates the shared `SelfTestLatency` hook family; Task 3 must wire `ShellThumbnailProviderAllowed` to the provider-allowed branch.
- **C3 [Task 1, correctness risk] — CONFIRMED at HEAD `045a1f773`:** `DiscardDeviceResources()` (`FolderView.Rendering.cpp:657`) nulls every pointer the recovery guards depend on (`_d2dContext`, `_d2dDevice`, `_d2dFactory`, `_d3dContext`, `_d3dDevice`) and releases both swap chains via `ReleaseSwapChain()`. `EnsureDeviceResources` early-returns only while the device/context/factory are non-null and `EnsureSwapChain` early-returns only while `!_d3dDevice`; `Render()` re-runs the ensure path every frame, so after discard+invalidate the device/factory/swap chain are all rebuilt — recovery does NOT silently no-op. No code change is needed. Still add a null-pointer assertion (e.g. `_d3dDevice`/`_d2dContext`/`_d2dFactory` are null immediately after the discard) in `folderView_render_device_loss_recovers` as a regression guard against a future edit dropping a `reset()`.
- **C4 [Tasks 0/5/6/7, why they exist]:** Measurement ground truth from the archive review — (a) `folderView_perf_large_folder_baseline` is **241 items** (`itemCount: 241`), not a large folder; (b) the only FolderView frame run with statistical mass is `2026-05-19_191156` at **n=6,171** (`folder.frame.total_us` p50≈64ms / p95≈85ms / p99≈112ms / max≈307ms) but it is **pre-sink-fix and inflated ~4–16×** — do NOT use it for budgets; (c) **no post-sink-fix run reaches the ≥200-sample target Task 5 requires** — among runs whose commit actually contains the sink fix (`54f21f3c8`, 2026-06-19) the largest is only ~n=20 (`2026-06-20_134225`) and the dedicated scroll/overlay cases produce ~2–9 frame samples/run; the bigger archives (n=55 at `2026-06-19_202636`, etc.) are PRE-sink-fix (`6d6fc137c`) and inflated; (d) **exactly one Release-build perf run exists** (`Specs/TestRuns/4cb089111a23/Commands/2026-05-19_170631`, n=107 `folder.frame.total_us` / 452 `folder.frame.*` rows) but every one of its 4621 rows is **mislabeled `build="Debug"`** despite `.build\x64\Release\` paths (the hardcoded-label bug, Task 0) and it is pre-sink-fix — no line anywhere carries `build="Release"`, so there is no usable post-sink-fix Release run; (e) all evidence is one machine. This is exactly what Task 0 and Tasks 5–7 must fix before any threshold is trustworthy.
- **C5 [Task 8]:** At plan audit time, no `folder.refresh.*` metric existed in any archive (a repo-wide scan matched only specs/prose, not emit sites). `folderView_perf_directory_change_storm` mutates ~200 created files, renames 100, deletes 100, plus 20 transient dirs (Commands.SelfTest.ViewCommands.cpp:18179-18305); steady state is **101 items** (`finalItemCount: 101`, :18317), **NOT 5,000**. It emitted `directorycache.post_refresh_count` (:18297), `folder.directory_change_storm_mutation_us` (:18303), and `folder.directory_change_storm_settle_us` (:18305) — **not** `folder.frame.*`. Mutation-to-paint refresh latency was unmeasured; Task 8 Step 1 is net-new instrumentation, not a tweak.

### Plan-level gaps found during review (2026-06-28)

These are about the plan's scope/sequencing, not individual code lines. Decide on
each before closeout.

- **G1 — Audit findings are NOT covered here.** This plan supersedes the two perf
  backlogs only. It does **not** cover the Tier-1 data-loss / security defects in
  `Specs/Reviews/FolderView-Audit.md` (directory move-onto-self destruction, drop
  point ignored → wrong-target MOVE, CF_HDROP/clipboard OOB reads,
  attacker-controlled `reserve` → `terminate`). Those are more urgent than
  performance. Track them in a separate remediation plan; do not let "this plan
  replaced the FolderView backlogs" imply they are handled.
  **Now tracked:** `Specs/Plans/WIP/Operation_IronLedger_FolderViewDataIntegrityDropClipboard_2026-06-28.md`
  (codename Iron Ledger) owns this cluster — D2/D3/D4/D5/D6/D7/D8/B4/B5/B6/B8/S1/S2/S3 plus
  6 new findings, grouped by root cause. D1 was re-verified as already-mitigated on the
  production path. R1's *fix* remains owned by this plan's G3 (Iron Ledger only verifies it).
- **G2 — The "any circumstance" config matrix is now scheduled.** Task 6 Step 6
  records build flavor, GPU/driver, refresh rate, DPI/scale, local-console/RDP
  status, WARP status, and blocked dimensions in every representative archive.
- **G3 — Fix the `_itemsFolder` worker torn-read while you are in that code.**
  Audit R1: `ProcessIconLoadQueue`/`ProcessThumbnailLoadQueue`
  (`FolderView.Icons.cpp:539` / `:911`) read the non-atomic `_itemsFolder`
  `std::filesystem::path` off the UI thread for perf detail while the UI thread
  move-assigns/clears it. Tasks 3 and 9 already touch this worker — snapshot the
  folder under `_enumerationMutex` at queue time (or use the per-request
  `fullPath`) and stop reading `_itemsFolder` off-thread.
- **G4 — Regression budgets (Task 7) must be machine-keyed; align to the existing
  convention by citation.** `Testing_PerformanceValidation.md` already requires
  same-machine comparison (lines 83, 87) and a same-machine archived baseline
  (line 170), and `Specs/TestRuns/README.md:64` already keys archive folders to a
  per-machine hash. None of these states a rule for a COMMITTED budget threshold
  file, so a committed `folder.frame.total_us.p95.max` will false-fail on
  slower/CI hardware. Task 7 must make budgets machine-keyed to a named machine
  hash (cite `Specs/TestRuns/README.md:64`); other machines run quality-only
  checks or scale by a recorded machine factor. Task 11 Step 2 must add this
  budget-file machine-keying rule to `Specs/Testing/Testing_PerformanceValidation.md`
  so it becomes a durable contract, not just a plan note.
- **G5 — The slow/virtual-FS harness is split into prerequisites and fixtures.**
  Task 0 creates the shared bounded latency hook family used by thumbnail, icon,
  and paste no-freeze tests. Task 6 Step 4 builds the representative slow/virtual
  fixture around those hooks. Do not add task-local delay APIs.

### Deep-review findings (verified 2026-06-28, multi-agent pass)

Code-cited findings from a second, adversarially-verified review pass. These refine the
C/G items above and back each task edit with exact `file:line` evidence.

- [2026-06-28 · Task 4 / Finding C1] CORRECTION: clipboard-paste selftests DO exist in Commands.SelfTest.ShellCommands.cpp (cmd_pane_clipboardPaste_uses_preferred_move_effect def :2046 reg :2997; cmd_pane_clipboardPasteShortcut_creates_unique_links def :2776 reg :3011; cmd_pane_clipboardPasteShortcut_rejects_missing_clipboard_paths def :2893 reg :3013). Plan C1's 0-hit grep was mis-scoped to ViewCommands.cpp. The creates_unique_links case is a load-bearing async-refactor regression gate. Evidence: RedSalamander/SelfTest/Commands/Commands.SelfTest.ShellCommands.cpp:2046,2776,2893,2997,3011,3013.
- [2026-06-28 · Task 4] CreateShellShortcut needs a COM apartment: it calls CoCreateInstance(CLSID_ShellLink) and IPersistFile, working today only because PasteShortcutFromClipboard runs on the COM-initialized UI thread. A worker must CoInitializeEx(COINIT_MULTITHREADED) like EnumerationWorker. Evidence: RedSalamander/FolderView.FileOps.cpp:279,303,816; RedSalamander/FolderView.Enumeration.cpp:276.
- [2026-06-28 · Task 4] PasteShortcutResult.generation must snapshot the existing _enumerationGeneration (FolderView.h:1412), captured on the UI thread before submit; EnumerateFolder() bumps it (Enumeration.cpp:985), so the staleness compare in OnPasteShortcutComplete must run before the refresh. _currentFolder is std::optional<path> (FolderView.h:848). Evidence: RedSalamander/FolderView.h:848,1412; RedSalamander/FolderView.Enumeration.cpp:985; RedSalamander/FolderView.FileOps.cpp:880-883.
- [2026-06-28 · Task 4] ReadPreferredDropEffectClipboard has an OOB read: GlobalLock then *effect at FolderView.FileOps.cpp:267 with no GlobalSize>=sizeof(DWORD) guard (Audit S3). Caller defaults to COPY via value_or. Evidence: RedSalamander/FolderView.FileOps.cpp:261-269,738.
- [2026-06-28 · Task 2] RecreateThemeBrushes() (FolderView.Rendering.cpp:269) is re-entered from SetTheme (FolderView.cpp:1044) with no prior discard and creates brushes via .addressof(); new _hoverBrush/_selectedItemTextBrush must be reset in its top reset block (alongside :286-287), not only in DiscardDeviceResources, or theme changes leak a live COM pointer. Evidence: RedSalamander/FolderView.Rendering.cpp:269,286-287; RedSalamander/FolderView.cpp:1044.
- [2026-06-28 · Task 2] Draw-loop guard should be a test-only counter (precedent _debugIncrementalSearchEffectUpdateCount FolderView.h:1422, incremented Rendering.cpp:1873, surfaced via DebugGetWarmPerfSnapshot FolderView.h:260, asserted Commands.SelfTest.ViewCommands.cpp:15933-15944) wrapping CreateSolidColorBrush inside DrawItem (Rendering.cpp:2126), not a Debug::Perf metric (Emit only writes JSONL/ETW, no in-process query — Common/Helpers.h:1404). Evidence: RedSalamander/FolderView.h:260,1422; RedSalamander/FolderView.Rendering.cpp:1873,2126.
- [2026-06-28 · Task 3] Cached-only (SIIGBF_INCACHEONLY) alone does not make teardown close-safe (Audit B1): StopEnumerationThread joins the worker (Enumeration.cpp:245), request_stop cannot interrupt an in-flight synchronous GetImage (Icons.cpp:99 via call site :1004), and ProcessThumbnailLoadQueue loops on _thumbnailLoadingActive without a stop token (Icons.cpp:915). Evidence: RedSalamander/FolderView.Enumeration.cpp:245; RedSalamander/FolderView.Icons.cpp:99,915,1004.
- [2026-06-28 · Task 3] Provider-allowed gate metric needs full snapshot plumbing: add counters to ThumbnailLoadStats (FolderView.h:1514), surface via ThumbnailDebugSnapshot (:429) and PaneViewOptionsDebugSnapshot (FolderWindow.h:832), copied at FolderWindow.FileSystem.Commands.Part.cpp:10562; the selftest must assert thumbnailShellProviderAllowedCount==0 from the snapshot. Evidence: RedSalamander/FolderView.h:419-450,1514; RedSalamander/FolderWindow.h:811-860; RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp:10562.
- [2026-06-28 · Task 3] No latency/block hook exists in the thumbnail path today: ExtractShellThumbnailBitmap takes no mode (Icons.cpp:80) and the ENABLE_TESTS branch only swaps outcomes (Icons.cpp:974-1001). Finding C2 needs TWO mechanisms (zero-provider-call counter + a blocking latency hook), not one assert-no-provider mode. Evidence: RedSalamander/FolderView.Icons.cpp:80-102,974-1001.
- [2026-06-28 · Task 9] PerformanceTests2 is a DynamicLibrary (PerformanceTests2.vcxproj ConfigurationType=DynamicLibrary), so PerformanceTests2.exe cannot exist; run via vstest.console.exe on PerformanceTests2.dll (Run-AllTests.ps1 does this). This plan's verify step must stay on `PerformanceTests2.dll`. Evidence: Tests/PerformanceTests2/PerformanceTests2.vcxproj:47; Tools/Run-AllTests.ps1:147,219-225.
- [2026-06-28 · Task 9] All 12 Step-2 icon/iconcache metrics already emit in code (sampled: QueryExtensions Enumeration.cpp:662; icons.queue_wait_to_dequeue_us Icons.cpp:603; icons.extract_us Icons.cpp:631); Step 2 is audit+summarization, not adding emits. Evidence: RedSalamander/FolderView.Enumeration.cpp:662; RedSalamander/FolderView.Icons.cpp:603,631.
- [2026-06-28 · Task 9] New pool icon workers must CoInitializeEx(COINIT_MULTITHREADED) themselves: existing extension/per-file pool callbacks (Enumeration.cpp:701-718,835-848) call SHGetFileInfoW without initializing COM in the callback. Evidence: RedSalamander/FolderView.Enumeration.cpp:276,701-718,835-848.
- [2026-06-28 · Task 9] Icon-index lookup blocks on cloud/offline recall: QuerySysIconIndexForPath (IconCache.cpp:1119) runs SHGetFileInfoW on the live path (:1156); the icon worker passes useFileAttributes=false (Enumeration.cpp:844). For offline/recall-flagged items it should pass SHGFI_USEFILEATTRIBUTES (no recall). Evidence: RedSalamander/IconCache.cpp:1119,1156; RedSalamander/FolderView.Enumeration.cpp:844.
- [2026-06-28 · Tasks 0/5/6/7 · Finding C4d] CORRECTION: exactly one Release-built perf run exists (Specs/TestRuns/4cb089111a23/Commands/2026-05-19_170631, 107 folder.frame.total_us / 452 total folder.frame.* rows), but all 4621 lines are mislabeled build=Debug despite .build\x64\Release\ paths; no line anywhere carries build=Release. It is pre-sink-fix and excluded by P2; no usable post-sink-fix Release run exists. Evidence: Specs/TestRuns/4cb089111a23/Commands/2026-05-19_170631/perf/perf_metrics.jsonl.
- [2026-06-28 · Task 0] The selftest perf build label is hardcoded L"Debug" in InitSelfTestRun (SelfTestCommon.cpp:1192); kRedSalamanderBuildFlavor (RedSalamander.cpp:6911-6917) is file-scope and not visible there. Without fixing this, metricQuality.buildConfiguration can never report Release. Evidence: RedSalamander/SelfTest/Common/SelfTestCommon.cpp:1192; RedSalamander/RedSalamander.cpp:6911-6917,7124.
- [2026-06-28 · Task 0] Existing perf analyzer Tools/Show-PerfRuns.ps1 already emits count/p50/p95/p99/max per metric (Measure-Metric, Get-RunFolders over perf_metrics.jsonl); extend it for a sample-count quality gate + non-zero exit + Debug/Release labeling rather than creating a new tool. AnalyzeTestRuns.ps1 reads suite-results JSON, not perf percentiles. Evidence: Tools/Show-PerfRuns.ps1:6,308,350,416.
- [2026-06-28 · Task 8 · Finding C5] CORRECTION: folderView_perf_directory_change_storm settles at 101 items (finalItemCount:101, :18317), not 5000, and emits directorycache.post_refresh_count (:18297), folder.directory_change_storm_mutation_us (:18303), folder.directory_change_storm_settle_us (:18305) — not folder.frame.*. Evidence: RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp:18297-18317.
- [2026-06-28 · Task 8] folder.refresh.request_to_paint_us must reuse the post-Present emit (EmitPendingInputToPaintMetricAfterPresent, Rendering.cpp:1921/1945) but the single-slot _pendingInputToPaintMetric (FolderView.h:939; overwritten FolderView.cpp:84-98) is already shared by scroll input-to-paint and navigation-to-paint (FolderWindow.cpp:962) — a refresh would collide last-writer-wins; give refresh its own slot. itemsPreserved already exists as a local (Enumeration.cpp:1562). Evidence: RedSalamander/FolderView.h:939; RedSalamander/FolderView.cpp:84-98,112-120; RedSalamander/FolderView.Rendering.cpp:1921,1945; RedSalamander/FolderWindow.cpp:962; RedSalamander/FolderView.Enumeration.cpp:1562.
- [2026-06-28 · Task 10] FolderItem has no stable identity and no per-item visual-state generation: refresh preservation keys on displayName (Enumeration.cpp:1566-1576), stableHash32 is a folder+name rainbow seed that changes on rename (FolderView.h:783, set Enumeration.cpp:485), unsortedOrder is reassigned to loop index each enumeration (Enumeration.cpp:1670). V2 must introduce a real identity. Evidence: RedSalamander/FolderView.h:783; RedSalamander/FolderView.Enumeration.cpp:485,1566-1576,1670.
- [2026-06-28 · Task 10] DXGI scroll-rect work must exclude/repaint screen-anchored overlays: incremental-search indicator (Rendering.cpp:1844,1949), error overlay (Rendering.cpp:1846), and hover/selection/focus rects; current scroll does full-client InvalidateRect(nullptr) (Interaction.cpp:395) so a content-only dirty rect is also required. A partial dirty-rect path + fullClientRenderCount instrumentation already exist (Rendering.cpp:1003-1025; FolderView.cpp:1033-1035). Evidence: RedSalamander/FolderView.Rendering.cpp:1003-1025,1844-1846,1949; RedSalamander/FolderView.Interaction.cpp:395; RedSalamander/FolderView.cpp:1033-1035.
- [2026-06-28 · Supersession] PerformanceBacklog_2026-06-28.md was never git-tracked; DxUi_FolderView_Monitor_FuturePerformanceIdeas_2026-05-20.md is already deleted in the working tree; no tracked file except this plan references either name; the three Done TextLayout pilots cite THIS plan as Parent backlog (lines 11/18/12). Plans 009/010 map to Tasks 9/8, not 6/7. Evidence: Specs/Plans/Done/FolderView_TextLayout_MetricPilot_2026-06-19.md:11; .../FolderView_LayoutPassDecomposition_MetricPilot_2026-06-19.md:18; .../FolderView_UpdateItemTextLayouts_Optimization_2026-06-19.md:12; plans/009-icon-pipeline-investigation.md; plans/010-watcher-incremental-refresh-investigation.md.
- [2026-06-28 · Any-circumstance] Coverage disposition: scheduled in Task 6/9 — cloud/OneDrive recall + offline placeholders (icon-index recall, IconCache.cpp:1156 / Enumeration.cpp:844), working-set memory at 50k/100k, theme/DPI/high-contrast relayout while scrolled via `folder.relayout_to_paint_us`, quick-search/sort at huge scale, select-all+scroll worst case for the brush path, and permission-denied/AV slow-then-fail items. Still deferred unless evidence reopens them: live drag-resize repaint storm and dual-pane concurrent heavy load beyond the 480-item `folderView_perf_iconcache_contention` fixture. Evidence: as cited.

### During implementation (append below)

- [2026-06-28 · Task 0] Evidence plumbing landed in task-scoped commit `6fc7851d1`: selftest perf JSONL build labels now use `Debug`/`Release`/`ASan Debug`, the shared `SelfTestLatency` hook family lives under `RedSalamander/SelfTest/Common/`, and `Tools/Show-PerfRuns.ps1` reports `Build`, `P95Quality`, and `P99Quality` with `-FailOnQuality` returning `2` for under-sampled requested metrics. Verification after commit: `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` passed with 0 warnings/errors (`.build/logs/msbuild-20260628_141515_274.log`); `folderView_debug_latency_hooks_consume_once` passed (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_141731`); `folderView_perf_scroll_render_stress` passed (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_141750`); analyzer on `folder.frame.total_us` reported `count=108`, `build=Debug`, `P95Quality=fail`, `P99Quality=fail`, and the hard gate exited `2` as expected. Evidence: `6fc7851d1`; `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_141750/perf/perf_metrics.jsonl`.
- [2026-06-28 · P2] Fresh same-machine pre-fix baselines captured after Task 0 on machine hash `4cb089111a23`, commit `4dd103b25`. Debug (`build=Debug`): `folderView_perf_scroll_render_stress` -> `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_141935`; `folderView_perf_overlay_invalidation_stress` -> `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_142015`; `folderView_perf_large_folder_baseline` -> `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_142035`; `folderView_perf_iconcache_contention` -> `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_142039`; `folderView_perf_directory_change_storm` -> `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_142048`. Test-enabled Release was built with `RSBuildEnableTests=true` and passed with one optimized-build warning (`C4883` in `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.cpp`, log `.build/logs/msbuild-20260628_142106_275.log`). Release (`build=Release`): `folderView_perf_scroll_render_stress` -> `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_142342`; `folderView_perf_overlay_invalidation_stress` -> `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_142454`; `folderView_perf_large_folder_baseline` -> `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_142352`; `folderView_perf_iconcache_contention` -> `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_142528`; `folderView_perf_directory_change_storm` -> `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_142529`. Note: `last_run` can carry prior custom metrics JSON files forward; use each run's `commands_results.json` case name plus `perf/perf_metrics.jsonl` build label to identify the actual case/build.
- [2026-06-28 · P3] Machine profile for the P2 baseline matrix: Windows 11 Pro `10.0.26200` build `26200`, x64; CPU AMD Ryzen 9 9950X3D 16 cores / 32 logical processors; physical memory `66155405312` bytes; motherboard Gigabyte X870E AORUS ELITE WIFI7 ICE. GPUs/displays from `Win32_VideoController`: AMD Radeon(TM) Graphics driver `32.0.21036.18` dated 2025-11-12, `2560x720@60`; NVIDIA GeForce RTX 5080 driver `32.0.16.1062` dated 2026-06-11, `3840x2160@120`. DPI registry `LogPixels=192` (200%), `Win8DpiScaling=0`; console session active for user `eric`, RDP listener present but not active; WARP present via `C:\Windows\System32\d3d10warp.dll` version `10.0.26100.8521`. Evidence commands: `Get-CimInstance Win32_VideoController`, `Win32_Processor`, `Win32_ComputerSystem`, `Win32_OperatingSystem`, `HKCU:\Control Panel\Desktop`, `query session`, and `Get-Item C:\Windows\System32\d3d10warp.dll`.
- [2026-06-28 · Task 1] Device-loss recovery landed in task-scoped commit `afbbf465d`: recoverable `EndDraw`/`Present1`/legacy `Present` failures now route through full device-resource discard, force a full repaint, invalidate the pane, emit `folder.render.device_loss_recovery_count`, and keep non-device-loss failures on the prior swap-chain retry path. The focused selftest injects `DXGI_ERROR_DEVICE_REMOVED` at both `EndDraw` and `Present`, verifies the pane repaints with a valid D2D target, verifies no rendering alert remains, and confirms the discard counter. Verification after commit: `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` passed with 0 warnings/errors (`.build/logs/msbuild-20260628_143959_042.log`); `folderView_render_device_loss_recovers` passed (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_144207`) with trace rows `discarded=yes` for `ID2D1DeviceContext::EndDraw` and `IDXGISwapChain1::Present1`; `folderView_perf_scroll_render_stress` passed (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_144221`). Analyzer on the scroll smoke still reports `folder.frame.total_us` `count=108`, `build=Debug`, `P95Quality=fail`, `P99Quality=fail`; this is expected until Task 5 increases sample counts.
- [2026-06-28 · Task 2] Draw-loop brush reuse landed in task-scoped commit `9506d9104`: `_hoverBrush` and `_selectedItemTextBrush` are cached D2D brushes, reset in `RecreateThemeBrushes()` and `DiscardDeviceResources()`, and updated inside `DrawItem` via `SetColor`; the `DrawItem` body contains no `CreateSolidColorBrush` calls or local `wil::com_ptr<ID2D1SolidColorBrush>` allocations. Verification after commit: `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` passed (`.build/logs/msbuild-20260628_145237_237.log`); `folderView_draw_item_brush_reuse_guard` passed (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_145443`) and asserts the wide-selection + hover transient-brush counter remains `0`; `folderView_perf_scroll_render_stress` passed (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_145458`, `folder.frame.total_us` count=108, build=Debug, P95Quality=fail/P99Quality=fail); `folderView_perf_overlay_invalidation_stress` passed (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_145511`, `folder.frame.total_us` count=226, build=Debug, P95Quality=pass/P99Quality=fail). The sample-quality caveat remains Task 5 scope.
- [2026-06-28 · Task 3] Thumbnail fast-first and close-safe behavior landed in task-scoped commit `6c26367b7`: normal visible thumbnail work now uses cached-only shell lookup (`SIIGBF_INCACHEONLY`), local-shell-backed file-system gating prevents shell/WIC extraction for virtual/non-local providers, WIC source frames above 64 MP are rejected before decode, provider-allowed shell work is value-only/abandonable with a bounded wait and stale-generation postback drop, and `SelfTestLatencyHooks` remains under `RedSalamander/SelfTest/Common/` with qualified includes. Verification after commit: `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` passed with 0 warnings/errors (`.build/logs/msbuild-20260628_152528_077.log`); `folderView_thumbnail_cached_only_no_close_stall` passed (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_152742`) and recorded `thumbnails.cached_extract_us`, `thumbnails.shell_cache_miss_count`, `thumbnails.shell_provider_allowed_count`, `thumbnails.shell_provider_timeout_count`, and `thumbnails.close_to_idle_us` with close-to-idle `9667us` and final pending count `0`; `folderView_thumbnail_scroll_stress` passed (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_152751`); `folderView_thumbnail_valid_images_shell_fail` passed (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_152757`). All three archives record git commit `6c26367b795e73bf63d603b310b286b077f7cc6d`.
- [2026-06-28 · Task 4] RED paste-shortcut responsiveness tests were added before the worker refactor and failed for the intended missing hook/boundary: `cmd_pane_clipboardPasteShortcut_returns_before_worker_complete` failed with "Paste Shortcut worker should consume the shared PasteShortcutSave latency hook" (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_153359`); `cmd_pane_clipboardPasteShortcut_close_does_not_wait_for_worker` failed with "Close-safe Paste Shortcut worker should consume the shared PasteShortcutSave latency hook" (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_153429`). Evidence: both archives record HEAD before `0dc592dcf`.
- [2026-06-28 · Task 4] Paste Shortcut worker refactor landed in task-scoped commit `0dc592dcf`: the command now reads CF_HDROP on the UI thread, submits value-only threadpool work, initializes COM MTA on the worker, consumes `SelfTestLatency::Point::PasteShortcutSave` immediately before each `CreateShellShortcut`, posts `PasteShortcutResult` via `PostMessagePayload`, handles completion on the UI thread, and adds the missing `GlobalSize >= sizeof(DWORD)` guard for `CFSTR_PREFERREDDROPEFFECT`. Implementation correction from testing: same-folder directory watcher refreshes can advance `_enumerationGeneration` before completion, so completion drops only when the current folder no longer matches the target folder; this preserves the existing focus-last-created-shortcut contract while still dropping navigate-away results. Verification after commit: Debug build passed with 0 warnings/errors (`.build/logs/msbuild-20260628_154345_708.log`); `cmd_pane_clipboardPasteShortcut_returns_before_worker_complete` passed (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_154551`, command return `5618us`, worker `2033536us`); `cmd_pane_clipboardPasteShortcut_close_does_not_wait_for_worker` passed (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_154603`, command return `6029us`, navigate-away `6720us`); `cmd_pane_clipboardPaste_uses_preferred_move_effect` passed (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_154608`); `cmd_pane_clipboardPasteShortcut_creates_unique_links` passed (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_154614`); `cmd_pane_clipboardPasteShortcut_rejects_missing_clipboard_paths` passed (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_154620`). All five green archives record git commit `0dc592dcfcd5a9247cfc07aae5708faac59a77f6`.
- [2026-06-28 · Task 5] Statistically useful harness landed in task-scoped commit `7a511bfeb`: `folderView_perf_scroll_render_stress` now repeats the existing Brief/Detailed/ExtraDetailed gesture pass until `folder.frame.total_us` and `folder.frame.present_us` reach the 200-sample p95 floor or the 4-pass cap fails the case; `folderView_perf_overlay_invalidation_stress` extends its bounded animation-cadence window and gates the same frame metrics; both artifacts write `metricQuality.folderFrameTotal.count`, `metricQuality.folderFramePresent.count`, `samplesEnoughForP95`, `samplesEnoughForP99`, `sampleMode`, and `buildConfiguration`; `Tools/Show-PerfRuns.ps1 -FolderViewPreset` now summarizes the core FolderView frame/scroll/icon/thumbnail/refresh metrics. Durable rules moved into `Specs/Testing/Testing_PerformanceValidation.md` and `Specs/Testing/Testing_TestCoverage.md`. Verification: Debug build passed with 0 warnings/errors (`.build/logs/msbuild-20260628_155746_161.log`); test-enabled Release build passed with existing optimized-build warning `C4883` in `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.cpp` and 0 errors (`.build/logs/msbuild-20260628_160123_348.log`). Post-commit archives all record git commit `7a511bfeb7d0a19bf8c4e686fa3f33f5a99a6534`: Debug scroll `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_160646` (`metricQuality` total/present `212/212`, `samplesEnoughForP95=true`, analyzer `folder.frame.total_us` count `212`, `build=Debug`, `P95Quality=pass`); Debug overlay `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_160751` (`357/357`, analyzer count `395`, `build=Debug`, `P95Quality=pass`); Release scroll `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_160653` (`212/212`, analyzer count `213`, `build=Release`, `P95Quality=pass`); Release overlay `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_160711` (`306/306`, analyzer count `341`, `build=Release`, `P95Quality=pass`). `samplesEnoughForP99=false` is expected until a future task intentionally captures at least 1000 frame samples. Note: as documented in P2, use `commands_results.json` to identify the actual case when stale custom metric JSON files are copied forward; `2026-06-28_160711` is the Release overlay run even though the archive also contains a carried-forward scroll metrics JSON.
- [2026-06-28 · Task 6] Representative scale/cold/slow/config fixtures landed across task-scoped commits `b6195d15`, `075e98d4`, `c51d1830`, and WARP switch/docs commit `4512fdcb`: the legacy 241-item "large" case is documented as a small local baseline; `folderView_perf_huge_folder_scale` covers 10,000 items by default and 50,000 behind `REDSALAMANDER_FOLDERVIEW_HUGE_PERF=1`; `folderView_perf_cold_first_visit` records first-visit icon-index, icon bitmap queue, settle, and forced thumbnail fallback fields; `folderView_perf_slow_virtual_provider` covers deterministic slow enumeration/icon/thumbnail and repeated failed lookup behavior; `folderView_perf_relayout_churn_while_scrolled` covers 10,000-item DPI/size/theme/high-contrast relayout while scrolled and records `fontRelayoutCovered=false`; every representative artifact records `environmentMatrix`. Verification: Debug build passed with 0 warnings/errors (`.build/logs/msbuild-20260628_173510_978.log`); test-enabled Release build passed with the existing optimized-build warning `C4883` in `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.cpp` and 0 errors (`.build/logs/msbuild-20260628_173718_265.log`). Post-commit `c51d1830` Debug archives: huge scale `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_172347`, cold first visit `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_172424`, slow virtual provider `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_172353`, relayout churn `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_172409`; post-commit `c51d1830` Release archives: huge scale `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_172839`, cold first visit `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_172848`, slow virtual provider `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_172857`, relayout churn `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_172909`. Forced WARP relayout evidence after `4512fdcb`: Debug `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_174104` and Release `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_174028`, both passed and archive `warpRunExecuted=true` with build labels confirmed by `Tools/Show-PerfRuns.ps1 -FolderViewPreset`. Caveats recorded in artifacts/log: RDP remains not-run on the local console; active high-DPI is covered; DXGI adapter description does not expose installed driver version; relayout-burst frame rows are not cited as p95 evidence because Task 5 owns statistically valid frame-p95 gates.
- [2026-06-28 · Task 7] Automated regression budgets are now wired through `--selftest-perf-budget=Specs\Testing\FolderViewPerfBudgets.json5`. The committed budget file is machine-keyed to `4cb089111a23`, scoped to test-enabled Release rows, records measured/source archive values for every threshold, and hard-gates scroll frame p95/input-to-paint/present, overlay frame p95/present, huge-folder working-set/filter/sort rows, slow-provider enumeration and repeated failed icon lookups, relayout churn latency, and thumbnail close/cached-provider counters. `folderView_perf_overlay_invalidation_stress` now drives deterministic `DebugWarmPaneRendering` frames so hidden/CI launches do not produce false zero-frame failures. Verification: Debug build passed with 0 warnings/errors (`.build/logs/msbuild-20260628_181748_075.log`); test-enabled Release build passed with existing C4883 optimized-build warning and 0 errors (`.build/logs/msbuild-20260628_181130_414.log`); expected-pass single-case budget run passed (`folderView_perf_scroll_render_stress`, `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_180538`); expected-fail scratch tiny budget failed with `FolderView perf budget folder.frame.total_us.p95 exceeded` and exit code 1 (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_180631`, scratch budget deleted); full six-case budgeted run passed (`6 passed / 0 failed`, `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_181518`). Durable budget contract is recorded in `Specs/Testing/Testing_PerformanceValidation.md`; coverage evidence and the deterministic overlay-rendering contract are recorded in `Specs/Testing/Testing_TestCoverage.md`.
- [2026-06-28 · Task 8] Refresh preservation measurement landed: same-folder refresh now emits aggregate `folder.refresh.preserve_count`, `folder.refresh.rebuild_count`, `folder.refresh.selection_preserve_count`, `folder.refresh.rename_transfer_count`, `folder.refresh.debounce_delay_ms`, `folder.refresh.enumeration_count`, and the post-present latency row `folder.refresh.request_to_paint_us` through a dedicated refresh pending slot. The red selftest first failed for the intended missing metric (`folderView_perf_refresh_preservation`, `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_184202`, `Missing folder.refresh.preserve_count metric`). Verification: Debug build passed with 0 warnings/errors (`.build/logs/msbuild-20260628_184739_282.log`); production Release build passed with 0 errors and two pre-existing C5245 optimized-build warnings in `BatchRenameWindow.cpp` (`.build/logs/msbuild-20260628_185254_939.log`); `folderView_perf_directory_change_storm` passed and archived all seven rows (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_184947`, final item count 101); `folderView_perf_refresh_preservation` passed and archived all seven rows while asserting create/delete/same-folder rename/burst, selection preservation, rename transfer, focus preservation, and active incremental-search preservation (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_185004`). Durable contracts are recorded in `Specs/UI/UI_FolderView.md`, `Specs/Testing/Testing_PerformanceValidation.md`, and `Specs/Testing/Testing_TestCoverage.md`.
- [2026-06-28 · Task 9] Icon-pipeline recall avoidance and metric coverage landed: per-file icon-index workers now carry file attributes, offline/recall items force attribute-only shell lookup and emit `icons.recall_avoided_count`, and the existing extension/per-file icon-index thread-pool callbacks initialize COM MTA before shell/IconCache calls. The red selftest first failed for the intended live-path recall gap (`folderView_perf_icon_pipeline_cold_slow`, `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_190646`, `live path lookups=4`). Green Debug evidence passed at `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_192653`; supporting Debug cases passed at `2026-06-28_192709`, `2026-06-28_192722`, and `2026-06-28_192733`; PerformanceTests2 icon enumeration tests passed through `vstest.console.exe`; test-enabled Release evidence passed at `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_193018`, `2026-06-28_193033`, `2026-06-28_193045`, and `2026-06-28_193058`. Concurrency is closed as measured no-op for this slice: both Debug and Release artifacts resolved a visible bitmap before the injected slow HICON extraction completed. Durable icon-pipeline contracts are recorded in `Specs/UI/UI_FolderView.md`, `Specs/Testing/Testing_PerformanceValidation.md`, and `Specs/Testing/Testing_TestCoverage.md`.
- [2026-06-28 · Task 10] Dirty-region gate landed only the safe no-op scroll slice: `folderView_perf_scroll_render_stress` now records product `WM_PAINT` deltas through `folder.scroll.product_paint_render_count`, `folder.scroll.product_paint_full_client_count`, and `folder.scroll.product_paint_dirty_rect_area_px` plus matching artifact fields. The red check first failed for the intended avoidable full-client repaint (`No-op scroll repainted 1 frame(s) for Brief / pass1:h-left-noop`, `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_194529`). Green Debug evidence passed at `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_195138`: repeated no-op boundary scrolls report `productPaintRenderCount=0`, `productPaintFullClientCount=0`, `productPaintDirtyAreaPx=0`; viewport-changing scrolls still report one full-client product paint (`productPaintFullClientCount=1`, dirty area `327148px`). Green test-enabled Release evidence passed at `Specs/TestRuns/4cb089111a23/Commands/2026-06-28_195621`: `folder.frame.total_us` p95 `8178us` and `folder.frame.present_us` p95 `7518us` both remain below the 20ms hard scroll budget, no-op boundary scrolls stay at zero product paints, and viewport-changing scrolls still full-client paint (`dirty_rect_area_px=920040`). Scroll rects are gate-not-met / not implemented until a correctness harness covers overlays, hover, selection, focus, resize, DPI, and device loss. Virtualization V2 is gate-not-met / not implemented: the latest budgeted Release matrix (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_181518`) keeps 10k scale below hard budgets (`working_set_bytes` p95 `126746624` < `268435456`, `working_set_bytes_per_item` p95 `12674` < `30000`, `folder.filter.keystroke_to_paint_us` p95 `13250us` < `60000`, `folder.sort_toggle_us` `15120us` < `60000`). Durable metric contracts are recorded in `Specs/UI/UI_FolderView.md`, `Specs/Testing/Testing_PerformanceValidation.md`, and `Specs/Testing/Testing_TestCoverage.md`.
- [2026-06-28 · Continuation checkpoint] Branch `codex/folderview-warpdrive` is rebased on `origin/master` and is 23 commits ahead / 0 behind at checkpoint HEAD `f357d7089`; `git status --short --branch` was clean after stopping the timed-out selftest process. Task 10's code/spec slice is committed as `e7aa01b60`. The post-rebase integration fix for `PerformanceTests2` is committed in `f357d7089`: `Tests/PerformanceTests2/PerformanceTests2.vcxproj` now compiles `RedSalamander/SelfTest/Common/SelfTestLatencyHooks.cpp`, because `PerformanceTests2.dll` imports `IconCache.cpp` and Task 9 added `SelfTestLatency::Consume(...)` calls there. Verification: `.\build.ps1 -ProjectName PerformanceTests2 -Configuration Debug` passed with 0 warnings/errors (`.build/logs/msbuild-20260628_200931_351.log`). A full Debug build inside `.\Tools\Run-AllTests.ps1 -Suite Full` also passed with 0 warnings/errors (`.build/logs/msbuild-20260628_201012_007.log`), but the full-suite run did not complete: the first attempt timed out after 30 minutes during `Commands`; the second attempt (`.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild`) timed out after 2 hours during `Commands`. Latest partial full-suite archive: CompareDirectories passed 163 / failed 0 / skipped 29 at `Specs/TestRuns/4cb089111a23/CompareDirectories/2026-06-28_204249`. The full FolderView perf prefix matrix passed before the post-rebase rebuild in Debug (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_200200`, 12 passed / 0 failed) and Release (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_200336`, 12 passed / 0 failed), but their metric rows still carry pre-rebase commit `847aca5c3d1361d54b5fbbc8135ce7e8c7686bd3`; rerun both matrices from current HEAD before checking Task 10 Step 4 or Task 11. Resume here: run `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2` (or isolate the hanging Commands case if it repeats), then finish the remaining suite coverage (`FileOps` plus the native/CppUnitTest/script portions normally driven by `-Suite Full`), rerun Debug and test-enabled Release `folderView_perf_` matrices from HEAD, cite the new archives, then move this plan to `Specs/Plans/Done/`.

---

## Status And Supersession

**Status:** WIP implementation plan.
**Created:** 2026-06-28.
**Audited HEAD:** `045a1f773`.

This plan consolidates the FolderView/DxUi performance backlogs:

- `Specs/Plans/WIP/DxUi_FolderView_Monitor_FuturePerformanceIdeas_2026-05-20.md` —
  already deleted from the working tree; its decisions are folded into the
  Carry-Forward Decisions section below.
- `Specs/Plans/WIP/FolderView_PerformanceBacklog_2026-06-28.md` — was a
  never-committed scratch note (no git history) and is not present in the repo;
  its content is likewise captured in the Carry-Forward section.

No file removal remains to be done after this plan lands. The FolderView
text-layout pilots that spawned from this work —
`Specs/Plans/Done/FolderView_TextLayout_MetricPilot_2026-06-19.md` (parent at
line 11), `Specs/Plans/Done/FolderView_LayoutPassDecomposition_MetricPilot_2026-06-19.md`
(line 18), and `Specs/Plans/Done/FolderView_UpdateItemTextLayouts_Optimization_2026-06-19.md`
(line 12) — already declare THIS plan as their Parent backlog, so no parent links
need repointing. Future FolderView performance work starts here.

Root plans `plans/009-icon-pipeline-investigation.md` and
`plans/010-watcher-incremental-refresh-investigation.md` are historical prompts,
not execution sources. Their remaining valid pieces are folded into Tasks 8
(refresh preservation, from plan 010) and 9 (icon pipeline, from plan 009) below.

---

## Why This Plan Exists

The prior FolderView performance work was good at one thing: preventing
speculative changes from shipping without local measurements. Keep that
discipline.

The missing piece is broader evidence and real-world stall coverage. The current
archive trail proves small, warm, local, mostly Debug scenarios. It does not yet
prove responsiveness under cold shell caches, slow shares, remote or virtual file
systems, huge folders, WARP/RDP, device loss, wide selections, or Release builds.

The current code also contains confirmed performance and responsiveness hazards
that a small warm benchmark cannot expose:

- `FolderView.Icons.cpp:98` calls shell thumbnail `GetImage` with
  `SIIGBF_THUMBNAILONLY | SIIGBF_BIGGERSIZEOK | SIIGBF_SCALEUP`, with no
  `SIIGBF_INCACHEONLY` fast path.
- `FolderView.Rendering.cpp:1876-1941` handles failed `EndDraw`/`Present` by
  recreating swap chains only, not by discarding and recreating D3D/D2D device
  resources for recoverable device-loss errors.
- `FolderView.Rendering.cpp:2251-2252` and `:2394` create D2D solid brushes
  inside the draw path for hover and selected text.
- `FolderView.FileOps.cpp:816-884` runs paste-shortcut creation on the UI thread,
  including repeated filesystem probes and `IPersistFile::Save`.
- At plan audit time no durable `folder.refresh.*` metric existed. Same-folder
  refresh already preserved some cached item state, but mutation-to-paint latency
  and preserve/rebuild counts were invisible until Task 8 landed them.
- Existing perf selftests mostly record metrics; they do not fail on p95/p99
  regressions.

---

## Carry-Forward Decisions From The Merged Specs

### Keep As Closed Unless Fresh Evidence Reopens Them

- DirectWrite item text-layout cache: closed as measured no-op after the
  2026-06-19 metric pilot.
- `UpdateItemTextLayouts` / `SetMaxWidth` / `SetMaxHeight` micro-optimization:
  closed as measured no-op after the PerfJsonl sink correction.
- Deeper IconCache association-cache redesign (this plan's own forward decision —
  it does not trace verbatim to either merged spec; the merged FuturePerformanceIdeas
  spec's only IconCache touchpoint is the `folderView_perf_iconcache_contention`
  wait/hold diagnostics entrypoint): kept closed. Reopen a redesign only if Task 9's
  icon-pipeline measurement shows the association-cache lock is the dominant
  bottleneck.
- Full redraw or flip-discard policy changes: blocked unless external trace
  evidence proves present policy is the bottleneck.

### Keep As Gated Follow-Ups

- Unified dirty-region and scroll-region model: gate on measured dirty-area,
  input-to-paint, or scroll repaint pressure after Tasks 1-7.
- FolderView Virtualization V2: gate on representative huge-folder evidence after
  no-freeze fixes and measurement improvements land.
- Perf lab with WPR/GPUView/PresentMon: required only for GPU, present-policy,
  compositor, or D2D/D3D batching claims.
- Monitor tail/scrollback renderer and DxUi animation/batching ideas from the old
  broad backlog are parked here as non-FolderView follow-ups. Do not implement
  them in this plan; create separate WIP specs if they become active.

### Explicitly Out Of Scope For This Plan

- Long paths (>MAX_PATH — no `\\?\` prefix is applied by `GetItemFullPath`
  (FolderView.cpp:131-132) before `SHCreateItemFromParsingName`/`CreateDecoderFromFilename`),
  junction/symlink loops, and very deep trees are NOT separately fixtured. Single-folder
  view does not recurse, so they are not a perf-stall risk for this plan's scope, and
  are not covered by the G1 audit-remediation plan (scoped to Tier-1 data-loss/security
  defects). If any >MAX_PATH icon/thumbnail parse failure is observed while running the
  cold/slow/huge fixtures in Tasks 6 and 9, record it as a new finding in the
  Implementation Findings Log rather than silently dropping the item.
- Theme/DPI/high-contrast/font change while scrolled is **in scope** via Task 6:
  `folderView_perf_relayout_churn_while_scrolled` adds `folder.relayout_to_paint_us`
  evidence over a ≥10k scrolled folder and preserves scroll/focus correctness.
- Live drag-resize repaint pressure over huge folders, and dual-pane concurrent heavy
  load beyond the existing 480-item `folderView_perf_iconcache_contention`: out of scope
  unless Task 6/9 evidence reopens them.

---

## Implementation Order

1. R0 readiness bootstrap: build-label fix, analyzer sample-quality gate, and
   shared latency hooks for shell/thumbnail/icon/paste no-freeze tests.
2. R1 no-freeze fixes: device loss, draw-loop brush churn, thumbnail fast-first,
   paste-shortcut off the UI thread.
3. R2 measurement truth: representative scale, cold/slow fixtures, enough frame
   samples, Release/test-enabled evidence, and regression gates.
4. R3 evidence-gated FolderView backlog: refresh metrics and icon pipeline
   measurement/concurrency decisions.
5. R4 larger architecture only if still justified: dirty/scroll regions,
   virtualization, and external trace lab.
6. R5 spec closeout: durable contracts move to `Specs/UI/` and `Specs/Testing/`;
   this plan moves to `Specs/Plans/Done/`.

Do not batch all tasks into one commit. Each task below has its own acceptance
boundary.

---

## Task 0: Bootstrap Evidence Labels, Analyzer Gates, And Shared Latency Hooks

**Files:**

- Modify: `RedSalamander/SelfTest/Common/SelfTestCommon.cpp`
- Modify: `Tools/Show-PerfRuns.ps1`
- Create: `RedSalamander/SelfTest/Common/SelfTestLatencyHooks.h`
- Create: `RedSalamander/SelfTest/Common/SelfTestLatencyHooks.cpp`
- Modify: `RedSalamander/RedSalamander.vcxproj`
- Modify: `RedSalamander/RedSalamander.vcxproj.filters`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- Spec update after green: `Specs/Testing/Testing_PerformanceValidation.md`

**Goal:** Make all later performance evidence and no-freeze selftests trustworthy
before fixing behavior: artifacts must identify the actual build flavor, the
analyzer must reject under-sampled p95/p99 claims, and slow shell/icon/paste paths
must be delayable by one deterministic bounded test hook family.

- [x] **Step 1: Fix selftest build labels** (`6fc7851d1`)

In `SelfTestCommon.cpp`, replace the hardcoded `L"Debug"` build argument at
`InitSelfTestRun` with a local helper that mirrors `RedSalamander.cpp`:

```cpp
namespace
{
#if defined(RS_ASAN_DEBUG_BUILD)
constexpr std::wstring_view kSelfTestBuildFlavor = L"ASan Debug";
#elif defined(_DEBUG)
constexpr std::wstring_view kSelfTestBuildFlavor = L"Debug";
#else
constexpr std::wstring_view kSelfTestBuildFlavor = L"Release";
#endif
}
```

Then pass `kSelfTestBuildFlavor` to:

```cpp
Debug::Perf::ConfigureJsonlOutput(perfPath, L"SelfTest", kSelfTestBuildFlavor, gitBranch, gitCommit, GetComputerHashName(), runId);
```

Do not include or depend on file-scope `kRedSalamanderBuildFlavor` from
`RedSalamander.cpp`; it is not visible here.

- [x] **Step 2: Extend the existing perf analyzer instead of adding a duplicate tool** (`6fc7851d1`)

Update `Tools/Show-PerfRuns.ps1` with these parameters:

```powershell
param(
    # existing parameters stay as-is
    [int]$MinimumSamplesForP95 = 200,
    [int]$MinimumSamplesForP99 = 1000,
    [switch]$FailOnQuality,
    [switch]$ShowBuildFlavor
)
```

For every metric group it prints, include:

```text
count=<n> p50=<...> p95=<...> p99=<...> max=<...> build=<Debug|Release|ASan Debug|mixed> p95Quality=<pass|fail> p99Quality=<pass|fail>
```

When `-FailOnQuality` is passed, exit non-zero if any requested metric has
`count < $MinimumSamplesForP95`. Treat p99 quality as informational unless the
caller also asks for p99 hard gating in Task 7. Preserve existing `-Run`,
`-CompareRun`, `-Trend`, `-Metric`, and `-Scenario` behavior.

- [x] **Step 3: Add one shared bounded latency hook family** (`6fc7851d1`)

Create `RedSalamander/SelfTest/Common/SelfTestLatencyHooks.h/.cpp`. The
implementation is active only under `ENABLE_TESTS`; if the files are included in
a non-test build, provide no-op stubs so Release/non-test builds do not gain
sleeps or extra state. The public contract must be:

```cpp
#pragma once

#include <chrono>
#include <cstdint>
#include <stop_token>

namespace SelfTestLatency
{
enum class Point : uint8_t
{
    ShellThumbnailProviderAllowed,
    IconExtractSystemIcon,
    PasteShortcutSave,
};

void SetNextDelay(Point point, std::chrono::milliseconds delay) noexcept;
void ClearAll() noexcept;
void Consume(Point point, std::stop_token stopToken = {}) noexcept;
uint64_t ConsumeCount(Point point) noexcept;
}
```

Implementation rules:

- store delays in `std::atomic<int64_t>` milliseconds, one slot per `Point`;
- `SetNextDelay` sets the next delay for exactly one consumption;
- `Consume` atomically exchanges the slot with `0`, increments the consume count,
  and sleeps in small slices (for example 5 ms) so a supplied `stop_token` can end
  the delay early;
- clamp every configured delay to a fixed selftest ceiling, `kMaxSelfTestDelayMs = 5000`;
- `ClearAll` resets delay slots and consume counts.

- [x] **Step 4: Record the required wiring contract** (`6fc7851d1`)

Add this comment block to `SelfTest/Common/SelfTestLatencyHooks.h` below the public API so every
later task wires the same mechanism consistently:

```cpp
// Wiring contract:
// - FolderView.Icons.cpp consumes ShellThumbnailProviderAllowed immediately before
//   any provider-allowed shell thumbnail lookup. Cached-only visible lookups must
//   not consume this hook.
// - IconCache.cpp consumes IconExtractSystemIcon around the expensive HICON/shell
//   extraction path used by ExtractSystemIcon.
// - FolderView.FileOps.cpp consumes PasteShortcutSave inside the async paste-
//   shortcut worker immediately before each CreateShellShortcut call. Clipboard
//   reading stays on the UI thread and must not consume this hook.
```

Task 3, Task 4, and Task 9 wire these points when they touch the corresponding
paths. They must not add separate `DebugSetNext*Delay*` APIs with different
semantics.

- [x] **Step 5: Add a smoke selftest for the hook family** (`6fc7851d1`)

Add `folderView_debug_latency_hooks_consume_once`. It should:

1. call `SelfTestLatency::ClearAll()`;
2. set a 25 ms delay for `ShellThumbnailProviderAllowed`;
3. call `SelfTestLatency::Consume(SelfTestLatency::Point::ShellThumbnailProviderAllowed)`;
4. assert elapsed time is at least 20 ms and below 500 ms;
5. call `Consume` for the same point again and assert the second call is below 50 ms;
6. assert `ConsumeCount(ShellThumbnailProviderAllowed) == 2`;
7. clear all hooks with `wil::scope_exit`.

- [x] **Step 6: Verify** (`6fc7851d1`)

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_debug_latency_hooks_consume_once --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4
.\Tools\Show-PerfRuns.ps1 -Metric folder.frame.total_us -MinimumSamplesForP95 200 -FailOnQuality
```

Expected: build and latency smoke test pass; the analyzer prints the actual build
flavor and exits non-zero for an under-sampled run when `-FailOnQuality` is used.
Record both outcomes in the Implementation Findings Log.

---

## Task 1: Recover From Device Loss Instead Of Staying Blank

**Files:**

- Modify: `RedSalamander/FolderView.Rendering.cpp`
- Modify: `RedSalamander/FolderView.h`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- Spec update after green: `Specs/UI/UI_FolderView.md`

**Goal:** A recoverable `EndDraw`/`Present` failure must discard D3D/D2D device
resources, force a full repaint, and invalidate the FolderView instead of only
recreating the swap chain.

- [x] **Step 1: Add a narrow failure classifier** (`afbbf465d`)

Add a local helper in `FolderView.Rendering.cpp` near the render helpers:

```cpp
[[nodiscard]] bool IsFolderViewDeviceLoss(HRESULT hr) noexcept
{
    return hr == D2DERR_RECREATE_TARGET || hr == DXGI_ERROR_DEVICE_REMOVED || hr == DXGI_ERROR_DEVICE_RESET || hr == DXGI_ERROR_DEVICE_HUNG;
}
```

- [x] **Step 2: Route failed EndDraw/Present through device discard** (`afbbf465d`)

For the failed `EndDraw`, `Present1`, and legacy `Present` branches, use this
shape:

```cpp
ReportError(L"IDXGISwapChain1::Present1", hrPresent);
if (IsFolderViewDeviceLoss(hrPresent))
{
    DiscardDeviceResources();
    _forceFullRenderOnNextPaint = true;
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
    return;
}

ReleaseSwapChain();
EnsureSwapChain();
return;
```

Use the correct operation label for each branch. Preserve the current fallback
behavior for non-device-loss failures.

- [x] **Step 3: Add test hooks for both failure sites** (`afbbf465d`)

Under `ENABLE_TESTS`, add one `FolderView` debug method that can force either the
`EndDraw` or `Present` failure path without relying on a real TDR:

```cpp
#ifdef ENABLE_TESTS
enum class DebugRenderFailurePoint : uint8_t
{
    EndDraw,
    Present,
};

void DebugForceNextRenderFailure(DebugRenderFailurePoint point, HRESULT hr) noexcept;
#endif
```

Store the point + HRESULT in test-only members. Consume `EndDraw` immediately
after the real `EndDraw` result is available and before the failure branch
decides what to do. Consume `Present` immediately before calling `Present1` or
legacy `Present`. The hook must reset after one use.

- [x] **Step 4: Add a Commands selftest** (`afbbf465d`)

Add `folderView_render_device_loss_recovers`. It should:

- open a local folder in a real pane,
- warm render once,
- call `DebugForceNextRenderFailure(DebugRenderFailurePoint::EndDraw, DXGI_ERROR_DEVICE_REMOVED)`,
- warm render and assert recovery,
- call `DebugForceNextRenderFailure(DebugRenderFailurePoint::Present, DXGI_ERROR_DEVICE_REMOVED)`,
- warm render again,
- assert the pane repaints and no rendering error overlay remains,
- assert a device-loss recovery counter or trace row was emitted.

- [x] **Step 5: Verify** (`afbbf465d`)

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_render_device_loss_recovers --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4
```

Expected: all commands exit `0`, and the new case archives its trace/results.

---

## Task 2: Remove Per-Item Brush Creation From The Draw Loop

**Files:**

- Modify: `RedSalamander/FolderView.h`
- Modify: `RedSalamander/FolderView.Rendering.cpp`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`

**Goal:** Hover and selected-text rendering must reuse device-created brushes and
update color through `SetColor`, matching the existing `_selectionBrush` and
`_focusBrush` pattern.

- [x] **Step 1: Add reusable brushes** (`9506d9104`)

Add members near the existing brushes:

```cpp
wil::com_ptr<ID2D1SolidColorBrush> _hoverBrush;
wil::com_ptr<ID2D1SolidColorBrush> _selectedItemTextBrush;
```

Reset them in both places that can recreate device/theme brushes:

- `DiscardDeviceResources()`;
- the top reset block in `RecreateThemeBrushes()` before calling
  `CreateSolidColorBrush(..., .addressof())`.

This avoids leaking or overwriting live COM pointers when `SetTheme()` calls
`RecreateThemeBrushes()` without a full device discard.

- [x] **Step 2: Create the brushes with other device resources** (`9506d9104`)

In the existing device-resource brush creation path, create both brushes:

```cpp
const HRESULT hrHoverBrush = _d2dContext->CreateSolidColorBrush(_theme.itemBackgroundHovered, _hoverBrush.addressof());
if (! CheckHR(hrHoverBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(hover)"))
{
    return;
}

const HRESULT hrSelectedTextBrush = _d2dContext->CreateSolidColorBrush(_theme.textSelected, _selectedItemTextBrush.addressof());
if (! CheckHR(hrSelectedTextBrush, L"ID2D1DeviceContext::CreateSolidColorBrush(selected item text)"))
{
    return;
}
```

- [x] **Step 3: Replace draw-loop brush creation** (`9506d9104`)

For hover:

```cpp
else if (isHovered && _hoverBrush)
{
    _hoverBrush->SetColor(_theme.itemBackgroundHovered);
    _d2dContext->FillRoundedRectangle(roundedBounds, _hoverBrush.get());
}
```

For selected text:

```cpp
if (item.selected && _selectedItemTextBrush)
{
    D2D1::ColorF selectedTextColor = selectionActive ? _theme.textSelected : _theme.textSelectedInactive;
    if (_theme.rainbowMode)
    {
        const float luminance =
            0.2126f * selectionBackgroundForContrast.r + 0.7152f * selectionBackgroundForContrast.g + 0.0722f * selectionBackgroundForContrast.b;
        selectedTextColor = luminance > 0.60f ? D2D1::ColorF(D2D1::ColorF::Black) : D2D1::ColorF(D2D1::ColorF::White);
    }
    _selectedItemTextBrush->SetColor(selectedTextColor);
    textBrush = _selectedItemTextBrush.get();
}
```

Do not create `wil::com_ptr<ID2D1SolidColorBrush>` locals inside `DrawItem` for
hover or selected text.

- [x] **Step 4: Add a structural selftest guard** (`9506d9104`)

Add a Commands selftest or extend `folderView_perf_scroll_render_stress` to
exercise a wide selection and hover. Use a test-only in-process counter, not a
`Debug::Perf` metric:

```cpp
#ifdef ENABLE_TESTS
std::atomic<uint64_t> _debugDrawItemTransientBrushCreateCount{0};
#endif
```

Increment this counter only immediately before any `CreateSolidColorBrush` call
that remains inside `DrawItem`. Surface it through an existing debug snapshot or
a focused accessor, reset it before the draw pass, and assert it remains `0`.
Do not introduce `render.draw_item_transient_brush_create_count` as an archived
perf metric; the guard is structural and must be queryable in-process.

- [x] **Step 5: Verify** (`9506d9104`)

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_overlay_invalidation_stress --selftest-timeout-multiplier=4
```

Expected: both cases pass, and the wide-selection guard proves no draw-loop brush
creation.

---

## Task 3: Make Thumbnail Extraction Fast-First And Close-Safe

**Files:**

- Modify: `RedSalamander/FolderView.Icons.cpp`
- Modify: `RedSalamander/FolderView.h`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- Spec update after green: `Specs/UI/UI_FolderView.md`

**Goal:** Folder navigation, thumbnail mode, and close/teardown must not wait on
unbounded shell thumbnail provider work. The first visible thumbnail pass should
use cached-only shell lookup, then fall back to WIC/icon without blocking the UI
on network or in-proc provider delays. Extraction must also be gated to
local-shell-backed file systems (no synthetic virtual-FS paths into the shell
namespace) and must reject oversized WIC source frames before decode.

- [x] **Step 1: Split shell thumbnail modes** (`6c26367b7`)

Add an enum in `FolderView.Icons.cpp`:

```cpp
enum class ShellThumbnailLookupMode : uint8_t
{
    CachedOnly,
    ProviderAllowed,
};
```

Change `ExtractShellThumbnailBitmap` to accept the mode and build flags like:

```cpp
SIIGBF flags = static_cast<SIIGBF>(SIIGBF_THUMBNAILONLY | SIIGBF_BIGGERSIZEOK | SIIGBF_SCALEUP);
if (mode == ShellThumbnailLookupMode::CachedOnly)
{
    flags = static_cast<SIIGBF>(flags | SIIGBF_INCACHEONLY);
}
```

- [x] **Step 2: Use cached-only lookup in the visible queue** (`6c26367b7`)

In `ProcessThumbnailLoadQueue`, call cached-only shell lookup for normal
thumbnail queue processing:

```cpp
hr = ExtractShellThumbnailBitmap(request.fullPath, request.targetPx, ShellThumbnailLookupMode::CachedOnly, shellBitmap);
```

If the cached-only shell call misses, continue to the existing WIC fallback for
likely image files, then the existing icon fallback. Do not run provider-allowed
shell lookup synchronously in the visible queue.

- [x] **Step 2b: Make teardown not block on in-flight shell extraction (Audit B1)** (`6c26367b7`)

The visible path is close-safe only as long as it never issues a provider-allowed
call. The fix must ALSO ensure teardown cannot hang if a provider-allowed
thumbnail lookup is introduced for a test hook or future background path.
Today `StopEnumerationThread` (FolderView.Enumeration.cpp:227-245) calls
`request_stop()` then move-assigns a fresh jthread whose destructor JOINS the
worker; `CancelThumbnailLoading` (FolderView.Icons.cpp:855) only flips
`_thumbnailLoadingActive`, and `ProcessThumbnailLoadQueue` (FolderView.Icons.cpp:907-1015)
loops on `_thumbnailLoadingActive` without checking the stop token, so a
synchronous shell call (Icons.cpp:1004 -> :99 `GetImage`) blocks the join for the
full SMB/handler timeout.

Implement the close-safe shape explicitly:

- do **not** detach a raw `std::thread`;
- do **not** let `StopEnumerationThread()` join a thread that can be inside a
  provider-allowed `GetImage`;
- keep the normal visible thumbnail queue cached-only;
- if a provider-allowed lookup is needed for the regression hook or a future
  background path, run that lookup in value-only abandonable work that captures
  only `{HWND hwnd, uint64_t generation, std::filesystem::path fullPath,
  uint32_t targetPx}` and COM-initializes its own thread;
- the enumeration/thumbnail worker waits only a bounded deadline for that work,
  checks its stop token before and after the wait, treats timeout/cancel as a
  cache miss, increments an abandoned/timeout counter, and falls back to WIC/icon;
- late results post back with `PostMessagePayload` and are dropped when the HWND is
  closed or the generation is stale.

Add an assertion to `folderView_thumbnail_cached_only_no_close_stall` (Step 4)
that, with `SelfTestLatency::Point::ShellThumbnailProviderAllowed` forcing a slow
provider-allowed call in flight, close/navigate returns within the bounded
deadline and `thumbnailPendingCount == 0`.

- [x] **Step 2c: Gate shell/WIC extraction to local-shell-backed file systems (Audit B2)** (`6c26367b7`)

`QueueThumbnailLoading`/`QueueIconLoading` currently gate only on visibility and
pass `GetItemFullPath(item)` verbatim into `SHCreateItemFromParsingName`
(FolderView.Icons.cpp:89) and `CreateDecoderFromFilename` (:145). For a virtual FS
(e.g. FileSystem7z) this parses a synthetic local-looking path against the real
shell namespace and can show a different real file's thumbnail. Add a local-shell
capability check (cache `NavigationLocation::IsFilePluginShortId(pluginShortId)`
from `SetFileSystem`, FolderView.cpp:558, into a member set during SetFileSystem);
when the current FS is not local-shell-backed, skip shell/WIC extraction and route
those items to the type-icon fallback. `thumbnails.shell_*`/WIC counters must
remain 0 for non-local FS.

- [x] **Step 2d: Cap WIC source dimensions before decode (Audit B3)** (`6c26367b7`)

`DecodeWicThumbnailPixels` (FolderView.Icons.cpp:133) clamps only the OUTPUT target
to `kMaxThumbnailPixelSize` (512); the source frame is checked only for non-zero
dimensions (:160-164) and the full source is fed to the scaler. Before decode,
reject sources whose `static_cast<uint64_t>(sourceWidth)*sourceHeight` exceeds a
ceiling (define `constexpr uint64_t kMaxThumbnailSourcePixels` e.g. `64u*1024u*1024u`),
returning `WINCODEC_ERR_BADIMAGE` so the worker falls back to the icon path instead
of decoding a decompression-bomb frame in-process.

- [x] **Step 3: Add metrics** (`6c26367b7`)

Add full snapshot plumbing (the Step 4 selftest and the roll-up gate can only read
stats through `PaneViewOptionsDebugSnapshot`):

- Add four counters to the `ThumbnailLoadStats` struct in `FolderView.h`
  (alongside `shellSuccess{0}` at line 1514):
  `std::atomic<uint64_t> shellCacheHit{0}; shellCacheMiss{0}; shellProviderAllowed{0}; shellProviderTimeout{0};`.
- Increment them in `ProcessThumbnailLoadQueue` (FolderView.Icons.cpp, next to the
  existing `shellSuccess.fetch_add` near line 1009): bump `shellCacheHit` on a
  CachedOnly hit, `shellCacheMiss` on a CachedOnly miss, `shellProviderAllowed`
  only if a provider-allowed lookup is ever issued on the visible path (must stay
  0), and `shellProviderTimeout` when abandonable provider work exceeds the
  bounded wait.
- Surface them through BOTH snapshots: add
  `shellCacheHitCount/shellCacheMissCount/shellProviderAllowedCount/shellProviderTimeoutCount` to
  `ThumbnailDebugSnapshot` (FolderView.h:419-450, next to `shellSuccessCount` at :429)
  populated in `DebugGetThumbnailSnapshot`; then add matching
  `thumbnailShellCacheHitCount/thumbnailShellCacheMissCount/thumbnailShellProviderAllowedCount/thumbnailShellProviderTimeoutCount`
  to `PaneViewOptionsDebugSnapshot` (FolderWindow.h:811-860, next to
  `thumbnailShellSuccessCount` at :832) and copy them next to
  FolderWindow.FileSystem.Commands.Part.cpp:10562.
- Also emit `thumbnails.shell_cache_hit_count`, `thumbnails.shell_cache_miss_count`,
  `thumbnails.shell_provider_allowed_count`, `thumbnails.shell_provider_timeout_count`,
  and `thumbnails.cached_extract_us` via `PerfEmitCounter`/`PerfEmitDuration` for
  archived artifacts. The selftest gate MUST assert
  `settled.thumbnailShellProviderAllowedCount == 0u` read from
  `PaneViewOptionsDebugSnapshot`.

- [x] **Step 4: Add a close/teardown regression case** (`6c26367b7`)

Add `folderView_thumbnail_cached_only_no_close_stall`. The case should:

1. enable thumbnail mode;
2. assert the visible (cached-only) path issues zero provider-allowed shell lookups
   via `settled.thumbnailShellProviderAllowedCount == 0u` (read from
   `PaneViewOptionsDebugSnapshot` — added in Step 3);
3. verify visible items settle through cached-shell, WIC, or icon fallback;
4. separately exercise a provider-allowed lookup made deterministically slow via
   `SelfTestLatency::Point::ShellThumbnailProviderAllowed`, then close/navigate away while that work is
   pending;
5. assert stale work is dropped, close completes within a bounded deadline, and
   `thumbnailPendingCount == 0`;
6. emit/record `thumbnails.close_to_idle_us` from the selftest so Task 7 can hard
   budget close-to-idle latency separately from frame metrics.

This requires BOTH a zero-provider-call assertion AND a blocking latency hook on the
provider-allowed branch — the existing `DebugThumbnailProviderMode` modes cannot
block. Wire the provider-allowed branch to the Task 0 `SelfTestLatency` hook; do
not rely on a mere assert-no-provider mode.

- [x] **Step 5: Verify** (`6c26367b7`)

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_thumbnail_cached_only_no_close_stall --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_thumbnail_scroll_stress --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_thumbnail_valid_images_shell_fail --selftest-timeout-multiplier=4
```

Expected: all cases pass; thumbnail artifacts show cached-only shell behavior and
no pending work after close/navigation.

> **Note:** `folderView_thumbnail_valid_images_shell_fail` is unaffected by the Step 2
> CachedOnly change. In `ForceShellFailureAllowWic` mode the branch at
> FolderView.Icons.cpp:995-999 sets `hr=ERROR_GEN_FAILURE` and `allowWicFallback=true`
> without ever calling `ExtractShellThumbnailBitmap`, so the new lookup mode never
> executes here. In real Shell mode a cached-only cache miss still sets
> `allowWicFallback=true` and (because `.bmp` is a likely WIC extension) falls through
> to WIC decode, so the WIC-success and zero-fallback assertions continue to hold. No
> extra protection for this path is required.

---

## Task 4: Move Paste Shortcut Creation Off The UI Thread

**Files:**

- Modify: `Common/WindowMessages.h`
- Modify: `RedSalamander/FolderView.FileOps.cpp`
- Modify: `RedSalamander/FolderView.cpp`
- Modify: `RedSalamander/FolderView.h`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ShellCommands.cpp` (co-locate the new case with the existing paste-shortcut cases `TestPaneClipboardPasteShortcutCreatesLinks` (ShellCommands.cpp:2776) and `TestPaneClipboardPasteShortcutRejectsMissingClipboardPaths` (:2893); reuse the paste-specific helpers `SetClipboardDropPathsForShellCommandTest` and `ReadShortcutTargetForShellCommandTest`)

**Goal:** `PasteShortcutFromClipboard` must return control to the UI promptly and
complete shortcut creation through value-only background work that can be
abandoned on close/navigation and posts results back with `PostMessagePayload`.

- [x] **Step 1: Add a completion message** (`0dc592dcf`)

Add a new FolderView message in the open `WM_APP + 0x300` range:

```cpp
inline constexpr UINT kFolderViewPasteShortcutComplete = WM_APP + 0x30A;
```

- [x] **Step 2: Add a result payload** (`0dc592dcf`)

Add a `PasteShortcutResult` struct to `FolderView.h` with:

```cpp
struct PasteShortcutResult
{
    uint64_t generation = 0; // snapshot of FolderView::_enumerationGeneration (FolderView.h:1412) taken on the UI thread before submitting the worker
    std::filesystem::path targetFolder;
    std::vector<std::filesystem::path> createdLinks;
    HRESULT firstFailure = S_OK;
    std::filesystem::path failedSource;
    uint64_t elapsedUs = 0;
};
```

- [x] **Step 3: Run shortcut creation on a worker** (`0dc592dcf`)

Keep clipboard reading on the UI thread. After `ReadFileDropClipboard`, run
shortcut creation on a worker.

Use value-only thread-pool work (`TrySubmitThreadpoolCallback`), not an owned
`std::jthread`. `IPersistFile::Save` is
synchronous and not stop-token-aware; an owned worker joined in `OnDestroy` would
move the UI freeze from command execution to close/teardown.

The callback can run after the window and the FolderView object are destroyed. It
MUST capture only by value:

- `std::vector<std::filesystem::path> sources`;
- `std::filesystem::path targetFolder`;
- `uint64_t generation`;
- `HWND hwndCopy` read from `_hWnd.get()` once on the UI thread.

It MUST NOT capture or dereference `this`. Reading `_hWnd`/`_currentFolder`/
`_fileSystem` from inside the callback would be a use-after-free.

Before submitting, capture the live generation on the UI thread:
`result.generation = _enumerationGeneration.load(std::memory_order_acquire);` (the
same counter at FolderView.h:1412 that enumeration/icon staleness checks use — do
not introduce a new counter).

The worker MUST initialize COM as MTA before calling `CreateShellShortcut`, exactly
as `EnumerationWorker` does (FolderView.Enumeration.cpp:276): at the top of the
worker body call `CoInitializeEx(nullptr, COINIT_MULTITHREADED)` (fail-fast on
failure) and hold a `wil::unique_couninitialize_call` for the worker's lifetime.
`CreateShellShortcut` calls `CoCreateInstance(CLSID_ShellLink,...)`
(FolderView.FileOps.cpp:279) and `IPersistFile::Save` (:303-304), which require a
COM apartment; the current code works only because `PasteShortcutFromClipboard`
runs on the already-COM-initialized UI thread (:816). A bare worker with no COM
apartment will fail every `CoCreateInstance` with `CO_E_NOTINITIALIZED`.

The worker calls `CreateShellShortcut` (a free function taking copies, safe
off-thread) and posts `PasteShortcutResult` via:

```cpp
auto payload = std::make_unique<PasteShortcutResult>(std::move(result));
static_cast<void>(PostMessagePayload(hwndCopy, WndMsg::kFolderViewPasteShortcutComplete, 0, std::move(payload)));
```

using the captured HWND value. If the window already reached `WM_NCDESTROY`,
`PostMessagePayload` deletes the payload and returns `false`; the worker ignores
that failure because the UI is gone.

Do not detach a raw `std::thread`.

- Opportunistic note (full fix tracked under Audit S3 / Finding G1): while editing
  `FolderView.FileOps.cpp`, you may add the missing clipboard bound check in
  `ReadPreferredDropEffectClipboard`. After the `GlobalLock`/null check, before
  `const DWORD result = *effect;` (FolderView.FileOps.cpp:267), insert
  `if (GlobalSize(handle) < sizeof(DWORD)) { GlobalUnlock(handle); return std::nullopt; }`
  — the caller already defaults to COPY via `value_or(DROPEFFECT_COPY)` (:738).
  `ReadFileDropClipboard` correctly stays on the UI thread (length-safe via
  `DragQueryFileW`, :220-231). If you skip it, leave it for the G1 security plan.

- [x] **Step 3b: Wire the shared paste-save delay hook** (`0dc592dcf`)

Consume the Task 0 shared hook inside the worker loop immediately before each
`CreateShellShortcut` call:

```cpp
#ifdef ENABLE_TESTS
SelfTestLatency::Consume(SelfTestLatency::Point::PasteShortcutSave);
#endif
```

The selftest sets and clears this via `SelfTestLatency::SetNextDelay(...)` /
`ClearAll()` with `wil::scope_exit`. Do not add a paste-specific delay API.

- [x] **Step 4: Handle completion on the UI thread** (`0dc592dcf`)

In `FolderView::WndProc`, add:

```cpp
case WndMsg::kFolderViewPasteShortcutComplete:
{
    auto result = TakeMessagePayload<PasteShortcutResult>(lParam);
    if (result)
    {
        OnPasteShortcutComplete(std::move(*result));
    }
    return 0;
}
```

`OnPasteShortcutComplete` must:

- drop the result (do nothing) if it is stale:
  `result.generation != _enumerationGeneration.load(std::memory_order_acquire)` OR
  the folder no longer matches (`! _currentFolder || result.targetFolder != *_currentFolder`
  — `_currentFolder` is `std::optional<std::filesystem::path>` at FolderView.h:848);
- otherwise, only when links were created, run the existing sync post-create sequence
  in this order: `DirectoryInfoCache::GetInstance().NotifyFolderContentsChanged(...)`,
  then `RememberFocusedItemForFolder(*_currentFolder, createdLinks.back().filename().wstring())`,
  then `EnumerateFolder()` (mirrors FolderView.FileOps.cpp:880-883). Because
  `EnumerateFolder()` bumps `_enumerationGeneration` (FolderView.Enumeration.cpp:985),
  the staleness compare above MUST be done first, before this refresh;
- show success/failure overlay on the UI thread,
- emit `clipboard.paste_shortcut_worker_us`.

- [x] **Step 5: Add responsiveness and close-safety cases** (`0dc592dcf`)

Add `cmd_pane_clipboardPasteShortcut_returns_before_worker_complete`. The case
should call
`SelfTestLatency::SetNextDelay(SelfTestLatency::Point::PasteShortcutSave, 500ms)`
to delay the worker's per-source shortcut creation and assert that the command
returns quickly while the worker later posts completion. The command must not
block message pumping for all sources.

Add `cmd_pane_clipboardPasteShortcut_close_does_not_wait_for_worker`. The case
should set a longer bounded `PasteShortcutSave` delay, invoke paste-shortcut, then
close/navigate away before the worker completes. Assert:

- close/navigation returns within the bounded deadline;
- no completion handler dereferences a destroyed `FolderView`;
- any late `PostMessagePayload` is safely dropped by the closed-HWND registry;
- stale generation/folder results do not refresh or focus the old folder.

- [x] **Step 6: Verify** (`0dc592dcf`)

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_clipboardPasteShortcut_returns_before_worker_complete --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_clipboardPasteShortcut_close_does_not_wait_for_worker --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_clipboardPaste_uses_preferred_move_effect --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_clipboardPasteShortcut_creates_unique_links --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_clipboardPasteShortcut_rejects_missing_clipboard_paths --selftest-timeout-multiplier=4
```

Expected: all five cases pass; the new returns_before_worker_complete case proves
command latency is separated from worker completion latency, the new close case
proves teardown does not wait for a slow save, and the two existing paste-shortcut
cases stay green after the worker refactor.

> **Note (Finding C1, corrected):** `cmd_pane_clipboardPaste_uses_preferred_move_effect`
> already exists (Commands.SelfTest.ShellCommands.cpp:2997), so the verify line that
> runs it is valid — keep it. Only `cmd_pane_clipboardPasteShortcut_returns_before_worker_complete`
> and `cmd_pane_clipboardPasteShortcut_close_does_not_wait_for_worker`
> are new work. Also run the existing `cmd_pane_clipboardPasteShortcut_creates_unique_links`
> (:3011) and `cmd_pane_clipboardPasteShortcut_rejects_missing_clipboard_paths` (:3013)
> as regression guards, since Task 4's worker refactor touches that exact path.

---

## Task 5: Make The Perf Harness Statistically Useful

**Files:**

- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- Modify: `Specs/Testing/Testing_PerformanceValidation.md`
- Modify: `Specs/Testing/Testing_TestCoverage.md`
- Prefer modify: `Tools/Show-PerfRuns.ps1` (existing percentile analyzer; only create `Tools/Analyze-FolderViewPerf.ps1` if it cannot be extended). Note: `Tools/AnalyzeTestRuns.ps1` reads suite-results JSON, not perf percentiles, so it is NOT the right tool.

**Goal:** FolderView perf claims must use enough samples and must include
test-enabled Release evidence before a performance improvement can be accepted.

- [x] **Step 1: Loop frame-producing gestures**

Update perf selftests so each frame metric family has enough samples for p95/p99
to mean something. Treat gesture-driven and animation-driven cases differently:

- Gesture-driven cases (`folderView_perf_scroll_render_stress`): wrap the existing
  scroll gesture in an outer loop and repeat until each target frame metric reaches
  ≥200 samples or a fixed iteration cap is hit; fail `metricQuality.samplesEnoughForP95`
  if the cap is reached first. A single Debug pass yields ~109 `folder.frame.total_us`
  samples, so ~2 passes reach 200 — size the cap accordingly (do NOT assume ~24
  frames/pass).
- Animation-driven cases (`folderView_perf_overlay_invalidation_stress`): there is NO
  input gesture to repeat — frames come from the overlay animation cadence under the
  time-bounded `UpdateWindow` loops. Raise the sample target by extending the bounded
  run-time, not by gesture repetition. A ~1.9s Debug run already yields ~120 overlay
  frames, so a modest run-time extension reaches ≥200; if the target is not reachable
  within a bounded deterministic run-time, set a lower cadence-justified target and
  record in the artifact that overlay samples are animation-cadence-bounded, not
  gesture-bounded.

Metric-quality fields to write into each artifact:

- `metricQuality.folderFrameTotal.count`
- `metricQuality.folderFramePresent.count`
- `metricQuality.samplesEnoughForP95`
- `metricQuality.samplesEnoughForP99`
- `metricQuality.buildConfiguration`

- [x] **Step 2: Add Release/test-enabled run instructions**

Task 0 owns the build-label fix. This step verifies and documents the Release
selftest workflow, and Task 5 Step 1's `metricQuality.buildConfiguration` field
must be sourced from the corrected JSONL `build` value, not a hardcoded literal.

Document and use:

```powershell
try {
    $env:RSBuildEnableTests='true'
    .\build.ps1 -ProjectName RedSalamander -Configuration Release
} finally {
    Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue
}
.\.build\x64\Release\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4
```

The artifact must identify Release versus Debug. If any Release-built row still
reports `build="Debug"`, stop and fix Task 0 before continuing.

- [x] **Step 3: Add the FolderView metric preset to the analyzer**

Prefer extending `Tools/Show-PerfRuns.ps1`, which already discovers
`Specs\TestRuns\**\perf\perf_metrics.jsonl` (`Get-RunFolders`), groups by the
`.metric` field, emits count/p50/p95/p99/max per metric (`Measure-Metric`), supports
`-Run`/`-CompareRun`/`-Trend`/`-Metric`/`-Scenario`, and already ranks the
FolderView/render/icons/iconcache families. Task 0 already adds sample-quality
verdicts, non-zero quality failure, and Debug/Release labeling. This step adds or
verifies the FolderView preset metric list below. Only create
`Tools/Analyze-FolderViewPerf.ps1` if `Show-PerfRuns.ps1` genuinely cannot be
extended; if so, state why in the Findings Log. Whichever tool is used must emit
count, p50, p95, p99, max, build flavor, and a pass/fail quality result for:

- `folder.frame.total_us`
- `folder.frame.present_us`
- `folder.frame.input_to_paint_us`
- `folder.frame.dirty_rect_area_px`
- `folder.scroll_input_to_paint_us`
- `folder.relayout_to_paint_us`
- `FolderView.IconLoading.ProcessQueue`
- `thumbnails.extract_us`
- `thumbnails.close_to_idle_us`
- `thumbnails.shell_provider_allowed_count`
- `thumbnails.shell_provider_timeout_count`
- `folder.refresh.request_to_paint_us`

- [x] **Step 4: Verify**

Run Debug and test-enabled Release for:

```powershell
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_overlay_invalidation_stress --selftest-timeout-multiplier=4
.\.build\x64\Release\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-timeout-multiplier=4
.\.build\x64\Release\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_overlay_invalidation_stress --selftest-timeout-multiplier=4
```

Expected: all pass and each archive has enough frame samples for the declared
percentiles.

---

## Task 6: Add Representative Scale, Cold, And Slow Fixtures

**Files:**

- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- Modify or create test helpers under `RedSalamander/SelfTest/Commands/`
- Modify: `Specs/Testing/Testing_TestCoverage.md`

**Goal:** FolderView performance coverage must include more than small warm local
folders.

- [x] **Step 1: Rename or reframe the current large-folder case** (`c51d1830`)

If `folderView_perf_large_folder_baseline` still uses a small fixture, either
increase it or rename the artifact description so it no longer claims to be the
large-folder evidence. The large-folder evidence for this plan should be a new
case:

```text
folderView_perf_huge_folder_scale
```

- [x] **Step 2: Add huge-folder scale** (`b6195d15`, `c51d1830`)

Add a scale case with at least:

- 10,000 item mode for routine local runs,
- 50,000 item mode behind a slow-suite or environment flag,
- artifact fields for item count, extension count, visible item count, enumeration
  time, sort time, first visible paint time, and frame samples;
- working-set and private bytes sampled before enumeration, after first paint, and
  after the scroll pass — using `SampleSelfTestWorkingSetBytes()`
  (Commands.SelfTest.ViewCommands.cpp:16066) or `CaptureProcessMemorySnapshot()`
  (FolderWindow.FileOperations.State.cpp:1871). Emit `folder.scale.working_set_bytes`
  (per sample point) plus a derived `folder.scale.working_set_bytes_per_item`;
- after first visible paint, invoke SelectAll then scroll the full vertical extent;
  during this all-selected scroll pass, record the Task 2 transient-brush-create
  counter (MUST remain 0 once Task 2 lands) alongside
  `folder.frame.total_us`/`folder.frame.present_us` p95 — every visible item takes the
  `item.selected` per-frame brush branch (FolderView.Rendering.cpp:2384), the worst
  case the brush fix targets;
- also exercise (a) a sort toggle and (b) an incremental quick-search typed
  char-by-char on the ≥10k set, recording `folder.sort_toggle_us` (reuse the existing
  key emitted at Commands.SelfTest.ViewCommands.cpp:15934 — do NOT invent a new key)
  and a new per-keystroke `folder.filter.keystroke_to_paint_us`. A 5000-item
  `folderView_perf_sort_toggle_stress` and a char-by-char quick-search case already
  exist — extend their pattern to huge scale, do not claim none exist.

Avoid writing 250,000 real files in the normal selftest root. For very large
model-scale tests, reuse the in-memory provider pattern in
`Tests/PerformanceTests2/FolderViewRefreshDuplicatePathPerfTest.cpp`
(`TestFilesInformation : IFilesInformation`, `DuplicatePathFileSystem : IFileSystem`,
synthetic `ReadDirectoryInfo` that builds a contiguous FileInfo buffer with no disk
I/O) to synthesize 10k/50k/250k items; or drive the pane through the existing
`builtin/file-system-dummy` plugin via `FindFileSystemPluginById` +
`FolderWindow::SetFileSystemPluginForPane`. Note the current PerformanceTests2 icon
fixture still writes ~1500 real files, so the no-disk path is net-new work even
though the `IFilesInformation` pattern exists to copy.

- [x] **Step 3: Add cold-cache first-visit coverage** (`b6195d15`, `c51d1830`)

Add a case:

```text
folderView_perf_cold_first_visit
```

It must use unique extensions and unique paths not warmed by earlier cases in the
same run. Record first enumeration, first icon-index lookup, first icon bitmap
queue, first thumbnail fallback, and first paint metrics.

- [x] **Step 4: Add slow or virtual file-system coverage** (`b6195d15`, `c51d1830`)

Add a named case:

```text
folderView_perf_slow_virtual_provider
```

Use a deterministic latency-injected provider or harness mode that can delay:

- directory enumeration,
- icon-index lookup,
- thumbnail lookup,
- file-existence probes,
- shortcut save (by reusing `SelfTestLatency::Point::PasteShortcutSave` from the
  ShellCommands paste tests, not by driving paste from this ViewCommands case),
- a per-item slow-then-fail mode (simulating permission-denied / AV-locked items)
  where the injected provider stalls then returns a failure, to prove
  repeatedly-failing items do not re-stall every refresh. This matters because
  IconCache deliberately does not cache failed lookups (IconCache.cpp:1163-1166
  returns nullopt without caching; the per-item call site is
  FolderView.Enumeration.cpp:844, awaited every enumeration pass), so a folder of
  denied/locked items re-issues the slow SHGetFileInfoW on every navigate/refresh.
  Emit `icons.repeated_failed_lookup_count` (failed per-item lookups recurring
  across successive passes for the same paths) and assert it stays bounded.

The delay must be test-controlled and bounded so the selftest is deterministic.
For shell/icon/thumbnail/paste slow points, reuse the Task 0 `SelfTestLatency`
hook family rather than adding new ad hoc delay APIs.

- [x] **Step 5: Add relayout churn while scrolled** (`b6195d15`, `c51d1830`)

Add a named case:

```text
folderView_perf_relayout_churn_while_scrolled
```

The case should:

1. open a synthetic ≥10k-item folder;
2. scroll to a non-zero vertical offset;
3. drive bounded DPI and theme relayout triggers through existing code paths
   (`OnDpiChanged` and `SetTheme`). Font relayout is not added here unless there is
   already a public selftest setter in the same file; if not, record that omission
   in the artifact as `fontRelayoutCovered=false`;
4. emit `folder.relayout_to_paint_us` from the relayout request through the first
   post-relayout present;
5. assert scroll position and focused item survive the relayout;
6. record frame sample-quality fields for the repaint burst.

This closes the earlier out-of-scope ambiguity: theme/DPI/high-contrast relayout
is now scheduled here for perf evidence. Live drag-resize and dual-pane heavy load
remain out of scope unless this case or Task 9 reopens them with evidence.

- [x] **Step 6: Add the environment/config matrix manifest** (`075e98d4`, `4512fdcb`)

Every Task 6 archive must include a compact `environmentMatrix` object with:

- build flavor (`Debug`, `Release`, or `ASan Debug`);
- GPU adapter name and driver version;
- display refresh rate and DPI/scale;
- local console versus RDP;
- WARP/software rendering availability and whether a WARP run was executed;
- high-DPI/high-scale run status;
- reason for any skipped matrix dimension.

Required for this plan: local-console Debug and test-enabled Release on the
developer machine, plus WARP if the machine can create the software adapter.
RDP, 4K/150%, 144Hz, and other hardware-dependent dimensions may be marked
`blocked` only with the missing environment stated explicitly in the archive and
Implementation Findings Log.

- [x] **Step 7: Verify** (`c51d1830`, `4512fdcb`)

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_huge_folder_scale --selftest-timeout-multiplier=6
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_cold_first_visit --selftest-timeout-multiplier=6
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_slow_virtual_provider --selftest-timeout-multiplier=6
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_relayout_churn_while_scrolled --selftest-timeout-multiplier=6
```

Expected: artifacts contain item counts, cold/huge/slow/relayout markers,
environment matrix fields, frame-sample quality fields, and archived perf rows.

---

## Task 7: Add Automated Regression Gates

**Files:**

- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- Create or modify: `Specs/Testing/FolderViewPerfBudgets.json5` if no budget file exists
- Modify: `Specs/Testing/Testing_PerformanceValidation.md`

**Goal:** A large FolderView regression should fail a targeted perf gate instead
of relying only on a human comparing archive folders.

- [x] **Step 1: Define soft and hard budgets**

Add budgets for test-enabled Release first. Debug budgets may be looser or
metric-quality-only. Start with conservative thresholds derived from fresh
same-machine baselines after Tasks 1-6, not from the pre-sink-fix archives.

Budget fields must include measured numeric values for:

- `folder.frame.total_us.p95.max`
- `folder.frame.input_to_paint_us.p95.max`
- `folder.frame.present_us.p95.max`
- `folder.scale.working_set_bytes.max`
- `folder.scale.working_set_bytes_per_item.max`
- `folder.filter.keystroke_to_paint_us.p95.max`
- `folder.sort_toggle_us.p95.max`
- `folder.relayout_to_paint_us.p95.max`
- `thumbnails.close_to_idle_us.max`
- `thumbnails.shell_provider_allowed_count.max` (must be `0` for the visible path)
- `thumbnails.shell_provider_timeout_count.max`
- `minimumSamples`

The budget file MUST record the machine hash it was derived from (matching the
`Specs/TestRuns/<MachineHash>/` archive it cites, precedent `Specs/TestRuns/README.md:64`);
other machines run sample-count/quality-only checks rather than these hard thresholds
(see G4).

Write the budget file only after Task 5 has produced fresh same-machine Debug
and test-enabled Release baselines, and cite the source archive for every
threshold in the same commit.

- [x] **Step 2: Add a compare mode**

Add command-line or environment control:

```text
--selftest-perf-budget=Specs\Testing\FolderViewPerfBudgets.json5
```

The selftest should fail only on hard budgets. Soft budgets should warn and
archive evidence.

- [x] **Step 3: Gate the first cases**

Gate:

- `folderView_perf_scroll_render_stress`
- `folderView_perf_overlay_invalidation_stress`
- `folderView_perf_huge_folder_scale`
- `folderView_perf_slow_virtual_provider`
- `folderView_perf_relayout_churn_while_scrolled`
- `folderView_thumbnail_cached_only_no_close_stall`

- [x] **Step 4: Verify**

Run one expected-pass budget and one local scratch expected-fail budget. Do not
commit the failing scratch budget.

```powershell
.\.build\x64\Release\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_scroll_render_stress --selftest-perf-budget=Specs\Testing\FolderViewPerfBudgets.json5 --selftest-timeout-multiplier=4
```

Expected: pass with normal budget; fail with intentionally tiny scratch budget.

---

## Task 8: Measure And Close Refresh Preservation

**Files:**

- Modify: `RedSalamander/FolderView.Enumeration.cpp`
- Modify: `RedSalamander/FolderView.cpp` (add a dedicated pending refresh-to-paint slot; do not reuse the single input-to-paint slot)
- Modify: `RedSalamander/DirectoryInfoCache.cpp`
- Modify: `RedSalamander/FolderView.h`
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- Reuse (no edit expected): `RedSalamander/FolderView.Rendering.cpp` emits the pending metric after Present at lines 1921 and 1945 via `EmitPendingInputToPaintMetricAfterPresent()`.
- Reuse / extend if routing the start through the pane: `RedSalamander/FolderWindow.cpp` and `RedSalamander/FolderWindow.FileSystem.cpp` hold the cross-async-boundary `*_to_paint_us` pattern (`PaneState::PendingNavigationToPaintMetric` set at FolderWindow.FileSystem.cpp:4581; handed to the view at FolderWindow.cpp:962).
- Spec update after green: `Specs/UI/UI_FolderView.md`

> The `folder.refresh.request_to_paint_us` metric is a request→paint span: the request
> originates in Enumeration.cpp but the duration is emitted only after Present in
> FolderView.Rendering.cpp via `EmitPendingInputToPaintMetricAfterPresent`
> (Rendering.cpp:1921/:1945), reusing the pending-metric facility in FolderView.cpp — a
> Files list of only Enumeration.cpp/DirectoryInfoCache.cpp/FolderView.h cannot reach it.

**Goal:** Turn the current same-folder refresh preservation path into measured,
guarded behavior.

- [x] **Step 1: Add aggregate refresh metrics**

Emit per-refresh aggregate metrics:

- `folder.refresh.preserve_count`
- `folder.refresh.rebuild_count`
- `folder.refresh.selection_preserve_count`
- `folder.refresh.rename_transfer_count`
- `folder.refresh.request_to_paint_us`
- `folder.refresh.debounce_delay_ms`
- `folder.refresh.enumeration_count`

Do not emit one row per item.

> **Pending-slot caveat (verified 2026-06-28):** `folder.refresh.request_to_paint_us`
> is a `*_to_paint_us` latency metric, but the existing input-to-paint mechanism stores
> only ONE pending metric in a single slot (`_pendingInputToPaintMetric`,
> FolderView.h:939; overwritten unconditionally in `RecordPendingInputToPaintStart`,
> FolderView.cpp:84-98; emitted once and reset in `EmitPendingInputToPaintMetricAfterPresent`,
> FolderView.cpp:112-120). That slot is ALREADY shared by scroll/focus input-to-paint and
> navigation-to-paint (FolderWindow.cpp:962). Do NOT route
> `folder.refresh.request_to_paint_us` through `RecordPendingInputToPaintStart` — a refresh
> in flight concurrent with a scroll/keyboard input-to-paint (or a navigation-to-paint)
> would overwrite each other last-writer-wins, silently dropping or mis-attributing one
> metric. Give refresh latency its own dedicated slot/timestamp (e.g. a
> `_pendingRefreshRequestTick` recorded when the refresh request is issued and consumed
> exactly once at the next post-present emit). If a dedicated slot is out of scope,
> explicitly document the last-writer-wins limitation in the Implementation Findings Log
> and in `Specs/Testing/Testing_PerformanceValidation.md`.

Wiring note: `folder.refresh.preserve_count` is the existing `itemsPreserved` local in
`ProcessEnumerationResult` (FolderView.Enumeration.cpp:1562 decl, incremented further
down), today only `Debug::Info`-logged — emit it as a metric, do not re-derive it.
`folder.refresh.rename_transfer_count` maps to the rename-transfer path fed by
`_pendingRefreshSelectionRenames` (moved into `refreshSelectionRenames` at :1557).
`ProcessEnumerationResult` currently has no perf-emit calls (perf scopes live only in the
worker), so this is the first metric emission added to the UI-thread result path.

- [x] **Step 2: Extend directory churn coverage**

Extend `folderView_perf_directory_change_storm` or add:

```text
folderView_perf_refresh_preservation
```

It should cover one create, one delete, one same-folder rename, and a bounded
burst. Assert selection/focus/incremental-search preservation and archive
preserve/rebuild counts.

- [x] **Step 3: Decide no-op versus implementation**

Close this task as measured no-op if current debounce, refresh-post coalescing,
and cached-state transfer already keep refresh latency and rebuild counts within
the accepted budget. Implement only the smallest measured extension if a specific
counter proves a problem.

Decision: the current preservation behavior passed correctness once tested, but the
metric family was absent, so Task 8 landed the smallest measured extension:
aggregate refresh metrics, a dedicated refresh-to-paint pending slot, and focused
coverage. No refresh behavior rewrite was needed.

- [x] **Step 4: Verify**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_directory_change_storm --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_refresh_preservation --selftest-timeout-multiplier=4
```

Expected: artifacts contain `folder.refresh.*` rows and correctness assertions
pass.

---

## Task 9: Measure The Icon Pipeline Before Adding More Concurrency

**Files:**

- Modify: `RedSalamander/FolderView.Enumeration.cpp`
- Modify: `RedSalamander/FolderView.Icons.cpp`
- Modify: `RedSalamander/IconCache.cpp`
- Modify: `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- Test: `Tests/PerformanceTests2/FolderIconEnumerationPerfTest.cpp`
- Test: `Tests/PerformanceTests2/FolderIconEnumerationDuplicatePathPerfTest.cpp`

**Goal:** Separate icon-index lookup, shell path lookup, HICON extraction, UI
bitmap conversion, and apply-to-items work before implementing additional
concurrency.

- [x] **Step 1: Preserve the current completed work**

Keep the existing parallel extension and per-file icon-index lookup in
`FolderView.Enumeration.cpp`. Do not regress path dedupe or MTA COM handling.

- [x] **Step 2: Audit the focused queue metrics (no new emits)**

All 12 metrics below already emit in code (verified 2026-06-28). This step adds NO new
emits and MUST NOT introduce new metric names — only confirm each reaches the perf
artifact and is summarized (count/p50/p95/p99/max) by the Task 5 analyzer. Audit each at
its emit site (sampled: `FolderView.ExecuteEnumeration.IconIndex.QueryExtensions`
FolderView.Enumeration.cpp:662; `icons.queue_wait_to_dequeue_us` FolderView.Icons.cpp:603;
`icons.extract_us` FolderView.Icons.cpp:631):

- `FolderView.ExecuteEnumeration.IconIndex.QueryExtensions`
- `FolderView.ExecuteEnumeration.IconIndex.QueryPerFileIcons`
- `FolderView.ExecuteEnumeration.IconIndex.BuildPerFilePaths`
- `FolderView.IconLoading.ProcessQueue`
- `FolderView.IconLoading.BatchUpdate`
- `FolderView.IconLoading.BitmapConversion`
- `icons.queue_wait_to_dequeue_us`
- `icons.extract_us`
- `icons.batch_update_scan_us`
- `iconcache.shgetfileinfo_us`
- `iconcache.lock_wait_slow_us`
- `iconcache.lock_hold_slow_us`

- [x] **Step 3: Add icon-heavy cold and slow fixtures**

Add a focused case with unique `.exe`, `.dll`, `.ico`, `.lnk`, `.url`, `.cpl`,
`.scr`, `.msc`, and `.ocx` paths. Include duplicate paths to prove dedupe and a
slow-injected path to prove one slow request cannot stall all visible icons once
the fix is implemented.

Wire the Task 0 shared hook before exercising the slow path:

```cpp
#ifdef ENABLE_TESTS
SelfTestLatency::Consume(SelfTestLatency::Point::IconExtractSystemIcon);
#endif
```

Place it around the expensive `ExtractSystemIcon` / HICON shell extraction path,
not in thumbnail-only code. `_debugThumbnailProviderMode` applies only to
thumbnail behavior and must not be reused for icon latency.

- [x] Add a cloud/offline-placeholder fixture and recall-avoidance contract
  (Audit-adjacent). Synthesize items flagged
  `FILE_ATTRIBUTE_OFFLINE` / `FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS` /
  `FILE_ATTRIBUTE_RECALL_ON_OPEN` via a test provider (no real OneDrive). Assert the
  visible icon-index lookup AND thumbnail pass do NOT trigger provider/recall I/O. Today
  the icon worker calls `QuerySysIconIndexForPath(work->fullPath.c_str(), 0, false)`
  (FolderView.Enumeration.cpp:844), which runs `SHGetFileInfoW` on the live path
  (IconCache.cpp:1156) and blocks on cloud recall for dehydrated placeholders. Gate
  `IconCache::QuerySysIconIndexForPath` (IconCache.cpp:1119) so that for items flagged
  offline/recall it queries with `SHGFI_USEFILEATTRIBUTES` (no recall) instead. Add a
  new `icons.recall_avoided_count` metric in this Step 3 fixture; do not add it to
  Step 2's "already exists / no new emits" audit list. If a synthesized offline/recall
  provider cannot be built, add an explicit out-of-scope note naming the manual OneDrive
  validation owner.

- [x] **Step 4: Implement bounded concurrency only if the data requires it**

If `FolderView.IconLoading.ProcessQueue` or `icons.extract_us` is a material
background bottleneck, use Windows thread-pool work with a small cap. Each new icon
worker must explicitly initialize COM as MTA at thread/callback entry (prefer
`[[maybe_unused]] auto coInit = wil::CoInitializeEx_failfast(COINIT_MULTITHREADED);`,
matching the contract in IconCache.h) before any `SHGetFileInfoW` / `IImageList::GetIcon`
call. Do NOT assume the Windows thread pool already provides an MTA: the existing
extension and per-file pool callbacks (FolderView.Enumeration.cpp:701-718 and :835-848)
call IconCache shell lookups WITHOUT initializing COM in the callback, so any new
pool-based icon worker must add the `CoInitializeEx` itself (as
`FolderView::EnumerationWorker` already does at FolderView.Enumeration.cpp:276). UI-thread
bitmap conversion stays on the UI thread. Results still return through
`PostMessagePayload`.

**Decision (2026-06-28):** closed as measured no-op for this slice. The focused
fixture delayed one HICON extraction and still resolved a visible bitmap icon
before the delayed extraction finished (`resolvedBeforeSlowExtractFinished=true`
in both Debug and test-enabled Release archives), so no additional icon
extraction concurrency was justified by current evidence. The existing
extension/per-file icon-index thread-pool callbacks were still corrected to
initialize COM MTA before shell/IconCache calls.

- [x] **Step 5: Verify**

Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_iconcache_contention --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_cold_first_visit --selftest-timeout-multiplier=6
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_slow_virtual_provider --selftest-timeout-multiplier=6
vstest.console.exe .\.build\x64\Debug\PerformanceTests2.dll /Tests:FolderIconEnumerationPerfTest,FolderIconEnumerationDuplicatePathPerfTest
# Alternatively run the whole CppUnitTest harness via .\Tools\Run-AllTests.ps1 -Suite Full, which drives PerformanceTests2.dll through vstest.console.exe.
```

Expected: focused metrics identify whether concurrency is accepted or closed as
measured no-op.

**Result (2026-06-28):** closed as measured no-op. Red coverage first failed at
`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_190646` because offline/recall
per-file icon lookup still consumed live path lookups. Green Debug evidence:
`folderView_perf_icon_pipeline_cold_slow`
`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_192653`
(`iconPathLiveLookupConsumeCount=0`, `recallAvoidedMetricRows=3`,
`thumbnailProviderAllowedConsumeCount=0`); supporting Debug cases:
`folderView_perf_iconcache_contention`
`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_192709`,
`folderView_perf_cold_first_visit`
`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_192722`, and
`folderView_perf_slow_virtual_provider`
`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_192733`. Debug build:
`.build/logs/msbuild-20260628_192440_115.log` (0 warnings/errors). The
PerformanceTests2 icon enumeration tests passed through `vstest.console.exe`.
Test-enabled Release build:
`.build/logs/msbuild-20260628_192752_991.log` (0 errors, existing C4883 warning
in File Operations selftests). Release evidence:
`folderView_perf_icon_pipeline_cold_slow`
`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_193018` (`build=Release`,
same recall/thumbnail invariants), plus support cases `2026-06-28_193033`,
`2026-06-28_193045`, and `2026-06-28_193058`. `Tools/Show-PerfRuns.ps1
-FolderViewPreset` summarizes the full Task 9 metric set including
`icons.recall_avoided_count`.

---

## Task 10: Reconsider Dirty Regions, Scroll Rects, And Virtualization

**Files:**

- Modify only after gates pass:
  - `RedSalamander/FolderView.Rendering.cpp`
  - `RedSalamander/FolderView.Selection.cpp`
  - `RedSalamander/FolderView.Interaction.cpp`
  - `RedSalamander/FolderView.Layout.cpp`
  - `RedSalamander/FolderView.h`
- Test:
  - `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`

**Goal:** Implement the older FuturePerformanceIdeas candidates only when the
new representative evidence shows they are still the best next move.

- [x] **Step 1: Gate dirty-region work**

Note: a partial dirty-rect path already exists end to end. The WM_PAINT invalid rect is
clamped to client bounds and drives DXGI Present1's `pDirtyRects`
(FolderView.Rendering.cpp:1003-1025, :1892). Test-only instrumentation already counts
full-client renders: `_debugLastRenderWasFullClient` / `_debugFullClientRenderCount`
(FolderView.Rendering.cpp:1018-1021), exposed to selftests via `DebugGetRenderingSnapshot()`
as `fullClientRenderCount` (FolderView.cpp:1033-1035). Use that existing
`fullClientRenderCount` as the gate metric: proceed only if archives show avoidable
full-client renders (e.g. the scroll handler's `InvalidateRect(_hWnd.get(), nullptr, FALSE)`
at FolderView.Interaction.cpp:395) or dirty area dominates `folder.frame.total_us` /
`folder.frame.input_to_paint_us`.

First slice result: reduce avoidable full-client `InvalidateRect(nullptr)` calls only on
measured no-op scroll paths. The implemented guard suppresses boundary scroll requests
that do not change `_horizontalOffset` / visible item range, and the selftest now asserts
zero product `WM_PAINT` frames for those steps. It does not change viewport-changing
scroll repaint behavior.

- [x] **Step 2: Gate scroll-rect work**

Proceed only if scroll gestures show repaint work dominates after no-freeze and
brush fixes. The first slice should use DXGI scroll rects only for a proven
same-size scroll path and must preserve correctness across resize, DPI, device
loss, theme/font changes, and full repaint boundaries. It must ALSO exclude or
repaint screen-anchored overlays from the scrolled region — the incremental-search
indicator (composited every frame at a fixed client-relative position,
FolderView.Rendering.cpp:1844/:1949), the alert/error overlay (FolderView.Rendering.cpp:1846),
and any in-progress hover/selection/focus rect — because DXGI
`pScrollRect`/`pScrollOffset` blits previously-presented pixels and would otherwise
scroll these fixed elements with the content and corrupt them. Note the current scroll
path does a full-client `InvalidateRect` (FolderView.Interaction.cpp:395), so the
scroll-rect path additionally requires switching scroll to a content-only dirty rect
rather than `InvalidateRect(nullptr)`.

**Decision (2026-06-28): gate not met — not implemented.** Latest test-enabled
Release evidence after the no-op scroll fix (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_195621`)
shows scroll frame p95s below the hard budget (`folder.frame.total_us` p95 `8178us`,
`folder.frame.present_us` p95 `7518us`, both under the 20ms budget), while
viewport-changing scrolls still full-client paint and no-op scrolls paint zero frames.
A DXGI scroll-rect implementation would now be a correctness-heavy feature without a
failing latency gate, so it is deferred until new evidence beats the risk threshold.

- [x] **Step 3: Gate virtualization V2**

Proceed only if huge-folder evidence shows visible work, row rebuilds, or
icon/detail hydration dominates after the narrower fixes. Note: FolderItem today has NO
stable identity and NO per-item visual-state generation. The refresh-preservation path
keys retained records solely on `displayName` (FolderView.Enumeration.cpp:1566-1576);
`stableHash32` is a folder+name-derived rainbow-rendering seed that changes on rename and
is collision-prone (FolderView.h:783, set at FolderView.Enumeration.cpp:485) — NOT an
identity; and `unsortedOrder` is reassigned to the loop index on every enumeration
(FolderView.Enumeration.cpp:1670). Introducing a stable item identity (e.g. NTFS file
id / FRN captured at enumeration) and a per-item visual-state generation field is net-new
design; scope it explicitly as part of V2 rather than implying these keys already exist.
Until then, do not treat `stableHash32` as an identity for retained-record matching.

**Decision (2026-06-28): gate not met — not implemented.** The latest budgeted
test-enabled Release matrix (`Specs/TestRuns/4cb089111a23/Commands/2026-06-28_181518`)
keeps the 10k scale case below the committed hard budgets: working set p95
`126746624` bytes is below `268435456`, working-set bytes/item p95 `12674` is below
`30000`, filter keystroke-to-paint p95 `13250us` is below `60000us`, and sort toggle
`15120us` is below `60000us`. V2 remains a real design item, but it needs stable item
identity and per-item visual-state generation before implementation; the current data
does not justify that scope.

- [ ] **Step 4: Verify**

Run the full FolderView perf matrix in Debug and test-enabled Release, then
`.\Tools\Run-AllTests.ps1 -Suite Full`.

---

## Task 11: Closeout And Spec Migration

**Files:**

- Modify: `Specs/UI/UI_FolderView.md`
- Modify: `Specs/Testing/Testing_PerformanceValidation.md`
- Modify: `Specs/Testing/Testing_TestCoverage.md`
- Move when complete: this plan to `Specs/Plans/Done/`

- [ ] **Step 1: Update authoritative contracts**

Move durable behavior into `Specs/UI/UI_FolderView.md`:

- cached-only first thumbnail behavior,
- device-loss recovery behavior,
- async paste-shortcut completion behavior,
- draw-loop brush reuse rule,
- refresh preservation metric contract,
- icon pipeline metric contract.

- [ ] **Step 2: Update testing contracts**

Move durable validation rules into testing specs:

- minimum sample count for p95/p99 claims,
- Release/test-enabled perf evidence requirement,
- budget-gated perf cases,
- representative scale/cold/slow fixture list,
- archive requirements under `Specs/TestRuns/`.

- [ ] **Step 3: Run final verification**

Run:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full
```

Run the full FolderView perf matrix in Debug and test-enabled Release. Archive
all evidence under `Specs/TestRuns/` and cite the paths in this plan before
moving it to Done.

---

## Acceptance Checklist

- [ ] Confirmed UI-thread stalls are fixed or explicitly blocked with evidence.
- [x] Device-loss recovery is covered by selftest. (`afbbf465d`)
- [x] Draw-loop brush creation is structurally guarded. (`9506d9104`)
- [x] Thumbnail visible path does not call unbounded shell providers. (`6c26367b7`)
- [x] Paste shortcut command latency is separated from shortcut creation latency. (`0dc592dcf`)
- [x] Huge, cold, slow, Debug, and test-enabled Release scenarios exist. (`c51d1830`, `4512fdcb`)
- [ ] Frame-producing perf cases have enough samples for declared p95/p99.
- [x] At least the core FolderView perf cases have automated regression budgets.
- [x] `folder.refresh.*` metrics exist and are archived.
- [x] Icon pipeline concurrency is either accepted with same-machine evidence or
  closed as measured no-op with current metrics.
- [ ] Dirty-region, scroll-rect, and virtualization work is implemented only if
  gated by fresh evidence.
- [x] Durable contracts are merged into `Specs/UI/` and `Specs/Testing/`.
- [ ] Final full-suite result and perf archive paths are recorded.

---

## 2026-06-30 14:05 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_140500_folderview_warpdrive_pause\
```

Resume from this archive, not the 13:01 archive. It includes the current tracked
patch, git status/name-status/stat, untracked list, WIP plan/baton snapshots,
copied focused/cluster/broad `.build\codex-runs` evidence, copied Debug build
logs, `CONTINUATION.md`, parsed run summary JSON, and SHA256 checksums.

New evidence since 13:01:

- Debug build passed after Preferences reverse-navigation diagnostics:
  `.build\logs\msbuild-20260630_131339_002.log`.
- Preferences reverse-navigation is green in focused and predecessor-cluster
  reruns:
  `.build\codex-runs\commands_prefs_category_reverse_focused_diag2_20260630`
  (1/0/0),
  `.build\codex-runs\commands_prefs_compare_to_reverse_cluster_diag2_retry_20260630`
  (12/0/0), and
  `.build\codex-runs\commands_prefs_350_to_reverse_cluster_diag2_retry_20260630`
  (33/0/0).
- Broad Commands moved past Preferences and failed at Shortcuts Escape:
  `.build\codex-runs\commands_broad_after_reverse_diag2_20260630`
  (496 passed / 1 failed / 279 skipped).
- Debug build passed after Shortcuts Escape diagnostics:
  `.build\logs\msbuild-20260630_133552_412.log`.
- Shortcuts Escape is green in focused and block reruns:
  `.build\codex-runs\commands_shortcuts_escape_focused_diag3_20260630`
  (1/0/0) and
  `.build\codex-runs\commands_shortcuts_block_to_escape_diag3_20260630`
  (12/0/0).
- Broad Commands now fails later at Compare Directories Options:
  `.build\codex-runs\commands_broad_after_shortcuts_diag3_20260630`
  (525 passed / 1 failed / 250 skipped).

Current active blocker:

```text
cmd_compare_directories_options_theme_cycle_keeps_surface_legible
Reason: Compare Directories options visible DX edit disappeared after the rainbow theme update.
```

Next implementation move: inspect and add diagnostics around the failing
Compare Options theme-cycle assertion in
`RedSalamander\SelfTest\Commands\Commands.SelfTest.CompareOptions.cpp` before
changing production behavior. Capture all visible ValuePattern states, expected
edit identity/value, options snapshot details, UIA/focus HWND state, and theme
settle state after the rainbow update.

---

## 2026-06-30 14:30 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_143026_folderview_warpdrive_pause\
```

Resume from this archive, not the 14:05 archive. It includes the current tracked
patch, git status/name-status/stat, untracked list, plan/baton snapshots,
copied `.build\codex-runs` evidence, copied Debug build log, WER crash events,
crash-dump manifest, `CONTINUATION.md`, parsed run summary JSON, and SHA256
checksums.

New evidence since 14:05:

- Debug build passed after the Compare Options helper/diagnostic patch:
  `.build\logs\msbuild-20260630_140542_081.log`.
- The Compare Options theme-cycle blocker is green:
  `.build\codex-runs\commands_compare_options_theme_focused_valuepattern_fix_20260630`
  (1/0/0) and
  `.build\codex-runs\commands_compare_options_cluster_valuepattern_fix_20260630`
  (10/0/0).
- Root cause fixed there: `CollectVisibleDescendantValuePatternState(...)`
  now scans all visible edit candidates until one exposes `UIA_ValuePatternId`;
  previously it failed on the first visible edit without ValuePattern even when
  a later visible edit had it.
- Broad Commands now gets past Compare Options but crashes with
  `-1073741819` (`0xC0000005`) at the Pack prompt case:
  `.build\codex-runs\commands_broad_after_valuepattern_fix_20260630`.
- Focused pack prompt and the immediate 3-case cluster pass:
  `.build\codex-runs\commands_pack_prompt_focused_after_valuepattern_fix_20260630`
  (1/0/0) and
  `.build\codex-runs\commands_pack_prompt_three_case_cluster_20260630`
  (3/0/0).
- Explicit 44-case suffix reproduces the native crash:
  `.build\codex-runs\commands_pack_prompt_suffix_600_643_names_20260630`.
- WER identifies both fresh crashes as `textinputframework.dll` access
  violations: exception `0xc0000005`, fault offset `0x000000000006a419`.
  Dump paths are recorded in
  `Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_143026_folderview_warpdrive_pause\crash-context\crash-dumps-manifest.json`.

Current active blocker:

```text
cmd_pane_pack_prompt_uses_dxui_unique_archive_and_packer_extensions
Only fails after the broad/suffix predecessor sequence.
Focused and immediate 3-case cluster runs pass.
Crash: textinputframework.dll 0xc0000005 at offset 0x6a419.
```

Next implementation move: inspect DxUi native text input / TextStore /
WindowHost teardown and the ArchivePackPrompt modal lifecycle, then add narrow
lifecycle tracing around pack prompt creation, native text input activation,
deactivation, and text-store teardown before patching behavior. Generate suffix
repro filters from explicit case names; numeric range syntax such as `600..643`
is treated as a literal case name by this runner path.

---

## 2026-06-30 15:40 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_154019_folderview_warpdrive_pause\
```

Resume from this archive, not the 14:55 archive. It includes the current tracked
patch, git status/name-status/stat, untracked list, WIP plan/baton snapshots,
copied small `.build\codex-runs` evidence, copied Debug build log,
`CONTINUATION.md`, manifest, and SHA256 checksums.

New evidence since 14:55:

- Debug build passed after the Pack prompt lifecycle diagnostics:
  `.build\logs\msbuild-20260630_150318_443.log`.
- The explicit Pack-prompt suffix is now green:
  `.build\codex-runs\commands_pack_prompt_suffix_628_643_lifecycle_trace2_20260630`
  (16/0/0).
- Broad Commands now gets past Pack prompt and fails later:
  `.build\codex-runs\commands_broad_after_pack_lifecycle_trace2_20260630`
  (551 passed / 1 failed / 224 skipped).

Current active blocker:

```text
cmd_pane_fileops_speedLimit_prompt_long_run_open_close_stays_stable
Reason: Custom speed-limit prompt ValuePattern should start with '64,0 KB' during cycle 2.
```

Next implementation move: inspect the speed-limit prompt long-run test,
`CollectVisibleDescendantValuePatternState(...)`, and
`FileOperationsSpeedLimitPromptWindow` snapshot/debug plumbing. If the focused
case does not immediately explain the mismatch, add diagnostics for cycle,
expected snapshot text, observed ValuePattern names/values, candidate counts,
prompt HWND, and focus/foreground HWNDs before changing product behavior.

---

## 2026-06-30 15:58 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-06-30_155822_folderview_warpdrive_pause\
```

Resume from this archive, not the 15:40 archive. It includes the current tracked
patch, git status/name-status/stat, untracked list, WIP plan/baton snapshots,
selected `.build\codex-runs` traces/results, the latest Debug build log,
`CONTINUATION.md`, manifest, and SHA256 checksums.

New evidence since 15:40:

- Debug build passed after diagnostic-only speed-limit ValuePattern context:
  `.build\logs\msbuild-20260630_154734_089.log`.
- The speed-limit blocker did not reproduce in focused, three-case, 21-case, or
  diagnostic focused reruns:
  `.build\codex-runs\commands_speedlimit_longrun_focused_20260630_resume`,
  `.build\codex-runs\commands_speedlimit_three_case_cluster_20260630_resume`,
  `.build\codex-runs\commands_fileops_issues_to_speedlimit_cluster_20260630_resume`,
  and `.build\codex-runs\commands_speedlimit_longrun_focused_diag_20260630_resume`.
- Broad Commands now fails earlier:
  `.build\codex-runs\commands_broad_speedlimit_diag_20260630_resume`
  (215 passed / 1 failed / 560 skipped).
- Focused credential secret-toggle and the 21-case connection/credential cluster
  are green:
  `.build\codex-runs\commands_credential_secret_toggle_focused_20260630_resume`
  and `.build\codex-runs\commands_connection_to_credential_toggle_cluster_20260630_resume`.

Current active blocker:

```text
cmd_connection_credential_prompt_pointer_click_toggles_secret_visibility
Reported reason: Credential prompt did not restore masked secret visibility on the second real click.
Actual trace evidence: masking was restored (`visible=0`, `checked=0`), but focus dropped to `None` (`focus=0`) after the second click in the broad run.
```

Next implementation move: add second-click diagnostics to
`TestConnectionCredentialPromptPointerClickTogglesSecretVisibility` in
`RedSalamander\SelfTest\Commands\Commands.SelfTest.Connections.cpp`. Capture the
second down/up snapshots, host/rect, final masked snapshot, and split the test
failure between "secret masking restored" and "focus stayed on toggle" before
changing product behavior.

---

## 2026-07-01 12:03 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_120312_folderview_warpdrive_pause\
```

Resume from this archive, not the 08:36 archive. It includes the tracked patch,
git status/name-status/stat, untracked list, WIP plan/baton snapshots, copied
July 1 `.build\codex-runs` evidence, copied Debug build logs, selected volatile
`SelfTest\last_run` files, crash-dump/WER manifests, and `CONTINUATION.md`.

New evidence since 08:36:

- Temporary icon hot-path diagnostics were removed from `IconCache.cpp` and
  `FolderView.Icons.cpp`; the diagnostic strings no longer match in those files.
- Debug rebuild after the General Preferences diagnostic hardening passed:
  `.build\logs\msbuild-20260701_090928_050.log`.
- Broad Commands is green:
  `.build\codex-runs\commands_broad_after_general_diag_20260701`
  (774 passed / 0 failed / 2 skipped).
- Full FileOperations is green:
  `.build\codex-runs\fileops_full_after_general_diag_20260701`
  (102 passed / 0 failed / 20 skipped).
- ViewerPETests is green:
  `.build\codex-runs\viewerpe_after_commands_fileops_20260701\viewerpe.log`.
- Tools Pester is green by log evidence:
  `.build\codex-runs\tools_pester_after_commands_fileops_20260701\pester.log`
  reports `Passed: 108 Failed: 0 Skipped: 0`; the wrapper exit 1 was a wrapper
  artifact after redirecting the `-PassThru` object.
- DxUiTests is green after the TSF teardown-order test guard was updated:
  `.build\codex-runs\dxui_tests_after_tsf_guard_update_20260701\dxui-tests.log`.

Current active blocker:

```text
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
Commands exits -1073740771 (0xC000041D) before commands\results.json exists.
Direct broad Commands is green, so this is currently Full-order/context sensitive.
```

Latest Full evidence:

```text
.build\codex-runs\full_suite_skipbuild_after_dxui_guard_20260701
Overall: 279 passed / 0 failed / 49 skipped / exit 1
CompareDirectories: 163 passed / 0 failed / 29 skipped
Commands: exit -1073740771, 0 cases recorded
FileOperations: 102 passed / 0 failed / 20 skipped
DxUiTests and every downstream standalone/Pester suite: exit 0
```

Next implementation move: inspect the latest Full-run Commands crash context in
the copied `run.log`, `SelfTest\last_run` snapshots, and crash-dump/WER
manifest. Reproduce with the smallest proven ordering (likely
CompareDirectories followed by Commands in Full context) before changing code.
After this blocker is fixed, rerun Full skip-build, then the final Debug and
test-enabled Release FolderView perf matrix plus `Tools\Show-PerfRuns.ps1
-FolderViewPreset -FailOnQuality -ShowBuildFlavor`. Do not move this plan to
Done until Full and the final perf matrix are both green and cited.

---

## 2026-07-01 12:09 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_120949_folderview_warpdrive_pause\
```

Resume from this archive. It supersedes the 12:03 archive because it also
copies the latest crash dump itself:

```text
crash-context\RedSalamander.exe.90544.dmp
```

The archive includes `git\tracked.patch`, branch/status/diff snapshots, WIP
plan and continuation baton snapshots, changed testing-spec snapshots, selected
green/red `.build\codex-runs` evidence, the latest Debug build log, WER rows,
the stackwalk output, the copied dump, `CONTINUATION.md`, and checksum/manifest
files. Continue with the Full-order Commands crash investigation; the next
move remains to inspect/reproduce the TSF/TextInputFramework crash around the
FileOperations popup menu helpers before changing code.

---

## 2026-07-01 12:32 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_123238_folderview_warpdrive_pause\
```

This supersedes the 12:09 archive. It includes the current tracked patch,
branch/status/diff snapshots, WIP and untracked continuation/remediation plan
snapshots, selected green evidence, the fresh broad-Commands failure log, the
latest Commands trace, `stackwalk_95456.txt`, `pdb_publics.txt`, and the copied
crash dump `RedSalamander.exe.95456.dmp`.

New evidence refined the suspected crash target: Compare followed by the
speed-limit prompt prefix is green, while broad Commands after Compare crashes
with `textinputframework.dll` `0xc0000005`. The latest trace stops after
starting `cmd_pane_navigation_create_directory_prompt_keeps_navigation_shell_stable`;
there is no `commands/results.json` because the process crashed. Resume by
inspecting and reproducing the create-directory prompt lifecycle / DxUi native
text input teardown path before changing product behavior. Do not move this
plan to Done until broad Commands, Full skip-build, and the final Debug plus
test-enabled Release FolderView perf matrices are green and cited.

---

## 2026-07-01 13:27 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_132759_folderview_warpdrive_pause\
```

This supersedes the 12:32 archive. It includes current git status, full tracked
patch, targeted Navigation diagnostics patch, WIP/continuation/remediation plan
snapshots, selected green evidence, latest broad Commands red artifacts, latest
Debug build logs, manifest, and `CONTINUATION.md`.

The old create-directory/TextInputFramework crash did not reproduce in focused,
prefix, or slice runs. Current broad Commands evidence is
`.build\codex-runs\commands_broad_after_nav_diag_20260701_resume`
(772 passed / 2 failed / 2 skipped). The active failures are
`cmd_pane_navigationView_disk_info_region_keyboard_activation_opens_menu` and
`cmd_pane_navigationView_edit_suggest_keyboard_routing`, both Navigation-view
focus/stability assertions. Resume by inspecting those two tests and their
NavigationView focus/debug helpers first, looking for stale snapshot formatting
or timing diagnostics before changing product behavior. Do not move this plan to
Done until broad Commands, Full skip-build, and the final Debug plus
test-enabled Release FolderView perf matrices are green and cited.

---

## 2026-07-01 13:41 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_134139_folderview_warpdrive_pause\
```

This supersedes the 13:27 archive. It includes current git status, full tracked
patch, untracked plan list, WIP/continuation/remediation plan snapshots, the
green NavigationView family run, latest broad Commands red artifacts, latest
Debug build log, manifest, and `CONTINUATION.md`.

The previous NavigationView broad-suite blocker is now green after
`ShowDiskInfoDropdown(...)` restores owner pane focus after the disk-info DxUi
menu closes. Evidence:
`.build\codex-runs\commands_navigationview_family_after_disk_focus_restore_20260701`
(16 passed / 0 failed) and
`Z:\src\RedSalamander\.build\logs\msbuild-20260701_133358_346.log` (0 warnings
/ 0 errors).

Current broad Commands evidence is
`.build\codex-runs\commands_broad_after_disk_focus_restore_20260701`
(227 passed / 1 failed / 548 skipped). The active failure is
`cmd_preferences_dialog_viewers_theme_cycle_keeps_surface_legible`, reason
`Preferences Viewers page did not settle before theme-cycle validation.` Resume
by inspecting `TestPreferencesDialogViewersThemeCycleKeepsSurfaceLegible(...)`
and the Viewers settle/snapshot helpers, then reproduce focused and with the
immediate predecessor cluster before patching. Do not move this plan to Done
until broad Commands, Full skip-build, and the final Debug plus test-enabled
Release FolderView perf matrices are green and cited.

---

## 2026-07-01 14:09 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_140931_folderview_warpdrive_pause\
```

This supersedes the 13:41 archive. It includes current git status, full tracked
patch, untracked list, WIP/continuation/remediation plan snapshots, selected
green Preferences and Connection Manager evidence, the previous broad
Connection Manager red, latest broad Find red artifacts, latest Debug build log,
manifest, SHA256 checksums, and `CONTINUATION.md`.

The previous Preferences Viewers blocker is green in focused/predecessor runs:
`Specs\TestRuns\4cb089111a23\Commands\2026-07-01_134530` (1 passed / 0
failed) and `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_134615` (3 passed
/ 0 failed).

The next broad Commands blocker at
`Specs\TestRuns\4cb089111a23\Commands\2026-07-01_135019` (204 passed / 1 failed
/ 571 skipped) was
`cmd_connection_manager_window_enter_from_dx_input_routes_default_connect`,
reason `Connection Manager did not create and select a new editable row before
Enter/default-button validation.` Focused, predecessor pair, and full
Connection Manager prefix reruns passed at `2026-07-01_135101`,
`2026-07-01_135143`, and `2026-07-01_135317`. A diagnostic-only Connection
Manager selftest patch is built in
`Z:\src\RedSalamander\.build\logs\msbuild-20260701_135517_041.log` (0 warnings
/ 0 errors).

Current broad Commands evidence is
`Specs\TestRuns\4cb089111a23\Commands\2026-07-01_140737` (445 passed / 1 failed
/ 330 skipped). The active failure is
`cmd_pane_find_dialog_restored_combined_view_state_copy_follows_visible_columns`,
reason `Find logical Name sort or widened reordered layout did not settle before
restored combined-view-state copy persistence.` Resume by inspecting/reproducing
that Find combined-view-state copy settle path first. If focused and immediate
predecessor-cluster repros pass, add narrow diagnostics around Find snapshot,
logical sort, column order/widths, copy payload/order, selected row, active
search text, and settle-loop timing before changing product behavior. Do not
move this plan to Done until broad Commands, Full skip-build, and the final
Debug plus test-enabled Release FolderView perf matrices are green and cited.

## 2026-07-01 14:25 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_142558_folderview_warpdrive_pause\
```

This supersedes the 14:09 archive. It includes current git status, full tracked
patch, focused diagnostic patch, WIP/continuation/remediation plan snapshots,
selected broad/narrow Commands archives, latest build logs, selected runner
stdout logs, and `CONTINUATION.md`.

The active broad Commands failure is still
`cmd_pane_find_dialog_restored_combined_view_state_copy_follows_visible_columns`
from `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_140737` (445 passed / 1
failed / 330 skipped). The important detail remains that the failure happens
before the persisted copy check: `Find logical Name sort or widened reordered
layout did not settle before restored combined-view-state copy persistence.`

Narrow repros are green:

- Focused case: `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_141539`.
- Immediate predecessor pair: `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_141609`.
- 26-case Find block: `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_141731`.
- Diagnostic focused case: `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_142342`.
- Diagnostic 26-case Find block: `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_142436`.

Debug build with the expanded Find assertion passed at
`Z:\src\RedSalamander\.build\logs\msbuild-20260701_142132_723.log` (0 warnings
/ 0 errors). The earlier launcher-level build timeout at
`Z:\src\RedSalamander\.build\logs\msbuild-20260701_141910_795.log` is archived
only for chronology and is not success evidence.

Resume by rerunning broad Commands with the diagnostic build:

```powershell
$log = ".build\logs\commands-broad-find-diag-20260701_<time>.out.log"
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 *> $log
```

If the same Find case fails, use the expanded message to identify the exact
failing predicate before patching. If broad Commands passes, rerun it once more
before trusting the pass, because the current blocker is broad-order and did not
reproduce in the narrower clusters. Do not move this plan to Done until broad
Commands, FileOps, ViewerPE, Tools Pester, Full skip-build, and the final Debug
plus test-enabled Release FolderView perf matrices are green and cited.

## 2026-07-01 15:11 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_151116_folderview_warpdrive_pause\
```

This supersedes the 14:59 archive. It includes current git status, full tracked
patch, untracked list, WIP/continuation/remediation plan snapshots, latest
runner/build logs, the exact 107-case filter, selected broad/narrow Commands
archives, copied `RedSalamander.exe.90048.dmp` and `RedSalamander.exe.74560.dmp`
dumps, manifest, SHA256 checksums, and `CONTINUATION.md`. The archive is large
because it intentionally carries forward the self-contained 14:48 TSF crash work
archive.

The latest 107-case FileOps+Navigation+Dialogs rerun failed cleanly at
`cmd_pane_navigation_go_to_root_directory_keeps_navigation_shell_stable`.
Evidence:
`.build\logs\commands-fileops-navigation-dialogs-to-createdir-rerun-20260701_150402.out.log`
and `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_150432` (66 passed / 1
failed / 40 skipped). The pre-diagnostic assertion saw `currentPath=''`,
`historyCount=0`, `itemCount=30`, and no popup/edit state, while
`selftest_run_trace.txt` shows `FolderView::ProcessEnumerationResult` reached
`displayedFolder='C:\'` with 30 items at generation 445.

A diagnostic-only Go To Root assertion patch was added in
`RedSalamander\SelfTest\Commands\Commands.SelfTest.Navigation.cpp`; it now
prints `panePath`, `expectedRoot`, `baselinePath`, history counts, folder-view
HWNDs, and focus HWND. Debug build
`Z:\src\RedSalamander\.build\logs\msbuild-20260701_150646_256.log` passed with
0 warnings / 0 errors.

Resume by rerunning the exact 107-case filter with the new root-navigation
diagnostics:

```powershell
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$filter = (Get-Content .build\logs\commands-fileops-navigation-dialogs-to-createdir-cases.txt) -join ','
$log = ".build\logs\commands-fileops-navigation-dialogs-to-createdir-rootdiag-$stamp.out.log"
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter $filter -TimeoutMultiplier 4 *> $log
```

If the same case fails, inspect the expanded fields before changing product
behavior. If `panePath='C:\'` and `currentPath=''`, inspect
`FolderWindow::SetFolderPath(Pane, const std::filesystem::path&)` and the
NavigationView `SetPath`/display-path update path. If focused HWNDs mismatch,
classify it as a focus/snapshot stability issue first. If the slice passes,
rerun it once more or broad Commands before trusting the pass. Do not move this
plan to Done until broad Commands, FileOps, ViewerPE, Tools Pester, Full
skip-build, and the final Debug plus test-enabled Release FolderView perf
matrices are green and cited.

## 2026-07-01 15:19 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_151949_folderview_warpdrive_pause\
```

This supersedes the 15:11 archive as the resume pointer. It is intentionally a
compact delta archive: it includes current git status, full tracked patch,
untracked list, WIP/continuation/remediation plan snapshots, latest runner/build
logs, the exact 107-case filter, the 15:16 failing Commands archive, the 15:18
focused passing Commands archive, manifest, SHA256 checksums, and
`CONTINUATION.md`. The previous 15:11 archive remains the heavy TSF crash/dump
bundle.

The exact 107-case FileOps+Navigation+Dialogs rerun with the root-navigation
diagnostic build now fails earlier at
`cmd_pane_navigation_change_directory_keeps_navigation_shell_stable`. Evidence:
`.build\logs\commands-fileops-navigation-dialogs-to-createdir-rootdiag-20260701_151636.out.log`
and `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_151702` (42 passed / 1
failed / 64 skipped). Failure reason:
`Change Directory did not enter live path edit mode cleanly.`

Focused Change Directory passes:
`.build\logs\commands-change-directory-focused-20260701_151804.out.log` and
`Specs\TestRuns\4cb089111a23\Commands\2026-07-01_151806` (1 passed / 0 failed /
0 skipped). Therefore the active blocker is order-sensitive leakage from the
predecessor FileOps/Selection/Navigation cases, not a focused Change Directory
failure.

A Release `RedSalamander.exe` process was running at archive time; check and
stop/wait for app processes before running new selftests.

Resume by narrowing the predecessor cluster first:

```powershell
$pair = 'cmd_pane_navigation_show_folders_history_keeps_navigation_shell_stable,cmd_pane_navigation_change_directory_keeps_navigation_shell_stable'
$log = ".build\logs\commands-change-directory-after-history-$(Get-Date -Format 'yyyyMMdd_HHmmss').out.log"
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter $pair -TimeoutMultiplier 4 *> $log
```

If that passes, try the context-menu + history + change-directory triple, then
the eight-case navigation predecessor block. If only the 107-case slice
reproduces, add diagnostic-only fields to
`TestChangeDirectoryKeepsNavigationShellStable` before changing product
behavior. Capture NavigationView snapshot fields, edit-mode/child-window flags,
current path/edit text, edit host/input HWNDs, debug enter-edit counters and
abort reason, refresh/item/selected counts, active/focused pane, focus HWND, and
current pane path.

Do not move this plan to Done until broad Commands, FileOps, ViewerPE, Tools
Pester, Full skip-build, and the final Debug plus test-enabled Release
FolderView perf matrices are green and cited.

## 2026-07-01 15:54 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_155455_folderview_warpdrive_tsf_prompt_pause\
```

This supersedes the 15:19 archive as the resume pointer. It contains current
git status, full tracked patch, untracked list, selected source snapshots,
latest focused/cluster/exact-107 Commands logs, the TSF diagnostic build log,
and `CONTINUATION.md`. The heavy crash evidence remains in
`Specs\TestRuns\4cb089111a23\Commands\2026-07-01_154113_fileops_navigation_dialogs_filter_prompt_av\`
and includes the copied `RedSalamander.exe.91588.dmp` dump plus WER evidence.

The Change Directory predecessor minimization no longer reproduces: the
history+change pair, context+history+change triple, and eight-case Navigation
block all passed. The exact 107-case FileOps+Navigation+Dialogs slice then got
past Change Directory and crashed later at
`cmd_pane_filter_prompt_live_dx_interaction` with `textinputframework.dll`
`0xc0000005` offset `0x000000000006a419`; no `commands_results.json` was
written. Focused filter-prompt, filter-prompt pair, and ChangeCase-to-filter
suffix reruns passed, so the active blocker remains order/timing-sensitive TSF
teardown around modal DxUi prompt windows.

A diagnostic-only `DeactivateNativeTextInputTsf()` trace-stage patch was built
successfully at `Z:\src\RedSalamander\.build\logs\msbuild-20260701_154919_197.log`,
and the exact 107-case slice passed with that diagnostic build at
`Specs\TestRuns\4cb089111a23\Commands\2026-07-01_155218` (107 passed / 0
failed / 0 skipped). Treat this as timing perturbation, not a fix.

Resume by first building/running the newly added DxUi source guard
`TestFolderViewDxUiPromptsDeactivateNativeTextInputBeforeDestroyWindow()` to
confirm RED, then remove the temporary TSF trace instrumentation, centralize
prompt close in the five `FolderWindow.FileSystem.cpp` modal prompt classes so
they call `_dxHost.SetFocusControl(nullptr)` and
`_dxHost.DeactivateTextInput(false)` before `_hWnd.reset()`, rerun DxUiTests,
rebuild RedSalamander Debug, rerun focused filter prompt, the filter-prompt
pair, the exact 107-case slice twice, then resume broad Commands/FileOps/
ViewerPE/Tools Pester/Full and final perf closeout. Do not move this plan to
Done until those gates and final Debug plus test-enabled Release FolderView perf
matrices are green and cited.

## 2026-07-01 16:19 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_161900_folderview_warpdrive_prompt_teardown_selectall_pause\
```

This supersedes the 15:54 archive as the resume pointer, while the 15:54 heavy
crash archive remains the dump/WER evidence bundle. It contains current git
status, full tracked patch, untracked list, WIP/baton snapshots, selected source
snapshots, latest build and Commands logs, copied focused/cluster/107-case
Commands archives, manifest, checksums, and `CONTINUATION.md`.

The temporary TSF trace instrumentation has been removed. The five
`FolderWindow.FileSystem.cpp` modal prompt classes now centralize prompt close
through helpers that call `_dxHost.SetFocusControl(nullptr)` before
`_hWnd.reset()`. The direct `_dxHost.DeactivateTextInput(false)` approach was
rejected by the app build because `WindowHost::DeactivateTextInput` is private;
`SetFocusControl(nullptr)` is the public boundary that deactivates the old text
input control.

Verification after this change:

- DxUiTests build green:
  `Z:\src\RedSalamander\.build\logs\msbuild-20260701_160504_430.log`.
- RedSalamander Debug build green:
  `Z:\src\RedSalamander\.build\logs\msbuild-20260701_160650_265.log`.
- Focused filter prompt green:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_160906`.
- Focused Select All green:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_161046`.
- Select All predecessor clusters green:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_161108` and
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_161132`.
- Exact 107-case slice green once:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_161233`.
- Exact 107-case slice then failed cleanly, without `textinputframework.dll`, at
  `cmd_pane_selection_select_all_keeps_navigation_shell_stable`:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_161310`.

A diagnostic-only Select All assertion expansion was added in
`RedSalamander\SelfTest\Commands\Commands.SelfTest.Navigation.cpp`, and the
Debug app build after that diagnostic patch passed at
`Z:\src\RedSalamander\.build\logs\msbuild-20260701_161448_476.log`.

The 16:19 attempt to recapture full `DxUiTests.exe` console output timed out at
the command wrapper and is archived only for chronology, not as green evidence.

Resume by rerunning the exact 107-case slice with the new Select All diagnostic
fields:

```powershell
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$filter = (Get-Content .build\logs\commands-fileops-navigation-dialogs-to-createdir-cases.txt) -join ','
$log = ".build\logs\commands-fileops-navigation-dialogs-to-createdir-selectall-diag-$stamp.out.log"
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter $filter -TimeoutMultiplier 4 *> $log
```

If the same Select All case fails and the expanded fields show the pane path,
baseline path, selection count, and focused FolderView are correct, inspect the
NavigationView snapshot/current-path path before changing product Select All
behavior. If the 107-case slice passes, rerun it once more before trusting the
pass, then resume broad Commands. Do not move this plan to Done until broad
Commands, FileOps, ViewerPE, Tools Pester, Full skip-build, and final Debug plus
test-enabled Release FolderView perf matrices are green and cited.

## 2026-07-01 16:32 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_163245_folderview_warpdrive_createdir_order_pause\
```

This supersedes the 16:19 archive as the resume pointer. It contains current git
status, full tracked patch, untracked list, source snapshots, WIP-plan snapshot,
latest build and Commands logs, copied focused/cluster/107-case Commands
archives, manifest, checksums, and `CONTINUATION.md`.

The exact 107-case FileOps+Navigation+Dialogs slice no longer crashes in
`textinputframework.dll` and the previously exposed Select All and Compare
Directories shell-stability failures no longer reproduce as the active blocker.
The latest exact 107-case slice reached the tail and failed cleanly at
`cmd_pane_createDirectory_prompt_live_dx_interaction`:

```text
Specs\TestRuns\4cb089111a23\Commands\2026-07-01_163009
```

Result: 105 passed / 1 failed / 1 skipped. Failure:

```text
Create-directory prompt should restore the default folder name 'New folder' after live DX cancel reopen.
```

Focused and immediate-pair reruns passed:

```text
Specs\TestRuns\4cb089111a23\Commands\2026-07-01_163037
Specs\TestRuns\4cb089111a23\Commands\2026-07-01_163057
```

Treat this as order-sensitive until proven otherwise. Resume by inspecting
`RunCreateDirectoryPromptModalCycle(...)` and
`TestCreateDirectoryPromptLiveDxInteraction(...)` in
`RedSalamander\SelfTest\Commands\Commands.SelfTest.Dialogs.cpp`, then inspect
the create-directory prompt implementation in
`RedSalamander\FolderWindow.FileSystem.cpp`. Narrow the predecessor suffix
before changing product behavior, or add diagnostic-only fields to the
create-directory live assertion that capture reopen text, selection range,
whether `New folder` already exists in the test root, cancel-cycle text, prompt
focus/child-window state, and the relevant NavigationView snapshot fields.

Several diagnostic-only navigation assertion expansions remain in
`RedSalamander\SelfTest\Commands\Commands.SelfTest.Navigation.cpp`. Keep them
while they are useful for the active reliability pass, but review and trim
diagnostic noise before final closeout. Do not move this plan to Done until
broad Commands, FileOps, ViewerPE, Tools Pester, Full skip-build, and final
Debug plus test-enabled Release FolderView perf matrices are green and cited.

## 2026-07-01 16:49 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_164918_folderview_warpdrive_tsf_crash_pause\
```

This supersedes the 16:32 archive as the resume pointer. It contains current git
status, diff stat, full tracked patch, copied latest build/test logs, copied
last-run traces, manifest, checksums, and `CONTINUATION.md`.

The create-directory live-prompt failure from the previous archive did not
reproduce in focused/suffix reruns. The exact 107-case
FileOps+Navigation+Dialogs slice then exposed an earlier intermittent FileOps
speed-limit prompt issue. `RedSalamander\FolderWindow.FileOperations.Popup.cpp`
now closes the speed-limit prompt through a helper that calls
`_dxHost.SetFocusControl(nullptr)` before `_hWnd.reset()`, matching the modal
prompt close discipline added in `FolderWindow.FileSystem.cpp`.

Verification after that focused change:

- RedSalamander Debug build green:
  `Z:\src\RedSalamander\.build\logs\msbuild-20260701_164131_076.log`.
- Speed-limit three-case cluster failed once with a UIA/live-text mismatch:
  `snapshotText='64,0 KB'`, but UIA observed `name/value='64,0 KBd'`:
  `.build\logs\commands-speedlimit-threecase-after-closefocus-20260701_164355.out.log`.
- The same speed-limit three-case cluster passed on immediate rerun:
  `.build\logs\commands-speedlimit-threecase-after-closefocus-rerun-20260701_164602.out.log`.
- The exact 107-case slice then crashed before producing a Commands archive:
  `.build\logs\commands-fileops-navigation-dialogs-to-createdir-after-speedlimit-closefocus-20260701_164622.out.log`.

Crash evidence:

```text
Faulting module: C:\WINDOWS\SYSTEM32\textinputframework.dll
Exception code: 0xc0000005
Fault offset: 0x000000000006a419
Dump: C:\Users\eric\AppData\Local\CrashDumps\RedSalamander.exe.46284.dmp
```

The last selftest trace reached
`cmd_pane_navigation_create_directory_prompt_keeps_navigation_shell_stable` and
ended after `FolderView::ProcessEnumerationResult: end`. `cdb.exe`/WinDbg was
not found on the checked PATH or under the Windows Kits debugger path, so no
stack is archived yet.

Resume by getting a stack for the fresh dump if a debugger is available, then
reproduce the smallest suffix that includes the speed-limit navigation-shell
prompt and create-directory navigation-shell prompt. If that suffix does not
reproduce, rerun the exact 107-case filter from
`.build\logs\commands-fileops-navigation-dialogs-to-createdir-cases.txt`. Treat
the `64,0 KBd` mismatch as evidence of shared DxUi/UIA/native-text state until
disproven; inspect `Common\DxUi\DxUi.NativeTextInput.cpp`,
`Common\DxUi\DxUi.Accessibility.cpp`, `Common\DxUi\DxUi.WindowHost.cpp`, and
prompt close paths before making another product change. Do not move this plan
to Done until broad Commands, FileOps, ViewerPE, Tools Pester, Full skip-build,
and final Debug plus test-enabled Release FolderView perf matrices are green
and cited.

## 2026-07-01 17:05 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_170543_folderview_warpdrive_associatefocus_pause\
```

This supersedes the 16:49 archive as the resume pointer. It contains current git
refs/status, diff stat, full tracked binary patch, untracked list, WIP/baton
snapshots, source snapshots, copied Commands narrowing logs, copied DxUiTests
build logs, recent crash-dump/WER manifests, manifest, checksums, and
`CONTINUATION.md`.

Additional work completed after the 16:49 archive:

- The 107-case crash was narrowed further. The minimal crashing recipe found so
  far is `[17..20,28,33,45..55,56..58]`, which reaches
  `cmd_pane_navigation_create_directory_prompt_keeps_navigation_shell_stable`
  after `cmd_pane_navigation_item_properties_window_keeps_navigation_shell_stable`
  and crashes in `textinputframework.dll` with exception code `0xc0000005` and
  fault offset `0x000000000006a419`.
- Several surrounding slices passed independently, so treat this as cumulative
  native-text/TSF/HWND churn rather than a single focused case until proven
  otherwise.
- A TDD guard was added to `Tests\DxUiTests\DxUiTests.NativeTextInput.cpp` for
  `ITfThreadMgr::AssociateFocus(...)` on native TSF activation and deactivation.
  RED evidence: DxUiTests failed with `FAILED: native TSF activate associates
  the host HWND with the active document manager`.
- `Common\DxUi\DxUi.NativeTextInput.cpp` now associates the host HWND with the
  pushed TSF document manager before `SetFocus(documentMgr.get())` and clears
  the HWND association before `documentMgr->Pop(TF_POPF_ALL)`.
- GREEN evidence after that change: DxUiTests Debug build
  `Z:\src\RedSalamander\.build\logs\msbuild-20260701_170302_597.log` passed and
  `.\.build\x64\Debug\DxUiTests.exe` exited 0 with `All DxUi tests passed.`.

This AssociateFocus change is not yet proven through RedSalamander Commands.
Resume by building RedSalamander Debug, then rerun the minimal crash recipe from
`CONTINUATION.md`. If it passes, rerun it once more, then rerun the exact
107-case filter from
`.build\logs\commands-fileops-navigation-dialogs-to-createdir-cases.txt`.
Do not move this plan to Done until broad Commands, FileOps, ViewerPE, Tools
Pester, Full skip-build, and final Debug plus test-enabled Release FolderView
perf matrices are green and cited.

## 2026-07-03 23:58 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_235801_folderview_warpdrive_full_red_compare_commands_baton\
```

This supersedes the 22:56 archive as the immediate resume pointer. It captures
the CompareDirectories status-diagnostic patch, Debug rebuild, focused and broad
Compare green runs, the new proper Full red run, latest `SelfTest\last_run`
top-level artifacts, current git metadata and patches, source snapshots for the
Compare and Commands red points, copied Compare/Commands suite archives, and
crash manifests.

Additional work completed after the 22:56 archive:

- Added narrow diagnostics to the
  `search_service_sqlite_ntfs_traversal_seed_stays_degraded` wait. This does
  not change behavior; it preserves the last successful service status and last
  `GetStatus` HRESULT in the failure reason.
- Rebuilt `RedSalamander` Debug:
  `.build\codex-runs\msbuild_redsalamander_compare_ntfs_status_diag_20260704.log`
  (0 warnings / 0 errors).
- Focused Compare target is green:
  `.build\codex-runs\runall_compare_ntfs_seed_focused_status_diag_20260704.log`
  (1 passed / 0 failed / 0 skipped).
- Broad CompareDirectories is green:
  `.build\codex-runs\runall_compare_broad_status_diag_20260704.log` and
  `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-03_231441`
  (208 passed / 0 failed / 30 skipped).
- The proper Full run with build is still red:
  `.build\codex-runs\full_after_compare_status_diag_20260704.log` exited 1
  with 1098 passed / 2 failed / 52 skipped / 1152 total. FileOperations,
  DxUiTests, ViewerPETests, Tools Pester, and downstream native/script suites
  passed. Remaining red gates:
  - CompareDirectories
    `search_service_sqlite_ntfs_traversal_seed_stays_degraded`: the diagnostic
    now proves this is not simply a slow warmup. Last status reports
    `running=false`, `completed=0`, `failed=1`, `indexedVolumes=0`,
    `lastFailure=true`, `failureHr=0x80070002`, and
    `lastGetStatusHr=0x80070002`.
  - Commands `cmd_pane_navigation_context_menu_keeps_navigation_shell_stable`:
    the navigation shell path/item/session state matches, but focus restoration
    is lost after context-menu close (`focusedFolderView=0x0`,
    `focusHwnd=0x0`).

Resume with systematic debugging on these two Full-order reds. For Compare,
start from the `ERROR_FILE_NOT_FOUND` startup-warmup failure path in
`Common\SearchServiceBroker.cpp` / `LocalSearchIndexCore::Repository::EnsureReadyForRoot`
and the forced NTFS traversal-seed path; do not paper over it with a longer
timeout. For Commands, focused-run
`cmd_pane_navigation_context_menu_keeps_navigation_shell_stable` and the nearby
Navigation prefix, then prove whether focus loss is selftest setup,
menu-close focus handoff, or another Full-order focus leak before patching.
Do not move this plan to Done until proper Full is green and the final Debug
plus test-enabled Release FolderView perf matrices are green and cited.

## 2026-07-03 22:56 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_225627_folderview_warpdrive_full_compare_ntfs_baton\
```

This supersedes the 21:59 archive as the immediate resume pointer. It captures
the cleaned DxUi pointer-root fix, the focused and broad DxUi green runs, the
focused and broad ViewerPE green runs, the latest proper Full red run, the
focused CompareDirectories rerun, current git status/diff metadata, no-TestRuns
staged/unstaged patches, source snapshots for the WIP plan, DxUi menu test,
ViewerPE test, CompareDirectories search/index case, `SearchServiceBroker`, and
`Run-AllTests.ps1`, plus `SelfTest\last_run` top-level output and crash
manifests.

Additional work completed after the 21:59 archive:

- Removed the temporary DxUi menu diagnostics and kept only the test-side
  `pointerRootSwitchArmed` guard in
  `Tests\DxUiTests\DxUiTests.Menu.cpp`. The guard ignores incidental setup
  root-switch callbacks from captured/live cursor movement while preserving the
  deliberate inside-popup hover probe.
- `DxUiTests` rebuild and verification are green:
  `.build\codex-runs\msbuild_dxuitests_menu_cleanup_20260704.log`,
  `.build\codex-runs\dxui_menu_suite_after_cleanup_20260704.log`, and
  `.build\codex-runs\dxui_direct_broad_after_menu_fix_20260704.log`.
- `ViewerPETests` focused and direct broad verification are green:
  `.build\codex-runs\viewerpe_vlc_tab_focus_after_dxui_green_20260704.log`
  and `.build\codex-runs\viewerpe_direct_broad_after_dxui_green_20260704.log`.
- The proper Full run with build is still red, but now only on
  CompareDirectories:
  `.build\codex-runs\full_after_dxui_viewerpe_green_20260704.out.log`
  exited 1 with 1099 passed / 1 failed / 52 skipped / 1152 total. The failing
  case is `search_service_sqlite_ntfs_traversal_seed_stays_degraded`, reason
  `Forced NTFS traversal-seed service did not finish startup warmup.`
- The same CompareDirectories case passed focused immediately afterward at
  `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-03_225200`
  (1 passed / 0 failed / 0 skipped, about 10.3s).

Resume by root-causing the Full-order CompareDirectories warmup timeout before
changing product code. Start in
`RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp`
and `Common\SearchServiceBroker.cpp`: inspect selftest filtering, the
`ForegroundSearchServiceProcess` request/status semantics, and the interaction
with the immediately prior cold-start case that sets
`REDSALAMANDER_SEARCH_SERVICE_STARTUP_WARMUP_DELAY_MS=5000`. If the broad
CompareDirectories rerun stays red, add targeted diagnostics to the failing
status wait before altering the timeout or predicate. Do not move this plan to
Done until broad CompareDirectories, proper Full, and the final Debug plus
test-enabled Release FolderView perf matrices are green and cited.

## 2026-07-03 21:59 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_215953_folderview_warpdrive_dxui_pointer_root_baton\
```

This supersedes the 21:49 archive as the immediate resume pointer, while still
depending on the 21:49 Full-run archive for the broad red-run evidence. It
captures the DxUi root-cause work performed after that archive: the failing
Menu test's popup HWND is already destroyed by timeout, and the menu trace
shows an incidental captured `WM_MOUSEMOVE` from the live OS cursor switching
the root from View to Plugins before the test posts its deliberate inside-popup
probe. The archive includes the trace, focused DxUi logs, current no-test-runs
staged/unstaged patches, focused DxUi/ViewPE/source snapshots, recent run-log
manifest, process snapshot, and a continuation note.

Current unverified code state: `Tests\DxUiTests\DxUiTests.Menu.cpp` contains a
test-side `pointerRootSwitchArmed` guard intended to ignore incidental setup
root-switch callbacks while preserving the deliberate inside-popup assertion.
Temporary diagnostics are still present (`DescribeContextMenuPopupDebugProbe`
and verbose failure text). Resume by building `DxUiTests`, rerunning
`DxUiTests.exe --suite=Menu`, then removing the temporary diagnostics if green
and rerunning focused plus broad DxUi before returning to the ViewerPE
fresh-child harness red. Do not move this plan to Done until DxUi, ViewerPE,
proper Full, and final Debug plus test-enabled Release FolderView perf matrices
are all green and cited.

## 2026-07-03 21:49 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_214957_folderview_warpdrive_full_red_dxui_viewerpe_baton\
```

This supersedes the 20:54 archive as the resume pointer. It contains current git
status/diff metadata, focused staged/unstaged patches, source snapshots for the
WIP plan, ViewerPE VLC test, DxUi menu test and menu implementation, inventory
Pester, and test coverage docs, plus copied run artifacts and the latest
`SelfTest\last_run` top-level logs/results.

Additional work completed after the 20:54 archive:

- Added ViewerPE VLC HUD child-window diagnostics in
  `Tests\ViewerPETests\ViewerPETests.cpp`, rebuilt `ViewerPETests`, and proved
  focused VLC, focused shell-combo, and direct broad ViewerPE were green before
  the next Full run.
- Corrected CompareDirectories test-inventory drift from 229 to 230 in
  `Tools\Tests\TestInventory.Tests.ps1` and
  `Specs\Testing\Testing_TestCoverage.md`; focused TestInventory Pester and
  full Tools Pester are green. Use the explicit `$result.FailedCount` wrapper as
  the trustworthy Tools Pester exit signal because the plain redirected command
  inherited a child negative-test `$LASTEXITCODE`.
- Ran the proper Full gate with `-TimeoutMultiplier 2`. It is still red with
  1098 passed / 2 failed / 52 skipped / 1152 total. CompareDirectories,
  Commands, FileOperations, and ToolsPester are green. Remaining red gates:
  `DxUiTests` and `ViewerPETests`.

Resume by root-causing these two Full-suite blockers before patching:

- `DxUiTests --suite=Menu` should start at
  `TestMenuBarHoverMessageSwitchesRootWhilePopupOverlapsMenuBar`; the Full log
  failed with `overlapping View popup exposes debug state`.
- `ViewerPETests` should start by rerunning direct broad and focused
  `TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly`; the Full parent
  wrapper failed only at the fresh-harness child exit-code assertion, while
  direct evidence before Full was green.

Do not move this plan to Done until both Full red gates are green, the final
Debug and test-enabled Release FolderView perf matrices are archived and cited,
durable spec migration is complete, and this plan is moved to
`Specs\Plans\Done\`.

## 2026-07-03 20:54 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_205410_folderview_warpdrive_compare_green_viewerpe_hud_baton\
```

This supersedes the 20:23 baton as the resume pointer. It contains the latest
Compare focused/broad logs, the failing focused ViewerPE VLC log, source
snapshots for the ViewerVLC/ViewerPE/Compare/Pester files to inspect next,
current git status/diff stats, focused diffs, and `CONTINUATION.md`.

Additional work completed after the 20:23 archive:

- Focused `decision_cache_eviction_budget_pending_wide_tree` reproduced the
  Full red with `pendingSubdirUpdates=9`, `contentPendingCompares=537`,
  `contentQueueSize=533`, and `contentInFlightSize=537`.
- The test-only root cause is the tiny-budget pending aggregation scenario
  pinning the left `keep` subtree while the right containing parent is `spill`;
  once `spill` is evicted, pending child aggregation can churn instead of
  settling. `RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.Cases.RuntimeAndRemote.cpp`
  now pins `keep` on the left and `spill` on the right for that case.
- Debug rebuild passed:
  `.build\codex-runs\msbuild_redsalamander_after_compare_pending_pin_spill_20260703.log`
  (0 warnings / 0 errors).
- Focused `decision_cache_eviction_budget_pending_wide_tree` is green:
  `.build\codex-runs\compare_decision_cache_pending_focused_after_pin_spill_20260703.log`
  (1 passed / 0 failed / 0 skipped).
- Focused `search_service_sqlite_ntfs_traversal_seed_stays_degraded` is green:
  `.build\codex-runs\compare_search_service_ntfs_seed_focused_after_full_red_20260703.log`
  (1 passed / 0 failed / 0 skipped).
- Broad CompareDirectories is green:
  `.build\codex-runs\compare_broad_after_pending_pin_spill_20260703.log`
  (208 passed / 0 failed / 30 skipped).
- Focused ViewerPE VLC is still red:
  `.build\codex-runs\viewerpe_vlc_tab_focus_focused_after_full_red_20260703.log`
  fails `ViewerVLC exposes a visible HUD child that can take focus` while the
  video child and overall visible-child count pass. `Plugins\ViewerVLC\ViewerVLC.h`
  confirms the product HUD class is exactly `RedSalamander.ViewerVLC.Hud`, so
  resume by proving enumeration depth, visibility/style, parent/owner, and
  timing rather than assuming a class-name typo.
- The second ViewerPE Full red,
  `TestViewerShellComboHostsLongRunOpenCloseStayStable`, has not yet been
  rerun focused after the Compare fix.
- ToolsPester inventory drift remains: CompareDirectories source now exposes
  230 `RunCase` registrations while `Tools\Tests\TestInventory.Tests.ps1` and
  `Specs\Testing\Testing_TestCoverage.md` still carry 229 in the failing path.

Resume with the ViewerPE VLC HUD diagnostic first, then run the ViewerPE shell
combo focused case, fix/rerun the Pester inventory drift, rerun broad
ViewerPETests and ToolsPester, then rerun Full. Do not move this plan to Done
until Full and the final Debug plus test-enabled Release FolderView perf
matrices are green and cited.

## 2026-07-03 20:23 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_202302_folderview_warpdrive_full_closeout_red\
```

This supersedes the 18:14 baton as the resume pointer. It contains compact
Full-suite evidence, selected green Commands/FileOps logs and archive summaries,
source snapshots, current git status/diff stats, and `CONTINUATION.md`. The
recursive `SelfTest\last_run\commands\work` fixture tree and raw dumps are
intentionally excluded from this archive; the useful result JSON/log/trace files
are preserved.

Additional work completed after the 18:14 archive:

- Root-caused the S3 plugin-configuration tab traversal red to a selftest wait
  predicate: `sendTab(...)` required `panelScrollPosY == 0` at every step even
  though the S3 form legitimately scrolls when focus reaches
  `Use virtual-hosted style addressing`. The rejected good sample caused an
  extra debug tab advance and the later `Max keys per request` sample.
- Patched only
  `RedSalamander\SelfTest\Commands\Commands.SelfTest.PluginConfig.cpp` to drop
  that per-step `panelScrollPosY == 0` predicate while preserving the exact
  focus kind/label and DxUi-vs-legacy host assertions.
- Debug build passed:
  `.build\codex-runs\msbuild_redsalamander_after_plugin_config_scroll_predicate_20260703.log`.
- Focused plugin-config tab traversal passed at
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_182138`.
- The first plugin-config local-window rerun at `2026-07-03_182158` ran zero
  cases because the filter file path was accidentally passed as the filter text;
  ignore it as a product/test result.
- Correct plugin-config local-window rerun passed at
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_182308`.
- Exact `0..370` prefix passed at
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_183006`.
- Broad Commands is green at
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_184648`
  (776 passed / 0 failed / 2 skipped).
- Standalone FileOps is green at
  `Specs\TestRuns\4cb089111a23\FileOps\2026-07-03_190019`
  (102 passed / 0 failed / 20 skipped).
- Full with build is still red:
  `.build\codex-runs\full_after_commands_fileops_green_20260703.log`
  / `last-run-top-level\run-all-tests-results.json` report 320 passed /
  4 failed / 50 skipped / 374 total.

Current Full blockers:

- CompareDirectories:
  `decision_cache_eviction_budget_pending_wide_tree` failed to drain pending
  subtree/content updates, and
  `search_service_sqlite_ntfs_traversal_seed_stays_degraded` did not finish
  startup warmup.
- ViewerPETests:
  `TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly` and
  `TestViewerShellComboHostsLongRunOpenCloseStayStable` failed inside their
  fresh-harness child runs.
- ToolsPesterTests:
  CompareDirectories inventory drift is inconsistent across source/docs:
  static RunCase expected 229 but found 230, while coverage spec expected 230
  but found 229.

Resume by running the two Compare failures focused, then the two ViewerPE
filters directly, then fixing the CompareDirectories inventory count drift in
`Tools\Tests\TestInventory.Tests.ps1`, `Tests\README.md`, and
`Specs\Testing\Testing_TestCoverage.md` as needed. After focused fixes, rerun
broad Compare, ViewerPETests, Tools Pester, and Full. Do not move this plan to
Done until Full and the final Debug plus test-enabled Release FolderView perf
matrices are green and cited.

## 2026-07-03 18:14 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_181403_folderview_warpdrive_plugin_config_baton\
```

This supersedes the 16:29 baton as the resume pointer. It contains the current
git refs/status, staged and unstaged patches excluding `Specs\TestRuns`, source
snapshots for the issues-pane focus patch and plugin-configuration/S3 focus-order
code, selected `.build\codex-runs` logs, copied Commands result archives, the
latest crash-folder inventory, a pointer to the 17:41 hang dump, and
`CONTINUATION.md`.

Additional work completed after the 16:29 archive:

- Hardened `cmd_pane_fileops_issues_pane_hide_restores_folder_focus` so the
  setup waits for selected issue state, logical DxUi grid focus, and native pane
  HWND focus before closing the issues pane.
- Debug build passed in
  `.build\codex-runs\msbuild_redsalamander_after_issues_focus_wait_20260703.log`.
- The focused issues-pane guard and local tail are green:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_164118` and
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_164146`.
- The Compare Options + issues-pane + speed-limit tail `517..552` is green:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_164607`
  (36 passed / 0 failed).
- A broad Commands run at
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_170344` failed with four
  order-sensitive cases; each passed focused/local-window reruns.
- A broad rerun then timed out/hung around the Connection Manager modeless
  connect path; the manual dump is
  `C:\Users\eric\AppData\Local\RedSalamander\Crashes\RedSalamander-hang-20260703-174154-p62052.dmp`.
  Focused, local-window, and prefix reruns for that area are green at
  `2026-07-03_174151`, `2026-07-03_174227`, and `2026-07-03_174418`.
- The second broad Commands run at
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_180131` again failed with
  four order-sensitive Preferences/Find cases; focused/local-window reruns are
  green at `2026-07-03_180208`, `2026-07-03_180249`, and `2026-07-03_180331`.
- The current best red point is the prefix `0..370` at
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_181043`: after 370 passed
  cases, `cmd_plugin_configuration_dialog_tab_traversal_live_dx_interaction`
  expected the S3 `Use virtual-hosted style addressing` toggle but sampled
  `Max keys per request` instead. The same tab-traversal case is green focused
  at `2026-07-03_181126` and green in the plugin-config local window at
  `2026-07-03_181205`, so this is still broad-order focus/snapshot behavior.

Resume by inspecting `RedSalamander\SelfTest\Commands\Commands.SelfTest.PluginConfig.cpp`
around the `sendTab`/`plugin-config advance` path and
`Plugins\FileSystemS3\FileSystemS3.h` for the S3 schema order. Prove whether the
test samples after an extra focus advance, the debug wrapper loses the
native/logical focused HWND before the wait predicate observes it, or the
product focus order changes after scroll. Do not weaken the tab-order assertion.
After the patch, rerun focused plugin-config tab traversal, the 11-case
plugin-config window, the exact `0..370` prefix, and then full broad Commands.
Do not move this plan to Done until broad Commands, proper Full, and the final
Debug plus test-enabled Release FolderView perf matrices are green and cited.

## 2026-07-03 16:29 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_162914_folderview_warpdrive_speedlimit_focus_baton\
```

This supersedes the 16:15 archive as the resume pointer. It contains current git
refs/status, branch ahead/behind state (`25` ahead / `4` behind
`origin/master`), no-TestRuns staged/unstaged patch snapshots, raw runner logs,
stable Commands archives, speed-limit tail filters, source snapshots, the latest
RedSalamander crash-handler sidecar pointer, manifest, and `CONTINUATION.md`.

Additional work completed after the 16:15 archive:

- Focused `cmd_pane_fileops_speedLimit_prompt_live_dx_interaction` passed at
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_162111`.
- The two-case speed-limit prompt window passed at
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_162136`.
- Tail `531..552` passed at
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_162221`.
- Wider tail `517..552` failed earlier, not at speed-limit:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_162631` reports
  35 passed / 1 failed at
  `cmd_pane_fileops_issues_pane_hide_restores_folder_focus`, reason
  `Issues-pane HWND should own Win32 focus before hide; focusHwnd=0x0`.
- No new RedSalamander crash-handler dump appeared after these probes. The
  latest paired crash sidecar remains
  `C:\Users\eric\AppData\Local\RedSalamander\Crashes\RedSalamander-20260703-161435-p81732.txt`,
  already dump-archived by the 16:15 baton.

Resume by treating the issues-pane/native-focus ordering as the active blocker:
inspect `FileOperationsIssuesPaneState::SelfTestFocusGrid`,
`FileOperationsIssuesPaneState::OnClose`, the show/hide wrapper, and
`TestFileOperationsIssuesPaneHideRestoresFolderFocus`; then run the focused
hide-focus case, `517..546`, and `531..546` one at a time. Patch only after
deciding whether the bug is selftest setup failing to establish native focus, or
product hide/close restoration missing a real logically-focused DxUi grid path.
Do not weaken the final assertion that hiding a focused issues pane restores
folder focus. Do not move this plan to Done until broad Commands, proper Full,
and final Debug plus test-enabled Release FolderView perf matrices are green and
cited from the current HEAD.

## 2026-07-03 16:15 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_161500_folderview_warpdrive_commands_broad_crash\
```

This supersedes the 15:55 archive as the resume pointer. It contains current git
refs/status, branch ahead/behind state (`25` ahead / `4` behind
`origin/master`), no-TestRuns staged/unstaged patches, source snapshots, copied
NavigationView recheck logs and result archives, copied broad Commands crash
log/last-run trace, paired RedSalamander crash dump/sidecar artifacts from
`C:\Users\eric\AppData\Local\RedSalamander\Crashes`, manifest, and
`CONTINUATION.md`.

Additional work completed after the 15:55 archive:

- The previous NavigationView full-path popup DPI red was rechecked and did not
  reproduce. The 2-case predecessor slice, 3-case predecessor slice, and exact
  17-case NavigationView prefix all passed (`Specs\TestRuns\4cb089111a23\Commands\2026-07-03_160054`,
  `2026-07-03_160114`, and `2026-07-03_160140`).
- The follow-up broad Commands run crashed before normal results were written:
  `.build\codex-runs\commands_broad_after_nav_prefix17_recheck_20260703.log`
  exited `-1073740771` (`0xC000041D`).
- The paired RedSalamander crash sidecar
  `RedSalamander-20260703-161435-p81732.txt` shows a TSF/native-input access
  violation while `FileOperationsSpeedLimitPromptWindow::ShowModal()` is inside
  `GetMessageW`; the active selftest case is
  `cmd_pane_fileops_speedLimit_prompt_live_dx_interaction`.

Resume with systematic debugging of the speed-limit prompt modal lifecycle:
inspect `FileOperationsSpeedLimitPromptWindow` OK/Cancel/close/`WM_NCDESTROY`
paths and the selftest debug helpers before patching. First reproduce focused,
then the 2-case speed-limit prompt window. If evidence confirms destructive
worker-thread prompt teardown, patch narrowly toward posted UI-thread
confirm/cancel/close commands, rebuild Debug, rerun focused/window/broad
Commands, then continue rebase/Full/perf closeout.

## 2026-07-03 15:55 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_155547_folderview_warpdrive_nav_focus_dpi_baton\
```

This supersedes the 15:05 archive as the resume pointer. It contains current git
refs/status, branch ahead/behind state (`25` ahead / `4` behind
`origin/master`), no-TestRuns staged/unstaged patches, source snapshots, copied
Commands/Compare logs and result archives, latest RedSalamander crash sidecars,
manifest, and `CONTINUATION.md`.

Additional work completed after the 15:05 archive:

- CompareDirectories shutdown crash root cause was fixed by explicitly waking
  and joining scan/content workers in `CompareDirectoriesSession` before
  content-compare state teardown. Verification passed: new worker-shutdown
  guard `Specs\TestRuns\4cb089111a23\CompareDirectories\2026-07-03_143301`,
  pending-wide-tree guard `2026-07-03_143335`, and full CompareDirectories
  `2026-07-03_144451` (208 passed / 0 failed / 30 skipped).
- The previous broad Commands status-bar failure passed focused and in its
  7-case navigation-shell window (`2026-07-03_150814`,
  `2026-07-03_150849`).
- The next broad Commands run exposed
  `cmd_pane_navigationView_menu_region_keyboard_activation_opens_menu`; focused
  and local menu windows passed, then the test was hardened with a passive
  condition-based focus-return wait. Build and 2/10-case verification passed
  (`2026-07-03_153142`, `2026-07-03_153211`).
- The next broad Commands run exposed
  `cmd_pane_navigationView_path_doubleClick_after_focusAddressBar_tab_traversal`
  with an already-valid edit surface but transient `focusTarget=None`; the test
  was hardened to wait on the real edit host/input/text/selection surface before
  forcing the tab handoff. Build and focused triplet passed
  (`2026-07-03_155244`).
- The current remaining red edge is
  `cmd_pane_navigationView_full_path_popup_edit_route` in the 17-case
  NavigationView prefix window (`2026-07-03_155328`): full-path popup remains
  visible but synthetic-DPI reflow waits time out with `dpi=96`, `client=0x0`,
  and `currentPath=''`. The same case passed focused immediately afterward
  (`2026-07-03_155401`).

Resume with systematic debugging of `TestPaneNavigationViewFullPathPopupEditRoute`
and `NavigationView::OnDpiChanged` / `UpdateFullPathPopupWindow`. First run
small predecessor slices around
`cmd_pane_navigationView_unfocused_pane_click_focuses_target_pane` before
patching. Do not weaken the DPI assertion until it is clear whether the broad
order is exposing product state, stale debug snapshot/HWND state, or a test wait
that assumes synthetic `WM_DPICHANGED` changes `GetDpiForWindow(...)`
synchronously. After broad Commands is green, preserve/stage the dirty worktree,
resync/rebase the branch against `origin/master`, and rerun closeout gates.

## 2026-07-01 22:47 Origin/Commands Recheck

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_223000_folderview_warpdrive_origin_commands_recheck\
```

This supersedes the 21:30 archive as the resume pointer. `git fetch origin`
had already completed successfully this turn; `origin/master...HEAD` remained
`0 25`, so `codex/folderview-warpdrive` is 25 commits ahead and 0 behind.
There is still no origin rebase to apply.

Fresh verification:

- Clean Debug rebuild passed:
  `.build\logs\msbuild-20260701_222239_273.log` (0 warnings / 0 errors).
- Focused repaint/edit guards passed after the selftest-harness hardening:
  `folderView_render_device_loss_recovers`,
  `folderView_dpi_change_repaints_both_panes`, and
  `cmd_pane_navigationView_edit_mode_survives_external_refresh`.
- Focused `cmd_pane_filter_prompt_live_dx_interaction` also passed.
- Broad Commands is still red before any Done move. The broad run exited before
  `commands\results.json` with process exit `-1073741819` (`0xC0000005`).
  WER identifies `textinputframework.dll`, exception `0xc0000005`, fault offset
  `0x000000000006a419`.
- The latest trace stops in `cmd_pane_filter_prompt_live_dx_interaction` after
  `pane_filter_prompt_live_dx: cycle 'live Cancel interaction' worker-done`,
  before the case returns.
- The reduced local prompt slice, Commands cases 602..627 ending at
  `cmd_pane_filter_prompt_live_dx_interaction`, passed `26/0/0`; this keeps the
  crash classified as suite-order/timing dependent rather than a deterministic
  focused filter-prompt failure.

Resume from the TSF/native text crash boundary. Keep the staged TSF/native text
work intact, inspect `Common\DxUi\DxUi.TextStoreACP.cpp` sink/advice,
`RequestLock`, disconnect/detach, and the modal prompt close paths in
`RedSalamander\FolderWindow.FileSystem.cpp`. Build the next reduced replay from
broader earlier slices that include FileOps/issues-pane and navigation-shell
churn, not only the local prompt 602..627 slice. If a reduced slice crashes,
rerun that exact slice once with `REDSALAMANDER_DXUI_TEXTINPUT_TRACE_FILE`, but
treat a traced pass as timing-perturbed evidence. Do not move this plan to Done
until broad Commands, FileOps, ViewerPE/ViewerVLC HUD, Tools Pester, Full
skip-build, final Debug plus test-enabled Release FolderView perf matrices, and
spec migration are all green/cited.

## 2026-07-01 21:30 Origin/Keyboard Recheck

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_213000_folderview_warpdrive_origin_keyboard_recheck\
```

This supersedes the 21:01 archive as the resume pointer. `git fetch origin`
completed successfully, and `HEAD...origin/master` is `25 0`, so
`codex/folderview-warpdrive` is 25 commits ahead of `origin/master` and 0
behind. There is no new origin work to rebase onto.

Fresh verification:

- Clean Debug rebuild passed:
  `.build\logs\msbuild-20260701_210811_496.log` (0 warnings / 0 errors).
- The prior Preferences Keyboard live-search blocker is no longer active:
  `.build\logs\commands-keyboard-live-clean-20260701_211042.out.log`
  passed `cmd_preferences_dialog_keyboard_live_search_dx_interaction`
  (1 passed / 0 failed / 0 skipped).
- Broad Commands fail-fast is still red before any Done move:
  `.build\logs\commands-broad-after-keyboard-clean-20260701_211102.out.log`
  exited before `commands\results.json` with process exit `-1073741819`
  (`0xC0000005`).
- WER identifies the crash as `textinputframework.dll`, exception
  `0xc0000005`, fault offset `0x000000000006a419`, report id
  `6c69fcac-19a3-4df9-b9eb-c79e71a806fd`.
- The trace stops while entering
  `cmd_pane_navigation_create_directory_prompt_keeps_navigation_shell_stable`.

Resume with the TSF/native text crash boundary around the create-directory
prompt. Do not continue chasing the Keyboard test unless it regresses again.
Temporary env-gated TSF trace instrumentation in
`Common\DxUi\DxUi.NativeTextInput.cpp` must still be removed or formalized
before final closeout. Do not move this plan to Done until broad Commands,
FileOps, ViewerPE/ViewerVLC HUD, Tools Pester, Full skip-build, final Debug
plus test-enabled Release FolderView perf matrices, and spec migration are all
green/cited.

## 2026-07-01 21:01 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_210100_folderview_warpdrive_origin_rebase_tests\
```

This supersedes the earlier 2026-07-01 pause entries as the resume pointer.
`codex/folderview-warpdrive` is rebased onto `origin/master`
(`origin/master` = `45ae2a9a90211aa71a961178661db3dfc36f9876`,
current HEAD = `275c040348a3994d45c6aef340333a35a0490cb6` before the
staged WIP is recommitted). The autostash backup remains available as
`stash@{0}: autostash`.

Rebase/build status:

- The post-rebase Debug build passed after conflict/include-order fixes:
  `.build\logs\msbuild-20260701_201933_303.log`.
- The focused ViewerSpace source guard now passes after compacting source text
  before exact-string checks:
  `.build\logs\commands-viewer-plugin-source-guard-after-compact-20260701_202243.out.log`.
- The broad Commands run then exposed a real Paste Shortcut long-path blocker:
  `.build\logs\commands-skipbuild-after-sourceguard-fix-20260701_202305.out.log`
  failed at `cmd_pane_clipboardPasteShortcut_creates_unique_links` after
  482 passed / 1 failed / 295 skipped.
- Paste Shortcut long-path creation/probing/verification was fixed in the
  staged worktree. Supporting Debug builds passed at
  `.build\logs\msbuild-20260701_203915_896.log`,
  `.build\logs\msbuild-20260701_204313_472.log`, and
  `.build\logs\msbuild-20260701_204633_170.log`.
- The focused Paste Shortcut case is now green:
  `.build\logs\commands-paste-shortcut-unique-focused-after-testverify-20260701_204858.out.log`
  (1 passed / 0 failed / 0 skipped).

Current active blocker:

- Broad Commands now fails in Preferences Keyboard live search:
  `.build\logs\commands-skipbuild-after-pasteshortcut-longpath-20260701_204917.out.log`
  (327 passed / 1 failed / 450 skipped), case
  `cmd_preferences_dialog_keyboard_live_search_dx_interaction`.
- Focused rerun also fails:
  `.build\logs\commands-pref-keyboard-live-search-focused-20260701_205952.out.log`,
  reason `Preferences Keyboard page visible DX search edit did not discard the
  pending search value after shell Cancel reopened the page.`

Resume with the focused Preferences Keyboard live-search failure. Compare it to
the analogous Plugins/Viewers/Themes pending-search discard cases, then inspect
`Common\DxUi\DxUi.NativeTextInput.cpp`; the temporary env-gated TSF trace
instrumentation is still present and must be removed or formalized before final
closeout. Do not move this plan to Done until broad Commands, FileOps,
ViewerPE/ViewerVLC HUD, Tools Pester, Full skip-build, and final Debug plus
test-enabled Release FolderView perf matrices are green and cited.

## 2026-07-01 17:44 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_174447_folderview_warpdrive_prevfocus_pause\
```

This supersedes the 17:23 archive as the resume pointer. It contains current git
refs/status, diff stat, full tracked binary patch, untracked list, WIP/baton
snapshots, source snapshots, copied DxUi/Commands logs, recent crash-dump/WER
manifests, manifest, checksums, and `CONTINUATION.md`.

Additional work completed after the 17:23 archive:

- Added a RED source-level DxUi guard requiring native TSF activation to keep
  the previous HWND focus document manager alive, and requiring deactivation to
  restore it instead of clearing the HWND association with `nullptr`.
- Implemented `_nativeTextInputTsfPreviousFocusDocumentMgr` in `DxUi`, captured
  the previous focus document manager from `ITfThreadMgr::AssociateFocus(...)`,
  restored it before `Pop(TF_POPF_ALL)`, then reset it after teardown.
- GREEN evidence: `Z:\src\RedSalamander\.build\logs\msbuild-20260701_172812_109.log`
  passed and `.\.build\x64\Debug\DxUiTests.exe` exited 0 with
  `All DxUi tests passed.`
- RedSalamander Debug rebuild passed:
  `Z:\src\RedSalamander\.build\logs\msbuild-20260701_173151_574.log`.
- The reduced crash recipe `[17..20,28,33,45..55,56..58]` passed twice after
  the previous-focus restore and archived as
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_173413` and
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_173442`.

Current blockers after the previous-focus restore:

- The exact 107-case filter no longer consistently crashes, but it is not green.
  First run failed `cmd_pane_createDirectory_prompt_live_dx_interaction` with
  `Create-directory prompt edit did not update after live interaction during
  live create-directory cancel`; the focused case passed immediately after.
- A tail run `[68..105]` crashed intermittently in
  `textinputframework.dll` (`0xc0000005`, fault offset
  `0x000000000006a419`) while entering
  `cmd_pane_filter_prompt_live_dx_interaction`; focused and narrowed bands
  passed, and the tail rerun passed.
- The exact 107-case rerun failed earlier, at
  `cmd_pane_navigation_refresh_keeps_navigation_shell_stable`, with a stable
  FolderView state (`itemCount=3`, `focusedItem='a.txt'`) but an empty
  NavigationView shell snapshot (`currentPath=''`, `historyCount=0`) after
  refresh.

Resume by diagnosing the navigation refresh shell failure before adding another
TSF change. Inspect the latest
`C:\Users\eric\AppData\Local\RedSalamander\SelfTest\last_run\trace.txt` around
`cmd_pane_navigation_refresh_keeps_navigation_shell_stable`, especially
`ResyncNavigationShellFromFolderView`, `SetFolderPath`, refresh, and
`NavigationView::DebugGetSnapshot` traces. Then run, one at a time:

```powershell
$cases = Get-Content '.build\logs\commands-fileops-navigation-dialogs-to-createdir-cases.txt'
$filter = @($cases[47]) -join ','
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter $filter -TimeoutMultiplier 4

$filter = @($cases[46..47]) -join ','
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter $filter -TimeoutMultiplier 4

$filter = @($cases[45..47]) -join ','
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter $filter -TimeoutMultiplier 4

$filter = @($cases[17..47]) -join ','
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter $filter -TimeoutMultiplier 4
```

If those are green, rerun the exact 107-case filter twice. Keep the temporary
TSF trace instrumentation only while diagnosing; remove it or convert it into a
proper opt-in diagnostic before final closeout. Do not move this plan to Done
until broad Commands, FileOps, ViewerPE, Tools Pester, Full skip-build, and
final Debug plus test-enabled Release FolderView perf matrices are green and
cited.

## 2026-07-01 19:38 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_193808_folderview_warpdrive_plugin_pointer_pause\
```

This supersedes the 19:13 archive as the resume pointer. It contains current git
refs/status/diff snapshots, copied WIP/baton specs, copied source snapshots for
the touched Panes/plugin-config selftests, copied Commands logs, copied relevant
Commands archives, latest `last_run` Commands evidence, process snapshot,
manifest/checksums, and `CONTINUATION.md`.

Additional work completed after the 19:13 archive:

- Exact no-trace `0..58` Commands rerun crossed the prior create-directory/TSF
  crash boundary and passed at
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_192004`.
- Broad Commands then advanced to the Preferences Panes theme-cycle area and
  failed at `cmd_preferences_dialog_panes_theme_cycle_keeps_surface_legible`
  (`Specs\TestRuns\4cb089111a23\Commands\2026-07-01_192923`, 309 passed /
  1 failed / 467 skipped). Focused Panes theme-cycle and the local Panes cluster
  passed, but the General+Panes cluster failed at
  `cmd_preferences_dialog_panes_page_uses_dxui_statics_and_toggles` with
  `Preferences navigation did not move to the Panes category`
  (`Specs\TestRuns\4cb089111a23\Commands\2026-07-01_193116`).
- Root cause for the Panes cluster failure was a remaining relative
  `VK_DOWN` navigation helper in
  `RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.ThemesGeneralPanes.cpp`;
  it now uses deterministic
  `DebugSelectPreferencesCategory(kPrefCategoryPanes)`, matching the other
  hardened Panes helpers.
- Rebuilt `RedSalamander` Debug cleanly at
  `Z:\src\RedSalamander\.build\logs\msbuild-20260701_193230_936.log`; the
  exact General+Panes red slice then passed at
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_193448`.
- Broad Commands after the Panes navigation fix now fails earlier in the plugin
  configuration/S3 dialog pointer path:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_193637` (176 passed /
  1 failed / 600 skipped). Current failure:
  `cmd_plugin_configuration_dialog_pointer_click_toggles_visible_dx_toggle`
  reports the visible DX `Use HTTPS` toggle did not restore after the second
  real pointer click (`expectedState=1 actualState=0`, focus still on
  `Default region`, `scrollY=0`, toggle rect `(116,5-167,35)`).

Resume with the plugin configuration pointer-click blocker. First run the
focused case
`cmd_plugin_configuration_dialog_pointer_click_toggles_visible_dx_toggle`; if it
passes, minimize the plugin-config cluster immediately preceding the broad
failure. Inspect
`RedSalamander\SelfTest\Commands\Commands.SelfTest.PluginConfig.cpp` around the
pointer-click helper before patching: prove whether the click target/coordinate
or settle wait is stale after prior plugin-config cases, or whether product
pointer/toggle routing regressed. If product behavior changes, add a failing
guard first. Temporary TSF trace instrumentation in
`Common\DxUi\DxUi.NativeTextInput.cpp` must still be removed or formalized
before final closeout. Do not move this plan to Done until broad Commands,
FileOps, ViewerPE, Tools Pester, Full skip-build, and final Debug plus
test-enabled Release FolderView perf matrices are green and cited.

## 2026-07-01 19:13 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_191333_folderview_warpdrive_tsf_lifetime_fix_pause\
```

This supersedes the 19:00 archive as the resume pointer. It contains current
git refs/status, diff stat, full tracked patch, WIP/source snapshots, the saved
`DxUiTests.exe --suite=NativeTextInput` green log, the exact no-trace `0..58`
Commands rerun after the TSF lifetime fix, and the focused common-folders pass.

Additional work completed after the 19:00 archive:

- Added a focused DxUi source-structure guard requiring
  `WindowHost::DeactivateNativeTextInputTsf()` to keep a local text-store COM
  reference while TSF-owned context/document-manager/thread-manager objects are
  released, then disconnect/detach/reset the text store afterward.
- Patched `Common/DxUi/DxUi.NativeTextInput.cpp` accordingly with
  `textStoreToDisconnect`; `DxUiTests` and `RedSalamander` Debug both rebuilt
  cleanly.
- `DxUiTests.exe --suite=NativeTextInput` is green and captured at
  `.build/logs/dxuitests-native-textinput-after-tsf-lifetime-fix-20260701_191333.out.log`.
- The exact no-trace `0..58` prefix did not crash after the lifetime fix, but
  fail-fast stopped before create-directory at
  `cmd_pane_navigation_nonstandard_menu_common_folders`
  (`Specs/TestRuns/4cb089111a23/Commands/2026-07-01_191106`, 36 passed / 1
  failed / 22 skipped). The focused common-folders case passes at
  `Specs/TestRuns/4cb089111a23/Commands/2026-07-01_191134`, so this is now an
  ordering/timing blocker before the TSF crash boundary can be considered fully
  proven.

Resume by minimizing the common-folders in-prefix failure, hardening the submenu
ready/wait path if it is selftest timing (or adding a focused RED guard first if
it is product state), then rerun the exact no-trace `0..58` boundary until it
reaches and passes the create-directory prompt. Temporary TSF trace
instrumentation in `Common/DxUi/DxUi.NativeTextInput.cpp` still must be removed
or formalized before final closeout.

## 2026-07-01 19:00 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_190022_folderview_warpdrive_tsf_boundary_pause\
```

This supersedes the 18:43 archive as the resume pointer. It contains current git
refs/status/diff snapshots, copied key `.build\logs` outputs and filters, copied
latest TSF-trace Commands run metadata, copied untracked WIP continuation specs,
manifest, checksums, and `CONTINUATION.md`.

Additional work completed after the 18:43 archive:

- Rebuilt `RedSalamander` Debug cleanly from current source after the reverted
  broad null-focus fallback:
  `.build\logs\msbuild-20260701_184856_655.log`.
- Focused issues-pane focus guard remained green:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_185112`.
- Previous reduced `8..16 + 17..28` red filter remained green:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_185204`.
- Exact `0..58` no-trace run now hard-crashes before results in
  `textinputframework.dll`, at
  `cmd_pane_navigation_create_directory_prompt_keeps_navigation_shell_stable`;
  targeted crash evidence is archived under
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_1853_prefix0_58_crash_after_rebuild`.
- Focused/suffix create-directory slices all pass, so the crash is
  order/load/timing dependent:
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_185656`,
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_185706`,
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_185716`, and
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_185727`.
- Exact `0..58` with
  `REDSALAMANDER_DXUI_TEXTINPUT_TRACE_FILE=.build\logs\dxui-tsf-prefix0-58-20260701_1904.trace.txt`
  passed `59/0/0` at
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_185819`; treat this as a
  timing-perturbed observation, not proof of correctness.

Resume from the TSF teardown boundary. Start by inspecting
`Common\DxUi\DxUi.TextStoreACP.cpp` sink/advice, `RequestLock`, `Disconnect()`,
`DetachHost()`, and ref-count release, plus the create-directory prompt close and
`WM_NCDESTROY` paths in `RedSalamander\FolderWindow.FileSystem.cpp`. Do not
start with a product patch: first produce a tighter failing guard or stronger
source-structure test for the TSF teardown hypothesis. Temporary TSF trace
instrumentation in `Common\DxUi\DxUi.NativeTextInput.cpp` must still be removed
or formalized before final closeout. Do not move this plan to Done until broad
Commands, FileOps, ViewerPE, Tools Pester, Full skip-build, and final Debug plus
test-enabled Release FolderView perf matrices are green and cited.

## 2026-07-01 18:43 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_184310_folderview_warpdrive_issuesfocus_pause\
```

This supersedes the 18:18 archive as the resume pointer. It contains current git
refs/status, diff stat, full tracked binary patch, untracked list, WIP/baton
snapshots, copied focused and reduced Commands logs/filters, copied relevant
Commands archives, process snapshot, and `CONTINUATION.md`.

Additional work completed after the 18:18 archive:

- Added focused RED guard
  `cmd_pane_fileops_issues_pane_hide_restores_folder_focus`, which failed before
  the product fix at
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_182937`.
- Implemented a narrow issues-pane hide focus handoff in
  `FileOperationState::ToggleIssuesPane()` and
  `FileOperationsIssuesPaneState::OnClose()`: when the pane or one of its
  children owns focus before hide, restore the active folder view through the
  existing `FolderWindow` focus helpers.
- The focused guard passed after the fix at
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_183342`, and the previous
  reduced `8..16 + 17..28` red filter passed twice.
- Broader navigation-shell runs still fail with correct selection/folder state
  but lost Win32 focus: `0..33` at
  `cmd_pane_selection_select_all_keeps_navigation_shell_stable` and `0..58` at
  `cmd_pane_selection_select_unselect_dialogs_keep_navigation_shell_stable`.
- A broader `GetFocus() == nullptr` fallback experiment proved too broad because
  it regressed the speed-limit prompt path; that experiment is reverted in
  source, but the Debug binary must be rebuilt before the next test run.

Resume by rebuilding `RedSalamander` Debug, rerunning the focused issues-pane
guard, then rerunning the old reduced filter. If those remain green, continue
root-cause on the selection-command/dialog focus loss; start in
`Commands.SelfTest.Navigation.cpp`, `RedSalamander.cpp` selection dispatch, and
the selection mask dialog implementation. Temporary TSF trace instrumentation in
`Common\DxUi\DxUi.NativeTextInput.cpp` must still be removed or formalized, and
static test inventory counts must be updated after the new guard is finalized.
Do not move this plan to Done until broad Commands, FileOps, ViewerPE, Tools
Pester, Full skip-build, and final Debug plus test-enabled Release FolderView
perf matrices are green and cited.

## 2026-07-01 18:18 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_181800_folderview_warpdrive_focusloss_pause\
```

This supersedes the 17:58 archive as the resume pointer. It contains current git
refs/status, diff stat, full tracked binary patch, untracked list, WIP/baton
snapshots, copied narrowed Commands logs/filters, process snapshot, and
`CONTINUATION.md`.

Additional work completed after the 17:58 archive:

- Create-directory/navigation-shell focused and suffix runs did not reproduce
  the hard TSF crash: case 58 alone, 57..58, 56..58, 55..58, 54..58, and the
  interesting prefix through 58 passed.
- `0..58` now fails without crashing at
  `cmd_pane_selection_select_unselect_dialogs_keep_navigation_shell_stable`:
  navigation state remains stable, but actual Win32 focus is lost
  (`focusedFolderView=0x0`, `focusHwnd=0x0`).
- The focused select/unselect case alone passes, and multiple prefix splits
  narrow the stable red point to FileOperations issues-pane churn followed by
  selection commands.
- The best repeatable reduced red is `8..16 + 17..28`, failing at
  `cmd_pane_selection_select_all_keeps_navigation_shell_stable` with correct
  selection/folder state but `focusedFolderViewMatch=no`. Dropping case 8 makes
  the slice pass; dropping case 16 does not. The cumulative culprit currently
  points at issues-pane focus retention across resized/reordered state,
  long-run open/close, and theme-cycle paths.
- No live `RedSalamander` or `DxUiTests` process was present at archive time.

Resume by inspecting `FileOperationState::ToggleIssuesPane()` and the
`FileOperationsIssuesPaneState` show/hide/destroy paths. Add a focused RED guard
for hiding/toggling the issues pane while its DxUi grid owns focus, then fix
product behavior so closing a focus-owning issues pane restores the active pane
folder view through the existing `FolderWindow` focus helpers. Do not patch the
selection tests to mask the lost-focus state. Temporary TSF trace
instrumentation in `Common\DxUi\DxUi.NativeTextInput.cpp` must still be removed
or formalized before final closeout.

## 2026-07-01 17:58 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_175828_folderview_warpdrive_tsf_crash_pause\
```

This supersedes the 17:44 archive as the resume pointer. It contains current git
refs/status, diff stat, full tracked binary patch, WIP/baton snapshots, focused
source snapshots, copied reduced-run logs, last-run traces, the WER report,
crash-dump manifest, copied relevant Commands archives, and `CONTINUATION.md`.

Additional work completed after the 17:44 archive:

- The archived refresh-shell failure at
  `cmd_pane_navigation_refresh_keeps_navigation_shell_stable` is not currently
  reproducible: case 47 alone, 46..47, 45..47, 17..47, and 0..47 all passed.
- The fresh 0..47 pass is archived as
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_175556`.
- The exact 107-case Commands rerun crashed before normal results with wrapper
  exit code `-1073741819` (`0xC0000005`). WER reports
  `textinputframework.dll`, exception `c0000005`, fault offset
  `000000000006a419`; the latest dump is
  `C:\Users\eric\AppData\Local\CrashDumps\RedSalamander.exe.101472.dmp`.
- The last Commands trace now stops while entering
  `cmd_pane_navigation_create_directory_prompt_keeps_navigation_shell_stable`,
  after compare-directories and item-properties navigation-shell cases passed.
- No live `RedSalamander` or `DxUiTests` process was present at archive time.

Resume with the narrowed create-directory navigation-shell crash boundary from
`CONTINUATION.md`: print `$cases[52..58]`, then run case 57, 56..57, 55..57,
54..57, `[17..20,28,33,45..57]`, and 0..57 one at a time. If a slice
reproduces the crash, rerun that exact slice once with
`REDSALAMANDER_DXUI_TEXTINPUT_TRACE_FILE` enabled. Root-cause prompt native
text activation/deactivation and HWND TSF association lifetime before adding
another TSF product change. Temporary TSF trace instrumentation must still be
removed or formalized before final closeout.

## 2026-07-01 17:23 Pause Archive Refresh

Fresh archive:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-01_172332_folderview_warpdrive_tsftrace_pause\
```

This supersedes the 17:05 archive as the resume pointer. It contains current git
refs/status, diff stat, full tracked binary patch, untracked list, WIP/baton
snapshots, source snapshots, copied DxUi/Commands logs, copied TSF trace,
recent crash-dump/WER manifests, manifest, checksums, and `CONTINUATION.md`.

Additional evidence after the 17:05 archive:

- RedSalamander Debug build passed after the AssociateFocus change:
  `Z:\src\RedSalamander\.build\logs\msbuild-20260701_171121_310.log`.
- The reduced crash recipe `[17..20,28,33,45..55,56..58]` passed once after
  AssociateFocus and archived as
  `Specs\TestRuns\4cb089111a23\Commands\2026-07-01_171351`.
- The immediate rerun of that same reduced recipe crashed with
  `textinputframework.dll` `0xc0000005` at fault offset
  `0x000000000006a419`:
  `Z:\src\RedSalamander\.build\logs\commands-speedlimit-select-selection-nav45-55-tail-after-associatefocus-rerun-20260701_171435.out.log`.
- Temporary env-gated TSF trace instrumentation was added to
  `Common\DxUi\DxUi.NativeTextInput.cpp`; the app rebuilt successfully in
  `Z:\src\RedSalamander\.build\logs\msbuild-20260701_171728_470.log`.
- The TSF-traced rerun crashed at the same module/exception/offset:
  `Z:\src\RedSalamander\.build\logs\commands-speedlimit-select-selection-nav45-55-tail-tsftrace-20260701_171945.out.log`.
- The copied TSF trace shows every directly traced TSF call returning `S_OK`,
  including `SetFocus(nullptr)`, `AssociateFocus(..., nullptr, ...)`,
  `Pop(TF_POPF_ALL)`, disconnect/detach/reset, and the final
  `deactivate.reset.after` line with all TSF pointers cleared.

Current root-cause lead: `ITfThreadMgr::AssociateFocus(...)` returns the
previous focus document manager through `ppdimPrev`, and the TSF docs say the
previous association should be restored by calling `AssociateFocus(...)` again
with that previous document manager. The current implementation associates the
HWND with the new document manager but deactivation clears the association with
`nullptr`; it does not store/restore the previous focus document manager. Resume
with a RED DxUi guard for capture/restore of the previous focus document
manager, then implement an owning WIL COM member such as
`_nativeTextInputTsfPreviousFocusDocumentMgr`, restore it before
`Pop(TF_POPF_ALL)`, rerun DxUiTests, rebuild RedSalamander Debug, and rerun the
reduced crash recipe twice before trying the exact 107-case filter.

Do not move this plan to Done until broad Commands, FileOps, ViewerPE, Tools
Pester, Full skip-build, and final Debug plus test-enabled Release FolderView
perf matrices are green and cited.
