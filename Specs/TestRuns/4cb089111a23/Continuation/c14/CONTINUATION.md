# FolderView Warp Drive Continuation Baton

Created: 2026-07-03 16:29 Europe/Paris

This supersedes the 16:15 baton for the next resume point:

```text
Specs\TestRuns\4cb089111a23\Continuation\2026-07-03_161500_folderview_warpdrive_commands_broad_crash\
```

## Branch State

- Repo: `Z:\src\RedSalamander`
- Branch: `codex/folderview-warpdrive`
- HEAD: `275c040348a3994d45c6aef340333a35a0490cb6`
- Relation to `origin/master`: `4` behind / `25` ahead at archive time.
- No live `RedSalamander` or `DxUiTests` process was observed at archive time.

## Why This Archive Exists

The 16:15 baton pointed at a broad Commands crash in the speed-limit prompt
live-Dx interaction area. The follow-up probes did not reproduce that crash in
focused or narrowed speed-limit windows. The useful new evidence is that the
wider predecessor slice now fails earlier, at an issues-pane focus precondition,
while the speed-limit cases still pass afterward.

Treat the current blocker as the issues-pane/native-focus ordering problem until
the speed-limit crash repeats with fresh evidence.

## New Evidence Since 16:15

| Run | Result | Meaning |
| --- | --- | --- |
| `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_162111` | 1 passed / 0 failed | Focused `cmd_pane_fileops_speedLimit_prompt_live_dx_interaction` did not reproduce the crash. |
| `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_162136` | 2 passed / 0 failed | Two-case speed-limit prompt window passed. |
| `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_162221` | 22 passed / 0 failed | Tail `531..552` passed, including issues-pane, conflict prompt, and both speed-limit prompt cases. |
| `Specs\TestRuns\4cb089111a23\Commands\2026-07-03_162631` | 35 passed / 1 failed | Wider tail `517..552` failed earlier at `cmd_pane_fileops_issues_pane_hide_restores_folder_focus`; both speed-limit prompt cases later passed. |

Failure text from `2026-07-03_162631`:

```text
Issues-pane HWND should own Win32 focus before hide; focusHwnd=0x0, pane=0x8920D2.
```

No new RedSalamander crash-handler dump appeared after the focused/tail probes.
The latest paired crash artifact remains:

```text
C:\Users\eric\AppData\Local\RedSalamander\Crashes\RedSalamander-20260703-161435-p81732.txt
```

The full dump for that crash was already copied by the 16:15 baton; this archive
copies the sidecar stack again and records the crash-folder listing.

## Archived Files

- Raw runner logs and filters:
  - `run-artifacts\commands_speedlimit_live_focused_after_broad_crash_20260703.log`
  - `run-artifacts\commands_speedlimit_two_case_after_broad_crash_20260703.log`
  - `run-artifacts\commands_speedlimit_tail_531_552_after_broad_crash_20260703.log`
  - `run-artifacts\commands_speedlimit_tail_531_552_after_broad_crash_20260703.filter.txt`
  - `run-artifacts\commands_speedlimit_tail_517_552_after_broad_crash_20260703.log`
  - `run-artifacts\commands_speedlimit_tail_517_552_after_broad_crash_20260703.filter.txt`
- Stable selftest archives copied under `commands-archives\`.
- Current git state, staged/unstaged diff stats, and no-`Specs/TestRuns` patch snapshots.
- Source snapshots for the speed-limit prompt, issues pane, file-ops selftests,
  view-command selftests, WIP plan, and crash-forensics skill.
- Crash sidecar text plus crash-folder listing.

## Resume Plan

1. Inspect the issues-pane focus path before patching:

```powershell
rg -n "SelfTestFocusGrid|OnClose|OnNcDestroy|ToggleIssuesPane|IssuesPane|GetFocus\\(" RedSalamander\FolderWindow.FileOperations*.cpp RedSalamander\SelfTest\Commands\Commands.SelfTest.FileOps.cpp
```

Read these first:

- `RedSalamander\FolderWindow.FileOperations.IssuesPane.cpp`
  - `FileOperationsIssuesPaneState::SelfTestFocusGrid`
  - `FileOperationsIssuesPaneState::OnClose`
  - `FileOperationsIssuesPaneState::OnNcDestroy`
- `RedSalamander\SelfTest\Commands\Commands.SelfTest.FileOps.cpp`
  - `TestFileOperationsIssuesPaneHideRestoresFolderFocus`
- The `FileOperationState` show/hide/toggle wrapper for the issues pane.

2. Reproduce in smallest useful order:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_pane_fileops_issues_pane_hide_restores_folder_focus" -TimeoutMultiplier 2 *> .build\codex-runs\commands_issues_hide_focus_focused_after_tail_red_20260703.log; exit $LASTEXITCODE
```

Then run predecessor slices one at a time, starting with `517..546` and
`531..546`. Do not run multiple `Run-AllTests.ps1` processes in parallel because
they share `SelfTest\last_run`.

3. Patch only after the focus contract is clear.

Likely decision point:

- If `SelfTestFocusGrid()` sets DxUi logical grid focus but does not ensure the
  native issues-pane HWND owns focus, harden the setup/wait so the test validates
  a real focused-pane hide scenario.
- If product hide/close restoration only checks native `GetFocus()` and misses a
  logically focused DxUi grid with null native focus, decide whether that is a
  real user-visible product path or just a selftest setup race.

Do not weaken the final behavior assertion that hiding a focused issues pane
restores folder focus.

4. After any patch, rerun:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_pane_fileops_issues_pane_hide_restores_folder_focus" -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "<517..552 filter from this archive>" -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter "cmd_pane_fileops_speedLimit_prompt_uses_dxui_surface,cmd_pane_fileops_speedLimit_prompt_live_dx_interaction" -TimeoutMultiplier 2
```

Then rerun broad Commands. If broad Commands is green, continue the normal
closeout gates: proper Full, final Debug plus test-enabled Release FolderView
perf matrices, durable spec migration, then move the plan to Done.

## Do Not Forget

- The branch is behind `origin/master`; recheck/rebase before final gates.
- The 16:15 broad crash stack is still relevant but currently not reproduced by
  focused speed-limit probes.
- Current plan remains WIP. Do not move it to Done until Full and final perf
  evidence are green and cited from the current HEAD.
