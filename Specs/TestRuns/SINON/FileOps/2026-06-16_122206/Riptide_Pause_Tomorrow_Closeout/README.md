# Riptide pause checkpoint - finish tomorrow

Date: 2026-06-16 12:22 local
Branch: `claude/zealous-poitras-ce383a`
Worktree: `Z:\src\RedSalamander\.claude\worktrees\zealous-poitras-ce383a`

This is a pause checkpoint, not final closeout evidence. Do not move the WIP plan to `Specs/Plans/Done/` until the remaining Commands regression is fixed and a fresh broad Full gate is green.

## Current state

- The Riptide FileOps work is implemented and has focused / broad FileOps evidence archived in the phase folders listed by the WIP plan.
- The authoritative File Operations spec has already been updated with lasting Riptide contracts.
- The previous Full closeout stalled at Compare case `oauth_authmode_roundtrip`.
- That Compare stall is root-caused: sandboxed settings writes to `%LOCALAPPDATA%\RedSalamander\Settings` fail with `0x80070005`; the test used success-style assumptions after `LoadSettings` returned `S_FALSE`, then hung/asserted instead of failing cleanly.
- `CompareDirectoriesEngine.SelfTest.Cases.CoreDiffs.cpp` now fails closed for `SaveSettings` / `LoadSettings` non-`S_OK` results and checks optional data before dereferencing.
- After that guard, sandboxed `oauth_authmode_roundtrip` exits quickly with the expected failure instead of hanging; unsandboxed `oauth_authmode_roundtrip` passes; unsandboxed Compare suite on a short `C:\R` root passed `125/0/24`.
- The broad `Full -SkipBuild` rerun on a short `C:\R` root still exited 1. FileOps passed inside that run (`93/0/20`), but several non-Riptide suites had failures.
- Isolated reruns narrowed those broad failures down: Compare local-search latency, five of six Commands cases, DxUi menu, Monitor ETW, and Tools Pester all pass focused after rerun / plan fix.
- `Tools/Tests/RunAllTestsPlan.Tests.ps1` was fixed to include `PluginContractTests` in the expected Full suite list; focused Pester now passes `91/0/0`.
- The current deterministic blocker is `cmd_pane_navigationView_edit_mode_survives_external_refresh`.

## Deterministic blocker

Failing case:

`cmd_pane_navigationView_edit_mode_survives_external_refresh`

Observed failure:

- The case enters NavigationView path edit mode, then changes the pane path only by case.
- Expected: path edit remains alive, typed text is preserved, focus stays on the path edit target, and the edit host/input HWNDs are unchanged.
- Actual: edit mode exits, the edit host/input HWNDs disappear, and the snapshot reports `focusTarget=0`, `editMode=false`, `visibleChildren=0`, `currentEdit=''`, `currentPath=''`.
- The visibility snapshot still shows the current NavigationView window is visible and the pane path differs only by case.

Likely fix direction:

- Inspect the path-refresh path in `NavigationView` / `FolderWindow::SetFolderPath`.
- Preserve edit mode when the refreshed pane path is equal to the previous pane path under the repo's no-case path comparison.
- Keep the existing behavior for a real path change: real path changes should retire path edit mode.
- Rebuild and rerun the exact Commands case before broad gates.

## Local archive contents

This folder contains the local restart packet:

- `git-status-short.txt`
- `git-status-ignored-short.txt`
- `git-diff-stat.txt`
- `git-diff-name-only.txt`
- `tracked-working-tree-diff.patch`
- `untracked-files-not-ignored.txt`
- `test-runs-file-inventory.txt`
- `wip-plan-snapshot.md`

The final-closeout attempt evidence is in:

`Specs\TestRuns\SINON\FileOps\2026-06-16_112812\Riptide_Final_Closeout\`

Important files there:

- `full-suite-croot-skipbuild-output.txt`
- `full-croot-failed-artifacts\`
- `compare-suite-croot-output.txt`
- `compare-croot-artifacts\`
- `isolate-cmd_pane_navigationView_edit_mode_survives_external_refresh-output.txt`
- `tools-pester-after-plan-fix-output2.txt`

## First commands tomorrow

Use a short NTFS selftest root for app/selftests and the sanitized `cmd.exe` wrapper when MSBuild or Run-AllTests is involved.

Reproduce the blocker:

```powershell
$pathValue = [Environment]::GetEnvironmentVariable('PATH', 'Process')
$selfRoot = 'C:\R\Riptide_Isolate_cmd_pane_navigationView_edit_mode_survives_external_refresh'
$cmd = 'set "Path=" & set "PATH=' + $pathValue + '" & set "REDSALAMANDER_SELFTEST_ROOT=' + $selfRoot + '" & "C:\Program Files\PowerShell\7\pwsh.exe" -NoProfile -ExecutionPolicy Bypass -File Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_pane_navigationView_edit_mode_survives_external_refresh -TimeoutMultiplier 2.0'
& cmd.exe /d /s /c $cmd
```

Build after the NavigationView fix:

```powershell
$pathValue = [Environment]::GetEnvironmentVariable('PATH', 'Process')
$cmd = 'set "Path=" & set "PATH=' + $pathValue + '" & "C:\Program Files\PowerShell\7\pwsh.exe" -NoProfile -ExecutionPolicy Bypass -File build.ps1 -ProjectName RedSalamander -Configuration Debug'
& cmd.exe /d /s /c $cmd
```

Then run:

```powershell
$pathValue = [Environment]::GetEnvironmentVariable('PATH', 'Process')
$selfRoot = 'C:\R\Riptide_Final_Closeout_20260616_Full2'
$cmd = 'set "Path=" & set "PATH=' + $pathValue + '" & set "REDSALAMANDER_SELFTEST_ROOT=' + $selfRoot + '" & "C:\Program Files\PowerShell\7\pwsh.exe" -NoProfile -ExecutionPolicy Bypass -File Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2.0'
& cmd.exe /d /s /c $cmd
```

## Done criteria tomorrow

1. Fix `cmd_pane_navigationView_edit_mode_survives_external_refresh`.
2. Build `RedSalamander` cleanly.
3. Pass the exact Commands blocker.
4. Pass the Commands suite or justify any unrelated isolated flake with rerun evidence.
5. Pass a fresh broad Full gate on a short NTFS selftest root.
6. Update the final-closeout archive README and the WIP plan with exact evidence.
7. Move the WIP plan to `Specs\Plans\Done\`.
8. Only then mark the Riptide goal complete.

## `.selftest-root-baseline`

`.selftest-root-baseline` is not present in this worktree at pause time. If it appears tomorrow, treat it as generated local selftest/root-baseline state unless a nearby README says otherwise; it is not required for Riptide closeout evidence.
