# FolderView WarpDrive PerfJsonl Crash Baton

Date: 2026-07-02 13:57 +02:00

## Summary

Two latest RedSalamander crash-handler sidecar callstack files show the same
root crash, not separate Archive Pack / column-audit failures:

- `C:\Users\eric\AppData\Local\RedSalamander\Crashes\RedSalamander-20260702-100018-p73832.txt`
- `C:\Users\eric\AppData\Local\RedSalamander\Crashes\RedSalamander-20260702-125817-p33336.txt`

Both stacks crash in:

```text
WideCharToMultiByte
Debug::detail::ConvertUtf8
Debug::detail::WritePerfJsonl
Debug::Perf::Scope::~Scope
FolderView::ProcessThumbnailLoadQueue
```

Root cause: `Debug::Perf::Scope` stored `_name` and `_detail` as
`std::wstring_view`. `FolderView::ProcessThumbnailLoadQueue()` passed a
temporary-backed detail:

```cpp
perf.SetDetail(request.fullPath.parent_path().native());
```

The path temporary was gone by the time `Scope::~Scope()` wrote JSONL, so
`WideCharToMultiByte(...)` read a dangling pointer.

## Code Changes In This Slice

- `Common/Helpers.h`
  - `Debug::Perf::Scope` now owns `_name` and `_detail` as `std::wstring`.
  - The strings are copied only when ETW or JSONL output is active, preserving
    the cheap no-sink path.
- `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp`
  - Archive/option prompt Confirm/Cancel debug helpers now post modal close
    commands instead of synchronously tearing down DxUi HWNDs from selftest
    worker stacks.
- `.agents/skills/red-salamander-crash-forensics/`
  - New reusable repo-local skill capturing RedSalamander crash triage habits.
  - The skill now checks `C:\Users\eric\AppData\Local\RedSalamander\Crashes`
    first for paired `.dmp`/`.txt` crash artifacts.

## Verification

- Build: `.\build.ps1 -ProjectName RedSalamander -Configuration Debug`
  - Log: `run-artifacts/msbuild-20260702_134221_182.log`
  - Result: passed, 0 warnings / 0 errors.
- Focused prior crash case:
  - Command: `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter folderView_column_widths_audit`
  - Log: `run-artifacts/folder_column_widths_audit_after_perfscope_20260702.log`
  - Result: passed, 1/0/0.
- FolderView prefix:
  - Command: `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter folderView_`
  - Log: `run-artifacts/folderview_prefix_after_perfscope_20260702.log`
  - Result: passed, 33/0/0.
- Broad Commands:
  - Command: `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2`
  - Log: `run-artifacts/commands_broad_after_perfscope_20260702.log`
  - Result: failed normally, 382 passed / 1 failed / 395 skipped.
  - Failing case: `cmd_preferences_dialog_category_tree_handles_reverse_keyboard_navigation`.
  - No new RedSalamander crash sidecar/dump was created; the latest crash files remain the pre-fix `p33336` pair.
- Focused broad-failure rerun:
  - Command: `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -TimeoutMultiplier 2 -CaseFilter cmd_preferences_dialog_category_tree_handles_reverse_keyboard_navigation`
  - Log: `run-artifacts/preferences_reverse_keyboard_focused_after_perfscope_20260702.log`
  - Result: passed, 1/0/0.

## Remaining To Move Done

- Broad Commands is still not green because the Preferences reverse-navigation
  case failed once in broad order, even though it passed focused immediately
  afterward. Resume with a broad rerun; if it repeats, diagnose the Preferences
  category-host focus loss with sidecar artifacts checked first.
- Full suite still must pass.
- Final Debug and test-enabled Release FolderView perf matrices still must be
  rerun from current HEAD and cited.
- Durable closeout still needs final spec migration and moving the WIP plan to
  `Specs/Plans/Done/`.

## Archive Contents

- `crash-sidecars/`
  - `RedSalamander-20260702-100018-p73832.txt`
  - `RedSalamander-20260702-125817-p33336.txt`
  - `red-salamander-crashes-manifest.txt`
- `run-artifacts/`
  - build and selftest logs listed above.
- `patches/unstaged-focused-diff.patch`
  - focused unstaged diff for this slice.
- `git-*.txt`
  - branch/head/status/diff-stat snapshots.
