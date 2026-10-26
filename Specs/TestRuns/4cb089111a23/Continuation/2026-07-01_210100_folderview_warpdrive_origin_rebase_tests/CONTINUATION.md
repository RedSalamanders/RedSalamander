# FolderView Warp Drive Continuation - Origin Rebase And Tests

Date: 2026-07-01 21:01 Europe/Paris

## Current State

- Worktree: `Z:\src\RedSalamander`
- Branch: `codex/folderview-warpdrive`
- HEAD: `275c040348a3994d45c6aef340333a35a0490cb6`
- Rebases onto origin/master: complete
- origin/master: `45ae2a9a90211aa71a961178661db3dfc36f9876`
- Backup stash: `stash@{0}: autostash` still exists
- Index: staged WIP remains large; the latest Paste Shortcut long-path fixes are staged too.
- Untracked WIP plans remain:
  - `Specs/Plans/WIP/DxUi_Uia_ContinuationBaton_2026-06-29.md`
  - `Specs/Plans/WIP/FolderView_WarpDrive_ContinuationBaton_2026-06-29.md`
  - `Specs/Plans/WIP/Operation_Tailwind_FolderViewWarpDrivePreMergeReviewRemediation_2026-06-30.md`

## What Was Done In This Checkpoint

- Fetched origin and rebased `codex/folderview-warpdrive` onto `origin/master`.
- Resolved commit-replay conflicts and autostash conflicts without dropping the autostash backup.
- Fixed post-rebase build breakages:
  - `ConnectionManagerWindow.cpp`: restored the correct `editedModelIndex` local.
  - `Commands.SelfTest.cpp`: kept Settings included before command-family splits.
  - `Commands.SelfTest.Preferences.cpp`: moved Preferences dispatch after the split Preferences case files.
- Fixed brittle ViewerSpace source guard after upstream alignment formatting changed:
  - `Commands.SelfTest.PluginConfig.cpp` now compacts source text before checking the load-bearing budget/cache strings.
- Found and fixed a real Paste Shortcut long-path regression:
  - Test fixture now creates long `.lnk` collision slots through the existing extended shell path helper.
  - Product shortcut slot probing now uses extended Win32 path checks instead of raw `std::filesystem::exists` for long candidates.
  - Test shortcut verification/load now uses extended shell paths.

## Verification

- Debug build after initial rebase fixes: `msbuild-20260701_201933_303.log` passed.
- Focused ViewerSpace source guard: `commands-viewer-plugin-source-guard-after-compact-20260701_202243.out.log` passed 1/0/0.
- Broad Commands after source guard fix: `commands-skipbuild-after-sourceguard-fix-20260701_202305.out.log` failed at `cmd_pane_clipboardPasteShortcut_creates_unique_links` after 482 passed.
- Debug build after Paste Shortcut fixture patch: `msbuild-20260701_203915_896.log` passed.
- Debug build after Paste Shortcut product long-path probe fix: `msbuild-20260701_204313_472.log` passed.
- Debug build after Paste Shortcut verifier update: `msbuild-20260701_204633_170.log` passed.
- Focused Paste Shortcut long-path case: `commands-paste-shortcut-unique-focused-after-testverify-20260701_204858.out.log` passed 1/0/0.
- Broad Commands after Paste Shortcut fix: `commands-skipbuild-after-pasteshortcut-longpath-20260701_204917.out.log` failed at `cmd_preferences_dialog_keyboard_live_search_dx_interaction` after 327 passed.
- Focused Preferences keyboard live-search case: `commands-pref-keyboard-live-search-focused-20260701_205952.out.log` failed with:
  - `Preferences Keyboard page visible DX search edit did not discard the pending search value after shell Cancel reopened the page.`

## Active Blocker

Do not move the FolderView Warp Drive plan to Done yet.

The current blocker is not the earlier S3/Plugin modal hang and not the Paste Shortcut path. The active red is:

```text
cmd_preferences_dialog_keyboard_live_search_dx_interaction
```

Focused failure:

```text
Preferences Keyboard page visible DX search edit did not discard the pending search value after shell Cancel reopened the page.
```

The broad Commands run can also fail around the same area with:

```text
Failed to focus the Preferences category host while navigating to the Keyboard page.
```

## Resume Plan

1. Inspect `TestPreferencesDialogKeyboardLiveSearchDxInteraction` in `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp`.
2. Compare the failing Keyboard case with the analogous Plugins/Viewers/Themes pending-search discard cases that already pass.
3. Review `Common/DxUi/DxUi.NativeTextInput.cpp`; temporary env-gated TSF trace instrumentation is still present and must be removed or formalized before final closeout.
4. Fix the focused Keyboard live-search failure first.
5. Rerun:
   - focused `cmd_preferences_dialog_keyboard_live_search_dx_interaction`
   - immediate Preferences predecessor cluster
   - broad Commands fail-fast
6. Then continue the Done gates: FileOps, ViewerPETests/ViewerVLC HUD, Tools Pester, Full skip-build, final Debug plus test-enabled Release FolderView perf matrices, durable spec migration, and finally move the WIP plan to Done.
