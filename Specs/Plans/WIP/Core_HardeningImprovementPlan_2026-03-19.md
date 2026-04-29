# Core Hardening and Improvement Plan

Last updated: 2026-03-19

Status: WIP

References:

- `Common/PlugInterfaces/NavigationMenu.h`
- `Common/PlugInterfaces/Viewer.h`
- `Specs/Plugins/Plugins_VirtualFileSystem.md`
- `Specs/Plugins/Plugins_ViewerPlugins.md`
- `Specs/Plans/Done/Core_CompareDirectoriesReview.md`
- `Specs/Plans/Done/FileSystem_RemediationPlan_2026-02-26.md`

## Closeout audit (`2026-04-25`)

Targeted code evidence shows parts of this slice have landed: `ViewerSqliteTests` now verifies that `SetCallback(nullptr, nullptr)` suppresses later `ViewerClosed` delivery, and the compare/Google Drive selftests cover registration-style navigation callback drain behavior. This umbrella remains WIP until the wider hardening scope is either finished or split into narrower Done/WIP plans.

Remaining closeout checklist:

- [ ] Confirm the registration-style callback drain contract across all in-repo `IViewer` and `INavigationMenu` providers, not only the currently covered SQLite viewer and Google Drive navigation paths.
- [ ] Finish or explicitly split the compare shutdown/threading hardening track, including conditional remote/S3 shutdown coverage where credentials are available.
- [ ] Consolidate duplicated HRESULT/system-message formatting helpers across the listed UI/status surfaces, or move that cleanup to a dedicated plan.
- [ ] Run/archive the listed validation commands after the remaining tracks are resolved.

## Summary

This umbrella plan tracks one hardening slice with four ordered tracks:

1. registration-style callback drain hardening
2. shutdown/threading hardening
3. error-handling plus localized-formatting discipline
4. compiler-warning plus regression-guard cleanup

The scope is intentionally additive and behavior-preserving. This slice does not introduce token-based unregister APIs.

## Track 1. Registration-Style Callback Drain Hardening

- Treat `SetCallback(callback, cookie)` as registration for `INavigationMenu` and `IViewer`.
- Treat `SetCallback(nullptr, nullptr)` as the official synchronous drain point.
- Freeze the contract as:
  - clearing the callback is idempotent
  - after clear returns, the previous callback must not be invoked again
  - queued/background work must either finish before clear returns or self-drop as stale
  - plugins must not hold locks across the drain wait if callback delivery can take those locks
- Keep `FileSystem.h` per-call callbacks on their existing call-return drain semantics.

## Track 2. Shutdown and Threading Hardening

- Keep compare window/app shutdown non-blocking while cleanup is released off the UI thread.
- Remove detached-thread fallback from compare cleanup scheduling.
- Keep the quiet-point ordering explicit: stop producers, request cancel, stop posting payloads, clear registration-style callbacks, then release instances and unload modules.
- Verify close behavior with focused compare window self-tests and existing command self-tests.

## Track 3. Error-Handling and Formatting Discipline

- Consolidate duplicated HRESULT/system-message formatting helpers used by:
  - connection manager
  - search status hinting
  - file-operation status/issue surfaces
  - folder-view diagnostic formatting
- Keep diagnostics on `Debug::Error` / `Debug::Warning`.
- Keep user-visible error/status detail based on localized system messages or explicit resources; do not reintroduce runtime printf-style formatting.

## Track 4. Compiler Warnings and Regression Guards

- Keep touched targets warning-clean after the helper consolidation and compare cleanup changes.
- Trim dead includes and helper duplication exposed by the sweep.
- Add at least one regression that proves clearing a viewer callback suppresses later close notification delivery.

## Validation

- `.\build.ps1 -ProjectName RedSalamander`
- `.\build.ps1 -ProjectName ViewerSqliteTests`
- `RedSalamander.exe --commands-selftest --selftest-case=app_windows_open_and_close_modeless`
- `ViewerSqliteTests.exe`

## Notes

- Navigation-menu callback delivery is currently synchronous in-tree; this slice documents and freezes the drain contract without adding new async navigation infrastructure.
- Remote compare shutdown coverage remains conditional on local connection profiles and secrets.
