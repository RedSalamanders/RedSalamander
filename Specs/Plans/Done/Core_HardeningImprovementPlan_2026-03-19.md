# Core Hardening and Improvement Plan

Last updated: 2026-06-19

Status: Done - split closeout

## Purpose

This umbrella hardening plan covered registration-style callback drain semantics, plugin shutdown quiet points, HRESULT/status formatting cleanup, and warning hygiene. As of 2026-06-19, the callback and quiet-point contracts are authoritative in public headers and subsystem specs, and the remaining HRESULT/status formatting cleanup has been split into a dedicated WIP plan.

This file is retained as the historical closeout record. Do not use it for new Compare Directories, viewer, navigation, file-operation, or status-formatting work.

## Closeout Decision

Closeout path: **Path B - Split And Close The Umbrella**

Completed:

- [x] Callback-drain contracts are authoritative in `Common/PlugInterfaces/Viewer.h`, `Common/PlugInterfaces/NavigationMenu.h`, `Specs/Plugins/Plugins_ViewerPlugins.md`, `Specs/Plugins/Plugins_VirtualFileSystem.md`, and `Specs/UI/UI_NavigationView.md`.
- [x] Shared plugin quiet-point wording is authoritative in `Specs/Plugins/Plugins_PluginAPI.md`.
- [x] Current viewer and navigation providers were re-audited on 2026-06-19.
- [x] Focused callback regression gates passed and were archived.
- [x] Compare shutdown/threading ownership remains with `Specs/Plans/WIP/Operation_Crosscut_CompareDirectoriesRemediation_SyncDataSafetyAndOptionsSimplification_2026-06-15.md`.
- [x] Residual HRESULT/status-format cleanup was split to `Specs/Plans/WIP/Core_HResultStatusFormattingCleanup_2026-06-19.md`.

Not completed here:

- Full HRESULT/status surface cleanup. It is intentionally owned by the successor WIP plan.
- A full-suite closeout run. Because the residual formatting work was split and no production code was changed for this closeout, the narrower callback-provider gates below are the archived closeout evidence for this umbrella.

## Source Of Truth

| Contract area | Authoritative owner | Closeout result |
|---------------|---------------------|-----------------|
| `IViewer::SetCallback(nullptr, nullptr)` synchronous drain | `Common/PlugInterfaces/Viewer.h`, `Specs/Plugins/Plugins_ViewerPlugins.md`, `Common/EmbeddedViewerBase.h`, `Common/Helpers.h` | Provider audit complete; focused viewer regression archived. |
| `INavigationMenu::SetCallback(nullptr, nullptr)` synchronous drain | `Common/PlugInterfaces/NavigationMenu.h`, `Specs/Plugins/Plugins_VirtualFileSystem.md`, `Specs/UI/UI_NavigationView.md` | Provider audit complete; active command-delivery provider regression archived. |
| Plugin quiet point / module unload ordering | `Specs/Plugins/Plugins_PluginAPI.md`, plugin-specific specs | Cross-check complete; wording remains consistent with callback-drain specs. |
| Compare Directories lifecycle, shutdown, queueing, and perf | `Specs/Core/Core_CompareDirectories.md` plus the active Crosscut plan | Not duplicated here. |
| HRESULT/system-message formatting | `Common/Helpers.h`, `Specs/Core/Core_Localization.md`, `.github/skills/compiler-warnings/SKILL.md` | Split to `Specs/Plans/WIP/Core_HResultStatusFormattingCleanup_2026-06-19.md`. |
| Warning policy | `.github/skills/compiler-warnings/SKILL.md`, project files | Build gates passed; no broad suppressions added. |

## Provider Audit - 2026-06-19

### Viewer Providers

Commands:

```powershell
rg -n "class .*EmbeddedViewerBase|NotifyViewerClosed\(" Plugins Common\EmbeddedViewerBase.h -g "*.h" -g "*.cpp"
rg -n "SetCallback\(IViewerCallback|RegistrationCallbackState<IViewerCallback>" Plugins Common -g "*.h" -g "*.cpp"
```

Findings:

- `Common/EmbeddedViewerBase.h` is the only `IViewerCallback` `SetCallback(...)` implementation in current in-tree viewer code.
- `EmbeddedViewerBase` routes callback registration and `NotifyViewerClosed()` through `RegistrationCallbackState<IViewerCallback>`.
- Current in-tree viewer providers inherit `EmbeddedViewerBase` and do not override `IViewer::SetCallback(...)` locally:
  - `ViewerText`
  - `ViewerSqlite`
  - `ViewerSpace`
  - `ViewerImgRaw`
  - `ViewerVLC`
  - `ViewerPE`
  - `ViewerWeb`
- `NotifyViewerClosed()` call sites remain in provider close paths and therefore use the shared generation/in-flight/stale-drop logic.

| Provider | Closeout state | Evidence |
|----------|----------------|----------|
| ViewerText | [x] | Inherits `EmbeddedViewerBase<ViewerText>`; no local `SetCallback` override. |
| ViewerSqlite | [x] | Inherits `EmbeddedViewerBase<ViewerSqlite>`; callback-clear regression archived. |
| ViewerSpace | [x] | Inherits `EmbeddedViewerBase<ViewerSpace>`; no local `SetCallback` override. |
| ViewerImgRaw | [x] | Inherits `EmbeddedViewerBase<ViewerImgRaw>`; no local `SetCallback` override. |
| ViewerVLC | [x] | Inherits `EmbeddedViewerBase<ViewerVLC>`; no local `SetCallback` override. |
| ViewerPE | [x] | Inherits `EmbeddedViewerBase<ViewerPE>`; no local `SetCallback` override. |
| ViewerWeb | [x] | Inherits `EmbeddedViewerBase<ViewerWeb>`; no local `SetCallback` override. |

### Navigation Providers

Commands:

```powershell
rg -n "SetCallback\(INavigationMenuCallback|TryCaptureNavigationMenuCallback|InvokeNavigationMenuCallback|NavigationMenuRequestNavigate|ExecuteMenuCommand" Plugins -g "*.h" -g "*.cpp"
.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter google_drive_navigation_menu_callback_clear_drains -TimeoutMultiplier 2.0
```

Findings:

- `FileSystem`, `FileSystemDummy`, `FileSystem7z`, `FileSystemCurl`, and `FileSystemS3` store navigation callbacks but currently return `E_NOTIMPL` from `ExecuteMenuCommand(...)`; they have no active callback delivery route to drain beyond static registration cleanup.
- `FileSystemGoogleDrive` has active menu command delivery through `TryCaptureNavigationMenuCallback(...)`, `InvokeNavigationMenuCallback(...)`, and `NavigationMenuRequestNavigate(...)`; its focused drain regression passed.
- `FileSystemMicrosoftDrive` has generation/in-flight helper machinery but `ExecuteMenuCommand(...)` currently returns `E_NOTIMPL`; add focused behavioral coverage if a command-delivery route is enabled.

| Provider | Closeout state | Evidence |
|----------|----------------|----------|
| FileSystem | [x] | Static audit: no active callback delivery path after clear. |
| FileSystemDummy | [x] | Static audit: no active callback delivery path after clear. |
| FileSystem7z | [x] | Static audit: no active callback delivery path after clear. |
| FileSystemCurl | [x] | Static audit: no active callback delivery path after clear. |
| FileSystemS3 | [x] | Static audit: no active callback delivery path after clear. |
| FileSystemGoogleDrive | [x] | `google_drive_navigation_menu_callback_clear_drains` passed and was archived. |
| FileSystemMicrosoftDrive | [x] | Static audit: helper machinery exists; no current command-delivery route from `ExecuteMenuCommand(...)`. |

Remote credential coverage remains conditional and belongs to `Specs/Testing/Testing_SelfTestRemoteCredentials.md`. It does not block this closeout because the only in-tree active navigation callback delivery path has deterministic selftest coverage.

## Quiet-Point Ownership

The shared plugin quiet-point order remains:

1. Stop producers.
2. Request shutdown or cancellation.
3. Stop posting payload messages.
4. Clear registration-style callbacks.
5. Release instances.
6. Unload modules only after no callback can still run.

Spec/header cross-check command:

```powershell
rg -n "plugin quiet point|stop producers|SetCallback\(nullptr, nullptr\)|callback drain|DrainPostedPayloadsForWindow" Specs\Plugins\Plugins_PluginAPI.md Specs\Plugins\Plugins_ViewerPlugins.md Specs\Plugins\Plugins_VirtualFileSystem.md Specs\UI\UI_NavigationView.md Common\PlugInterfaces\Viewer.h Common\PlugInterfaces\NavigationMenu.h
```

Result:

- Viewer and navigation public headers document `SetCallback(nullptr, nullptr)` as the synchronous drain point.
- Viewer, VFS, and NavigationView specs document stale-drop/drain behavior after callback clear.
- `Specs/Plugins/Plugins_PluginAPI.md` keeps module quiet-point ownership distinct from instance close, callback clear, cancellation, and COM release.
- HWND payload receiver requirements remain in the posted-payload contract: initialize on create and drain in `WM_NCDESTROY`.

Compare-specific shutdown, queueing, and perf work is not re-owned here. The active owner is `Specs/Plans/WIP/Operation_Crosscut_CompareDirectoriesRemediation_SyncDataSafetyAndOptionsSimplification_2026-06-15.md`.

## HRESULT/Status Formatting Split

Audit command:

```powershell
rg -n "FormatMessageW|FormatHResult\(|FormatStatusText\(|sprintf_s|swprintf_s|C4774" RedSalamander Common Plugins -g "*.cpp" -g "*.h" -g "*.rc" -g "*.vcxproj"
```

Findings:

- `Common/Helpers.h` is the only production source using `FormatMessageW` for shared HRESULT/system-message text.
- `RedSalamander/FolderViewInternal.h` still has a trivial `FormatHResult(HRESULT)` alias over `FormatHResultMessage(...)`.
- `RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp` still has a trivial `FormatStatusText(HRESULT)` alias over `FormatHResultMessage(...)`.
- No `sprintf_s`, `swprintf_s`, or `C4774` hits were found in the audited production paths.

Decision:

- The broader numeric HRESULT/status audit contains diagnostics and domain-data cases that must be classified before changing user-visible text.
- That work is larger than an umbrella-plan closeout and now belongs to `Specs/Plans/WIP/Core_HResultStatusFormattingCleanup_2026-06-19.md`.

## Validation Evidence

Build gates:

| Gate | Command | Result | Evidence |
|------|---------|--------|----------|
| Build app | `.\build.ps1 -ProjectName RedSalamander` | Passed; strict scan found no `warning <code>` or `error <code>` diagnostics. | `.build/logs/msbuild-20260619_173726_912.log` |
| Build viewer test | `.\build.ps1 -ProjectName ViewerSqliteTests` | Passed; strict scan found no `warning <code>` or `error <code>` diagnostics. | `.build/logs/msbuild-20260619_174118_873.log` |

Focused regression gates:

| Gate | Command | Result | Evidence |
|------|---------|--------|----------|
| Viewer callback regression | `.\.build\x64\Debug\ViewerSqliteTests.exe` | Passed; stdout includes `clearing the viewer callback before close does not deliver ViewerClosed immediately`, `viewer does not invoke ViewerClosed after SetCallback(nullptr, nullptr) returns`, and `ViewerSqliteTests passed.` | `Specs/TestRuns/7d3a1247382a/ViewerSqliteTests/2026-06-19_174259/` |
| Navigation callback regression | `.\Tools\Run-AllTests.ps1 -Suite Compare -SkipBuild -CaseFilter google_drive_navigation_menu_callback_clear_drains -TimeoutMultiplier 2.0` | Passed: `1 passed / 0 failed / 0 skipped`. | `Specs/TestRuns/7d3a1247382a/CompareDirectories/2026-06-19_174206/` |
| App close/modeless smoke | `.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter modeless_window_ownership -TimeoutMultiplier 2.0` | Passed: `1 passed / 0 failed / 0 skipped`. | `Specs/TestRuns/7d3a1247382a/Commands/2026-06-19_174244/` |

Plan correction:

- The previous WIP validation table named the stale case filter `app_windows_open_and_close_modeless`. The current Commands selftest case is `modeless_window_ownership`, so the closeout evidence uses the real case name.

## Regression Guards

- [x] Viewer callback clear suppresses later `ViewerClosed` delivery.
- [x] Navigation callback clear/drain is covered for the provider with an active command-delivery path.
- [x] Shared quiet-point wording is documented in plugin specs and public callback headers.
- [x] No production formatting changes were made in this closeout; formatting regression guards belong to `Core_HResultStatusFormattingCleanup_2026-06-19.md`.

## Final Notes

- Future registration-style callback changes must update the relevant public header and plugin spec, not only an implementation file.
- Future navigation providers that add real `ExecuteMenuCommand(...)` callback delivery must add focused clear/drain selftest coverage.
- Future Compare shutdown/threading/perf work belongs to the Crosscut Compare remediation plan and `Specs/Core/Core_CompareDirectories.md`.
- Future HRESULT/status text cleanup belongs to `Specs/Plans/WIP/Core_HResultStatusFormattingCleanup_2026-06-19.md`.
