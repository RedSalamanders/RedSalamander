# Win32 Inventory

Last updated: 2026-10-17

Status: WIP - narrow RAII/lifetime/DPI audit plan. Evidence anchors were re-read against the tree on 2026-10-17. This file is that verification, not the 2026-07-02 snapshot. CHK-1, CHK-3, CHK-4, and CHK-6 are still open, so this plan stays in `Specs/Plans/WIP/`.

Original 2026-05-18 inventory snapshot removed 2026-07-02 (predates BatchRename/FileSystemMtp/WarpDrive; see git history of this file).

## Scope

This plan covers four bounded Win32 audits: owned-window manual destroy calls, manual `Release()` sites that were recorded in `Common/DxUi/DxUi.Accessibility.cpp`, raw `CoInitializeEx`/`CoUninitialize` normalization, and `WM_DPICHANGED_AFTERPARENT` coverage for child-hosted windows.

The supporting counts are regex-assisted, then checked by reading the cited lines on 2026-10-17. They are useful for locating hot spots, but each item still needs code review before changing behavior.

Intentional platform Win32 stays out of these four items: the message loop, shell calls, WIL wrappers, and top-level HWND shells (`WS_OVERLAPPEDWINDOW`, `WS_POPUP`, dialogs) that correctly handle `WM_DPICHANGED`. A short "verified elsewhere" note below records real Win32 UI that other specs already own.

## Central Win32 Infrastructure

Verified present in this repository on 2026-10-17:

- `Common/WindowMessages.h` is the custom message registry. Keep all `WM_APP`/`WM_USER` IDs there unless a component-local constant is intentionally private.
- `Common/Helpers.h` owns `PostMessagePayload(...)`, `TakeMessagePayload<T>(...)`, `InitPostedPayloadWindow(...)`, and `DrainPostedPayloadsForWindow(...)` (definitions from about line 2389).
- `Common/Win32CallbackHelpers.h` centralizes no-throw wrappers for `SetWindowLongPtrW`, `CallWindowProcW`, dialog creation, and WndProc hook install/restore.
- `Common/ViewerFileComboHost.h` standardizes viewer file-combo host subclassing.
- `Common/WindowSizing.h` and `Common/WindowBackdropPolicy.h` carry reusable DPI/min-track/backdrop helpers.
- `Common/ModalWindowShell.h` creates owned modal shells (`WS_CAPTION | WS_SYSMENU | WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS` at line 22).

DxUi is an external restored dependency, not an in-tree `Common/DxUi/` tree. `Build/RedSalamander.DxUi.props` imports `DxUi.resolved.$(Platform).props` when that file exists and `$(DxUiRoot)\Build\DxUi.Consumer.props` when `DxUiRoot` is set. `RequireRestoredDxUi` fails the build when `DxUiRoot` is empty and tells the builder to run `Tools/Restore-DxUi.ps1`. There is no `Common/DxUi/` directory in this repo, so `Common/DxUi/DxUi.Accessibility.cpp`, `DxUi.Menu.cpp`, and `DxUi.WindowHost.cpp` are not sources this plan can cite as in-tree files.

The durable per-monitor DPI contract for DxUi hosts and hybrid shells is `Specs/UI/UI_DxUiWinUIDesign.md` section 2.9 (required behavior at lines 312-317). CHK-6 tracks the remaining in-tree gaps against that contract. It does not restate it.

## Existing Guardrails

- WIL RAII is widely used for handles and COM (`wil::unique_*`, `wil::com_ptr`, `wil::scope_exit`).
- Custom message IDs are centralized in `Common/WindowMessages.h`.
- Posted heap payloads use `PostMessagePayload(...)` and `TakeMessagePayload<T>(...)`. A 2026-10-17 search found no `PostMessageW(... .release())` payload handoff. CHK-5 stays dropped; see below.
- `WM_DPICHANGED_AFTERPARENT` is handled in 11 in-tree files (listed under CHK-6). The 2026-07-02 "grew from 6 files to 11" list is stale: it counted comment-only mentions in `BatchRenameWindow.cpp` and `FindFilesWindow.cpp`, and it counted `DxUi.Menu.cpp` / `DxUi.WindowHost.cpp`, which are not in this repo. The current handler list is the CHK-6 coverage list.
- `std::thread::detach()` is not a checklist item. `AGENTS.md` already owns the policy. The 2026-05-18 sentence that production had no such detach is stale. The 2026-10-17 search found these production sites, each with an adjacent comment for a bounded-wait or process-lifetime worker: `RedSalamanderMonitor/RedSalamanderMonitor.cpp:2667` and `:2889`, `RedSalamanderMonitor/EtwListener.cpp:253`, `RedSalamander/SettingsHotReload.cpp:458`, and `Plugins/ViewerSpace/ViewerSpace.cpp:8023`. Other `.detach()` hits in that search are COM or stream ownership transfers, not threads.

## Live Audit Checklist

Anchors below are from the 2026-10-17 code read. Line numbers are the current tree.

- [ ] **CHK-1 (NEXT ACTION)** - Replace `DestroyWindow(_hWnd.get())` on owned `wil::unique_hwnd` members with `reset()`-based close (or explicit ownership transfer).

  Product sites from the 2026-07-02 anchors are already `reset()`. Those old lines (`ViewerText.cpp:592` and `:602`, `RedSalamander.cpp:760` and `:1152`) no longer match the code:

  - `Plugins/ViewerText/ViewerText.cpp:597` and `:607` call `_hWnd.reset()` from `ViewerTextPromptWindow::Confirm` / `Cancel`. The member is `wil::unique_hwnd _hWnd` at line 619.
  - `RedSalamander/RedSalamander.cpp:808` and `:1148` call `_hWnd.reset()` from `AboutDialogWindow::Close` and `FatalErrorDialogWindow::Close`. The members are `wil::unique_hwnd _hWnd` at lines 815 and 1157.
  - The product rename prompt does the same: `RedSalamander/FolderViewInternal.h:666` (`Confirm`) and `:676` (`Cancel`).

  Repo-wide search for `DestroyWindow` on a `.get()` HWND found one remaining file. It is still CHK-1 work:

  - `Tests/PerformanceTests2/Performance.Tests.FolderViewInternal.Access.h:772` and `:782`, in the test `FolderViewRenamePromptWindow::Confirm` / `Cancel`. `_hWnd` is `wil::unique_hwnd` at line 846. `WM_NCDESTROY` at lines 623-627 calls `_hWnd.release()` when the wrapper still owns that HWND. That release can avoid a second destroy on the synchronous path; the call is still the banned owner pattern, and this test copy has not been moved to `_hWnd.reset()` like `FolderViewInternal.h`.

  Leave this box unchecked until that test site is gone. It is the remaining bounded CHK-1 edit.

- [ ] **CHK-3** - The 2026-07-02 item was an audit of 18 manual `->Release()` sites in `Common/DxUi/DxUi.Accessibility.cpp`. That file is not in this repository. `Common/DxUi/` is absent. DxUi is the external root described under Central Win32 Infrastructure (`Build/RedSalamander.DxUi.props`).

  This repo cannot close CHK-3 against `Common/DxUi`. Do not mark the item done from an empty search. Re-scope it to an upstream DxUi audit of `DxUi.Accessibility.cpp` (UI Automation provider plumbing: COM-boundary return `Release()` versus owning raw COM members that should be `wil::com_ptr`). Treat each upstream site as COM-boundary code, not assumed safe or unsafe without local review.

  In-tree `->Release()` calls exist in plugin factories and COM lifetime helpers. They are outside this item. A 2026-10-17 search found no `AddRef();` followed by `.attach(`.

- [ ] **CHK-4** - Normalize raw `CoUninitialize()` against `wil::CoInitializeEx` where the lifetime is scope-shaped. Keep an explicit comment where manual COM lifetime is required (window-lifetime STA, process STA, or a small RAII type whose destructor is the uninitialize point).

  2026-10-17 count of raw `CoUninitialize()`: **26 calls in 17 files**. The 2026-07-02 "27 calls across 18 files" and "10 `wil::CoInitializeEx` uses" figures are stale.

  WarpDrive Tasks 4 and 9 are checked done in `Specs/Plans/Done/Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`. The old "do not edit `FolderView.FileOps.cpp` / `FolderView.Enumeration.cpp`" lock is retired. Those files no longer call `CoUninitialize()`; their current sites are in the already-paired list below (`FolderView.FileOps.cpp:1037` / `:1048`, `FolderView.Enumeration.cpp:232` / `:238`, and `wil::CoInitializeEx_failfast` at `FolderView.Enumeration.cpp:711` and `:848`).

  Raw `CoUninitialize()` still paired with raw `CoInitializeEx` (the open set):

  - `RedSalamander/FolderWindow.FileOperations.cpp:1240` init, `:1245` uninit (`OpenLocationInFileExplorerCallback`, uninit only when `SUCCEEDED`).
  - `RedSalamander/FolderWindow.cpp:381` init, `:387` uninit (`ResolveShellLinkTarget`).
  - `RedSalamander/FolderView.cpp:877` init in `OnCreate`. Uninit at `:280` (`Destroy`) and `:1026` (`OnDestroy`), gated by `_coInitialized`. Same two functions also call `OleUninitialize()` at `:285` and `:1021` when `_oleInitialized` (`OleInitialize` at `:892`). This lifetime spans the window, so a function-scope `wil::CoInitializeEx` does not fit without storing the guard. `RPC_E_CHANGED_MODE` is already documented at lines 882-885.
  - `RedSalamander/RedSalamander.cpp:7937` init, `:7958` uninit (process STA around startup; `RPC_E_CHANGED_MODE` logged at 7944-7946).
  - `RedSalamander/SplashScreen.cpp:674` init, `:684` uninit (splash worker thread).
  - `RedSalamander/NavigationView.Rendering.cpp:1080` init, `:1086` uninit (`DriveInfoWorker`).
  - `RedConfigure/RedConfigureSplashScreen.cpp:514` init, `:523` uninit (splash worker thread).
  - `Plugins/FileSystem/FileSystem.cpp:577` init, `:583` uninit (`ResolveShellLinkTargetForProperties`).
  - `Plugins/FileSystem/FileSystem.FileOps.cpp:9066` init, `:9072` uninit (`RunShellOperationOnDedicatedStaThread`); `:9575` / `:9581` (`DeleteToRecycleBin`); `:9856` / `:9862` (batched recycle). These are dedicated STA scopes for `IFileOperation`.
  - `Plugins/FileSystemMtp/FileSystemMtp.Device.cpp:84` init, `:92` uninit (`ComInitialization` constructor / destructor). The queue worker itself uses `wil::unique_couninitialize_call` in `FileSystemMtp.Core.cpp` (below). This per-call helper is the remaining raw uninit in the MTP plugin.
  - `Plugins/ViewerImgRaw/ViewerImgRaw.Decode.cpp:958` init, `:964` uninit (`DecodeImageToBgraWic`; comment at 951-957 already says uninit only when this call succeeded).
  - `Plugins/ViewerImgRaw/ViewerImgRaw.Export.cpp:169` init, `:175` uninit (`ShowExportSaveDialog`); `:812` init, `:818` uninit (export worker).
  - `Tests/ViewerPETests/ViewerPE.Tests.Scope.Tests.ViewerSpaceAnalysisHost.cpp:100` init, `:373` uninit (gated by `uninitializeCom`).
  - `Tests/ViewerPETests/ViewerPE.Tests.Runner.cpp:16068` uninit and `:16069` init on `promptProbeThread`. The `wil::scope_exit` that calls `CoUninitialize()` is armed before `CoInitializeEx`, and the cleanup is unconditional.
  - `Tests/ViewerPETests/ViewerPE.Tests.Runner.cpp:16372` init, `:16382` uninit (process entry, uninit only when `SUCCEEDED`).
  - `Tests/ViewerSqliteTests/ViewerSqlite.Tests.Runner.cpp:467` init, `:478` uninit; `:574` init, `:585` uninit (UIA helpers, uninit only when `SUCCEEDED`).
  - `RedSalamander/SelfTest/Commands/Commands.Tests.ShellCommands.cpp:1061` init, `:1067` uninit (`CreateShellLinkForShellCommandTest`).
  - `RedSalamander/SelfTest/Commands/Commands.Tests.Settings.cpp:6642` init, `:6653` uninit (`CollectUiaSelectionPatternState`); `:6738` init with failure-path uninit at `:6751` and `:6762`, and destructor uninit at `:6791` (`UiaThreadContext`).

  Already off raw `CoUninitialize()`. These are the comparison set, not the open call list:

  - `wil::CoInitializeEx`: `Plugins/FileSystem/FileSystem.Search.cpp:741` and `:2151`; `RedSalamander/SelfTest/FileOperations/FileOperations.Scope.Tests.ReviewRemediation.cpp:68`; `Tests/PerformanceTests2/Performance.Tests.FolderViewRefreshDuplicatePathPerfTest.cpp:914`.
  - `wil::CoInitializeEx_failfast`: `RedSalamander/FolderWindow.FileOperations.State.cpp:6900` and `:12243`; `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.ItemResultReducer.cpp:1537`; `RedSalamander/FolderView.Enumeration.cpp:711` and `:848`; `Plugins/FileSystem/FileSystem.FileOps.cpp:657` and `:14300`.
  - Raw `CoInitializeEx` plus `wil::unique_couninitialize_call`: `RedSalamander/FolderWindow.FileSystem.Commands.cpp:13811` / `:13817` and `:14042` / `:14048`; `RedSalamander/FolderView.FileOps.cpp:1037` / `:1048`; `RedSalamander/FolderView.Enumeration.cpp:232` / `:238`; `RedSalamander/FolderView.Icons.cpp:993` / `:994`; `RedSalamander/BatchRenameWindow.cpp:768` / `:775` and `:1396` / `:1403`; `RedSalamander/CompareDirectoriesEngine.cpp:3843` / `:3849` and `:4095` / `:4101`; `RedSalamander/SelfTest/Commands/Commands.Tests.BatchRename.cpp:2663` / `:2668` and `:2748` / `:2753`; `RedSalamanderMonitor/ColorTextView.cpp:3678` / `:3684` and `:4554` / `:4560`; `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp:2285` / `:2306`; `Plugins/ViewerPE/ViewerPE.cpp:172` / `:178`; `Plugins/Terminal/TerminalRuntimeLoader.cpp:47` with `wil::unique_couninitialize_call _coUninitialize` at `TerminalRuntimeLoader.h:35`; `Tests/PerformanceTests2/Performance.Tests.FolderView.Enumeration.Access.cpp:248` / `:254`.

  Converting a raw init that already has `wil::unique_couninitialize_call` over to `wil::CoInitializeEx` is the same normalization, and it is separate from the 26 raw `CoUninitialize()` calls above.

- [ ] **CHK-6** - Systematic `WM_DPICHANGED_AFTERPARENT` audit of child/custom windows with DPI-sensitive chrome. Add `WM_DPICHANGED_AFTERPARENT` where the window is child-hosted and needs the parent DPI transition. The durable rule is `Specs/UI/UI_DxUiWinUIDesign.md` section 2.9 item 3. Top-level windows (`WS_OVERLAPPEDWINDOW`, owned `WS_POPUP`, dialogs) correctly handle `WM_DPICHANGED`; those are listed at the end of this item so they are not re-opened as gaps.

  In-tree handlers of `WM_DPICHANGED_AFTERPARENT` on 2026-10-17 (11 files):

  - Child: `RedSalamander/FolderView.cpp:832` (created `WS_CHILD` at line 239); `RedSalamander/NavigationView.cpp:586` (`WS_CHILD` at line 455); `RedSalamander/FunctionBar.cpp:528` (`WS_CHILD` at line 418); `RedSalamander/FolderWindow.StatusBar.cpp:966` (created `WS_CHILD` in `RedSalamander/FolderWindow.cpp:1809-1814`).
  - Child, always: `Plugins/Terminal/Terminal.cpp:4646` (`WS_CHILD` at line 2653). This proc does not also handle `WM_DPICHANGED`.
  - Child hosts inside ViewerSpace: `Plugins/ViewerSpace/ViewerSpace.AnalysisHost.cpp:646` (`WS_CHILD` at line 153); `Plugins/ViewerSpace/ViewerSpace.NavigationHost.cpp:419` (`WS_CHILD` at line 68).
  - Dual-mode viewer that already handles both messages: `Plugins/ViewerSpace/ViewerSpace.cpp:3158` (`WM_DPICHANGED`) and `:3159` (`WM_DPICHANGED_AFTERPARENT`). Embedded creation is `WS_CHILD` at line 13411.
  - Top-level windows that also branch on `WM_DPICHANGED_AFTERPARENT`: `RedSalamander/Ui/AlertOverlayWindow.cpp:1240-1241` (created `WS_POPUP` at line 2505); `RedSalamander/FolderWindow.ItemProperties.cpp:1317` (DxUi-handled path; the window is `WS_OVERLAPPEDWINDOW` at line 1155, and the switch handles `WM_DPICHANGED` at line 1340); `RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp:1778-1779` (DxUi-handled path; the window is `WS_OVERLAPPEDWINDOW` at line 1915, and the switch handles `WM_DPICHANGED` at line 1812).

  Open gaps. These windows are created `WS_CHILD` in embedded or always-child mode and do not handle `WM_DPICHANGED_AFTERPARENT`. A parent DPI change delivers `WM_DPICHANGED_AFTERPARENT` to the child; `WM_DPICHANGED` is the top-level message.

  Dual-mode viewers that handle only `WM_DPICHANGED` (standalone overlapped path is fine; embedded `WS_CHILD` path is the gap):

  - `Plugins/ViewerText/ViewerText.cpp:4765`. Embedded `WS_CHILD` at line 9700; standalone `WS_OVERLAPPEDWINDOW` at lines 9736 and 9790.
  - `Plugins/ViewerWeb/ViewerWeb.cpp:2927`. Style selection at line 2622.
  - `Plugins/ViewerSqlite/ViewerSqlite.cpp:840`. Style selection at line 973.
  - `Plugins/ViewerPE/ViewerPE.cpp:626`. Style selection at line 2986.
  - `Plugins/ViewerImgRaw/ViewerImgRaw.cpp:995`. Style selection at line 3710.

  Always-child surface that handles only `WM_DPICHANGED`:

  - `RedSalamanderMonitor/ColorTextView.cpp:5284`. `ColorTextView::Create` uses `WS_CHILD` at line 206. The monitor main window is top-level and handles `WM_DPICHANGED` at `RedSalamanderMonitor/RedSalamanderMonitor.cpp:4926`. That handler synthesizes `WM_DPICHANGED` to the toolbar and status children at lines 4620 and 4625. It does not send `WM_DPICHANGED_AFTERPARENT` to `ColorTextView`.

  Same gap, no DPI message at all, on child windows that read DPI while painting:

  - `Plugins/ViewerVLC/ViewerVLC.cpp` main proc handles `WM_SIZE` at line 1957 and has no `WM_DPICHANGED` or `WM_DPICHANGED_AFTERPARENT`. Embedded creation is `WS_CHILD` at line 2747. Child video / HUD / overlay / seek windows are `WS_CHILD` at lines 2528, 2539, 2544, and 2549. Paint and layout call `GetDpiForWindow` (for example lines 3363, 3592, 3689, 3799, and 3917).
  - `Plugins/ViewerText/ViewerText.Text.cpp:2632` and `Plugins/ViewerText/ViewerText.Hex.cpp:1492` handle `WM_SIZE` only. Both windows are `WS_CHILD` (`ViewerText.cpp:5052-5053` and `:5060-5061`) and call `GetDpiForWindow` throughout paint and layout. They follow the parent; the parent gap is the `ViewerText.cpp:4765` row above.

  In-tree DxUi child thunks forward every message into external `WindowHost::HandleMessage` and do not themselves branch on `WM_DPICHANGED_AFTERPARENT`. Whether that external host implements section 2.9 is outside this repo: `RedSalamanderMonitor/RedSalamanderMonitor.cpp:2360-2387` (toolbar and status created `WS_CHILD` at lines 2552 and 2604); `RedSalamander/Preferences.Dialog.cpp:1364-1393` (`PreferencesDxShellHostWndProc`, child created at line 1457). The preferences dialog itself is top-level and handles `WM_DPICHANGED` at line 5332, then reflows the page host.

  Verified top-level `WM_DPICHANGED`-only handlers. These are expected and are not CHK-6 defects:

  - `RedSalamander/RedSalamander.cpp:659` (About) and `:947` (Fatal Error), both `ModalWindowShell` popups; main window `:12220`.
  - `RedSalamander/FolderViewInternal.h:431` and `:445` (rename prompt, `WS_POPUP` at line 343). The PerformanceTests2 mirror is the same shape at `Performance.Tests.FolderViewInternal.Access.h:591` and `:605`.
  - `Plugins/ViewerText/ViewerText.cpp:414` and `:438` (prompt, `WS_POPUP` at line 307).
  - `RedSalamander/ShortcutsWindow.cpp:942` (`WS_OVERLAPPEDWINDOW` at line 819).
  - `RedSalamander/CommandPaletteWindow.cpp:322` (`WS_POPUP` at line 239).
  - `RedSalamander/FloatingTerminalWindow.cpp:327` (`WS_OVERLAPPEDWINDOW` at line 195).
  - `RedSalamander/ConnectionManagerWindow.cpp:4269` (`WS_OVERLAPPEDWINDOW` at line 1249).
  - `RedSalamander/CompareDirectoriesWindow.cpp:375` (`WS_OVERLAPPEDWINDOW` at line 538).
  - `RedSalamander/BatchRenameWindow.cpp:2521` (`WS_OVERLAPPEDWINDOW` at line 2398). The comment at line 2527 is not a handler. On `WM_DPICHANGED` the window forwards DPI into the child `NavigationView`, which has its own `WM_DPICHANGED_AFTERPARENT` handler.
  - `RedSalamander/FindFilesWindow.cpp:7396` (`WS_OVERLAPPEDWINDOW` at line 2556). The comment at line 7410 is not a handler. Same explicit child forward as Batch Rename.
  - `RedSalamander/SplashScreen.cpp:498` (`WS_POPUP` at line 717); `RedConfigure/RedConfigureSplashScreen.cpp:439` (`WS_POPUP` at line 550).
  - `RedSalamander/ConnectionCredentialPromptDialog.cpp:820` (`WS_POPUP` at line 327).
  - `RedSalamander/Preferences.Dialog.cpp:5332` (`CreateDialogParamResourceNoThrow` of `IDD_PREFERENCES` at lines 5569-5570).
  - `RedSalamander/FolderWindow.LeftoverCleanup.cpp:195` (`WS_POPUP` at line 949).
  - `RedSalamander/FolderWindow.FileSystem.cpp:999`, `:1581`, `:2268`, `:2814`, `:3511` (popups created at lines 871, 1452, 2149, 2695, and 3376).
  - `RedSalamander/FolderWindow.FileSystem.Commands.cpp:1263`, `:3568`, `:5055`, `:5720`, `:9511`, `:9928` (popups created at lines 1121, 3426, 4919, 5584, 13151, and 13463).
  - `RedSalamander/FolderWindow.FileOperations.Popup.cpp:3236` and `:3262` (speed-limit prompt, `WS_POPUP` at line 2546); popup proc `:16742` (`WS_POPUP` at line 16842).

### Dropped checklist items (2026-07-02 folder review)

CHK-2 (long WndProc/DialogProc switch structure), CHK-5 (`PostMessagePayload` teardown contract for touched senders/receivers), CHK-7 (localization placeholder rules for resource edits), and CHK-8 (focused selftests plus archived perf evidence for responsiveness-affecting refactors) were dropped from this plan: the policy is already codified in `.github/skills` (win32-wndproc, wil-raii, localization) and `AGENTS.md` - these are ongoing coding standards, not discrete deliverables.

2026-10-17 recheck, still not reopened: no `PostMessageW` call passes `.release()` as the payload, and no `AddRef();` is followed by `.attach(`.

## Verified elsewhere (not these four items)

These are real Win32 surfaces. Closed UI plans and the audit scripts own them. They are not CHK-1/3/4/6 work.

- File Operations failure chrome still builds `STATIC` / `BUTTON` children and a message font: `RedSalamander/FolderWindow.FileOperations.Popup.cpp:14068-14072`, `CreateFontIndirectW` at line 14136, `WM_SETFONT` at line 14142. Owner: `Specs/Plans/Done/UI_RemainingWin32UiDependencyRetirementPlan.md`.
- Plugin config is still a dialog: `RedSalamander/ManagePluginsDialog.cpp:3997-3998` and `:4090-4091` (`DialogBoxParamResourceNoThrow` of `IDD_PLUGIN_CONFIG`). Owner: the same retirement plan and `Specs/Plans/Done/UI_Win32CommonControlsAudit_2026-05-13.md`.
- Preferences is a modeless dialog template, not `DialogBoxParam`: `RedSalamander/Preferences.Dialog.cpp:5569-5570` (`IDD_PREFERENCES`). Owner: `Specs/Plans/Done/UI_PreferencesWin32Removal.md`.
- Other modal dialog shells still in tree: `Plugins/ViewerText/ViewerText.Encoding.cpp:555-556` (`IDD_VIEWERTEXT_ENCODING_PICKER`); `RedSalamanderMonitor/RedSalamanderMonitor.cpp:4657` (`DialogBox` of `IDD_ABOUTBOX`).
- Inventory scripts, not this plan: `Tools/Audit-RemainingWin32UiDependencies.ps1`, `Tools/Audit-VisibleNativeSurfaces.ps1`, `Tools/Audit-ComctlReportSurfaces.ps1`.

## Verification Notes

No build was run for the 2026-10-17 revision because the change is documentation only. Minimum closeout verification for future code changes from this plan should include:

- `.\build.ps1 -ProjectName RedSalamander`
- Targeted viewer/plugin/test project builds for the touched area.
- Relevant command selftest case(s) from `RedSalamander/SelfTest/Commands`.
- `git diff --check`

## Closeout Criteria

This WIP can move to `Specs/Plans/Done/` after CHK-1, CHK-3, CHK-4, and CHK-6 are either completed or split into narrower WIP plans, and any durable Win32 contract changes are merged into `AGENTS.md`, `.github/skills/win32-wndproc/SKILL.md`, `.github/skills/wil-raii/SKILL.md`, or the relevant `Specs/<Domain>/` document.

CHK-3 is not completable inside this repo against `Common/DxUi`. Closing it means re-scoping the item to the upstream DxUi tree, or splitting that audit into its own plan, and recording that decision here. Marking the box done because the in-tree file is gone would hide the original audit.
