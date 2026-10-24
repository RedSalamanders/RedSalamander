# UI Win32 Common Controls Audit and Removal Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Inventory the remaining Win32 common-control usage across `RedSalamander`, `RedSalamanderMonitor`, `RedConfigure`, `Common`, and all `Plugins`, then provide a removal checklist.

**Architecture:** Keep top-level Win32 windows, message loops, DPI, shell integration, and non-visible OS interop, but remove app-owned visible comctl32 controls from product UI. Where a surface is still native but not a comctl32 common control, keep it in a separate adjacent backlog so the comctl removal stays measurable.

**Tech Stack:** C++23, Win32 HWND shell, WIL RAII, Direct2D/DirectWrite, DxUi `WindowHost`, `.rc` resources, command self-tests, audit scripts.

---

Last updated: 2026-05-13

Status: Done - visible comctl listview/tooltip surfaces removed

## Scope

Requested product surface:

- `RedSalamander`
- `RedSalamanderMonitor`
- `RedConfigure` (the project matching the requested "RedConfig")
- `Common`
- `Plugins`

Included in the audit:

- Source, headers, resource scripts, manifests, and project/resource compile entries under the requested product surface.
- Shared `Common/DxUi` and `Common/Common` code, plus every first-party plugin folder under `Plugins`.
- Comctl32/common-control classes and APIs: `SysListView32`, `SysTreeView32`, `msctls_*`, `ToolbarWindow32`, `STATUSCLASSNAME`, `TOOLBARCLASSNAME`, `WC_*` control classes, `ListView_*`, `TreeView_*`, `ImageList_`, `IImageList`, `SHGetImageList`, `InitCommonControls`, and Common-Controls v6 manifests.

Separated but not counted as comctl32 common controls:

- Standard USER32 dialog controls such as `Static`, `Edit`, `Button`, `ComboBox`, `LTEXT`, `EDITTEXT`, `PUSHBUTTON`, and `DEFPUSHBUTTON`.
- Standard scroll bars on custom HWNDs using `SetScrollInfo`, `GetScrollInfo`, `SB_VERT`, and `SB_HORZ`.
- Custom owner windows that only use HWND hosting, custom window classes, or DxUi-rendered controls.

## Audit Commands

Fresh source searches were run on 2026-05-13 from repo root:

```powershell
rg -n --glob '*.{cpp,h,rc,vcxproj,manifest}' -i "commctrl|InitCommonControls|ICC_|COMCTL32|Microsoft.Windows.Common-Controls|SysListView32|SysTreeView32|SysHeader32|SysTabControl32|SysLink|msctls_[A-Za-z0-9_]+|ToolbarWindow32|ReBarWindow32|STATUSCLASSNAME|TOOLBARCLASSNAME|REBARCLASSNAME|WC_LISTVIEW|WC_TREEVIEW|WC_HEADER|WC_TABCONTROL|WC_COMBOBOXEX|WC_LINK|WC_NATIVEFONTCTL|PROGRESS_CLASS|TRACKBAR_CLASS|UPDOWN_CLASS|DATETIMEPICK_CLASS|MONTHCAL_CLASS|HOTKEY_CLASS|ANIMATE_CLASS|ListView_|TreeView_|Header_|TabCtrl_|ImageList_|HIMAGELIST|LVS_|LVN_|LVIF_|TVS_|TVN_|TVIF_|TCS_|TCN_|HDN_|HDS_|PBM_|TB_|TBM_|TTM_|UDM_|DTM_|MCM_" Common Plugins RedSalamander RedSalamanderMonitor RedConfigure
```

```powershell
rg -n --glob '*.{cpp,h,rc}' 'InitCommonControls|#include <[Cc]omm[Cc]trl\.h>|#pragma comment\(lib, "[Cc]omctl32\.lib"\)|SysListView32|ListView_|LVS_|LVN_|LVIF_|LVIS_|LVNI_|msctls_statusbar32|SysTreeView32|TOOLTIPS_CLASS|TTM_|TOOLINFO|TTS_|TTF_|HIMAGELIST|ImageList_|SHGetImageList|IImageList|SHIL_' Common Plugins RedSalamander RedSalamanderMonitor RedConfigure
```

Existing audit helpers:

```powershell
powershell -ExecutionPolicy Bypass -File .\Tools\Audit-VisibleNativeSurfaces.ps1
powershell -ExecutionPolicy Bypass -File .\Tools\Audit-ComctlReportSurfaces.ps1
powershell -ExecutionPolicy Bypass -File .\Tools\Audit-RemainingWin32UiDependencies.ps1 -AsMarkdown
```

Observed helper status:

- `Tools/Audit-VisibleNativeSurfaces.ps1` passes and reports an empty narrow visible-native table.
- `Tools/Audit-ComctlReportSurfaces.ps1` passes across `Common`, `Plugins`, `RedSalamander`, `RedSalamanderMonitor`, and `RedConfigure`.
- `Tools/Audit-RemainingWin32UiDependencies.ps1 -AsMarkdown` currently reports separate broad Win32 UI findings in HDC/HFONT/font-message categories; those are outside the visible comctl listview/tooltip closeout.
- The expanded manual scan found one additional live common-control surface outside `RedSalamander`: `Plugins/ViewerSpace/ViewerSpace.cpp` creates a `TOOLTIPS_CLASSW` tracking tooltip and drives it with `TTM_*` messages.
- `Plugins/ViewerText`, `Plugins/ViewerWeb`, `Plugins/ViewerPE`, and `Plugins/ViewerImgRaw` include `commctrl.h` and link `comctl32`, but the manual scan did not find live comctl32 window classes or `TTM_`/`ListView_`/`TreeView_` API use in those files.
- The include/link scan also found stale or notification-only `commctrl.h` surfaces in `RedSalamander` headers and split implementation files. Those are tracked below as cleanup candidates, not visible-control blockers.
- `NMHDR`/`WM_NOTIFY` alone is not counted as a live common-control window. It is still a `commctrl.h` contract surface if the later target becomes "no commctrl header dependency."
- Broad `WC_` searches can produce false positives such as `WC_ERR_INVALID_CHARS` in UTF-8 conversion code. The inventory below excludes those non-UI hits.

## Pre-Removal Comctl32 Inventory

The inventory below records the surfaces found by the audit before the removal pass. Closeout status is summarized in
`Specs/UI/UI_VisibleComctlAudit.md`.

| Project | Surface | Files | Classification | Removal target |
| --- | --- | --- | --- | --- |
| `Common/DxUi` | Shared commctrl include/link | `Common/DxUi/DxUi.WindowHost.cpp:17`, `Common/DxUi/DxUi.WindowHost.cpp:32`, `Common/DxUi/DxUi.TextInput.cpp:8`, `Common/DxUi/DxUi.Grid.cpp:16` | **Shared dependency include/link**, no live comctl32 class found in the searched lines | Remove includes/link if a compile proves no comctl32 symbols are still required; keep hidden RichEdit text bridge separate because it is not a comctl32 control. |
| `Common/Common` | Common-control classes/API | None found in source/resources | **No live comctl32 common-control surface found** | No comctl removal needed in `Common/Common`. |
| `RedSalamander` | Shared Directories report list | Former `IDD_SHARED_DIRECTORIES` resource rows | **Removed**: `RedSalamander.SharedDirectoriesWindow` now hosts a `DxUi::Grid` and no `SysListView32` report control remains. | Closed for visible comctl scope. |
| `RedSalamander` | Shared Directories listview runtime helpers | Former `ListView_*` helpers in `FolderWindow.FileSystem.Commands.Part.cpp` | **Removed**: row insert/select/activation now uses the DxUi model/delegate path. | Closed for visible comctl scope. |
| `RedSalamander` | Native-control self-test classifier | `RedSalamander/FolderWindow.FileSystem.cpp` | **Guard only**: native-control classification remains a negative assertion surface, not a live Shared Directories dependency. | Keep as regression guard. |
| `RedSalamander` | Stale or notification-only `commctrl.h` includes outside live controls | `RedSalamander/CompareDirectoriesWindow.Internal.h:28`, `RedSalamander/FolderViewInternal.h:29`, `RedSalamander/ManagePluginsDialog.cpp:41`, `RedSalamander/Ui/AlertOverlayWindow.h:11` | **Header cleanup candidate**: no live common-control class/API found in these files during the expanded scan; `FolderWindowInternal.h` no longer pulls `<commctrl.h>`. | Remove or replace the remaining includes only after a focused build proves each translation unit has no `commctrl.h` symbol dependency. |
| `RedSalamander` | Native status-bar guard | `RedSalamander/FolderWindow.StatusBar.cpp:1271`, `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp:9901` | **Guard only**: checks `msctls_statusbar32` is not used | Keep as regression guard until the broader native-control audit script owns this check. |
| `RedSalamander` | Legacy Preferences tree guard | `RedSalamander/Preferences.Dialog.cpp:5480`, `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ChromeAndPlugins.cpp:75-76` | **Guard only**: checks `SysTreeView32` is not visible | Keep as regression guard until Preferences closeout is folded into the automated audit. |
| `RedSalamander` | `NavigationView` common-control init/include surface | `RedSalamander/NavigationView.cpp:15`, `RedSalamander/NavigationView.cpp:414-415`, `RedSalamander/NavigationView.Breadcrumb.cpp:5`, `RedSalamander/NavigationView.Edit.cpp:5`, `RedSalamander/NavigationView.FullPathPopup.cpp:5`, `RedSalamander/NavigationView.Interaction.cpp:5`, `RedSalamander/NavigationView.Menus.cpp:8`, `RedSalamander/NavigationView.Rendering.cpp:6` | **Residual initialization/header dependency**: calls `InitCommonControls()` but no matching tooltip/common-control window was found in `NavigationView` | Remove after confirming no common-control tooltip window is created by this path; then trim split-file includes. |
| `RedSalamander` | `WM_NOTIFY` custom notification plumbing | `RedSalamander/CompareDirectoriesWindow.Options.cpp:1696` plus custom FolderWindow status-bar notification structs | **Notification contract only**: custom/native message routing, not evidence of an app-owned common-control window by itself; FolderWindow status-bar clicks no longer require `NMHDR`/`NMMOUSE`. | Keep for visible-comctl removal. Revisit only if the target expands to removing `commctrl.h` contracts entirely. |
| `RedSalamander` | Shell image-list dependency | `RedSalamander/IconCache.cpp:12`, `RedSalamander/IconCache.h:37`, `RedSalamander/IconCache.h:272-275`, `RedSalamander/IconCache.cpp:592-627`, `RedSalamander/IconCache.cpp:809` | **Non-visible comctl32 shell-image-list dependency**: `IImageList`, `SHGetImageList`, `SHIL_*` | Keep if the goal is only visible common-control removal; replace with a non-comctl shell/WIC extraction path only if the goal becomes full comctl32 dependency removal. |
| `RedSalamander` | Common-Controls v6 manifest | `RedSalamander/res/exe.manifest:59-68` | **Activation/dependency manifest**, not a visible control | Remove only after all visible comctl and any standard themed native controls that rely on it are retired. |
| `RedSalamanderMonitor` | Common-control classes/API | None found in app source/resources | **No live comctl32 common-control surface found** | No comctl removal needed for Monitor source. |
| `RedSalamanderMonitor` | Common-Controls v6 manifest | `RedSalamanderMonitor/res/exe.manifest:57-65` | **Activation/dependency manifest**, not a visible control | Keep while Monitor still has standard dialog controls; remove only if those dialogs move to DxUi/custom UI and visual styles are no longer needed. |
| `RedConfigure` | Common-control classes/API | None found in app source/resources | **No live comctl32 common-control surface found** | No comctl removal needed for RedConfigure source. |
| `RedConfigure` | Common-Controls v6 manifest | `RedConfigure/res/exe.manifest:5-9` | **Activation/dependency manifest**, not a visible control | Candidate for removal after confirming RedConfigure does not need themed native controls beyond menu/top-level HWND. |
| `Plugins/ViewerSpace` | Tracking tooltip | `Plugins/ViewerSpace/ViewerSpace.cpp` | **Removed**: hover details paint as an in-canvas Direct2D/DirectWrite tooltip overlay with no `TOOLTIPS_CLASSW`, `TOOLINFOW`, or `TTM_*` path. | Closed for visible comctl scope. |
| `Plugins/ViewerText` | Comctl include/link and `NMHDR` notification handler with no live common-control use found | `Plugins/ViewerText/ViewerText.cpp:25`, `Plugins/ViewerText/ViewerText.cpp:42`, `Plugins/ViewerText/ViewerText.cpp:4623`, `Plugins/ViewerText/ViewerText.cpp:6901`, `Plugins/ViewerText/ViewerText.h:380` | **Compile/link cleanup candidate**; file combo host is a standard `Static` DxUi host, not a comctl control | Remove include/link if a focused build proves unused; keep or replace `NMHDR` routing according to the later dependency target. |
| `Plugins/ViewerWeb` | Comctl include/link with no live common-control use found | `Plugins/ViewerWeb/ViewerWeb.cpp:24`, `Plugins/ViewerWeb/ViewerWeb.cpp:41` | **Compile/link cleanup candidate**; file combo host is a standard `Static` DxUi host, not a comctl control | Remove include/link if a focused build proves unused. |
| `Plugins/ViewerPE` | Comctl include/link with no live common-control use found | `Plugins/ViewerPE/ViewerPE.cpp:16`, `Plugins/ViewerPE/ViewerPE.cpp:27` | **Compile/link cleanup candidate**; file combo host is a standard `Static` DxUi host, not a comctl control | Remove include/link if a focused build proves unused. |
| `Plugins/ViewerImgRaw` | Comctl include/link with no live common-control use found | `Plugins/ViewerImgRaw/ViewerImgRaw.cpp:18`, `Plugins/ViewerImgRaw/ViewerImgRaw.cpp:38` | **Compile/link cleanup candidate**; visible host hits are standard `Static` DxUi hosts | Remove include/link if a focused build proves unused. |
| `Plugins/ViewerVLC` | `NMHDR` notification handler with no live common-control use found | `Plugins/ViewerVLC/ViewerVLC.h:105`, `Plugins/ViewerVLC/ViewerVLC.cpp:3239` | **Notification contract only**: no `commctrl.h` include/link or live common-control class/API hit found in the expanded scan | Keep for visible-comctl removal. Revisit only if the target expands to removing `NMHDR` contracts entirely. |
| `Plugins` remaining projects | Common-control classes/API | No `Sys*`, `msctls_*`, `TOOLTIPS_CLASS`, `ListView_*`, `TreeView_*`, `ImageList_*`, or `commctrl.h` hits found in the expanded scan | **No live comctl32 common-control surface found** | No comctl removal needed unless future scans add findings. |

## Adjacent Standard Win32 Native Controls

These are not comctl32 common controls, but they are still native Win32 controls that may matter for a later full native-control removal pass.

| Project | Surface | Files | Current role |
| --- | --- | --- | --- |
| `RedSalamander` | Preferences dialog shell buttons and hosts | `RedSalamander/RedSalamander.rc:454-463`, culture copies | Dialog template still owns OK/Cancel/Apply buttons and DxUi host placeholders. |
| `RedSalamander` | Connection credential prompt | `RedSalamander/RedSalamander.rc:474-487`, culture copies | Native modal prompt with `LTEXT`, `EDITTEXT`, and `PUSHBUTTON` controls. |
| `RedSalamander` | Compare Directories banner/options controls | `RedSalamander/CompareDirectoriesWindow.cpp:1472-1517`, `RedSalamander/CompareDirectoriesWindow.Options.cpp:1887-1913` | Standard `Static`, `Button`, and `Edit` controls plus a custom progress spinner window. The spinner is custom, not `msctls_progress32`. |
| `RedSalamander` | Folder command-line controls | `RedSalamander/FolderWindow.FileSystem.Navigation.Part.cpp:205-244` | Standard hidden `STATIC` label and `EDIT` control, with `DEFAULT_GUI_FONT` and `WM_SETFONT`; this is also an unallowlisted broad Win32 audit finding. |
| `RedSalamander` | Manage Plugins configuration fallback controls | `RedSalamander/ManagePluginsDialog.cpp:3204-3569` | Standard `Static`, `Edit`, `Button`, and `ComboBox` controls for schema-driven fallback editing. |
| `RedSalamander` | Plugin config resource fallback | `RedSalamander/PluginManagerResources.rc:5-12` | Native dialog template with `Static` placeholder and OK/Cancel buttons. |
| `RedSalamanderMonitor` | Find panel | `RedSalamanderMonitor/ColorTextView.cpp:4942-4972` | Standard `STATIC`, `EDIT`, `BUTTON`, and `COMBOBOX` controls. |
| `RedSalamanderMonitor` | About/message/confirm dialogs | `RedSalamanderMonitor/RedSalamanderMonitor.rc:145-173`, culture copies | Native resource dialogs with `LTEXT`, `EDIT`, and buttons. |
| `RedConfigure` | Main UI | `RedConfigure/Main.cpp:87`, `RedConfigure/RedConfigureRoot.cpp` | Top-level Win32 window with DxUi content. No standard child controls were found in the searched app source/resources. |
| `Plugins/ViewerText` | File combo DxUi host | `Plugins/ViewerText/ViewerText.cpp:4758-4762` | Standard `Static` child window hosting DxUi file combo chrome. |
| `Plugins/ViewerWeb` | File combo DxUi host | `Plugins/ViewerWeb/ViewerWeb.cpp:1580-1584` | Standard `Static` child window hosting DxUi file combo chrome. |
| `Plugins/ViewerPE` | File combo DxUi host | `Plugins/ViewerPE/ViewerPE.cpp:893-897` | Standard `Static` child window hosting DxUi file combo chrome. |
| `Plugins/ViewerImgRaw` | File combo DxUi host | `Plugins/ViewerImgRaw/ViewerImgRaw.cpp:1486-1500` | Standard `Static` child window hosting DxUi file combo chrome. |
| `Plugins/ViewerVLC` | Video/static host surfaces | `Plugins/ViewerVLC/ViewerVLC.cpp:2445-2446`, `Plugins/ViewerVLC/ViewerVLC.cpp:3063-3064` | Standard `Static` child windows for video fallback and test/debug wheel child; not comctl32. |

## Removal Checklist

### Task 1: Make the Audits Complete for the Requested Product Surface

**Files:**

- Modify: `Tools/Audit-ComctlReportSurfaces.ps1`
- Modify: `Tools/Audit-VisibleNativeSurfaces.ps1` if its scope should also include Monitor, RedConfigure, Common, and Plugins
- Modify: `Specs/UI/UI_VisibleComctlAudit.md`

- [x] Extend `Tools/Audit-ComctlReportSurfaces.ps1` source roots from `RedSalamander`, `Plugins` to at least `Common`, `Plugins`, `RedSalamander`, `RedSalamanderMonitor`, and `RedConfigure`.
- [x] Add explicit inventory rows for the current Shared Directories `SysListView32` report surface and the `ViewerSpace` `TOOLTIPS_CLASSW` surface, or intentionally fail until Task 2 and Task 3 remove them.
- [x] Add resource-template patterns for `"SysListView32"` and listview styles (`LVS_`) so `.rc` templates cannot bypass the audit.
- [x] Add tooltip patterns (`TOOLTIPS_CLASSW`, `TOOLINFOW`, `TTS_*`, `TTM_*`, `TTF_*`) so plugin tooltip usage cannot bypass the audit.
- [x] Keep guard-only references (`msctls_statusbar32`, `SysTreeView32`) classified separately from live controls.
- [x] Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\Tools\Audit-ComctlReportSurfaces.ps1
```

Expected closeout state: the audit reports no live visible comctl32 listview/report or tooltip surfaces.

### Task 2: Replace the Shared Directories `SysListView32`

**Files:**

- Modify: `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp`
- Modify: `RedSalamander/FolderWindow.FileSystem.cpp`
- Modify: `RedSalamander/FolderWindow.h`
- Modify: `RedSalamander/RedSalamander.rc`
- Modify: `RedSalamander/Lang/fr-FR/RedSalamander-fr-FR.rc`
- Modify: `RedSalamander/Lang/ja-JP/RedSalamander-ja-JP.rc`
- Test: `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`

- [x] Replace `IDD_SHARED_DIRECTORIES` with a DxUi-owned dialog/window or a custom HWND containing a `DxUi::Grid`.
- [x] Preserve the existing four columns: shared name, local path, type, and remark.
- [x] Preserve existing actions: Open Path, Manage, Close, double-click/Open activation, empty-state text, and access-denied/error text.
- [x] Delete the `SysListView32` resource line from all three language resources.
- [x] Remove listview helper functions: `ConfigureSharedDirectoriesListView`, `SetSharedDirectoriesListSubItem`, `InsertSharedDirectoriesListRow`, `GetSharedDirectoriesListSelectedIndex`, and `SelectSharedDirectoriesListRow`.
- [x] Replace `LVN_ITEMCHANGED`/`NM_DBLCLK` handling with DxUi selection and activation callbacks.
- [x] Update `DebugGetSharedDirectoriesDialogSnapshot`, `DebugSelectSharedDirectoriesDialogRow`, and `DebugInvokeSharedDirectoriesDialogOpenPath` to read the DxUi model state instead of the listview HWND.
- [x] Update `TestSharedDirectoriesShowsSyntheticRowsOpensPathsAndReportsAccessDenied` to assert that no visible `SysListView32` child exists while preserving the existing synthetic-row/open-path/access-denied coverage.
- [x] Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_settings_shared_directories --selftest-fail-fast --selftest-timeout-multiplier=2
```

If the exact self-test filter name differs, use the existing case containing `TestSharedDirectoriesShowsSyntheticRowsOpensPathsAndReportsAccessDenied`.

### Task 3: Replace the `ViewerSpace` Tooltip Common Control

**Files:**

- Modify: `Plugins/ViewerSpace/ViewerSpace.cpp`
- Modify: `Plugins/ViewerSpace/ViewerSpace.h`
- Test: focused ViewerSpace command/selftest coverage if available

- [x] Replace `_hTooltip` and `TOOLTIPS_CLASSW` creation with a plugin-owned Direct2D/DxUi tooltip layer or custom owner-drawn popup.
- [x] Preserve current behavior: node hover text, tracked positioning, immediate display timing, max width near 420 DIPs, theme colors, and hide-on-empty-node behavior.
- [x] Remove `TOOLINFOW`, `TTS_*`, `TTF_*`, and `TTM_*` usage.
- [x] Remove `#include <commctrl.h>` from `Plugins/ViewerSpace/ViewerSpace.cpp` if no other comctl symbol remains.
- [x] Run:

```powershell
.\build.ps1 -ProjectName ViewerSpace -Configuration Debug
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
```

### Task 4: Remove Stale Comctl Initialization, Headers, and Linkage

**Files:**

- Modify: `RedSalamander/NavigationView.cpp`
- Review/modify: `RedSalamander/NavigationView.*`
- Review/modify: `Common/DxUi/DxUi.WindowHost.cpp`
- Review/modify: `Common/DxUi/DxUi.TextInput.cpp`
- Review/modify: `Common/DxUi/DxUi.Grid.cpp`
- Review/modify: `RedSalamander/CompareDirectoriesWindow.Internal.h`
- Review/modify: `RedSalamander/FolderViewInternal.h`
- Review/modify: `RedSalamander/FolderWindowInternal.h`
- Review/modify: `RedSalamander/ManagePluginsDialog.cpp`
- Review/modify: `RedSalamander/Ui/AlertOverlayWindow.h`
- Review/modify: `Plugins/ViewerText/ViewerText.cpp`
- Review/modify: `Plugins/ViewerText/ViewerText.h`
- Review/modify: `Plugins/ViewerWeb/ViewerWeb.cpp`
- Review/modify: `Plugins/ViewerPE/ViewerPE.cpp`
- Review/modify: `Plugins/ViewerImgRaw/ViewerImgRaw.cpp`
- Review/modify: `Plugins/ViewerVLC/ViewerVLC.cpp`
- Review/modify: `Plugins/ViewerVLC/ViewerVLC.h`

- [x] Remove `NavigationView::OnCreate`'s `InitCommonControls()` call if no common-control tooltip/window creation remains on that path.
- [x] Remove `#include <commctrl.h>` from NavigationView files that no longer need comctl declarations.
- [x] Review `RedSalamander` stale-header candidates (`CompareDirectoriesWindow.Internal.h`, `FolderViewInternal.h`, `FolderWindowInternal.h`, `ManagePluginsDialog.cpp`, `Ui/AlertOverlayWindow.h`) and remove `commctrl.h` wherever a focused build proves it is not required.
- [x] Treat `NMHDR`/`WM_NOTIFY` routes as dependency cleanup, not visible-control cleanup. Do not replace them unless the intended end state is full `commctrl.h` contract removal.
- [x] Try removing `#include <commctrl.h>` from `Common/DxUi/DxUi.WindowHost.cpp`, `Common/DxUi/DxUi.TextInput.cpp`, and `Common/DxUi/DxUi.Grid.cpp`; keep only the includes that a build proves are still required.
- [x] Try removing `#pragma comment(lib, "comctl32.lib")` from `Common/DxUi/DxUi.WindowHost.cpp`; keep it only if a remaining source dependency requires comctl32.
- [x] Try removing `#include <commctrl.h>` and `#pragma comment(lib, "comctl32")` from `ViewerText`, `ViewerWeb`, `ViewerPE`, and `ViewerImgRaw`; keep only any include/link that a focused build proves is still required.
- [x] Review `ViewerText` and `ViewerVLC` `NMHDR` notification handlers separately from live common-control migration.
- [x] Run:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
.\build.ps1 -ProjectName RedConfigure -Configuration Debug
.\build.ps1 -ProjectName RedSalamanderMonitor -Configuration Debug
.\build.ps1 -ProjectName ViewerText -Configuration Debug
.\build.ps1 -ProjectName ViewerWeb -Configuration Debug
.\build.ps1 -ProjectName ViewerPE -Configuration Debug
.\build.ps1 -ProjectName ViewerImgRaw -Configuration Debug
```

### Task 5: Decide Whether Shell Image Lists Count as Removal Scope

**Files:**

- Review: `RedSalamander/IconCache.h`
- Review: `RedSalamander/IconCache.cpp`
- Update: `Specs/UI/UI_VisibleComctlAudit.md`

- [x] The selected goal is visible common-control removal; keep `IImageList`/`SHGetImageList` documented as non-visible shell icon interop.
- Full comctl32 dependency removal, including replacing shell image-list extraction, is explicitly out of scope for this closeout and remains a separate decision.
- [x] Record the decision in `Specs/UI/UI_VisibleComctlAudit.md` so future audits do not conflate visible controls with shell image extraction.

### Task 6: Keep Adjacent Standard Controls in a Separate Backlog

**Files:**

- Update: `Specs/UI/UI_VisibleNativeAudit.md`
- Update or create a later WIP plan for full standard-control removal

- [x] Do not block the comctl32 closeout on standard `Static`, `Edit`, `Button`, or `ComboBox` controls.
- [x] Track the command-line label/edit controls separately because the broad Win32 audit currently reports their `HFONT`/`WM_SETFONT` usage as unallowlisted.
- [x] Track Monitor's find panel and modal resource dialogs separately if the desired end state is "no native child controls" rather than "no comctl32 controls."
- [x] Track Manage Plugins fallback controls separately unless Task 2 expands into a broader native-control migration.
- [x] Track plugin standard `Static` host windows separately if the desired end state is "no native child controls" rather than "no comctl32 controls."

### Task 7: Closeout Verification

**Files:**

- Modify: `Specs/UI/UI_VisibleComctlAudit.md`
- Move this plan to: `Specs/Plans/Done/UI_Win32CommonControlsAudit_2026-05-13.md`

- [x] Run the complete product-surface common-control search:

```powershell
rg -n --glob '*.{cpp,h,rc,vcxproj,manifest}' -i "SysListView32|SysTreeView32|SysHeader32|SysTabControl32|SysLink|msctls_[A-Za-z0-9_]+|ToolbarWindow32|ReBarWindow32|STATUSCLASSNAME|TOOLBARCLASSNAME|REBARCLASSNAME|WC_LISTVIEW|WC_TREEVIEW|WC_HEADER|WC_TABCONTROL|WC_COMBOBOXEX|WC_LINK|WC_NATIVEFONTCTL|PROGRESS_CLASS|TRACKBAR_CLASS|UPDOWN_CLASS|DATETIMEPICK_CLASS|MONTHCAL_CLASS|HOTKEY_CLASS|ANIMATE_CLASS|TOOLTIPS_CLASS|TOOLINFO|TTS_|TTF_|ListView_|TreeView_|Header_|TabCtrl_|ImageList_|HIMAGELIST|LVS_|LVN_|LVIF_|TVS_|TVN_|TVIF_|TCS_|TCN_|HDN_|HDS_|PBM_|TB_|TBM_|TTM_|UDM_|DTM_|MCM_" Common Plugins RedSalamander RedSalamanderMonitor RedConfigure
```

- [x] Run the focused audit helper:

```powershell
powershell -ExecutionPolicy Bypass -File .\Tools\Audit-ComctlReportSurfaces.ps1
```

- [x] Run the broad native audit to ensure no new broad Win32 UI regressions were introduced:

```powershell
powershell -ExecutionPolicy Bypass -File .\Tools\Audit-RemainingWin32UiDependencies.ps1 -AsMarkdown
```

- [x] Run focused Shared Directories command coverage, focused ViewerSpace coverage, and any updated DxUi tests.
- [x] Archive audit output under `Specs/TestRuns/<machine>/Audit/<date>_win32_common_controls_audit_closeout/`.
- [x] Update `Specs/UI/UI_VisibleComctlAudit.md` with the final inventory and allowed non-visible dependencies.

## Closeout Result

- `RedSalamander` Shared Directories now uses a `DxUi::Grid` in `RedSalamander.SharedDirectoriesWindow`; the localized
  `IDD_SHARED_DIRECTORIES` templates and `SysListView32` report control were removed.
- `Plugins/ViewerSpace` now paints node hover details as an in-canvas Direct2D tooltip overlay; the `TOOLTIPS_CLASSW` window,
  `TOOLINFOW`, and `TTM_*` path were removed.
- `Tools/Audit-ComctlReportSurfaces.ps1` now scans `Common`, `Plugins`, `RedSalamander`, `RedSalamanderMonitor`, and
  `RedConfigure` for visible listview/tooltip common-control regressions.
- Stale `commctrl.h`/`comctl32` include and link references were removed from `Common/DxUi`, NavigationView split files,
  and the viewer plugins that did not need them.
- Remaining commctrl-adjacent references are non-visible: `RedSalamander/IconCache` shell image-list interop, `FolderWindow`
  plugin `NMHDR` notification shims, the app Common-Controls v6 manifests, and historical comments in Compare
  Directories / NavigationView specs.
- Closeout audit output is archived at `Specs/TestRuns/SINON/Audit/2026-05-13_win32_common_controls_audit_closeout/`.

## Current Conclusion

- No visible `SysListView32`, `WC_LISTVIEWW`, `ListView_*`, `LVS_*`, `LVN_*`, `TOOLTIPS_CLASSW`, `TOOLINFOW`, `TTM_*`,
  `TTN_*`, `TTS_*`, or `TTF_*` references remain in `Common`, `Plugins`, `RedSalamander`, `RedSalamanderMonitor`, or
  `RedConfigure`.
- `RedSalamanderMonitor` and `RedConfigure` still have no confirmed live comctl32 common-control surface in app source/resources.
- `RedSalamander/IconCache` remains the explicit non-visible shell-image-list dependency; replacing it is a separate
  full-comctl32-dependency decision, not a visible UI cleanup.
