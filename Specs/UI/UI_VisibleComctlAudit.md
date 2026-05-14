# Visible Comctl Report-Surface Audit

Last updated: 2026-05-13

## Scope

This document tracks visible Win32 common-control report/list and tooltip surfaces across `Common`, `Plugins`, `RedSalamander`,
`RedSalamanderMonitor`, and `RedConfigure`.

It is intentionally narrower than the broader visible-native cleanup backlog: this audit closes visible comctl32 listview/tooltip
surface inventory, not every standard USER32 child control or non-visible shell interop dependency.

Validation helper:

- `Tools/Audit-ComctlReportSurfaces.ps1`
- `Tools/Audit-RemainingWin32UiDependencies.ps1` for the broader HFONT/GDI/native-control cleanup inventory

## Current Inventory

| File | Surface | Bucket | Current path | Notes |
| --- | --- | --- | --- | --- |
| (none) | Visible comctl32 listview/report/tooltip surfaces | `done` | n/a | No `SysListView32`, `WC_LISTVIEWW`, `ListView_*`, `LVS_*`, `LVN_*`, `TOOLTIPS_CLASSW`, `TOOLINFOW`, `TTM_*`, `TTN_*`, `TTS_*`, or `TTF_*` references remain in the audited product surface. |

Shared Directories now uses a `DxUi::Grid` hosted by `RedSalamander.SharedDirectoriesWindow`; the localized `IDD_SHARED_DIRECTORIES`
resource templates were removed. ViewerSpace now paints hover details as a Direct2D overlay instead of creating a `TOOLTIPS_CLASSW`
tracking tooltip.

Non-visible shell icon extraction remains deliberately out of this visible-control audit. `RedSalamander/IconCache` still uses
`SHGetImageList` / `IImageList` for shell icon interop; treat that as a separate full-comctl32-dependency decision.
`FolderWindowInternal.h` no longer pulls `<commctrl.h>` for status-bar click plumbing; any remaining `NMHDR`-shaped plugin
notification handlers are not visible comctl32 listview/tooltip surfaces and belong to a later full dependency cleanup if desired.

## Current Conclusion

- There is no remaining audited visible `SysListView32` / `WC_LISTVIEWW` / tooltip common-control surface in `Common`, `Plugins`,
  `RedSalamander`, `RedSalamanderMonitor`, or `RedConfigure`.
- `Tools/Audit-ComctlReportSurfaces.ps1` is the narrow guard for visible listview/tooltip comctl regressions.
- `Tools/Audit-RemainingWin32UiDependencies.ps1` remains the broader backlog guard for HDC/HFONT/font-message/native-child-control
  cleanup and is expected to report categories outside this visible comctl scope.
- Standard USER32 controls such as `Static`, `Edit`, `Button`, and `ComboBox` are tracked outside this comctl audit.
- Closeout audit output is archived at
  `Specs/TestRuns/SINON/Audit/2026-05-13_win32_common_controls_audit_closeout/`.
