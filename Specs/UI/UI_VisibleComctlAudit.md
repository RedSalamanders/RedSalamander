# Visible Comctl Report-Surface Audit

Last updated: 2026-04-26

## Scope

This document tracks the remaining `SysListView32` / `WC_LISTVIEWW` report-surface references in the repo.

It is intentionally narrower than the broader visible-native cleanup backlog: this audit closes the report-surface inventory item, not the full native-surface retirement plan.

Validation helper:

- `Tools/Audit-ComctlReportSurfaces.ps1`
- `Tools/Audit-RemainingWin32UiDependencies.ps1` for the broader HFONT/GDI/native-control cleanup inventory

## Current Inventory

| File | Surface | Bucket | Current path | Notes |
| --- | --- | --- | --- | --- |
| (none) | (Connection Manager fully migrated to single-canvas DxUi `ConnectionManagerWindow.{h,cpp}`; the legacy `IDD_CONNECTION_MANAGER` template was deleted in Phase 12 of `Specs/Plans/Done/UI_ConnectionManagerSingleCanvasPlan.md`) | `done` | n/a | No `SysListView32` or other comctl class is created by the Connection Manager path; the live list is a `DxUi::Grid` inside the single canvas. |

Connection Manager has no live `IDD_CONNECTION_MANAGER` dialog template path and no visible legacy comctl owner-draw controls. The single DxUi window host is responsible for UIA, keyboard traversal, pointer interaction, and theme repaint behavior.

## Current Conclusion

- There is no remaining audited `SysListView32` / `WC_LISTVIEWW` surface that is a required visible native exception on the validated live DX paths.
- The only remaining reference is legacy dialog-resource fallback scaffolding, not accepted visible end-state UI.
- The former Preferences `Keyboard`, `Plugins`, `Themes`, and `Viewers` listview fallback tokens are gone from source; their live paths are DX-owned and no longer appear in this report-surface inventory.
- The broader visible-native cleanup is closed by `Specs/Plans/Done/UI_RemainingWin32UiDependencyRetirementPlan.md`; its first automated baseline is archived at `Specs/TestRuns/4cb089111a23/Audit/2026-04-25_182415_remaining_win32_ui_baseline/`.
- The latest same-machine report-surface audit is archived with the broad and narrow native audits at `Specs/TestRuns/ac3bbb87f7dd/Audit/2026-04-26_220000_dxui_remaining_native_audit_closeout/`.
