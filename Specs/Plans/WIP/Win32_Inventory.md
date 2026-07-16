# Win32 Inventory

Last updated: 2026-07-02

Status: WIP - narrow RAII/lifetime/DPI audit plan (2026-07-02 folder review). This document is no longer a codebase-wide Win32 inventory; it now tracks only the four live audit items in the checklist below. Evidence anchors were re-verified by an independent code-verification pass on 2026-07-02.

Original 2026-05-18 inventory snapshot removed 2026-07-02 (predates BatchRename/FileSystemMtp/WarpDrive; see git history of this file).

## Scope

This plan covers four bounded Win32 audits: owned-window manual destroy calls, manual `Release()` sites in `Common/DxUi/DxUi.Accessibility.cpp`, raw `CoInitializeEx`/`CoUninitialize` normalization, and `WM_DPICHANGED_AFTERPARENT` coverage for child-hosted windows.

The supporting counts are regex-assisted, not a semantic compiler analysis. They are useful for locating hot spots, but each item still needs code review before changing behavior.

## Central Win32 Infrastructure

- `Common/WindowMessages.h` is the custom message registry. Keep all `WM_APP`/`WM_USER` IDs there unless a component-local constant is intentionally private.
- `Common/Helpers.h` owns `PostMessagePayload(...)`, `TakeMessagePayload<T>(...)`, `InitPostedPayloadWindow(...)`, and `DrainPostedPayloadsForWindow(...)`.
- `Common/Win32CallbackHelpers.h` centralizes no-throw wrappers for `SetWindowLongPtrW`, `CallWindowProcW`, dialog creation, and WndProc hook install/restore.
- `Common/ViewerFileComboHost.h` standardizes viewer file-combo host subclassing.
- `Common/WindowSizing.h` and `Common/WindowBackdropPolicy.h` carry reusable DPI/min-track/backdrop helpers.
- `Common/DxUi/*` is the shared retained DirectX UI layer sitting on top of HWNDs, message routing, native menu interop, accessibility, native text input, and D2D/DWrite rendering.

## Existing Guardrails

- WIL RAII is widely used for handles and COM (`wil::unique_*`, `wil::com_ptr`, `wil::scope_exit`).
- Custom message IDs are centralized in `Common/WindowMessages.h`.
- Posted heap payloads mostly use `PostMessagePayload(...)` and `TakeMessagePayload<T>(...)`.
- `WM_DPICHANGED_AFTERPARENT` coverage grew organically from 6 files (2026-05-18: `FolderView`, `NavigationView`, `FunctionBar`, `FolderWindow.StatusBar`, `Ui/AlertOverlayWindow`, `FolderWindow.ItemProperties`) to 11 files (2026-07-02 folder review: added `BatchRenameWindow.cpp`, `FindFilesWindow.cpp`, `ViewerSpace.cpp`, `DxUi.Menu.cpp`, `DxUi.WindowHost.cpp`). See CHK-6.
- Production scan did not find a clear `std::thread(...).detach()` usage; `detach()` hits in production were COM pointer ownership transfers or out-params, while raw detached threads appeared only in selftest code (2026-05-18 scan; not re-verified in 2026-07-02 folder review).

## Live Audit Checklist

Anchors below refreshed by the 2026-07-02 code-verification pass.

- [ ] **CHK-1 (NEXT ACTION)** - Replace `DestroyWindow(_hWnd.get())` on owned `wil::unique_hwnd` members with `reset()`-based close (or explicit ownership transfer). Refreshed anchors:
  - `Plugins/ViewerText/ViewerText.cpp:592`
  - `Plugins/ViewerText/ViewerText.cpp:602`
  - `RedSalamander/RedSalamander.cpp:760`
  - `RedSalamander/RedSalamander.cpp:1152`

  This is the only fully specified, bounded item in the plan; do it first.
- [ ] **CHK-3** - Audit the 18 manual `->Release()` sites in `Common/DxUi/DxUi.Accessibility.cpp` (UI Automation provider plumbing). Some are likely COM-boundary return semantics, but owning raw COM members or avoidable manual ref-counting should move to `wil::com_ptr`. Treat each site as COM-boundary code, not assumed safe or unsafe without local review.
- [ ] **CHK-4** - Normalize the 27 raw `CoUninitialize()` calls across 18 files against the 10 existing `wil::CoInitializeEx` uses; keep explicit comments where manual COM lifetime is required. EXCLUDE the worker initializations in `FolderView.FileOps.cpp` and `FolderView.Enumeration.cpp` - those are owned by WarpDrive Tasks 4/9; do not create conflicting edits.
- [ ] **CHK-6** - Systematic `WM_DPICHANGED_AFTERPARENT` audit of child/custom windows with DPI-sensitive chrome that handle only `WM_DPICHANGED`; add `WM_DPICHANGED_AFTERPARENT` where the window is child-hosted and needs parent DPI transitions. Coverage grew organically from 6 to 11 files (see Existing Guardrails), but the systematic audit has still not been done.

### Dropped checklist items (2026-07-02 folder review)

CHK-2 (long WndProc/DialogProc switch structure), CHK-5 (`PostMessagePayload` teardown contract for touched senders/receivers), CHK-7 (localization placeholder rules for resource edits), and CHK-8 (focused selftests plus archived perf evidence for responsiveness-affecting refactors) were dropped from this plan: the policy is already codified in `.github/skills` (win32-wndproc, wil-raii, localization) and `AGENTS.md` - these are ongoing coding standards, not discrete deliverables.

## Verification Notes

No build was run for this plan revision because the change is documentation only. Minimum closeout verification for future code changes from this plan should include:

- `.\build.ps1 -ProjectName RedSalamander`
- Targeted viewer/plugin/test project builds for the touched area.
- Relevant command selftest case(s) from `RedSalamander/SelfTest/Commands`.
- `git diff --check`

## Closeout Criteria

This WIP can move to `Specs/Plans/Done/` after CHK-1, CHK-3, CHK-4, and CHK-6 are either completed or split into narrower WIP plans, and any durable Win32 contract changes are merged into `AGENTS.md`, `.github/skills/win32-wndproc/SKILL.md`, `.github/skills/wil-raii/SKILL.md`, or the relevant `Specs/<Domain>/` document.
