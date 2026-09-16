# DxUi integration guide

RedSalamander links the public [DxUi library](https://github.com/RedSalamanders/DxUi)
as `DxUi.lib`. `Dependencies/DxUi.lock.json` selects an exact revision. Shared controls,
window hosting, typography, input and accessibility implementation belong in that repository.
RedSalamander owns its application windows, commands, settings, localization and plugin adapters.

The canonical [library documentation](https://github.com/RedSalamanders/DxUi/blob/main/docs/README.md)
contains control examples and the gallery. Read the documentation at the locked revision when
checking a shipped build; main may describe a later version. Product behavior is governed by
[UI_DxUiSharedGrid](../Specs/UI/UI_DxUiSharedGrid.md),
[UI_DxUiWinUIDesign](../Specs/UI/UI_DxUiWinUIDesign.md) and the owning window specification.

## Build and update

Root `build.ps1` restores the exact source and builds the archive into a consumer-owned output
root. For an IDE build, first run `Tools/Restore-DxUi.ps1 -Platform x64 -Configuration Debug`.
Every consumer imports `Build/RedSalamander.DxUi.props` and `.targets`; do not add library source
files to an application project or include private `src` headers.

Use public includes such as `<DxUi/DxUi.h>` and `<DxUi/Typography.h>` and the `DxUi` namespace.
The six profiles are Debug, Release and ASan Debug for x64 and ARM64. Matching compiler, SDK,
CRT and sanitizer settings are enforced by the dependency identity. ARM64 cross-builds establish
compilation only; native execution is separate evidence.

To adopt an update, change the exact lock on a branch, restore, build and run the product
regressions. Fix a library regression in DxUi, update the candidate pin and repeat. Roll back
by restoring the previous lock and rebuilding. Availability notices are advisory and never
edit the lock. The complete contract is in [Build_Toolchain](../Specs/Build/Build_Toolchain.md).

## Product adapters

| File | Product responsibility |
| --- | --- |
| `RedSalamander/DxUiThemePalette.h` | Convert application theme settings into the library palette. |
| `Common/ViewerDxUiTheme.h` | Convert plugin `ViewerTheme` into neutral DxUi colors. |
| `Common/ViewerFileComboHost.h` | Viewer file-selection combo behavior and callbacks. |
| `Common/ModalWindowShell.h` | Modal ownership, owner enable/restore and library message-loop delegation. |
| `Tests/ProductUiTests/` | Product helpers, palettes, viewer adapters and lifetime regressions. |
| `Tests/RedConfigureTests/` | Configuration-window behavior and localized retained controls. |

Use `RedSalamander::MakeThemePaletteFromViewerTheme` from `ViewerDxUiTheme.h` for a viewer;
plugin ABI types do not belong in the shared library. Supply product-owned strings through the
public control API and logical control tree, including inputs nested in controls such as TagPicker.

## Lifetime and rendering

Keep the `DxUi::WindowHost` alive for the HWND it attaches to. Forward the documented host
messages; disconnect retained models, delegates and callbacks before detaching the host destroys
its controls. Forward `WM_NCDESTROY` through any saved subclass procedure before discarding the
procedure state. Preferences tests exercise repeated close/open cycles and all page categories.

DxUi owns its native hosting and animation mechanisms. Application-specific scheduling remains
with its application owner. Do not reintroduce the retired diagnostics, window-message names or
application animation dispatcher as library dependencies. Each statically linked plugin owns its
module's registrations; do not pass DxUi C++ objects across the plugin ABI or unload live UI.

Native text transport supports large valid selections with allocation/Unicode checks and one
clipboard-open attempt. Embedded editing keeps its separate text-size ceiling. Product acceptance
includes focus, selection, clipboard, IME, UI Automation, DPI, localization and teardown behavior.
Visual resemblance alone does not qualify an update.

Keep layout, allocation and rasterization out of composition paths. Reuse bounded resources,
avoid hidden animation work, and compare before/after runs on the same fixture and machine.
Unstable baselines or unexplained regressions remain open findings.

## Validation and ownership

Run the canonical product entrypoint, for example:

```powershell
./Tools/Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Configuration Debug -Platform x64
```

The Full plan runs ProductUiTests and product integration suites against the pinned library.
Shared control suites and galleries run from the DxUi repository using its `test.ps1` and
`gallery.ps1`; the retired in-tree DxUiTests executable is not a product validation surface.
Library passes do not replace product tests. Required hardware/manual and native-platform checks
must be recorded separately from unavailable or skipped capabilities.

The [I19 plan](../Specs/Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md)
records completed local x64 Debug/Release adoption and the separate platform/release handoff. Old implementations and tests remain available through
Git history; there is no second editable shared-control tree in this checkout.
