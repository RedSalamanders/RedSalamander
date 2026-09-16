# File Operations UI review — revised 2026-09-16

The requested capture blockers are resolved. **Incoming and Existing always stack vertically** in the revised proposal and UI contract. The product redesign remains proposed, owned by [I26](../UI_FileOperationsReview_2026-09-16.md).

Start with the [revised proposal](proposal.md) and [filterable screenshot gallery](gallery.html).

- **437 native screenshots / 124 named scenarios:** 128 original English images plus 309 French captures on the two physical displays.
- **Three revised French concepts:** vertical comparison, short-window comparison, and progress with visible controls. The original English concepts are retained as superseded references.
- [Scenario inventory](scenario-inventory.md): complete image map, resolved capture gaps and qualification limits.
- [Capture manifest](capture-manifest.json): per-image dimensions/hashes, locale, display, DPI, source and build receipts.
- [Validation](validation.md): focused builds, native cases, interaction checks and document integrity.

## Reproduce with the test harness

Use the existing runner and an initialized, marked test sandbox. These opt-in cases stay outside the broad behavioral suite.

```powershell
$env:RSBuildEnableTests = 'true'
.\Tools\Run-AllTests.ps1 -Suite FileOps -CaseFilter FileOps_VisualGallery -Platform x64 -Configuration Debug -TestRoot C:\RedSalamander.Perf -TimeoutMultiplier 3
```

For real keyboard/pointer states, explicitly select the notified foreground lane:

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -CaseFilter FileOps_VisualGalleryInteraction -Platform x64 -Configuration Debug -TestRoot C:\RedSalamander.Perf -TimeoutMultiplier 3
```

The interaction lane warns before taking focus, holds the runner's desktop lease, checks foreground and pointer ownership, captures real Tab/Shift+Tab, hover and pressed states, releases the pointer outside the action, and restores cursor and foreground focus. Default captures do not activate windows. No Computer Use or desktop screenshot is used.

The runner prints its run directory. Preserve `artifacts/selftest/last_run/fileops/visual-gallery/` before the next run replaces disposable evidence. Scenario/build provenance accompanies the PNGs. The [fixture](../../../../RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.VisualGallery.cpp) and [adjacent-surface driver](../../../../RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.VisualGallerySurfaces.cpp) use the shared [capture helper](../../../../Tests/TestSupport/WindowScreenshot.h). See [Tests documentation](../../../../Tests/README.md).

## Evidence and scope

DISPLAY1 is 144 DPI / 150%; DISPLAY5 is 96 DPI / 100% with a short 720-pixel physical height. Both use genuine native window rendering. The interaction case moves the same HWND between them and back. No display scaling settings were changed.

The static French catalog includes live confirmations, speed validation, destination history, all conflict classes, Issues empty/populated/selected/scrolled states, and recovery menus. Actual screenshots use production rendering with synthetic presentation or diagnostic fixtures; they do not prove provider execution. Generated mockups cannot establish pixel fit or accessibility.

Capture success does not mean the present UI has no clipping. The proposal documents the observed failures and the required fixes. Full-suite, 200% DPI, Japanese/RTL, complete UIA/assistive-technology and performance/platform qualification are outside this focused review. The harness-only workflow and its authorized focus exception are persisted in both repositories' `AGENTS.md` files.
