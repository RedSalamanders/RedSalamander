# File Operations UI review — revised 2026-10-01

Latest published reviewed captures: [filterable V55 gallery](implementation/2026-10-01-v55-final/gallery.html)
and its [acceptance scope](implementation/2026-10-01-v55-final/README.md),
with **562 static and ten protected-interaction captures** at physical 96/144 DPI. All **512 unique
image hashes** have retained visual-review evidence, including sixteen newly inspected images.
The long Unicode/UNC path detail now keeps its caret inside the viewport and exposes complete
copy/close actions. Same-window DPI, keyboard, hover and pressed checks pass.

V65 x64 build status: passed, with zero warnings or errors. Release Full/Fresh `20261001T115346Z-104056-8e2cfb6a0b8049218c622dd915a2697e` records
2211 passed, 0 failed, 46 skipped (2257 total); the [V65 packet](../../../TestRuns/4cb089111a23/Validation/2026-10-01_i26-v65-final/README.md)
and [I26 checkpoint](../../Done/UI_FileOperationsReview_2026-09-16.md) hold the bound evidence and scope. The independent
A78/B19 resource review records acceptance for I26 with stated limits: the final matched Embedded presentation and logical-resource workload, with all eight processes and outliers retained; the earlier ProgramDatabase pair stays separate, without whole-product or driver-memory claims.
This gallery retains original V55 capture provenance; it does not claim native 192-DPI or external AT.
Dated slice notes and earlier failures remain historical evidence. Consumer ARM64/ASan, I19 and H4
retain separate owners.

Earlier September 23 Release slices: [readable, scrollable confirmations](implementation/2026-09-23-confirmation-layout-release/README.md)
and [compact waiting cards plus separate Issues facts](implementation/2026-09-23-waiting-and-issues-release/README.md).
The latter 424-frame harness run passes at native 96/144 DPI, as do 48 same-task transition witnesses,
36 waiting-layout checks and fresh Release PR validation. The packet records exact source/build
hashes and visual-review scope. [Fresh Debug qualification](implementation/2026-09-23-current-ui-debug/README.md)
also passes the 424-frame gallery and PR suite. All body pixels match Release; only the Debug title-bar
icon differs. The [speed-selector correction](implementation/2026-09-23-speed-selector/README.md) now
passes paired Release and identical-source Debug galleries: 438 captures each, all 136 caption
checks and 12 same-window changes. [Readable terminal results and completed groups](implementation/2026-09-23-terminal-group/README.md)
now pass 494-frame Release/Debug galleries and fresh PR validation in both profiles. All 28 group
caption checks, 56 result-summary checks and 12 same-window summary changes pass; the packet retains
the failing baseline and 36 focused originals. Final Full, native interaction/AT, resource and combined library-pin acceptance
remain open. Older measurements and failed attempts below remain historical.

The requested capture blockers are resolved. **Incoming and Existing always stack vertically.**
All six product decisions are accepted; [I26](../../Done/UI_FileOperationsReview_2026-09-16.md) now contains
the ACTIVE implementation and qualification plan. Throughput starts expanded
and remembers collapse; mockup 06 illustrates the collapsed state after a user choice. Conflict,
accepted decision and confirmed result receive distinct text, icons and semantic colors.

Start with the [revised proposal](proposal.md) and [filterable screenshot gallery](gallery.html).

The [September 20 interaction review](interaction-review-2026-09-20/README.md) incorporates the user's
reported clipping, repeated waiting/name text and missing pending-item inspection. It adds 309 fresh
French harness captures at native 144/96 DPI, a 32-step interaction inventory, and a
[responsive proposal](mockups/07-conflict-single-name-french.html) with one filename and stacked folders.
The original gallery below remains historical evidence. Passing capture cases do not accept the UI.

- **437 native screenshots / 124 named scenarios:** 128 original English images plus 309 French captures on the two physical displays.
- **Three revised French concepts:** vertical comparison, short-window comparison, and progress with visible controls. The original English concepts are retained as superseded references.
- [Scenario inventory](scenario-inventory.md): complete image map, resolved capture gaps and qualification limits.
- [Capture manifest](capture-manifest.json): per-image dimensions/hashes, locale, display, DPI, source and build receipts.
- [Validation](validation.md): focused builds, native cases, interaction checks and document integrity.

## Reproduce with the test harness

Implementation update, 2026-09-19: [ordinary action-row repair and 300 fresh French captures](implementation/2026-09-19-ordinary-actions/README.md).
The subsequent [four direct conflict actions / Remplacer slice](implementation/2026-09-19-conflict-actions/README.md)
adds focused policy/native integration evidence and 300 further French captures. Narrow labels still
clip: measured wrapping and the broader redesign remain unfinished.
The [throughput disclosure slice](implementation/2026-09-19-throughput-disclosure/README.md) adds
per-card collapse and persistence with focused Debug/Release state checks. Its Release captures
are readable; Debug loses the legacy drawing layer and is explicitly rejected as visual acceptance.
The subsequent [capture validation slice](implementation/2026-09-19-throughput-capture-validation/README.md)
adds mandatory pixel witnesses for that fixture and eight readable Debug/Release captures, including
the corrected destination path. The intermittent original failure is retained; layout gates stay open.
Both displays reported 144 DPI during that historical run; the September 23 captures again exercise
native 96/144 DPI. Neither run alone completes the redesign. The September 16 gallery below remains the original review baseline.

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

The runner prints its run directory. Preserve `artifacts/selftest/last_run/fileops/visual-gallery/` before the next run replaces disposable evidence. Scenario/build provenance accompanies the PNGs. The [fixture](../../../../RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.VisualGallery.cpp) and [adjacent-surface driver](../../../../RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.VisualGallerySurfaces.cpp) use the shared [capture helper](../../../../Tests/TestSupport/TestSupport.Tests.WindowScreenshot.h). See [Tests documentation](../../../../Tests/README.md).

## Evidence and scope

In the September 16 evidence, DISPLAY1 is 144 DPI / 150%; DISPLAY5 is 96 DPI / 100% with a short
720-pixel physical height. Both use genuine native window rendering. The historical interaction case
moves the same HWND between them and back. No display scaling settings were changed. September 19's
implementation packs report both displays at 144 DPI; fresh mixed-DPI acceptance remains open.

The static French catalog includes live confirmations, speed validation, destination history, all conflict classes, Issues empty/populated/selected/scrolled states, and recovery menus. Actual screenshots use production rendering with synthetic presentation or diagnostic fixtures; they do not prove provider execution. Generated mockups cannot establish pixel fit or accessibility.

Capture success does not mean the present UI has no clipping. The proposal documents the observed failures and the required fixes. Full-suite, 200% DPI, Japanese/RTL, complete UIA/assistive-technology and performance/platform qualification are outside this focused review. The harness-only workflow and its authorized focus exception are persisted in both repositories' `AGENTS.md` files.
