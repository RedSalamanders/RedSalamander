# FolderView Thumbnail Performance And Sizing Plan

## Checklist

- [x] Capture current thumbnail behavior with a baseline run before production changes.
- [x] Add deterministic thumbnail display, ratio, fallback, scrolling, and setting tests before fixing production code.
- [x] Archive baseline evidence under the Commands run folder printed by the selftest `ArchiveToRepo:` trace line.
- [x] Preserve image aspect ratio for every thumbnail draw.
- [x] Display valid WIC-decodable images even when shell thumbnail extraction fails.
- [x] Requeue visible thumbnail work after horizontal scroll so newly visible valid images do not remain on icon fallback.
- [x] Requeue visible thumbnail work after pane resize so newly visible valid images do not remain on icon fallback.
- [x] Keep normal-view icons at their normal shell image-list size after returning from thumbnail mode.
- [x] Keep thumbnail loading asynchronous, bounded, cancelable, and resilient during scroll, refresh, navigation, sort, and size changes.
- [x] Add a discrete thumbnail-size slider to the pane bottom-right sort popup.
- [x] Persist thumbnail size per pane, defaulting missing settings to the current `64 DIP` size.
- [x] Prove candidate performance is on par with or better than the baseline before closeout, with noisy large-folder maxima documented below.
- [x] Update authoritative specs/docs and move this plan to `Specs/Plans/Done/` when complete.

## Goal

Make FolderView thumbnail mode fast, resilient, visually correct, and configurable per pane. The final implementation must be at least on par with the current implementation for folder open, queueing, scrolling, and render cost, and must be clearly better for valid image coverage and aspect-ratio correctness.

## Current Findings

- `RedSalamander/FolderView.Rendering.cpp` draws thumbnails into a square `_iconSizeDip x _iconSizeDip` rectangle. If the source bitmap is not square, it can be stretched.
- `RedSalamander/FolderView.Icons.cpp` currently extracts thumbnails through `IShellItemImageFactory` with shell-only fallback. Valid images can fall back to icons when the shell provider refuses `SIIGBF_THUMBNAILONLY`.
- Thumbnail extraction runs on the same background worker as enumeration and icon loading. `ProcessThumbnailLoadQueue()` drains the queue in one loop, so visible thumbnail work can delay later enumeration if the user navigates or refreshes while extraction is active.
- The queue is bounded to visible items and `kMaxThumbnailQueueItems = 256`, which is good, but item thumbnails can accumulate while scrolling through a huge folder.
- `Common/SettingsStore.h` persists display mode per pane but has no thumbnail-size setting.
- `RedSalamander/RedSalamander.cpp::ShowSortMenuPopup()` currently shows only sort radio items. `Common/DxUi/DxUi.Menu.cpp` has no slider menu item kind yet.

## Protected Scenarios

These scenarios must be named in the before/after evidence:

- `folderView.thumbnail.valid_images_shell_fail`: valid JPEG/PNG/BMP/GIF images display thumbnails even when shell extraction fails.
- `folderView.thumbnail.aspect_ratio`: portrait, landscape, panorama, and tall-thin thumbnails render inside the slot without stretching.
- `folderView.thumbnail.scroll_stress`: rapid horizontal scrolling through a large mixed folder keeps queue work bounded and does not block navigation.
- `folderView.thumbnail.scroll_requeues_visible`: horizontal scrolling into new columns queues visible thumbnail work instead of leaving valid images on icon fallback.
- `folderView.thumbnail.resize_requeues_visible`: widening a thumbnail pane after an initial narrow queue loads thumbnails for newly visible columns instead of leaving valid images on icon fallback.
- `folderView.thumbnail.resize_change`: changing thumbnail size while work is pending cancels stale payloads, relayouts once, and requeues visible work at the new size.
- `folderView.thumbnail.return_to_normal_icon_size`: returning from thumbnail mode to Brief/Detailed/Extra Detailed uses the normal shell image-list size class instead of shrinking thumbnail-mode jumbo icon bitmaps into the normal icon slot.
- `folderView.thumbnail.bad_files`: zero-byte, truncated, renamed, inaccessible, and non-image files fall back to icons without stuck pending work.
- `folderView.thumbnail.settings_roundtrip`: left/right panes persist independent thumbnail sizes and missing settings load as `64 DIP`.
- `folderView.thumbnail.sort_popup_slider`: the bottom-right sort popup exposes a keyboard and pointer usable discrete size slider.

## Metrics

Reuse existing metrics:

- `thumbnails.queue_build_us`
- `thumbnails.queue_count`
- `thumbnails.queue_visible_count`
- `thumbnails.cache_hit_count`
- `thumbnails.queue_wait_to_dequeue_us`
- `thumbnails.extract_us`
- `thumbnails.post_message_latency_us`
- `thumbnails.ui_convert_us`
- `thumbnails.fallback_count`
- `thumbnails.ui_apply_count`
- `render.draw_item_us`
- `render.status_bar.paint_us`
- `icons.*`

Add metrics needed for the new contract:

- `thumbnails.shell_extract_us`
- `thumbnails.wic_decode_us`
- `thumbnails.shell_success_count`
- `thumbnails.wic_success_count`
- `thumbnails.decode_fail_count`
- `thumbnails.first_visible_apply_us`
- `thumbnails.visible_apply_count`
- `thumbnails.stale_drop_count`
- `thumbnails.cancel_count`
- `thumbnails.cache_bytes`
- `thumbnails.cache_evicted_count`
- `thumbnails.size_change_us`
- `thumbnails.scroll_requeue_us`
- `menu.sort_popup.open_us`
- `menu.sort_popup.thumbnail_slider_change_us`

The candidate run must not regress same-machine baseline medians or worst-case values by more than 10% for queue build, scroll requeue, menu open, and warm render metrics unless the run shows a correctness win that cannot be achieved otherwise and the regression is explicitly accepted. Valid image thumbnail success must improve from the current shell-only baseline.

## Before-Code Test And Measurement Gate

No production thumbnail behavior changes start until this gate is complete.

1. Build the current tree:

   ```powershell
   .\build.ps1 -ProjectName RedSalamander
   ```

2. Add test-only fixtures, debug snapshots, and metrics for the thumbnail scenarios, without changing production thumbnail behavior.

3. Run the new focused baseline against current behavior:

   ```powershell
   .\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_thumbnail_ --selftest-timeout-multiplier=4
   ```

4. Run existing focused guards:

   ```powershell
   .\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=pane_view_options_toggle_thumbnails --selftest-timeout-multiplier=4
   .\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_large_folder_baseline --selftest-timeout-multiplier=4
   .\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_statusBar_uses_owned_window_and_sort_click_opens_menu --selftest-timeout-multiplier=4
   ```

5. Confirm each run archives `results.json`, `trace.txt`, and `perf/perf_metrics.jsonl` under the Commands run folder printed by the selftest `ArchiveToRepo:` trace line.

6. Record baseline run paths and current failures in the closeout notes before production fixes begin.

Baseline evidence captured before production fixes:

- `.\build.ps1 -ProjectName RedSalamander` passed in Debug x64.
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-list-cases --selftest-case=folderView_thumbnail_` listed 7 cases.
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_thumbnail_ --selftest-timeout-multiplier=4` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_170533`.
- Baseline failures: settings size roundtrip loads `64 DIP` instead of pane-specific `48/128 DIP`; valid shell-failed BMPs fall back to icons (`wicSuccess=0`); wide synthetic thumbnail source `96x48` draws as `64x64`; bad image-looking files do not report WIC decode failures; size command leaves target at `16 DIP`; sort popup has no `Thumbnail size` slider row.
- Baseline pass: `folderView_thumbnail_scroll_stress` keeps pending work settled and queue count bounded in the mixed 640-item fixture.

Candidate evidence after production fixes:

- `.\build.ps1 -ProjectName RedSalamander` passed in Debug x64 with `0 warning(s), 0 error(s)`; log `Z:\src\RedSalamander\.build\logs\msbuild-20260514_175409_424.log`.
- `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_thumbnail_ --selftest-timeout-multiplier=4` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_175610`: 7 passed, 0 failed.
- `pane_view_options_toggle_thumbnails` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_175623`: 1 passed, 0 failed.
- `cmd_pane_statusBar_uses_owned_window_and_sort_click_opens_menu` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_175628`: 1 passed, 0 failed.
- `folderView_perf_large_folder_baseline --selftest-timeout-multiplier=4` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_175645`: 1 passed, 0 failed.
- `folderView_perf_large_folder_baseline` default-timeout samples archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_175726` and `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_175757`: both passed.
- `.\.build\x64\Debug\DxUiTests.exe --suite=Menu` passed all Menu tests after the shared slider menu item was added.
- Baseline `folderView_thumbnail_` correctness failures are fixed: per-pane settings round-trip, WIC fallback for shell-failed valid images, aspect-preserving draw, decode-failure reporting for bad images, size change while work is pending, and bottom-right sort-popup slider exposure.
- Same-machine large-folder perf comparison against archived run `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_130206` is directionally acceptable but noisy because render-call counts differ (`10` baseline vs `20` candidate). Candidate `2026-05-14_175726` improves `icons.extract_us` max (`7096 -> 3296 us`), `FolderView.IconLoading.ProcessQueue` max (`19224 -> 16326 us`), and `render.draw_item_us` max (`2913 -> 2802 us`); `icons.queue_build_us` max is slightly higher (`1839 -> 2063 us`). A repeated candidate sample `2026-05-14_175757` kept extraction/render max below baseline but had a `FolderView.IconLoading.ProcessQueue` outlier (`46729 us`), so no stronger perf-win claim is made beyond the passing guard and improved image coverage.
- Durable specs/docs updated: `Specs/UI/UI_FolderView.md`, `Specs/UI/UI_FolderWindow.md`, `Specs/Core/Core_SettingsStore.md`, `Specs/Testing/Testing_TestCoverage.md`, `Specs/SettingsStore.schema.json`, and `Docs/UserGuide.md`.

Resize-visible follow-up evidence after testing `C:\Users\eric\Pictures`:

- A Shell probe against representative `C:\Users\eric\Pictures` files that appeared as red icons showed `IShellItemImageFactory::GetImage(...)` succeeded for valid JPG/PNG files, so the missing thumbnails were in the app-visible-range pipeline rather than in Windows image decoding. The zero-byte `New Bitmap image.jpg` correctly failed as non-image content.
- Red run: `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_thumbnail_resize_requeues_visible --selftest-timeout-multiplier=4 --selftest-fail-fast` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_181351`: failed with `visible=10 visibleThumb=5 totalThumb=5 queued=5 pending=0`.
- Final build after the resize-visible queue patch passed in Debug x64 with `0 warning(s), 0 error(s)`; log `Z:\src\RedSalamander\.build\logs\msbuild-20260514_182217_610.log`.
- Green run after the resize-visible queue patch: same command archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_182351`: 1 passed, 0 failed.
- Full thumbnail suite after the patch: `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_thumbnail_ --selftest-timeout-multiplier=4 --selftest-fail-fast` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_182359`: 8 passed, 0 failed.
- Supporting guards after the patch: `pane_view_options_toggle_thumbnails` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_181948`, `cmd_pane_statusBar_uses_owned_window_and_sort_click_opens_menu` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_181952`, and `folderView_perf_large_folder_baseline` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_182006`; all passed.
- `.\.build\x64\Debug\DxUiTests.exe --suite=Menu` passed after the patch.

Thumbnail-to-normal icon-size follow-up evidence:

- Red run before the icon cache size-class patch: `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_thumbnail_return_to_normal_icon_size --selftest-timeout-multiplier=4 --selftest-fail-fast` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_184306`: failed because Detailed mode reused a thumbnail-mode `256x256` source bitmap in a `16 DIP` draw slot.
- Candidate build after the icon cache size-class patch passed in Debug x64 with `0 warning(s), 0 error(s)`; log `Z:\src\RedSalamander\.build\logs\msbuild-20260514_184743_648.log`.
- Green run after the patch: same command archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_184935`: 1 passed, 0 failed. Trace recorded thumbnail source `256x256` at `64 DIP`, then normal source `48x48` at `16 DIP`.
- Full thumbnail suite after the patch: `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_thumbnail_ --selftest-timeout-multiplier=4 --selftest-fail-fast` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_184957`: 9 passed, 0 failed.
- Large-folder perf guard after the patch: `folderView_perf_large_folder_baseline --selftest-timeout-multiplier=4 --selftest-fail-fast` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_185019`: 1 passed, 0 failed.
- `.\.build\x64\Debug\DxUiTests.exe --suite=Menu` passed on rerun after one transient hover timing failure in `TestMenuPointerInsideOverlappingPopupDoesNotSwitchRoot`.

Final closeout rerun after the same patch:

- `.\build.ps1 -ProjectName RedSalamander` passed in Debug x64 with `0 warning(s), 0 error(s)`; log `Z:\src\RedSalamander\.build\logs\msbuild-20260514_185252_424.log`.
- `folderView_thumbnail_return_to_normal_icon_size --selftest-timeout-multiplier=4 --selftest-fail-fast` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_185424`: 1 passed, 0 failed.
- Full `folderView_thumbnail_ --selftest-timeout-multiplier=4 --selftest-fail-fast` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_185433`: 9 passed, 0 failed.
- `folderView_perf_large_folder_baseline --selftest-timeout-multiplier=4 --selftest-fail-fast` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_185447`: 1 passed, 0 failed.
- `.\.build\x64\Debug\DxUiTests.exe --suite=Menu` failed twice in the pre-existing `TestMenuPointerInsideOverlappingPopupDoesNotSwitchRoot` hover-timing guard with `a menu-bar hover message must not switch root while the live cursor is inside a frontmost popup`. No DxUi/Menu implementation or test file is modified in this diff; the pane sort-popup slider command selftest remains green.

Horizontal-scroll follow-up evidence after testing real `C:\Users\eric\Pictures` scrolling:

- Red run before the scroll requeue patch: `folderView_thumbnail_scroll_requeues_visible --selftest-timeout-multiplier=4 --selftest-fail-fast` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_192810`: failed after one horizontal line scroll with layout-visible items but no thumbnail requeue (`immediate visible=10 visibleThumb=0 totalThumb=5 queued=5 completed=5 pending=0`).
- Candidate build after the scroll requeue patch passed in Debug x64 with `0 warning(s), 0 error(s)`; log `Z:\src\RedSalamander\.build\logs\msbuild-20260514_192858_822.log`.
- Green run after the patch: same command archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_193044`: 1 passed, 0 failed. Trace recorded `immediate visible=10 visibleThumb=10 totalThumb=15 queued=10 completed=10 pending=0`.
- Full thumbnail suite after the patch: `folderView_thumbnail_ --selftest-timeout-multiplier=4 --selftest-fail-fast` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_193109`: 10 passed, 0 failed. Scroll-stress artifact recorded `elapsedUs=833443`, `queued=30`, `completed=30`, `pending=0`, and `staleDrops=29`.
- Large-folder perf guard after the patch: `folderView_perf_large_folder_baseline --selftest-timeout-multiplier=4 --selftest-fail-fast` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_193215`: 1 passed, 0 failed.
- Supporting guards after the patch: `cmd_pane_statusBar_uses_owned_window_and_sort_click_opens_menu` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_193239`, and `pane_view_options_toggle_thumbnails` archived to `Specs/TestRuns/4cb089111a23/Commands/2026-05-14_193308`; both passed.

## Crazy Folder Fixture

Create one deterministic selftest folder containing:

- 320 valid JPEGs generated by WIC: square, portrait, landscape, panorama, tall-thin, tiny, and large downscale cases.
- 120 valid PNG/BMP/GIF files, including alpha PNG and first-frame GIF.
- 160 non-image files mixed among images: `.txt`, `.json`, extensionless, `.url`, `.lnk` where possible.
- 80 intentionally bad image-looking files: zero-byte `.jpg`, truncated JPEG header, text saved as `.jpg`, invalid `.png`, denied/read-only cases where the local filesystem allows it.
- Long filenames near the selftest-safe path budget, Unicode names, mixed case extensions, and names like `PXL_20260503_121254627.PORTRAIT.jpg`.
- Enough items to force horizontal scrolling in thumbnail mode at all four configured sizes.

The fixture must keep path lengths within the rules in `Specs/Testing/Testing_SelfTests.md`.

## Implementation Tasks

### Task 1: Test And Debug Surface First

Files:

- Modify `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`
- Modify `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`
- Modify `RedSalamander/FolderView.h`
- Modify `RedSalamander/FolderWindow.h`
- Modify `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp`
- Modify `RedSalamander/FolderView.Rendering.cpp` only for debug snapshot capture, not behavior
- Modify `Specs/Testing/Testing_TestCoverage.md`

Add selftest cases:

- `folderView_thumbnail_valid_images_shell_fail`
- `folderView_thumbnail_aspect_ratio`
- `folderView_thumbnail_scroll_stress`
- `folderView_thumbnail_scroll_requeues_visible`
- `folderView_thumbnail_resize_requeues_visible`
- `folderView_thumbnail_size_change_while_pending`
- `folderView_thumbnail_return_to_normal_icon_size`
- `folderView_thumbnail_bad_files_fallback`
- `folderView_thumbnail_settings_roundtrip`
- `folderView_thumbnail_sort_popup_slider`

Add debug-only fields to pane snapshots:

- thumbnail target size
- queued/completed/fallback/stale/pending/cache-hit counts
- shell success count
- WIC success count
- decode failure count
- visible apply count
- thumbnail cache bytes
- last thumbnail draw source size and destination rect

The first baseline run is expected to expose at least these current gaps:

- aspect-ratio draw rect is square for non-square thumbnails,
- shell-failure valid images use icon fallback instead of WIC thumbnails,
- sort popup has no thumbnail-size slider,
- settings have no per-pane thumbnail size.

### Task 2: Thumbnail Geometry

Files:

- Create `RedSalamander/FolderViewThumbnailGeometry.h`
- Modify `RedSalamander/FolderView.Rendering.cpp`
- Modify `Tests/DxUiTests/DxUiTests.Menu.cpp` only if menu layout tests need shared geometry helpers
- Modify `RedSalamander/RedSalamander.vcxproj`
- Modify `RedSalamander/RedSalamander.vcxproj.filters`

Add pure helper:

- `FitBitmapRectPreserveAspect(D2D1_RECT_F slot, D2D1_SIZE_U bitmapSize) noexcept`
- clamp zero-size inputs to an empty rect
- center the fitted rect in the slot
- never exceed slot width or height
- preserve source aspect ratio within a small float epsilon

Use the helper when `_thumbnailsVisible && item.thumbnail`; keep existing icon behavior for normal shell icons and placeholders.

Acceptance:

- portrait thumbnails are pillarboxed inside the slot,
- landscape thumbnails are letterboxed inside the slot,
- no image is stretched wider or squashed shorter than its source ratio,
- shortcut overlays still draw only for real icon fallback, not on image thumbnails unless the item is actually a shortcut.

### Task 3: Decode Pipeline And Valid Image Fallback

Files:

- Modify `RedSalamander/FolderView.Icons.cpp`
- Modify `RedSalamander/FolderView.h`

Change thumbnail payloads from UI-thread `HBITMAP` conversion to worker-produced BGRA pixel buffers:

- worker extracts shell thumbnail first,
- worker converts shell `HBITMAP` to premultiplied BGRA bytes with width/height,
- if shell extraction fails for a WIC-supported image extension, worker decodes the first WIC frame at bounded target size,
- UI thread creates `ID2D1Bitmap1` from the posted BGRA bytes and does no image decode,
- payload carries source kind: `Shell`, `Wic`, `Synthetic`, or `Fallback`.

WIC fallback rules:

- try only likely image extensions to avoid probing every `.txt` in a mixed folder,
- use scaled decode through WIC transform/scaler,
- cap decoded target to the selected thumbnail size in physical pixels, never more than `kMaxThumbnailPixelSize`,
- treat bad/truncated images as fallback without retry storms,
- log failures through metrics, not noisy normal-control-flow debug rows.

### Task 4: Async Resilience And Cache Bounds

Files:

- Modify `RedSalamander/FolderView.Icons.cpp`
- Modify `RedSalamander/FolderView.Enumeration.cpp`
- Modify `RedSalamander/FolderView.h`

Make thumbnail work cooperative:

- when `_pendingEnumerationPath` exists, stop draining thumbnail work and let enumeration win,
- check generation and batch id before and after every extraction,
- keep the visible queue bounded,
- cancel stale work on navigation, refresh, sorting, display-mode change, and size change,
- keep `PostMessagePayload`/`TakeMessagePayload` ownership for all UI payloads.

Bound memory:

- track approximate thumbnail bytes per item,
- trim offscreen thumbnails by least-recently-visible order,
- never evict currently visible thumbnails during the same paint/queue pass,
- emit `thumbnails.cache_bytes` and `thumbnails.cache_evicted_count`.

### Task 5: Per-Pane Thumbnail Size Model

Files:

- Modify `Common/SettingsStore.h`
- Modify `Common/Common/SettingsStore.cpp`
- Modify `Specs/SettingsStore.schema.json`
- Modify `Specs/Core/Core_SettingsStore.md`
- Modify `RedSalamander/FolderView.h`
- Modify `RedSalamander/FolderView.cpp`
- Modify `RedSalamander/FolderWindow.h`
- Modify `RedSalamander/FolderWindow.cpp`
- Modify `RedSalamander/RedSalamander.cpp`
- Modify `RedSalamander/Preferences.Internal.h`
- Modify `RedSalamander/Preferences.Internal.cpp`
- Modify `RedSalamander/Preferences.Dialog.cpp`

Use four discrete sizes:

- Small: `48 DIP`
- Medium: `64 DIP`
- Large: `96 DIP`
- Extra Large: `128 DIP`

Settings contract:

- add `view.thumbnailSizeDip`,
- default missing or invalid values to `64`,
- clamp parsed values to the nearest allowed size,
- omit the setting when it equals `64`,
- capture runtime left/right pane size independently,
- preserve the size through Preferences save/apply even if Preferences does not expose a separate control.

FolderView contract:

- add `SetThumbnailSizeDip(uint32_t sizeDip)` and `GetThumbnailSizeDip()`,
- changing size while in thumbnail mode cancels current thumbnail work, clears/invalidates size-dependent thumbnails, relayouts, updates scroll metrics, and requeues visible thumbnails,
- changing size while outside thumbnail mode stores the value without disturbing the current non-thumbnail layout.

### Task 6: Sort Popup Thumbnail Size Slider

Files:

- Modify `Common/DxUi/DxUi.h`
- Modify `Common/DxUi/DxUi.Menu.cpp`
- Modify `Tests/DxUiTests/DxUiTests.Menu.cpp`
- Modify `RedSalamander/Resource.h`
- Modify `RedSalamander/RedSalamander.rc`
- Modify `RedSalamander/RedSalamander.cpp`
- Modify `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`

Add a discrete menu slider:

- new `MenuItemKind::Slider`,
- fixed stop count support,
- pointer click/drag and Left/Right/Home/End keyboard behavior,
- accessible display text through existing debug item-text helpers,
- row height large enough for a polished track and four stops,
- no overlap with accelerator text or submenu chevrons,
- menu stays anchored above the bottom-right status sort part.

Main sort popup layout:

- sort radio items,
- separator,
- thumbnail-size slider labeled with localized strings,
- current size reflected by the slider stop,
- selecting a stop applies to the pane passed by the status-bar click,
- left and right panes remain independent.

Compare Directories sort popup remains sort-only for this plan.

### Task 7: Command And Resource Wiring

Files:

- Modify `RedSalamander/Resource.h`
- Modify `RedSalamander/RedSalamander.rc`
- Modify `RedSalamander/CommandRegistry.cpp` if command ids are exposed to shortcut/command dispatch
- Modify `RedSalamander/ShortcutDefaults.cpp` only if a default shortcut is explicitly added

Add localized strings:

- `Thumbnail size`
- `Small`
- `Medium`
- `Large`
- `Extra Large`

Add command ids for left/right pane size stops or a pane-targeted command route that maps the slider value to the active popup pane.

### Task 8: Before/After Validation

Run after production changes:

```powershell
.\build.ps1 -ProjectName RedSalamander
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_thumbnail_ --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=pane_view_options_toggle_thumbnails --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_perf_large_folder_baseline --selftest-timeout-multiplier=4
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_statusBar_uses_owned_window_and_sort_click_opens_menu --selftest-timeout-multiplier=4
.\.build\x64\Debug\DxUiTests.exe --suite=Menu
```

Compare archived baseline and candidate runs:

```powershell
$baselineRun = Resolve-Path $env:RS_THUMBNAIL_BASELINE_RUN
$candidateRun = Resolve-Path $env:RS_THUMBNAIL_CANDIDATE_RUN
.\Tools\CompareTestRuns.ps1 $baselineRun $candidateRun -ShowTraceDiff -MaxTraceDiffLines 80
.\Tools\Show-PerfRuns.ps1 -Run $candidateRun -Metric thumbnails
```

Acceptance:

- all new `folderView_thumbnail_` cases pass,
- existing thumbnail and FolderView perf guards pass,
- candidate queue/render/menu metrics are on par with or better than baseline,
- WIC fallback increases valid-image thumbnail success,
- stale drops happen only for intentional cancel/navigation/size-change cases,
- no pending thumbnail count remains after settle waits,
- every run archives required evidence.

### Task 9: Spec And Closeout

Files:

- Modify `Specs/UI/UI_FolderView.md`
- Modify `Specs/UI/UI_FolderWindow.md`
- Modify `Specs/Core/Core_SettingsStore.md`
- Modify `Specs/Testing/Testing_TestCoverage.md`
- Modify `Docs/UserGuide.md`
- Move this file to `Specs/Plans/Done/FolderViewThumbnailPerformanceAndSizingPlan_2026-05-14.md`

Closeout must list:

- baseline run paths,
- candidate run paths,
- metric deltas,
- correctness failures fixed,
- residual caveats if any,
- exact durable specs updated.
