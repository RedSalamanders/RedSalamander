# ViewerImgRaw Plugin Specification

## Overview

**ViewerImgRaw** (`builtin/viewer-imgraw`) is a built-in **viewer plugin** that displays:
- **camera RAW images** using **LibRaw**
- **common image formats** using **Windows Imaging Component (WIC)**, with **JPEG decoding via libjpeg-turbo** (not WIC)

Goals:
- fast first paint (prefer thumbnails when available)
- asynchronous decoding (do not block UI thread)
- good “Other Files” navigation UX (ViewerText-style)
- respect host theme (colors, dark/high-contrast, rainbow, DPI)

## Invocation (Host Integration)

ViewerImgRaw is invoked via the standard viewer association mechanism:
- **Shortcut**: `F3` (View)
- Host chooses the viewer by file extension using `extensions.openWithViewerByExtension` (see `Specs/Plugins/Plugins_ViewerPlugins.md` and `Specs/Core/Core_SettingsStore.md`).

Deployment:
- The host loads the built-in DLL from `<exeDir>\\Plugins\\ViewerImgRaw.dll`; required runtime DLLs are deployed next to the plugin DLL.

## ViewerOpenContext Contract

ViewerImgRaw uses:
- `ViewerOpenContext.fileSystem` (active filesystem instance)
- `ViewerOpenContext.fileSystemName` (optional display name)
- `ViewerOpenContext.focusedPath` (filesystem-internal UTF-16 file path)
- `ViewerOpenContext.otherFiles` + `focusedOtherFileIndex` for navigation

`selectionPaths` is currently unused by ViewerImgRaw.

Path semantics:
- Paths are treated as **opaque filesystem-internal strings** and may not be valid Win32 paths.
- ViewerImgRaw MUST read through `IFileSystemIO::CreateFileReader` and MUST NOT use Win32 file APIs for file access.
- A reader seek to the beginning must succeed and report position zero. Each `Read` request is capped at `1 MiB`; a failed read or a returned count larger than the requested buffer is rejected before ViewerImgRaw advances an offset or exposes bytes to a decoder. ViewerImgRaw reads exactly the declared `GetSize()` byte count and then requires an EOF probe to return zero bytes; early EOF, trailing bytes, size drift, and impossible provider counts are invalid provider data.
- Main image sources are capped at `1 GiB`; sidecar JPEG sources are capped at `64 MiB`. These source-byte ceilings apply before proportional allocation and are independent of the decoded-frame limits below.

## Supported Extensions

ViewerImgRaw is intended to be associated with:

Baseline WIC formats (built-in codecs):
`.bmp`, `.dib`, `.gif`, `.ico`, `.jpe`, `.jpeg`, `.jpg`, `.png`, `.tif`, `.tiff`, `.hdp`, `.jxr`, `.wdp`

Camera RAW formats (LibRaw):
`.3fr`, `.ari`, `.arw`, `.bay`, `.braw`, `.crw`, `.cr2`, `.cr3`, `.cap`, `.data`, `.dcs`, `.dcr`, `.dng`, `.drf`, `.eip`, `.erf`, `.fff`, `.gpr`, `.iiq`, `.k25`, `.kdc`, `.mdc`, `.mef`, `.mos`, `.mrw`, `.nef`, `.nrw`, `.obm`, `.orf`, `.pef`, `.ptx`, `.pxn`, `.r3d`, `.raf`, `.raw`, `.rwl`, `.rw2`, `.rwz`, `.sr2`, `.srf`, `.srw`, `.x3f`

Notes:
- WIC may decode additional formats when third-party codecs are installed; ViewerImgRaw may still succeed even if an extension is not listed here.

## Configuration

ViewerImgRaw exposes a JSON configuration schema (`GetConfigurationSchema`) and accepts configuration via `SetConfiguration`.

Decode & navigation keys (defaults):
- `halfSize` (`true`): decode at half resolution for faster load and lower memory use (LibRaw).
- `preferThumbnail` (`true`): open in Thumbnail mode by default (uses sidecar JPEG when present, otherwise embedded thumbnail when available, otherwise falls back to full RAW decode).
- `useCameraWb` (`true`): use camera white balance (LibRaw).
- `autoWb` (`false`): enable LibRaw auto white balance.
- `zoomOnClickPercent` (`50`): temporary zoom level (percent) while the left mouse button is held down on the image.
- `prevCache` (1, `0..8`): number of previous images to keep decoded in memory.
- `nextCache` (1, `0..8`): number of next images to keep decoded in memory.

Export (WIC encoder) keys (defaults):
- `exportJpegQualityPercent` (`90`): JPEG encoder quality (mapped to WIC `ImageQuality`).
- `exportJpegSubsampling` (`0`): WIC `WICJpegYCrCbSubsamplingOption` (`0..4`).
- `exportPngFilter` (`0`): WIC `WICPngFilterOption` (`0..6`).
- `exportPngInterlace` (`false`): PNG interlace (WIC `InterlaceOption` when supported).
- `exportTiffCompression` (`0`): WIC `WICTiffCompressionOption` (`0..7`).
- `exportBmpUseV5Header32bppBGRA` (`true`): BMP V5 header BGRA option (WIC `EnableV5Header32bppBGRA` when supported).
- `exportGifInterlace` (`false`): GIF interlace (WIC `InterlaceOption` when supported).
- `exportWmpQualityPercent` (`90`): JPEG XR (WMP container) quality (mapped to WIC `ImageQuality`).
- `exportWmpLossless` (`false`): JPEG XR lossless (WIC `Lossless` when supported).

Notes:
- `prevCache` and `nextCache` are independent; `0` disables caching in that direction.
- Cached images are stored as decoded BGRA frames. Neighbor cache retention and in-flight prefetch share the decoded-byte budget described below; configured item counts are upper bounds, not permission to exceed that budget.

## UI / UX

Layout:
- **Header**: in standalone mode only, filename dropdown (combo box) listing `otherFiles` when `otherFileCount > 1`; the collapsed combo uses compact chrome (28 DIP control height with minimal themed padding) so the header leaves more vertical space for image data while still keeping the combo fully inside the header background. Embedded preview mode hides this standalone header/dropdown and renders the image surface against the preview/tab background.
- **Content**: decoded image (fit-to-window by default, with manual zoom controls).
- **Scrollbars**: when not in Fit-to-Window and the image exceeds the viewport, standard Win32 scrollbars are shown and can be used to pan (no scrollbars in Fit mode); ranges are based on the oriented (EXIF+user) image size at the current zoom.
- **Status bar** (owner-drawn, themed):
  - left segment: `Loading…` (with LibRaw stage/percent when available), navigation position (e.g. `3/17`), current label, status message
  - right segment (when an image is displayed): displayed source (`RAW` / `JPG` / `THUMB`), oriented dimensions, zoom percent, and adjustment flags (`Ori*`, `B…`, `C…`, `G…`, `Gray`, `Neg`)
- **Exif & orientation**:
  - Exif is extracted from RAW (LibRaw) and from JPEG APP1 Exif (sidecar JPEGs and embedded RAW JPEG thumbnails). For non-JPEG WIC inputs, ViewerImgRaw reads the first frame's orientation metadata from the WIC metadata query reader.
  - EXIF orientation is applied at render time; user transforms (rotate/flip/reset) compose on top of the source orientation.
  - JPEG/TIFF scalar metadata used for orientation, Exif-IFD traversal, and ISO must have the expected TIFF type and `count == 1`; malformed wrong-type, non-scalar, or oversized counts are ignored.
  - Rotate/flip/reset recenters the image (pan offsets reset).

RAW + sidecar pairing:
- When both a RAW file and a `.jpg/.jpeg` with the same base name exist in the `otherFiles` list, ViewerImgRaw shows a single combo entry formatted like: `afile.raf | afile.jpg`.
- Next/Previous/First/Last navigation operates on these pairs as single items.
- When the viewer is in **Thumbnail** mode and a sidecar JPEG exists, ViewerImgRaw displays the sidecar JPEG instead of opening/decoding the RAW.

Menu (DxUi-hosted from the hidden native menu model):
- File: Refresh (`F5`), Export..., Exit (`Esc`)
- Other Files: Next / Previous / First / Last
- View: Fit to Window / Actual Size / Toggle Fit↔100%, Zoom In/Out/Reset, Transform (rotate/flip/reset), Adjust (brightness/contrast/gamma + grayscale/negative), Image Source (RAW / Thumbnail), Show Exif Overlay

When the shared `RedSalamander.DxNativeMenuBar` host attaches successfully, the window detaches its live native `HMENU`, adopts that detached handle as the single hidden RAII menu-model owner, and renders the visible top menu bar through DxUi. If attach fails, the handle remains solely window-owned and the native menu stays visible as the functional fallback. `Alt`, `F10`, and menu mnemonics continue to route through the active menu bar.

Keyboard shortcuts (ViewerText-aligned where meaningful):
- `Esc`: cancel active RAW decoding if loading; otherwise close viewer
- File-dropdown and menu keyboard use stays local to the focused combo/menu; `Esc` from that chrome returns focus to the image surface instead of closing. After peer-file selection or command activation, focus returns to the image surface.
- `F5`: refresh current file
- `Right` / `PgDn` / `Space`: next file
- `Left` / `PgUp` / `Backspace`: previous file
- `Home` / `End`: first/last file
- `Ctrl+F` / `Double Click`: fit to window
- `F`: toggle Fit ↔ 100%
- `1` / `Ctrl+Double Click`: actual size (100%)
- `key left of Backspace` / `key right of 0` / `0`: zoom in/out/reset (keyboard-layout independent; menu displays the current keyboard-layout glyph)
- `Ctrl+S`: export
- `I`: toggle Exif overlay
- `R`: rotate clockwise
- `Ctrl+R` / `Shift+R`: rotate counterclockwise
- `H` / `V`: flip horizontal / vertical
- `O`: reset orientation
- `G`: toggle grayscale
- `N`: toggle negative
- `Ctrl+Alt+Up/Down`: brightness ±
- `Ctrl+Alt+Left/Right`: contrast ±
- `Ctrl+Alt+PgUp/PgDn`: gamma ±
- `Ctrl+Arrow keys`: pan image (when zoomed)
- Mouse:
  - `Ctrl+Left click (hold)`: transient zoom/unzoom to `zoomOnClickPercent` (release restores previous zoom + pan; panning is clamped to image bounds)
  - left drag: pan image when zoomed
  - wheel: zoom smoothly around cursor
  - `Ctrl+Wheel`: brightness adjust
  - `Ctrl+Shift+Wheel`: contrast adjust

## Loading / Decode Pipeline

Decode runs on background threads:
- **WIC images**: non-JPEG formats decoded via WIC to 8-bit BGRA (first frame) on a background worker; JPEG uses libjpeg-turbo. The first frame's WIC orientation metadata is preserved and applied by the common render-orientation path, including when a RAW extension falls back to WIC decoding.
  - For large/progressive JPEGs, ViewerImgRaw may display a scaled preview first, then replace it with the full decoded frame (this is not scanline-level progressive rendering).
- **RAW fast preview (thumbnail mode)**:
  - if a sidecar `.jpg/.jpeg` exists for the current RAW pair, it is decoded via libjpeg-turbo and displayed without reading the RAW.
  - otherwise, ViewerImgRaw attempts an embedded thumbnail via `LibRaw::unpack_thumb()` and `raw.imgdata.thumbnail`; JPEG thumbnails are decoded via libjpeg-turbo and displayed immediately.
- **Image source selection**:
  - `RAW (decoded)`: for RAW inputs, a thumbnail may be shown first while the full RAW is decoded; then ViewerImgRaw switches to the full RAW.
  - `Thumbnail`: displays the best available thumbnail source; if no thumbnail exists it falls back to full RAW decode. A successfully decoded embedded thumbnail is delivered directly as one final success instead of a non-final result followed by a default failure. RAW mode may publish that embedded thumbnail as a non-final preview before the full decode; a progressive sidecar JPEG may independently publish a scaled preview before its final full thumbnail.
- **Progress reporting**:
  - ViewerImgRaw installs a LibRaw progress handler via `LibRaw::set_progress_handler(...)` during full RAW decode and forwards updates to the UI.
  - Progress posts remain allocation-free: the request generation travels in `WPARAM`, while the signed stage and percent values are packed into `LPARAM`. The UI applies an update only while that exact generation is still loading, so queued progress from navigation or cancellation cannot overwrite the new request's status.
  - The loading overlay shows a progress line and bar under the spinner when percent information is available.
  - Pressing `Esc` while the image surface is focused and a RAW decode/load is active invalidates the current open request, drops pending decode entries, stops the loading UI, and leaves the viewer open.
- **Full RAW image**: ViewerImgRaw decodes the full RAW using:
  - `LibRaw::open_buffer()` + `unpack()` + `dcraw_process()` + `dcraw_make_mem_image()`
  - output is converted to 8-bit BGRA for display

Main-open scheduling:
- Each viewer owns one shared scheduler state with at most one active main-open callback and one replaceable pending request. Repeated navigation coalesces the pending slot to the latest request; it never starts unbounded parallel main reads/decodes.
- Every active request checks its generation immediately before each whole-file read and again before expensive WIC, JPEG, embedded-thumbnail, LibRaw, or fallback decode work. A stale active request releases its source/result storage and advances to the latest pending request.
- Navigation, cancellation, `Close()`, and window teardown never wait for a filesystem provider or decoder. They invalidate the generation and drop the pending request; the active request owns the viewer/filesystem/module lifetimes until it reaches a cancellation boundary and returns.
- At most one active main request owns source/decode storage. The production active-worker ceiling used for evidence is the conservative sum of one `1 GiB` main source, one `64 MiB` sidecar source, and two `256 MiB` BGRA frames. A result already transferred to the UI payload registry is outside that active-worker ceiling and remains governed by request identity plus the window payload-drain contract.
- `viewer.imgraw.open.queue.replaced_count` increments only when a new request overwrites an already occupied pending slot. In `ENABLE_TESTS`, `ViewerImgRawResourceDebugSnapshot.replacedMainDecodeCount` is the matching cumulative count, while `activeMainDecodeCount` and `pendingMainDecodeCount` expose the production one-plus-one bound.

Decoded-image resource policy:
- One overflow-safe policy applies to WIC, libjpeg-turbo, LibRaw full output, and embedded bitmap/JPEG thumbnails before ViewerImgRaw allocates or copies owned BGRA bytes.
- A decoded frame is limited to a `16384`-pixel dimension, `64 * 1024 * 1024` pixels (about 67 megapixels), and `256 MiB` of BGRA. Width, height, row bytes, pixel count, and total output bytes must all validate.
- LibRaw's exposed output dimensions are checked after unpack and before `dcraw_process()`, then checked again against the actual processed output before the BGRA allocation.
- Embedded thumbnails must declare positive bounded dimensions. `tlength` must fit within the source file; bitmap thumbnails must contain the overflow-safe expected packed sample bytes, while JPEG thumbnails also have a fixed `64 MiB` compressed-byte ceiling. The decoded JPEG header is independently checked by the same frame policy.

Neighbor prefetch:
- After a successful decode, ViewerImgRaw prefetches missing neighbors according to `prevCache`/`nextCache` and keeps decoded frames for fast navigation.
- Prefetch prioritizes the active Image Source mode; when a sidecar exists it may be used as the thumbnail source.
- Neighbor cache frames and in-flight prefetch reservations share a per-viewer `384 MiB` speculative decoded-byte budget. The actively opened image is exempt so speculative work cannot reject or delay the user's main open; when it becomes a neighbor it must obtain budget or be evicted.
- Reservations are released on decode failure, budget rejection, stale request generation, failed publication, cache eviction/clear, and teardown.
- Request-generation advancement and the final prefetch generation check/cache publication are serialized by the cache mutex. Cancellation or configuration-driven cache clearing therefore linearizes either before publication, causing the stale result to be rejected, or after publication, causing the entry and its reservation to be removed; stale work cannot repopulate a cleared cache.
- Perf evidence uses `viewer.imgraw.decode.accepted_bytes`, `viewer.imgraw.decode.rejected_bytes`, `viewer.imgraw.prefetch.accepted_bytes`, `viewer.imgraw.prefetch.rejected_bytes`, and `viewer.imgraw.resource.decoded_bytes`. Main-open queue evidence additionally records `viewer.imgraw.open.queue.active_workers`, `viewer.imgraw.open.queue.pending_count`, `viewer.imgraw.open.queue.replaced_count`, `viewer.imgraw.open.queue.wait_us`, `viewer.imgraw.open.queue.stale_before_read`, `viewer.imgraw.open.queue.stale_before_decode`, `viewer.imgraw.open.resource.source_bytes`, `viewer.imgraw.open.resource.decoded_bytes`, and `viewer.imgraw.open.resource.owned_peak_limit`.

Threading notes:
- ViewerImgRaw decodes through Windows threadpool callbacks and creates a fresh LibRaw instance per decode.
- Async open, neighbor prefetch, and export work must acquire an owning `ViewerImgRaw.dll` module pin before queueing. Threadpool callbacks transfer that pin with `FreeLibraryWhenCallbackReturns` at callback entry. Pin or queue failure never runs work synchronously: prefetch is abandoned, while an active open leaves the loading state and shows a terminal localized error.
- Final open results normally use `PostMessagePayload`. If the result allocation or payload post fails for the current generation, the worker clears its in-flight marker and records an allocation-free terminal slot in the shared scheduler. The existing loading timer polls that slot, ends loading, and shows a localized terminal alert without the worker reading or mutating a destroyed/recycled window. `ENABLE_TESTS` exposes `REDSALAMANDER_VIEWERIMGRAW_FORCE_OPEN_RESULT_ALLOCATION_FAILURE` and `REDSALAMANDER_VIEWERIMGRAW_FORCE_OPEN_RESULT_POST_FAILURE` for deterministic coverage.
- `RedSalamanderPluginShutdown` is the required quiet point for unregistering the ImgRaw window class and resetting its WIL-owned class brushes. `DllMain` performs process-attach initialization only. `RedSalamanderPluginCanUnloadNow` remains false until quiet-point cleanup succeeded, every created `ViewerImgRaw` COM instance has completed destruction, and every queued main-open, prefetch, and export callback has returned; an unopened retained instance or a main-open callback blocked in provider `Read` therefore keeps the unload vote false. `RedSalamanderPluginRetainModuleUntilProcessExit` is false because callback-return module pins and this unload gate provide the lifetime contract.
- The project links against vcpkg's re-entrant LibRaw library (`raw_r`); OpenMP may be enabled when LibRaw is built with OpenMP support (see `libraw[openmp]` in `vcpkg.json`) and is detectable via `LIBRAW_USE_OPENMP`.
- The dependency floor is LibRaw 0.22.1 or newer. Do not lower the `vcpkg.json` minimum or ship a resolved build below that floor; decoder security fixes are part of the ViewerImgRaw input-safety contract.

Export:
- File → Export... writes the currently displayed frame via WIC encoding.
- Supported export formats: PNG, JPEG, TIFF, BMP, GIF, JPEG XR (WMP container); the encoder is selected primarily by the output file extension (and by the dialog filter when no extension is provided).
- Encoder options are configured via the `export*` configuration keys and are applied via `IPropertyBag2` when supported by the selected encoder.
- Atomic export reserves a collision-safe sibling staging file with `CREATE_NEW`, encodes to its long-path-aware Win32 path, releases the WIC file graph, and commits with `MoveFileExW(..., MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)`. Failure deletes only the staging file; valid destinations beyond `MAX_PATH` retain the same atomicity guarantee.
- Every reader, decode, export validation, encoder, staging, commit, success, and terminal-failure message shown to the user is loaded from the ViewerImgRaw resource owner (base plus cs-CZ/fr-FR/ja-JP/sk-SK satellites). User text does not carry the diagnostic `ViewerImgRaw:` prefix; that prefix remains valid only for debug logging.

## Rendering / Theme

- Rendering uses Direct2D (Hwnd render target) + DirectWrite.
- A centered loading overlay is shown after a ~200ms delay and uses theme accent colors; in `rainbowMode` the accent is derived from the current file name.
- When a preview image is already displayed while decoding continues, the loading overlay is drawn with lower opacity so the preview remains visible.
- Exif overlay (when enabled) renders in the bottom-right of the content area and uses theme/rainbow-aware accent tinting.
- Menus are owner-drawn to respect theme colors (including selection colors and high-contrast handling).

## Testing Contract

- `TestViewerImgRawLatestWinsExactReaderAndCloseSafety` exercises the production one-active/one-replaceable-pending scheduler, requires exactly one pending replacement and no read of the superseded middle request, and verifies the `1 MiB` request ceiling. It also requires the module unload vote to remain false for an unopened retained COM instance, proves current-generation progress applies while stale-generation progress is ignored, and covers the `1 GiB` source cap, initial-seek mismatch, premature EOF, trailing data, impossible over-return, allocation/post terminal fallbacks, recovery, and sub-500 ms close while a provider `Read` remains blocked. The blocked path requires exactly one `ViewerClosed`, retained provider/viewer lifetime, a false runtime unload vote, a mapped callback module until the provider returns, and exact retirement after callback return.
- `TestViewerImgRawResourceBudgetAndLongPathExport` pauses prefetch after its optimistic generation check, invalidates and clears the cache, then requires the retiring worker to leave zero cache entries, in-flight markers, and speculative bytes. This guards the generation/cache-publication linearization in addition to the decoded-resource and long-path export bounds.
- `ENABLE_TESTS` exposes monotonic `lastPreviewApplyOrdinal` and `lastFinalApplyOrdinal` values in `ViewerImgRawResourceDebugSnapshot`; zero means that result kind was not applied. `TestViewerImgRawEmbeddedThumbnailTerminalSequencing` requires Thumbnail mode to apply one final embedded-thumbnail success with no preview ordinal or warning, while RAW mode must apply one preview ordinal strictly before one final full-image ordinal, again without a terminal warning.
