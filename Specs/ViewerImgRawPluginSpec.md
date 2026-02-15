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
- Host chooses the viewer by file extension using `extensions.openWithViewerByExtension` (see `Specs/PluginsViewer.md` and `Specs/SettingsStoreSpec.md`).

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
- Cached images are stored as decoded BGRA frames; memory use scales with image size and cache counts.

## UI / UX

Layout:
- **Header**: filename dropdown (combo box) listing `otherFiles` when `otherFileCount > 1`.
- **Content**: decoded image (fit-to-window by default, with manual zoom controls).
- **Scrollbars**: when not in Fit-to-Window and the image exceeds the viewport, standard Win32 scrollbars are shown and can be used to pan (no scrollbars in Fit mode); ranges are based on the oriented (EXIF+user) image size at the current zoom.
- **Status bar** (owner-drawn, themed):
  - left segment: `Loading…` (with LibRaw stage/percent when available), navigation position (e.g. `3/17`), current label, status message
  - right segment (when an image is displayed): displayed source (`RAW` / `JPG` / `THUMB`), oriented dimensions, zoom percent, and adjustment flags (`Ori*`, `B…`, `C…`, `G…`, `Gray`, `Neg`)
- **Exif & orientation**:
  - Exif is extracted from RAW (LibRaw) and from JPEG APP1 Exif (sidecar JPEGs and embedded RAW JPEG thumbnails).
  - EXIF orientation is applied at render time; user transforms (rotate/flip/reset) compose on top of the source orientation.
  - Rotate/flip/reset recenters the image (pan offsets reset).

RAW + sidecar pairing:
- When both a RAW file and a `.jpg/.jpeg` with the same base name exist in the `otherFiles` list, ViewerImgRaw shows a single combo entry formatted like: `afile.raf | afile.jpg`.
- Next/Previous/First/Last navigation operates on these pairs as single items.
- When the viewer is in **Thumbnail** mode and a sidecar JPEG exists, ViewerImgRaw displays the sidecar JPEG instead of opening/decoding the RAW.

Menu (themed/owner-drawn like ViewerText):
- File: Refresh (`F5`), Export..., Exit (`Esc`)
- Other Files: Next / Previous / First / Last
- View: Fit to Window / Actual Size / Toggle Fit↔100%, Zoom In/Out/Reset, Transform (rotate/flip/reset), Adjust (brightness/contrast/gamma + grayscale/negative), Image Source (RAW / Thumbnail), Show Exif Overlay

Keyboard shortcuts (ViewerText-aligned where meaningful):
- `Esc`: dismiss alert (if visible) or close viewer
- `F5`: refresh current file
- `Right` / `PgDn` / `Space`: next file
- `Left` / `PgUp` / `Backspace`: previous file
- `Home` / `End`: first/last file
- `Ctrl+F` / `Double Click`: fit to window
- `F`: toggle Fit ↔ 100%
- `1` / `Ctrl+Double Click`: actual size (100%)
- `+` / `-` / `0`: zoom in/out/reset (virtual-key based; menu displays the current keyboard-layout glyph)
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
- **WIC images**: non-JPEG formats decoded via WIC to 8-bit BGRA (first frame) on a background thread; JPEG uses libjpeg-turbo.
  - For large/progressive JPEGs, ViewerImgRaw may display a scaled preview first, then replace it with the full decoded frame (this is not scanline-level progressive rendering).
- **RAW fast preview (thumbnail mode)**:
  - if a sidecar `.jpg/.jpeg` exists for the current RAW pair, it is decoded via libjpeg-turbo and displayed without reading the RAW.
  - otherwise, ViewerImgRaw attempts an embedded thumbnail via `LibRaw::unpack_thumb()` and `raw.imgdata.thumbnail`; JPEG thumbnails are decoded via libjpeg-turbo and displayed immediately.
- **Image source selection**:
  - `RAW (decoded)`: for RAW inputs, a thumbnail may be shown first while the full RAW is decoded; then ViewerImgRaw switches to the full RAW.
  - `Thumbnail`: displays the best available thumbnail source; if no thumbnail exists it falls back to full RAW decode.
- **Progress reporting**:
  - ViewerImgRaw installs a LibRaw progress handler via `LibRaw::set_progress_handler(...)` during full RAW decode and forwards updates to the UI.
  - The loading overlay shows a progress line and bar under the spinner when percent information is available.
- **Full RAW image**: ViewerImgRaw decodes the full RAW using:
  - `LibRaw::open_buffer()` + `unpack()` + `dcraw_process()` + `dcraw_make_mem_image()`
  - output is converted to 8-bit BGRA for display

Neighbor prefetch:
- After a successful decode, ViewerImgRaw prefetches missing neighbors according to `prevCache`/`nextCache` and keeps decoded frames for fast navigation.
- Prefetch prioritizes the active Image Source mode; when a sidecar exists it may be used as the thumbnail source.

Threading notes:
- ViewerImgRaw decodes on background `std::thread` workers and creates a fresh LibRaw instance per decode.
- The project links against vcpkg's re-entrant LibRaw library (`raw_r`); OpenMP may be enabled when LibRaw is built with OpenMP support (see `libraw[openmp]` in `vcpkg.json`) and is detectable via `LIBRAW_USE_OPENMP`.

Export:
- File → Export... writes the currently displayed frame via WIC encoding.
- Supported export formats: PNG, JPEG, TIFF, BMP, GIF, JPEG XR (WMP container); the encoder is selected primarily by the output file extension (and by the dialog filter when no extension is provided).
- Encoder options are configured via the `export*` configuration keys and are applied via `IPropertyBag2` when supported by the selected encoder.

## Rendering / Theme

- Rendering uses Direct2D (Hwnd render target) + DirectWrite.
- A centered loading overlay is shown after a ~200ms delay and uses theme accent colors; in `rainbowMode` the accent is derived from the current file name.
- When a preview image is already displayed while decoding continues, the loading overlay is drawn with lower opacity so the preview remains visible.
- Exif overlay (when enabled) renders in the bottom-right of the content area and uses theme/rainbow-aware accent tinting.
- Menus are owner-drawn to respect theme colors (including selection colors and high-contrast handling).
