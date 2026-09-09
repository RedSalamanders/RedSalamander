# Advisor Plan 014 - ViewerImgRaw export pixel snapshot

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/014-viewerimgraw-export-uaf-snapshot-pixels.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/014-viewerimgraw-export-uaf-snapshot-pixels.md`
- **Live owner:** Specs/Plans/Done/Operation_Farsight_ViewerPluginsRemediation_ExportSafetyDecodeReliabilityWebSecurityAndTextGeometry_2026-06-16.md
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 014: ViewerImgRaw — snapshot export pixels before the modal save dialog (fix use-after-free / corrupt export)

> **Frozen historical instructions (do not execute)**: This plan is archived; do not follow these steps. Run every verification
> command and confirm the expected result before moving to the next step. If anything
> in the "STOP conditions" section occurs, stop and report — do not improvise.
> Historical: status-row updates no longer apply; root `plans/` was deleted.
>
> **Drift check (run first)**: `git diff --stat b274022d9..HEAD -- Plugins/ViewerImgRaw/ViewerImgRaw.Export.cpp Plugins/ViewerImgRaw/ViewerImgRaw.Decode.cpp`
> If either in-scope file changed since this plan was written, compare the "Current
> state" excerpts against the live code before proceeding; on a mismatch, treat it as
> a STOP condition.

## Status

- **Priority**: P1 (ship-blocker — data-corrupting use-after-free)
- **Effort**: S
- **Risk**: LOW
- **Depends on**: none
- **Category**: bug (concurrency / lifetime)
- **Planned at**: commit `b274022d9`, 2026-06-16

## Why this matters

`ViewerImgRaw::BeginExport` captures a **raw pointer to the cached image and a C++
reference directly into that image's pixel vector**, then opens a modal `IFileSaveDialog`
whose nested message pump dispatches the viewer's posted messages. A background RAW-decode
or neighbor-prefetch completion that fires during the dialog calls `OnAsyncOpenComplete`,
which does `image->rawBgra = std::move(result->bgra)` (destroying the referenced vector's
contents) and can reset the owning `unique_ptr`; `ClearImageCache` can destroy the
`CachedImage` outright. After the dialog returns, `BeginExport` copies from the now
dangling/empty reference. Opening a RAW file and immediately exporting it — while the
full-resolution decode is still running in the background, the normal case — reads freed
heap or silently writes a **blank/corrupt image to the user's chosen path**. It is
non-deterministic and hard to reproduce, which is exactly why it must be fixed by design.

## Current state

- `Plugins/ViewerImgRaw/ViewerImgRaw.Export.cpp` — `BeginExport` (the export entry point).
  The dangerous capture and the late copy:

```cpp
// ViewerImgRaw.Export.cpp:599-603  (inside BeginExport, on the UI thread)
const CachedImage* image         = _currentImage;
const bool exportingThumb        = IsDisplayingThumbnail();
const uint32_t w                 = exportingThumb ? image->thumbWidth : image->rawWidth;
const uint32_t h                 = exportingThumb ? image->thumbHeight : image->rawHeight;
const std::vector<uint8_t>& bgra = exportingThumb ? image->thumbBgra : image->rawBgra;   // reference into the cache
// ... w/h/empty guard ...
// ViewerImgRaw.Export.cpp:648  — MODAL, pumps the message loop:
const std::optional<ExportSaveDialogResult> save = ShowExportSaveDialog(hwnd, defaultFormat, suggested);
// ... extension handling ...
// ViewerImgRaw.Export.cpp:691-692  — copies from a reference that the pump may have invalidated:
std::vector<uint8_t> pixels;
pixels = bgra;
```

- The mutation that invalidates the reference, reachable from the dialog's pump via
  `kAsyncOpenCompleteMessage` → `OnAsyncOpenComplete`:

```cpp
// ViewerImgRaw.Decode.cpp:2886-2902  (OnAsyncOpenComplete, UI thread — these writes run WITHOUT _cacheMutex held;
//                                     the lock at :2839 was already released at :2846)
image->rawWidth  = result->width;
image->rawHeight = result->height;
image->rawBgra   = std::move(result->bgra);     // <-- destroys the contents `bgra` references
// ...
if (cachingEnabled) { _currentImageOwned.reset(); }   // <-- can free the CachedImage `image` points at
_currentImage    = image;
```

```cpp
// ViewerImgRaw.Decode.cpp:1389-1397  (ClearImageCache, UI thread)
{
    std::scoped_lock lock(_cacheMutex);   // guards ONLY the two containers below
    _inflightDecodes.clear();
    _imageCache.clear();          // <-- destroys cached CachedImage objects
}                                 // <-- lock RELEASED here (:1393)
_currentImageOwned.reset();   // <-- destroys the non-cached current image (NO lock held)
_currentImage = nullptr;      // <-- (NO lock held)
```

- After line 692 the owned `pixels` copy is what gets moved into the async export work
  item (`pixels = std::move(pixels)` at `:719`), so once the snapshot is owned the rest of
  the path is already safe.

**Why the fix works (the load-bearing invariant)**: every mutation shown above runs on the
**UI thread**, and it reaches this code only *re-entrantly* — the modal `ShowExportSaveDialog`
pumps the message loop, which dispatches the already-posted `kAsyncOpenCompleteMessage` into
`OnAsyncOpenComplete` (and a `ClearImageCache` can run the same way). The fix is therefore to
**deep-copy the pixel bytes into an owned `std::vector` *before* the dialog pumps**. Once the
bytes are owned, any re-entrant mutation of the cache is harmless.

**Do NOT take `_cacheMutex` for the copy.** `_cacheMutex` (declared `ViewerImgRaw.h:319` as a
plain non-recursive `std::mutex`) guards only the cache *containers* (`_imageCache`,
`_inflightDecodes`). As the two excerpts above show, it is **not** held when `rawBgra`,
`thumbBgra`, `_currentImage`, or `_currentImageOwned` are written — so locking the snapshot
read would NOT make it safe against those writes, and would be asymmetric with the existing
lock-free readers/writers of these fields for zero benefit. Safety comes entirely from copying
before the pump on the single UI thread, not from any lock.

## Commands you will need

| Purpose | Command | Expected on success |
|---|---|---|
| Build plugin | `.\build.ps1 -ProjectName ViewerImgRaw` | exit 0, 0 warnings, 0 errors |
| Full build | `.\build.ps1` | exit 0 |
| Viewer integration tests | `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` | all pass (long; serialized by a machine-wide mutex — if it exits 3, another selftest run holds the mutex; retry later) |

## Scope

**In scope** (modify only these):
- `Plugins/ViewerImgRaw/ViewerImgRaw.Export.cpp`

**Out of scope** (do NOT touch):
- `ViewerImgRaw.Decode.cpp` mutation logic — the fix is entirely on the export side.
- The async export work-item dispatch (`AddRef`/`ctx`/`queued==0`) — that leak is fixed
  separately by plan 017; do not change it here, but do not regress it either.

## Git workflow

- Branch: `advisor/014-viewerimgraw-export-uaf`
- Conventional-commit message, e.g. `fix(ViewerImgRaw): snapshot export pixels before modal save dialog`
- Do NOT push or open a PR unless the operator instructs it.

## Steps

### Step 1: Snapshot width/height/pixels into owned locals before the dialog

In `BeginExport`, replace the reference capture at lines 599-603 so that the pixel bytes
are **copied into an owned `std::vector<uint8_t>`** and width/height into plain locals
(no lock needed — see "Why the fix works" above), *before* the `w == 0 || h == 0 ||
bgra.empty()` guard and well before `ShowExportSaveDialog`. Target shape:

```cpp
uint32_t w = 0;
uint32_t h = 0;
std::vector<uint8_t> pixels;                 // owned snapshot — survives the modal dialog
const bool exportingThumb = IsDisplayingThumbnail();
const CachedImage* image  = _currentImage;   // UI-thread-affine; non-null here per the guard at the top of BeginExport
if (image != nullptr)                        // defensive re-check; no pump runs between that guard and here
{
    w      = exportingThumb ? image->thumbWidth : image->rawWidth;
    h      = exportingThumb ? image->thumbHeight : image->rawHeight;
    pixels = exportingThumb ? image->thumbBgra : image->rawBgra;   // deep copy BEFORE the modal dialog pumps
}

if (w == 0 || h == 0 || pixels.empty())
{
    // ... existing IDS_VIEWERRAW_EXPORT_NO_IMAGE alert + return ...
}
```

Then delete the later `std::vector<uint8_t> pixels; pixels = bgra;` at lines 691-692 (the
data is already in `pixels`). Leave the `pixels = std::move(pixels)` capture in the work
item lambda as-is.

**Confirm** there is no remaining use of the `image` raw pointer or the `bgra` reference
*after* `ShowExportSaveDialog`. Run
`grep -n "image->\|bgra" Plugins/ViewerImgRaw/ViewerImgRaw.Export.cpp` and read the hits:

- Lines ~356, ~362, ~477 are the **`EncodeImageToFile` helper's `bgra` parameter** — a
  *different* function that lives *above* `BeginExport` (which starts at line 571). These are
  unrelated; leave them.
- Inside `BeginExport`, the only `image->`/`bgra` tokens must be in the snapshot block *above*
  the `if (w == 0 ...)` guard. The old `pixels = bgra;` (was line 692, *after* the dialog) must
  be gone.

A flat grep cannot prove "after the dialog," so **also open `BeginExport` and read from the
`ShowExportSaveDialog` call to the end of the function**, confirming no `image` or `bgra`
identifier appears in that span.

**Verify**: `.\build.ps1 -ProjectName ViewerImgRaw` → exit 0, no new warnings.

### Step 2: Confirm only owned locals cross the dialog

Read the final `BeginExport`. Only the owned `w`, `h`, and `pixels` locals may be referenced
after `ShowExportSaveDialog`; the raw `image` pointer and the `bgra` reference must not appear
past the snapshot block. (No `_cacheMutex` lock should be involved — see the invariant above. If
you nonetheless added one, its `{ }` scope must close *before* `ShowExportSaveDialog` and must
not wrap the dialog or the work-item submission.)

**Verify**: `.\build.ps1` → exit 0.

## Test plan

- The existing viewer integration harness is `Tests/ViewerPETests/ViewerPETests.cpp`
  (builds `ViewerPETests.exe`, registered in the `Full` suite). Add a regression case
  there modeled on the existing isolated-viewer cases: open a RAW (or any image) in
  ViewerImgRaw, trigger the export command (`IDM_VIEWERRAW_*` export id), and confirm no
  crash and that the produced file is non-empty. A fully deterministic repro of the race
  is impractical from the UI harness; the high-value automated check is "export of a
  freshly-opened image does not crash and writes a non-zero-length file".
- If adding the case is not feasible without new harness plumbing, STOP and report; the
  code fix is the priority and is verifiable by inspection (no `image`/`bgra` use after the
  dialog) plus the build + full-suite gates below.
- Verification: `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` → all pass.

## Done criteria

ALL must hold:

- (archived, not live)  `.\build.ps1 -ProjectName ViewerImgRaw` exits 0 with 0 warnings/0 errors.
- (archived, not live)  `.\build.ps1` exits 0.
- (archived, not live)  In `BeginExport`, no `image->...` access or `bgra` reference exists after
      `ShowExportSaveDialog` (`grep` within the function returns none).
- (archived, not live)  The pixel data exported comes from an owned `std::vector` snapshot deep-copied **before**
      `ShowExportSaveDialog` (so the modal pump cannot invalidate it); no `image`/`bgra` is read
      after the dialog.
- (archived, not live)  `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` passes (or, if mutex-blocked,
      re-run until it executes and passes).
- (archived, not live)  No files outside `ViewerImgRaw.Export.cpp` modified (`git status`).
- (archived, not live)  `plans/README.md` status row updated.

## STOP conditions

Stop and report (do not improvise) if:

- The "Current state" excerpts don't match the live code (drift since `b274022d9`).
- `IsDisplayingThumbnail()` or `HasDisplayImage()` no longer exists or has changed semantics,
  or the top-of-function `if (! HasDisplayImage() || ! _currentImage)` guard is gone — the
  snapshot relies on `_currentImage` being non-null and stable until the dialog.
- You find any code path that writes `image->rawBgra`/`thumbBgra`, `_currentImage`, or
  `_currentImageOwned` from a **non-UI thread** (e.g. the decode worker writing the cache
  directly instead of posting an `AsyncOpenResult`). That would break the UI-thread-affine
  invariant this fix depends on — report instead of proceeding.
- The full-suite gate fails for a reason that looks related to this change.

## Maintenance notes

- Any future code that reads `_currentImage`/`CachedImage` pixel buffers and then pumps
  messages (modal dialog, `DoModal`, `MsgWaitForMultipleObjects`, nested loop) has the same
  hazard — snapshot first.
- Reviewer should confirm the snapshot is a deep copy (an owned `std::vector`, not a `const&`
  or pointer) taken before `ShowExportSaveDialog`, and that no `image`/`bgra` is read after it.
  The safety property is "copy before the pump," not any lock — don't accept a change that drops
  the early copy on the theory that a mutex protects the read.
- Related: plan 016 makes the export write atomic; plan 017 fixes the work-item dispatch
  leak in the same function. They are independent and can land in any order.
