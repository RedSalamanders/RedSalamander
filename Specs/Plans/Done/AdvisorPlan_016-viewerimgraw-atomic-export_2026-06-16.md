# Advisor Plan 016 - ViewerImgRaw atomic export

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/016-viewerimgraw-atomic-export.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/016-viewerimgraw-atomic-export.md`
- **Live owner:** Specs/Plans/Done/Operation_Farsight_ViewerPluginsRemediation_ExportSafetyDecodeReliabilityWebSecurityAndTextGeometry_2026-06-16.md
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 016: ViewerImgRaw — atomic export (temp file + rename) so a failed encode never destroys the original

> **Frozen historical instructions (do not execute)**: Follow step by step; run every verification and confirm the
> expected result before continuing. Historical only. Root plans/ was deleted; do not execute.
>
> **Drift check (run first)**: `git diff --stat b274022d9..HEAD -- Plugins/ViewerImgRaw/ViewerImgRaw.Export.cpp`
> If changed, compare excerpts to live code first; mismatch ⇒ STOP.

## Status

- **Priority**: P1 (data-loss: overwrites can be destroyed mid-encode)
- **Effort**: M
- **Risk**: MED (touches the file-writing path; must preserve all existing format/encoder logic)
- **Depends on**: none (independent of plan 014, but both touch `BeginExport`/the export worker — land one, rebase the other)
- **Category**: bug (data safety)
- **Planned at**: commit `b274022d9`, 2026-06-16

## Review corrections — 2026-06-21

A multi-lens adversarial review (against HEAD; excerpts verified byte-identical to base
`b274022d9`, drift confined to plan 014's snapshot region, not tripped) found the design sound but
the recipe wrong in three ways — all folded into the steps below:

1. **BLOCKER — `stream.reset()` alone leaves the file handle open.** The encoder (initialized with
   `stream.get()`) and frame keep the `IWICStream` (and its OS handle) alive, so `MoveFileExW` would
   fail with a sharing violation on *every* export. Fix: release `frame`/`encoder`/`source`/`bitmap`/
   `paletteToSet` before `stream.reset()` and the rename (Step 2).
2. **HIGH — `noexcept` hazard.** The function is `noexcept`; the original `std::filesystem::path`
   path math could throw. Fix: derive the directory with `std::wstring` (the function's existing
   exception profile) and clean up with `DeleteFileW`; **no** `std::filesystem`, **no** new include
   (Step 1).
3. **Cleanup ordering.** Arm `cleanupTemp` *before* the WIC objects so it destructs *after* them
   (handle closed before `DeleteFileW`) (Step 1). Also: drop dead `MOVEFILE_COPY_ALLOWED`, capture
   `GetLastError()` once, and resolve the test-plan vs done-criterion #6 tension in favour of
   cover-by-inspection (see Test plan).

## Why this matters

`EncodeBgraToImageFileWic` opens the **user's chosen destination directly** with
`stream->InitializeFromFilename(outputPath, GENERIC_WRITE)`, which creates/truncates that
file before any pixels are written. Any later failure — disk full, device removed, encoder
error, GIF palette quantization failure, very large dimensions — returns after the file is
already zero-length or partially written. The save dialog uses `FOS_OVERWRITEPROMPT`, so the
user may have confirmed overwriting an existing image; **that original is destroyed the moment
`InitializeFromFilename` runs, with no recovery.** A file-manager's export must never be able
to corrupt the very file it is writing.

## Current state

```cpp
// ViewerImgRaw.Export.cpp:397-402  (EncodeBgraToImageFileWic — opens dest in place)
hr = stream->InitializeFromFilename(outputPath.c_str(), GENERIC_WRITE);   // create/truncate NOW
if (FAILED(hr)) { outStatusMessage = std::format(L"...open export file..."); return hr; }
```

```cpp
// ViewerImgRaw.Export.cpp:546-565  (failures after the file is already truncated)
hr = frame->WriteSource(source.get(), nullptr);  if (FAILED(hr)) { ...; return hr; }
hr = frame->Commit();                            if (FAILED(hr)) { ...; return hr; }
hr = encoder->Commit();                          if (FAILED(hr)) { ...; return hr; }
return S_OK;
```

On failure, `OnAsyncExportComplete` (`:786`) only shows a warning alert — it does not delete
the partial file or restore the original.

**Exemplar already in this codebase** (ViewerWeb writes its export to a temp then renames):

```cpp
// ViewerWeb.cpp:4119-4137  (pattern to mirror)
GetTempFileNameW(tempDir, L"rsw", 0, tempName);            // temp file
// ... write to tempPath ...
MoveFileExW(tempPath.c_str(), newPath.c_str(), MOVEFILE_REPLACE_EXISTING);  // atomic replace on success
```

For an **image export**, prefer a temp file in the **same directory as the destination** (so
the rename is same-volume and atomic, and a huge image isn't written to `%TEMP%` on a
different/smaller drive). `ReplaceFileW` (or `MoveFileExW(MOVEFILE_REPLACE_EXISTING)`) does the
swap; delete the temp on any failure via `wil::scope_exit`.

## Commands you will need

| Purpose | Command | Expected |
|---|---|---|
| Build plugin | `.\build.ps1 -ProjectName ViewerImgRaw` | exit 0, 0 warnings/errors |
| Full build | `.\build.ps1` | exit 0 |
| Tests | `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` | all pass |

## Scope

**In scope**: `Plugins/ViewerImgRaw/ViewerImgRaw.Export.cpp`
**Out of scope**: the pixel snapshot (plan 014) and dispatch leak (plan 017) — do not regress
them; the encoder/format option logic (`GUID_ContainerFormat*`, palette, `encoderOptions`) must
be preserved byte-for-byte — only the *file open + final rename* changes.

## Git workflow

- Branch: `advisor/016-viewerimgraw-atomic-export`
- Message: `fix(ViewerImgRaw): export atomically via temp file + rename to protect originals`
- Do NOT push/PR unless instructed.

## Steps

### Step 1: Encode to a sibling temp path

In `EncodeBgraToImageFileWic`, compute a temp path in the destination's directory and pass it
to `InitializeFromFilename` instead of `outputPath`. Use `GetTempFileNameW(destDir, L"rsi", 0, ...)`
where `destDir` is the destination's directory.

Two corrections over the naive recipe (see "Review corrections — 2026-06-21" at the top):
1. **`noexcept`-safe path math.** `EncodeBgraToImageFileWic` is declared `noexcept`. Derive the
   directory with `std::wstring` (the function's existing exception profile — it already uses
   `std::format`/`std::wstring`), **not** `std::filesystem::path` (whose constructor/`parent_path()`
   are non-`noexcept`). Keep the trailing separator (`substr(0, lastSep + 1)`) so the result is the
   canonical `GetTempPath`-style directory argument; if there is no separator, report (don't
   overwrite in place).
2. **Arm `cleanupTemp` BEFORE the WIC factory/stream are created.** A `wil::scope_exit` declared
   *after* the `stream` `com_ptr` would destruct *before* it, so its `DeleteFileW` would run while
   the WIC file handle is still open (sharing violation → leaked temp). Declaring it first makes it
   destruct **last**, after every WIC object (and thus the file handle) is gone. Use `DeleteFileW`
   (a `noexcept` C API) for cleanup — no `std::filesystem::remove`, no `<system_error>`.

```cpp
const size_t lastSeparator = outputPath.find_last_of(L"\\/");
if (lastSeparator == std::wstring::npos)
{
    outStatusMessage = L"ViewerImgRaw: Export path has no directory component.";
    return E_INVALIDARG;
}
const std::wstring destDirectory = outputPath.substr(0, lastSeparator + 1);   // keep trailing separator

wchar_t tempNameBuffer[MAX_PATH]{};
if (GetTempFileNameW(destDirectory.c_str(), L"rsi", 0, tempNameBuffer) == 0)
{
    const DWORD lastError = GetLastError();   // capture once
    outStatusMessage = std::format(L"...create temp export file... (hr=0x{:08X})", ...);
    return HRESULT_FROM_WIN32(lastError);
}

const std::wstring tempPath = tempNameBuffer;
bool committed              = false;
auto cleanupTemp            = wil::scope_exit([&tempPath, &committed]() noexcept
{
    if (! committed)
    {
        static_cast<void>(DeleteFileW(tempPath.c_str()));   // best-effort; original untouched
    }
});
// ... CreateStream, then InitializeFromFilename(tempPath.c_str(), GENERIC_WRITE) ...
// ... all existing encoder/frame setup, WriteSource, frame->Commit, encoder->Commit ...
```

No new `#include` is needed: `GetTempFileNameW`/`DeleteFileW`/`MAX_PATH`/`MoveFileExW` come from
`<windows.h>` (via `Helpers.h`), `wil::scope_exit` is already used in this file, and `<filesystem>`
is **not** required. Keep every existing failure path returning its current `hr`/message;
`cleanupTemp` removes the temp automatically on those returns.

**Verify**: `.\build.ps1 -ProjectName ViewerImgRaw` → exit 0.

### Step 2: Atomically replace the destination on success

After `encoder->Commit()` succeeds, release the **entire WIC object graph that pins the stream**
(not just the stream), then rename temp → destination and set `committed = true` only after a
successful rename.

**Critical correction (see top-of-file note):** `encoder->Initialize(stream.get(), ...)` makes the
`IWICBitmapEncoder` hold an AddRef on the `IWICStream` for its lifetime, and the `frame` holds the
encoder. `encoder->Commit()` flushes the bytes but does **not** close the file handle. Resetting
*only* `stream` therefore leaves the handle open (the encoder still pins it), and `MoveFileExW`
fails with `ERROR_SHARING_VIOLATION` on every export. Release `frame`, `encoder` (and, harmlessly,
`source`/`bitmap`/`paletteToSet`) **before** `stream.reset()`:

```cpp
// encoder->Commit() flushed the bytes, but the encoder still holds the IWICStream (and the frame
// holds the encoder), so the file handle stays open. Release the whole graph before renaming.
frame.reset();
encoder.reset();
source.reset();
bitmap.reset();
paletteToSet.reset();
stream.reset();

if (MoveFileExW(tempPath.c_str(), outputPath.c_str(),
                MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH) == 0)
{
    const DWORD lastError = GetLastError();   // capture once (std::format below can clobber it)
    outStatusMessage = std::format(L"ViewerImgRaw: Failed to finalize export (hr=0x{:08X}).",
                                   static_cast<unsigned long>(HRESULT_FROM_WIN32(lastError)));
    return HRESULT_FROM_WIN32(lastError);   // cleanupTemp removes the temp; original untouched
}
committed = true;   // temp has been renamed onto the destination — do not delete it
return S_OK;
```

`MOVEFILE_COPY_ALLOWED` is **dropped**: the temp is a same-directory sibling (same volume), so the
rename is a pure atomic metadata move; `COPY_ALLOWED` would only matter for a cross-volume move,
which would be a *non-atomic* copy+delete that defeats the point. `MoveFileExW(MOVEFILE_REPLACE_EXISTING)`
(not `ReplaceFileW`) is correct because the destination may be brand-new (`ReplaceFileW` requires it
to already exist). Capture `GetLastError()` once — reading it twice risks the intervening `std::format`
clobbering the thread-error.

**Verify**: `.\build.ps1` → exit 0.

## Test plan

**Decision: cover-by-inspection (option C), matching the merged plan 014 precedent.** A real
ViewerPETests case is **not feasible** without new plumbing: `EncodeBgraToImageFileWic` has
internal linkage (anonymous namespace) and is reachable only through `BeginExport`, which first
drives a **modal** `IFileSaveDialog` before the destination is known. `ViewerPETests` talks to the
plugin only via the `IViewer` COM vtable, which exposes no export/encode/save entry point, so it
cannot reach the encode without either automating the modal dialog or adding a production seam —
both of which would violate done-criterion "Only `ViewerImgRaw.Export.cpp` modified". This is the
exact wall plan 014 hit and waived.

Therefore:
- The failure-path safety (`cleanupTemp` + `committed` gate, full handle release before rename) is
  verified **by reading the final function** and by adversarial multi-agent review.
- The **Full suite** (incl. `ViewerPETests`, which loads `ViewerImgRaw.dll` and exercises the
  viewer) is the regression gate: `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild`.
- **Deferred follow-up** (shared with plan 014): add a ViewerImgRaw export regression test once a
  non-dialog export seam or modal-dialog automation exists.

## Done criteria

ALL must hold:

- (archived, not live)  `.\build.ps1 -ProjectName ViewerImgRaw` exits 0, 0 warnings/0 errors.
- (archived, not live)  `.\build.ps1` exits 0.
- (archived, not live)  `EncodeBgraToImageFileWic` writes to a temp file and only `MoveFileExW`/`ReplaceFileW`s
      over the destination after `encoder->Commit()` succeeds.
- (archived, not live)  Every failure path leaves the destination untouched and removes the temp (verify the
      `wil::scope_exit` cleanup + `committed` flag logic by reading the final function).
- (archived, not live)  `grep -n "InitializeFromFilename(outputPath" Plugins/ViewerImgRaw/ViewerImgRaw.Export.cpp`
      returns nothing (the dest is no longer opened directly).
- (archived, not live)  `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` passes.
- (archived, not live)  Only `ViewerImgRaw.Export.cpp` modified.
- (archived, not live)  `plans/README.md` updated.

## STOP conditions

- Excerpts don't match live code (drift).
- The WIC stream cannot be released before rename without restructuring the function beyond
  this file — report rather than leaking the handle or forcing a sharing-violation retry loop.
- `GetTempFileNameW` in the destination directory fails for a legitimate reason (e.g. the
  destination is itself a virtual/non-Win32 path) — in that case report; the export target for
  ViewerImgRaw is a real save-dialog path, so this should not happen, but do not silently fall
  back to overwriting in place.

## Maintenance notes

- The same "write temp, fsync/commit, atomic rename, delete-temp-on-failure" pattern should be
  the template for any future save/export in the viewer plugins (ViewerWeb already follows it).
- Reviewer: confirm the temp lands on the destination volume (same-dir temp), the stream handle
  is closed before rename, and no failure path can leave the destination truncated.
