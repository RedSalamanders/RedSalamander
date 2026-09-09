# Advisor Plan 020 - ViewerText hex short-read

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/020-viewertext-hex-short-read.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/020-viewertext-hex-short-read.md`
- **Live owner:** Specs/Plans/Done/Operation_Farsight_ViewerPluginsRemediation_ExportSafetyDecodeReliabilityWebSecurityAndTextGeometry_2026-06-16.md
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 020: ViewerText hex view — handle short reads so the buffer never shows phantom zero bytes

> **Frozen historical instructions (do not execute)**: Follow step by step; run every verification before continuing.
> Historical only. Root plans/ was deleted; do not execute.
>
> **Drift check (run first)**: `git diff --stat b274022d9..HEAD -- Plugins/ViewerText/ViewerText.Hex.cpp`
> If changed, compare excerpts to live code first; mismatch ⇒ STOP.

## Status

- **Priority**: P1 (data integrity — viewer fabricates file content)
- **Effort**: S
- **Risk**: LOW
- **Depends on**: none
- **Category**: bug
- **Planned at**: commit `b274022d9`, 2026-06-16
- **Status**: **DONE** — executed 2026-06-21 on branch `advisor/020-viewertext-hex-short-read`
  (opus, ultracode multi-agent review). **Scope expanded beyond the plan after review** (operator-approved):
  see "Execution record" below.

## Execution record (2026-06-21)

A 5-dimension adversarial review of this plan (drift, Step-1 correctness, Step-2 feasibility,
RefillHexCache sibling, test harness) surfaced two things the plan missed, so the fix shipped
**three** source edits, not just Step 1:

1. **Step 1 (as planned)** — `LoadHexData` in-memory loop now `resize`s `_hexBytes` down to the bytes
   actually read (`ViewerText.Hex.cpp`). ✅
2. **Step-1-introduced heap OOB read (new, REQUIRED)** — Step 1 alone makes the search-highlight paint
   path read past the end of `_hexBytes`: `searchMask`/`visibleBytes` are sized from `_fileSize`
   (`ViewerText.Hex.cpp:~967`) while `searchBytesPtr` points into the now-shorter `_hexBytes`
   (`:~984`); the scan loop (`:~1016`) then reads `searchBytesPtr[0..searchMask.size())`. Fixed by
   clamping `searchMask` to `_hexBytes.size() - searchMaskStartOffset` in the in-memory branch. Without
   this, Step 1 turns a correctness bug into a memory-safety bug.
3. **Dominant-path fix (new, REQUIRED for the fix to matter)** — `LoadHexData` is the *minority* path.
   `_hexBytes` is normally filled by the async preload in `ViewerText.cpp` (preload loop `~:5497-5514`,
   moved to `_hexBytes` at `~:5799`); `LoadHexData` only runs when `_hexBytes.empty()`
   (`ViewerText.cpp:7832/7846`). The preload loop had the identical short-read-no-resize-down bug and
   now `resize`s down to bytes read. The sibling **fallback** loop (`~:5588-5607`) got the same
   resize-down for consistency, but note it is currently **unreachable dead code** (pre-existing: its
   block is gated by `if (FAILED(result->hr)...)` immediately after `result->hr = S_OK` — the dead
   hex-fallback is tracked separately by plan 028). This expands the fix beyond the plan's stated
   `ViewerText.Hex.cpp`-only scope — operator-approved.

**Step 2** (effective-size scroll/click/item-count bounds): **deferred** — STOP condition confirmed.
There is no single hex-document-size helper; `_fileSize` is threaded through ~25+ geometry sites
(several feed 32/64-bit offset formatting that must not change), and after the fix the residual is
purely cosmetic (a scrollbar/clickable region a few empty rows too long; click/drag already bail on
`validBytes == 0`). Not worth the blast radius.

**RefillHexCache / streaming path**: confirmed **already correct** (`_hexCacheValid = read`, the actual
bytes — `ViewerText.Hex.cpp:~2389`); STOP condition #3 did not trigger; left untouched.

**Test**: `TestViewerTextHexShortReadDropsPhantomTailBytes` added to `Tests/ViewerPETests/ViewerPETests.cpp`
(+ a test-only `ShortReadFileReader` and an opt-in `BuiltinFileSystemStub::EnableShortRead`, both
default-off so sibling tests are unaffected). A fixture reports 16 bytes via `GetSize()` but the reader
yields only 8; the test asserts the hex view's `visibleByteCount == 8` (not 16) through the
`ViewerTextDebugSnapshot` contract. The reader caps by absolute file **position** (not a cumulative
counter) so it survives the encoding BOM probe + the loader's `Seek(0)`. **Regression value proven
empirically**: the test FAILS on a build with the preload fix disabled (`visibleByteCount` settles at
16) and PASSES with it enabled. The plan's original assertion (a) ("search for `\x00` returns no match
in [M,N)") was dropped — it has no observable path through the debug contract; the in-memory search bug
is covered transitively (a shrunk `_hexBytes` has no phantom zeros to scan).

**Coverage note (cover-by-inspection for part 2).** The automated test exercises parts 1+3 (the
resize-to-bytes-read on the dominant preload path). Part 2 (the search-mask OOB clamp) is **not**
automatically covered: driving it needs an *active hex search* over a short-read buffer, and the hex
search needle is entered through the modal Find prompt (the same modal-dialog wall plans 014/016 hit).
Part 2 is therefore verified by inspection plus the adversarial diff review (memory-safety lens: "ship",
high confidence — the scan is provably bounded by `_hexBytes.size()` after the clamp, the subtraction
cannot underflow, and no other consumer indexes `_hexBytes` with an `_fileSize`-derived length). A
follow-up could add a search-over-short-read smoke test under ASan/AppVerifier if a non-modal way to set
the hex needle is introduced.

## Why this matters

`LoadHexData` resizes the in-memory buffer to the **reported** file size up front, then fills it
from `IFileReader::Read` in a loop. On a short read (the loop `break`s when `read == 0`) it does
**not** resize the buffer down to the bytes actually read, leaving a zero-padded tail. Every
downstream consumer treats `_hexBytes.size()` as authoritative, so the viewer renders fabricated
`00` bytes past the true end of data, lets the user select them, and an in-memory search for `00`
returns **phantom matches that aren't in the file**. `IFileReader::Read` does not guarantee
`bytesRead == bytesRequested` on success, so this triggers on truncation races and is *normal* on
virtual/remote filesystems (Curl/S3/GoogleDrive). For a hex viewer — a tool people use precisely
to see exact bytes — silently inventing content is a serious correctness bug.

## Current state

```cpp
// ViewerText.Hex.cpp:2093-2122  (LoadHexData, in-memory path for files <= kMaxHexLoadBytes)
if (_fileSize <= kMaxHexLoadBytes && _fileSize <= static_cast<uint64_t>(std::numeric_limits<size_t>::max()))
{
    _hexBytes.resize(static_cast<size_t>(_fileSize));          // sized to REPORTED size

    uint64_t ignored     = 0;
    const HRESULT seekHr = _fileReader->Seek(0, FILE_BEGIN, &ignored);
    if (FAILED(seekHr)) { _hexBytes.clear(); return seekHr; }

    size_t offset = 0;
    while (offset < _hexBytes.size())
    {
        const unsigned long want = static_cast<unsigned long>(std::min<size_t>(256 * 1024, _hexBytes.size() - offset));
        unsigned long read       = 0;
        const HRESULT readHr     = _fileReader->Read(_hexBytes.data() + offset, want, &read);
        if (FAILED(readHr)) { _hexBytes.clear(); return readHr; }
        if (read == 0) { break; }                              // <-- short read: tail stays zero-padded
        offset += read;
    }
    // <-- BUG: no resize-down to `offset`; _hexBytes keeps _fileSize zero-filled elements
}
```

Why it corrupts every reader:
```cpp
// ViewerText.Hex.cpp:2293-2305  (ReadHexBytes — uses _hexBytes.size() as truth)
if (! _hexBytes.empty())
{
    const uint64_t total = static_cast<uint64_t>(_hexBytes.size());   // == _fileSize, not bytes read
    if (offset >= total) return 0;
    const size_t available = _hexBytes.size() - start;
    const size_t take      = std::min(destSize, available);
    memcpy(dest, _hexBytes.data() + start, take);                     // returns phantom 0x00 for the padded tail
    return take;
}
```
`FormatHexLine`, the in-memory search (`FindHexNeedleForwardInMemory`), and click mapping all flow
through `_hexBytes`/`ReadHexBytes`, so all of them see the phantom tail.

## Commands you will need

| Purpose | Command | Expected |
|---|---|---|
| Build plugin | `.\build.ps1 -ProjectName ViewerText` | exit 0, 0 warnings/errors |
| Full build | `.\build.ps1` | exit 0 |
| Tests | `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` | all pass |

## Scope

**In scope**: `Plugins/ViewerText/ViewerText.Hex.cpp` (`LoadHexData`, and a small consistency tweak
if you choose Step 2).
**Out of scope**: the streaming-cache path (`RefillHexCache`, the `else` branch) — it reads
on-demand per window and is not part of this bug; the text-mode renderer.

## Git workflow

- Branch: `advisor/020-viewertext-hex-short-read`
- Message: `fix(ViewerText): resize hex buffer to bytes actually read on short read`
- Do NOT push/PR unless instructed.

## Steps

### Step 1: Resize the buffer to bytes actually read

After the read loop, shrink `_hexBytes` to the number of bytes actually read so its size equals the
available data (never the zero-padded reported size):

```cpp
    // ... after the while loop, still inside the in-memory branch ...
    if (offset < _hexBytes.size())
    {
        Debug::Warning(L"ViewerText: hex load short read ({} of {} bytes); showing available data.",
                       static_cast<uint64_t>(offset), _fileSize);
        _hexBytes.resize(offset);    // drop the phantom zero tail
    }
```

This alone eliminates the integrity bug: `ReadHexBytes` now reports `available` based on real
bytes, `FormatHexLine` renders only real bytes, and in-memory search no longer matches the padding.
(If `offset == 0`, `_hexBytes` becomes empty and `ReadHexBytes` falls through to the streaming path,
which correctly returns 0 for an unreadable file.)

**Verify**: `.\build.ps1 -ProjectName ViewerText` → exit 0.

### Step 2 (recommended): Keep scroll/click bounds consistent with loaded size

After Step 1, the scroll range and click guard still use `_fileSize` (see `UpdateHexViewScrollBars`
and `OnHexMouseDown`'s `offset >= _fileSize` check), so the user can scroll/click into rows past the
real data (cosmetic — those rows read 0 bytes, no phantom content). To make the document model fully
consistent, have the in-memory path expose the loaded length. The least-invasive approach: where the
hex view computes the total document size for scroll/item-count/click bounds, use
`(! _hexBytes.empty()) ? _hexBytes.size() : _fileSize` instead of bare `_fileSize`.

Read `UpdateHexItemCount`, `UpdateHexViewScrollBars`, and `OnHexMouseDown` and confirm the size each
uses. If the change is localized (a single helper that returns the effective hex document size),
apply it. If it would touch many sites with non-trivial `_fileSize` semantics, STOP and report —
Step 1 already fixes the correctness bug; Step 2 is polish and must not destabilize offset display
(`FormatFileOffset` still uses `_fileSize` for 32- vs 64-bit formatting, which is fine).

**Verify**: `.\build.ps1` → exit 0.

## Test plan

- Add a case to `Tests/ViewerSqliteTests` is wrong target — use `Tests/ViewerPETests/ViewerPETests.cpp`
  (it drives ViewerText, including hex mode and the `IViewerTextDebugSnapshot` debug contract used by
  existing ViewerText cases). Construct a file whose backing reader returns fewer bytes than the
  reported size (a stub `IFileReader` that reports size N but `Read`s only M<N), open it in hex mode,
  and assert that searching for `\x00` does not return matches located in the [M, N) region, and that
  rows past offset M show no data. If wiring a short-read stub reader is not feasible in the harness,
  add a focused test that calls into the hex read path with such a reader (model after how existing
  ViewerText tests obtain the viewer/debug snapshot).
- If neither is feasible without new harness plumbing, STOP and report; the fix is verifiable by
  inspection (buffer size now equals bytes read) plus the build + full-suite gates.
- Verification: `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` → all pass.

## Done criteria

ALL must hold:

- [x] After `LoadHexData`'s in-memory read loop, `_hexBytes.size()` equals the bytes actually read
      (never the zero-padded reported size).
- [x] **And** the dominant async-preload + fallback loops in `ViewerText.cpp` do the same resize-down
      (added during execution — `LoadHexData` is the minority path).
- [x] **And** the search-highlight paint path clamps its scan to `_hexBytes.size()` so Step 1's shrink
      cannot cause a heap out-of-bounds read (added during execution).
- [x] A short read is logged (`Debug::Warning`) and does not leave phantom zeros.
- [~] (Step 2) scroll/click bounds: **deferred** — STOP condition confirmed (no localized helper;
      residual is cosmetic only). See Execution record.
- [x] `.\build.ps1 -ProjectName ViewerText` and `.\build.ps1` exit 0, 0 warnings.
- [x] `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` — run; result recorded in `plans/README.md`
      `execute 020` log (no new failures attributable to this change; pre-existing/environmental only).
- [~] Only `ViewerText.Hex.cpp` modified — **intentionally exceeded** (operator-approved): also
      `ViewerText.cpp` (dominant-path fix) and `Tests/ViewerPETests/ViewerPETests.cpp` (regression test).
- [x] `plans/README.md` updated.

## STOP conditions

- Excerpts don't match live code (drift).
- Step 2 would require broad changes to `_fileSize` semantics — land Step 1 only and report.
- You find the streaming-cache path (`RefillHexCache`) has the *same* short-read padding bug — note
  it and report (it may warrant a sibling fix), but do not expand this plan's scope to it without
  confirming.

## Maintenance notes

- The invariant to preserve: **the hex buffer's size is the count of bytes actually available, not
  the filesystem-reported size.** Any future change to the load loop must keep it.
- Reviewer: confirm `_hexBytes.size()` can never exceed bytes read, and that an empty buffer (total
  failure) cleanly falls through to the streaming path.
