# Advisor Plan 021 - ViewerPE non-waiting parse

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/021-viewerpe-no-ui-thread-join.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/021-viewerpe-no-ui-thread-join.md`
- **Live owner:** Specs/Plans/Done/Operation_Farsight_ViewerPluginsRemediation_ExportSafetyDecodeReliabilityWebSecurityAndTextGeometry_2026-06-16.md
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 021: ViewerPE — non-waiting latest-wins parsing on slow or hostile providers

## Status

- **State**: DONE — implementation, focused proof, consolidated archive, and repository-wide Full gate green (2026-07-12)
- **Priority**: P2 (UI hang on slow or blocked virtual/network providers)
- **Effort**: L after reliability closeout expansion
- **Risk**: MED (worker, window, and DLL lifetime)
- **Category**: performance / concurrency / provider reliability
- **Planned at**: commit `b274022d9`, 2026-06-16

## Historical defect

ViewerPE stored its parser in a `std::jthread`. Reassigning or clearing the member from refresh,
navigation, or `WM_NCDESTROY` requested stop and then joined on the UI thread. `IFileReader::Read` is
not cancellable, so one blocked cloud/network read could freeze close indefinitely. The old parser
also used 16 MiB chunks and had no module lifetime that would make a naive detach safe.

## 2026-07-12 implementation reconciliation

The original proposal to copy another viewer's self-retaining callback shape was refined. ViewerPE
now owns a detached scheduler state whose worker inputs are independent of the viewer object:

- one parse is active and one replaceable latest request may be pending;
- close invalidates request/window generations and drops pending ownership without waiting;
- whole-file allocation is rejected above 256 MiB and reads are chunked at no more than 1 MiB;
- the reader must seek to byte zero and return position zero;
- successful `GetSize` is an exact commitment: early EOF, trailing bytes, and over-return are errors;
- completion carries request ID plus per-window identity, preventing a stale or recycled-HWND apply;
- current-result allocation or payload-post failure records an allocation-free terminal slot consumed
  by a bounded UI timer, so loading always terminates;
- the callback transfers its module pin with `FreeLibraryWhenCallbackReturns` at callback entry.

No worker callback dereferences `ViewerPE`, no `std::jthread` remains, and close/navigation never join.

## Reconciled scope

Implementation and proof necessarily span `ViewerPE.cpp`/`.h`, the shared Debug snapshot in
`Common/WindowMessages.h`, `Tests/ViewerPETests`, source contracts, performance instrumentation, and
authoritative ViewerPE/plugin documentation. The original “only `ViewerPE.cpp`/`.h` modified” done
criterion is obsolete as of 2026-07-12.

Rendered PE report semantics are out of scope except where localization work in plan 029 applies.

## Implementation requirements

1. Keep the detached scheduler at one active + one replaceable latest pending request.
2. Check generation before every read and bounded parsing stage; stale work exits silently.
3. Use the exact-reader contract and the 256 MiB / 1 MiB limits above.
4. Post ordinary completion with `PostMessagePayload`; preserve payload drain on `WM_NCDESTROY`.
5. Use the allocation-free scheduler fallback for current allocation/post failure.
6. Keep the DLL mapped through callback return, not merely through the last callback statement.
7. Never add a UI-thread wait as a cleanup fallback.

## Focused proof

`TestViewerPELatestWinsAndCloseDoesNotWaitForBlockedRead` covers:

- size-cap rejection before source reads;
- nonzero seek result, short data, trailing data, and impossible over-return;
- one active request, middle-request replacement, and latest request completion;
- forced result-allocation and payload-post terminal fallback;
- `Close()` below 500 ms while the provider remains blocked;
- module mapping while blocked and clean unload after the callback-return boundary.

Existing standalone/embedded parse and navigation cases remain green. Scoped ViewerPE and
ViewerPETests builds are zero-warning.

Consolidated proof is archived at
`Specs/TestRuns/4cb089111a23/Viewers/2026-07-12_105903_farsight_viewer_closeout/`: all 13 focused
cases passed with zero conditional skips, and the archive contains 1,426 performance metric records.

## Done criteria

- [x] No UI-thread `join()` executes during open, refresh, navigation, close, or destroy.
- [x] One active + one replaceable latest pending request is enforced.
- [x] Worker state outlives close safely without dereferencing the viewer.
- [x] Exact reader, 256 MiB allocation ceiling, and ≤1 MiB chunks are enforced.
- [x] Request + window identity rejects stale/recycled-window completion.
- [x] Allocation/post failure reaches exactly one terminal current-request state.
- [x] Callback-return-safe module lifetime and clean focused unload are proven.
- [x] Scoped zero-warning builds and focused regressions pass.
- [x] Expanded diagnostics/test/spec scope is recorded; the old two-file-only criterion is retired.
- [x] `plans/README.md` reflects this reconciliation.
- [x] Consolidated Farsight metrics/archive evidence is refreshed.
- [x] `.\Tools\Run-AllTests.ps1 -Suite Full` passes in the final closeout workspace (`Specs/TestRuns/4cb089111a23/Continuation/2026-07-12_205636_farsight_full_green/`).

## STOP conditions

- A worker needs to dereference the viewer after close.
- Module lifetime cannot be held through callback return.
- A current allocation/post failure can leave loading active.
- Any proposed fallback reintroduces a UI-thread wait.

## Maintenance notes

The viewers share safety invariants, not an interchangeable scheduler. Plan 030 deliberately retired
a generic async helper because PE's detached latest-wins state differs from ImgRaw, Text, Web, VLC,
and SQLite ownership/resource contracts.
