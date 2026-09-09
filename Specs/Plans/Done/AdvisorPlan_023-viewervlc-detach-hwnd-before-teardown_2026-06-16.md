# Advisor Plan 023 - ViewerVLC retained drawable HWNDs

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/023-viewervlc-detach-hwnd-before-teardown.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/023-viewervlc-detach-hwnd-before-teardown.md`
- **Live owner:** Specs/Plans/Done/Operation_Farsight_ViewerPluginsRemediation_ExportSafetyDecodeReliabilityWebSecurityAndTextGeometry_2026-06-16.md
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 023: ViewerVLC — retain hidden drawable HWNDs until asynchronous stop/release completes

> **Frozen historical instructions (do not execute)**: The filename and the original “detach HWND” title are historical. The
> `set_hwnd(player, nullptr)` design was rejected after validating the live libVLC integration
> contract. Do not implement it. This 2026-07-12 reconciliation is the required architecture: hide
> and retain the parent/video HWNDs, clean the complete VLC state asynchronously, and destroy the HWNDs
> only after identity-bound cleanup reaches its quiet point.

## Status

- **State**: DONE — corrected implementation, focused cleanup, installed-VLC proof, consolidated archive, and repository-wide Full gate green (2026-07-12)
- **Priority**: P2 (intermittent close-time crash, worse with hardware decoders)
- **Effort**: L after teardown/fallback/unload expansion
- **Risk**: MED–HIGH (drawable, window, worker, and DLL lifetime)
- **Category**: concurrency / lifetime
- **Planned at**: commit `b274022d9`, 2026-06-16

## Historical defect

Ordinary Win32 teardown destroyed the video child while async cleanup could still be inside
`libvlc_media_player_stop` or release. libVLC's video-output thread could therefore touch a dead HWND.
The original advisor plan proposed calling `libvlc_media_player_set_hwnd(player, nullptr)` from
`WM_DESTROY`. That prescription is rejected: the authoritative integration contract does not permit a
null drawable while the player/vout is live, and `WM_DESTROY` is already too late to build a robust
identity/lifetime protocol.

## 2026-07-12 corrected architecture

### Close path

`Close()` begins a deferred close and returns without waiting:

1. invalidate load generation and stop UI/loading/HUD timers and input producers;
2. exit fullscreen and hide the owning parent, video child, HUD, and overlay surfaces;
3. retain the parent/video HWND identities in their `wil::unique_hwnd` owners;
4. move every drawable-owning `VlcState` whole into asynchronous stop/release;
5. count pending load and cleanup work;
6. destroy/reset the owning window only after load/cleanup/deferred queues reach zero.

No production call passes `nullptr` to `libvlc_media_player_set_hwnd`.

### State and module lifetime

The worker stops the player and releases player → media → instance → module in the complete moved
`VlcState`. Module pins transfer with `FreeLibraryWhenCallbackReturns` at callback entry. Ordinary
completion payloads own the viewer reference, HWND, request ID, per-window identity, and module pin.
Only the matching live generation decrements its counters or completes close; stale completion owns
and retires its state without touching a recycled/newer HWND.

### Persistent cleanup dispatcher

`Open()` pre-creates a persistent `PTP_WORK` cleanup dispatcher and its module pin. It owns an
allocation-free intrusive deferred-state queue. If cleanup-result allocation, callback submission,
payload posting, or a per-operation module pin fails, the complete `VlcState` is spliced into that
queue. `SubmitThreadpoolWork` guarantees progress without allocating another work item, waiting on the
UI thread, retrying from a timer, or requiring the closing HWND to survive. The dispatcher:

- holds a viewer reference while scheduled;
- acquires/transfers a callback-return module pin;
- drains every deferred state and records allocation-free completion aggregates;
- uses bounded `SendMessageTimeoutW` only to notify a still-matching identity-bound HWND;
- remains safe when an embedded owner destroys its parent during delayed cleanup.

Residual stale/undeliverable load-result states route through this same dispatcher rather than
stopping inline in a destructor or on the UI thread.

## Reconciled scope

This correction spans `ViewerVLC.cpp`/`.h`, shared Debug contracts, resources, ViewerPETests, source
contracts, performance instrumentation, and the authoritative ViewerVLC spec. The original
“only `ViewerVLC.cpp` modified” criterion is obsolete as of 2026-07-12. The historic filename is kept
only so existing references remain valid.

## Required proof

The focused delayed-cleanup regression must cover:

- non-blocking close with parent/video surfaces hidden but still identity-valid;
- request + window identity rejection under close/reopen and HWND-reuse pressure;
- exactly-once accepted cleanup and viewer-close notification;
- forced cleanup-result allocation, callback-submit, and payload-post failure;
- persistent-dispatcher progress with no surviving HWND dependency;
- destruction of an embedded parent while delayed dispatcher cleanup is active;
- module mapping through callback return and clean unload after all work retires;
- real playback when VLC is installed.

The injected cleanup/identity/unload path is implemented and focused-green. Consolidated proof is
archived at `Specs/TestRuns/4cb089111a23/Viewers/2026-07-12_105903_farsight_viewer_closeout/`: all
13 focused cases passed with zero conditional skips and 1,426 performance metric records. The
installed-VLC playback/Tab/HUD case ran rather than skipping and passed.

## Done criteria

- [x] The rejected `set_hwnd(player, nullptr)` teardown design is absent and documented as forbidden.
- [x] Close stops producers, hides surfaces, and retains owning parent/video HWNDs until cleanup quiets.
- [x] `VlcState` moves whole through asynchronous stop/release with correct destruction order.
- [x] Request + per-window identity prevents stale/recycled-HWND completion.
- [x] A persistent pre-created `PTP_WORK` dispatcher drains allocation-free deferred cleanup.
- [x] Allocation/submit/post/module-pin failures never stop VLC synchronously on the UI thread.
- [x] Callback-return-safe module lifetime and delayed embedded-parent destruction are focused-tested.
- [x] Close/destroy notification and owning-HWND destruction are exactly once for accepted close.
- [x] Scoped zero-warning build and focused cleanup/config/source-contract regressions pass.
- [x] Expanded shared/test/resource/spec scope is recorded; the old one-file-only criterion is retired.
- [x] `plans/README.md` reflects the corrected retained-drawable design.
- [x] Refreshed installed-VLC playback/Tab/HUD proof is green where VLC is present.
- [x] Consolidated VLC/Farsight cleanup metrics archive is recorded.
- [x] `.\Tools\Run-AllTests.ps1 -Suite Full` passes in the final closeout workspace (`Specs/TestRuns/4cb089111a23/Continuation/2026-07-12_205636_farsight_full_green/`).

## STOP conditions

- A proposal destroys/releases the drawable before stop/release completes.
- A proposal passes a null drawable while the player/vout is live.
- Failure handling performs synchronous player stop/release on the UI thread.
- Deferred cleanup needs allocation, a retry timer, or a surviving HWND to make progress.
- The plugin module could unmap before a cleanup callback returns.

## Maintenance notes

The invariant is: **hide and retain drawable identity; stop/release asynchronously; destroy only at
the quiet point**. External owner destruction is an abnormal path that still must leave state and DLL
lifetime safe, but it does not justify weakening the normal deferred-close ordering.
