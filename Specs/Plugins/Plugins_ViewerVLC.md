# ViewerVLC Plugin Specification

## Overview

`ViewerVLC.dll` exposes `builtin/viewer-vlc`, the optional libVLC-backed audio/video viewer. The plugin owns the RedSalamander window, loading, HUD, focus, configuration, diagnostics, and teardown contracts; libVLC supplies media decode and playback.

The viewer supports standalone and embedded preview windows. A compatible embedded `Open()` reuses the existing viewer/window and, when the resolved VLC install and instance arguments are unchanged, keeps the existing libVLC instance/player and switches media with `libvlc_media_player_set_media`.

## Configuration Boundary

Configuration is stored under `plugins.configurationByPluginId["builtin/viewer-vlc"]`. Supported keys include the VLC install path and auto-detection flag, caching values, playback rate, hardware decode mode, audio/video output modules, audio visualization, volume/mute persistence, and `extraArgs`.

- `avcodecHw` is restricted to the schema values `any`, `none`, `dxva2`, and `d3d11va`.
- `audioVisualization` is restricted to the schema values.
- `videoOutput` and `audioOutput` are optional bounded module tokens. They may contain only ASCII letters, digits, `_`, `-`, and `.`, may not start with `-` or `.`, and may not contain whitespace, control characters, option syntax, or path separators.
- `extraArgs` is trusted-user configuration. It is never derived from a media file or filesystem metadata. Arguments must be bounded, well-formed `--name[=value]` options. Options that change plugin/search/config paths or request externally selected interfaces, modules, codecs, access/demux chains, stream output, filters, service discovery, keystores, or input slaves are rejected. `extraArgs` also may not override the separately validated vout, aout, hardware-decode, or audio-visual fields.
- Dangerous-option matching removes one leading `no-` case-insensitively before applying the deny list, so negated or mixed-case forms cannot bypass the boundary. File logging, logfile destinations, and pidfile options are rejected in both positive and `no-` forms.
- Signed and unsigned JSON numbers are clamped in their native domains before conversion to the bounded 32-bit caching, playback-rate, and volume settings. An unsigned value above `INT64_MAX` must not wrap negative before clamping.
- Invalid option configuration returns `E_INVALIDARG` without replacing the last accepted configuration.

## Asynchronous Load and Identity

The UI thread may copy the path, immutable configuration, current reusable-instance identity, and window identity into a work item. All potentially blocking media/install work runs on the module-pinned worker:

- local-file existence/type probes;
- configured-install normalization and validation;
- registered VLC install paths under the machine/user VideoLAN registration and Windows `FOLDERID_ProgramFiles` / `FOLDERID_ProgramFilesX86` locations. Auto-detection never calls `SearchPathW`, searches the current directory or process `PATH`, or derives a global search root from environment variables;
- libVLC argument construction;
- full-path `LoadLibraryExW`, export resolution, and `libvlc_new`.

`TrySubmitThreadpoolCallback` failure records a localized terminal UI error directly on the submitting UI thread. It must not run any probe, DLL load, export resolution, or `libvlc_new` synchronously as a fallback.

Every posted load/close result owns a non-null `ViewerVLC.dll` module pin for its entire queued lifetime. Failure to acquire that pin is terminal and routes through the payload-free completion path instead of posting an unpinned payload. Payload-post failure is also terminal: any undelivered player-owning state is transferred to the cleanup dispatcher, a small request/window-identity completion is recorded in a mutex-protected allocation-free aggregate, and a bounded `SendMessageTimeoutW` notification is attempted for a surviving HWND. The UI decrements the matching pending count and either shows the localized terminal load error or completes close; a failed payload post may not strand loading, a hidden HWND, or `ViewerClosed` notification.

The test-only debug snapshot exposes the outstanding asynchronous load-work count in addition to accepted/stale completion counters. Cleanup-gate timing assertions must begin only after unrelated startup/load work has reached zero, so the gate measures retirement cleanup rather than worker-pool scheduling of earlier opens.

Every load has a monotonically increasing request ID and captures a nonzero per-window identity. Every result owns its `VlcState` until the UI accepts it. Acceptance requires the same viewer, current window identity, and current request ID. A stale result is stopped/released without attaching a drawable or starting playback. The parent window remains alive while a close is waiting for already-accepted work, so posted results can be drained through the identity check rather than targeting a recycled HWND.

When the worker resolves the same normalized install and instance-argument key as the live reusable instance, it returns a reuse decision without constructing another libVLC instance. The UI revalidates that a live instance still exists before switching media; if the instance was retired while the result was in flight, the viewer starts a new full asynchronous load instead of treating the missing reuse target as success.

## DLL Search and Module Lifetime

The configured directory is validated first when present. Auto-detection then accepts only a validated registered VideoLAN install or the validated `VideoLAN\VLC` child of a Windows Program Files known folder. `libvlc.dll` is loaded by full path with `LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_DEFAULT_DIRS`. ViewerVLC never mutates the process DLL directory with `SetDllDirectoryW`, `AddDllDirectory`, or related restore APIs.

Queued load and cleanup work pins `ViewerVLC.dll`. Each threadpool callback transfers its work pin at callback entry with `FreeLibraryWhenCallbackReturns`, so releasing the work context cannot unmap callback code before return. The pre-created cleanup dispatcher has a persistent module pin and acquires/transfers a callback-return pin on every invocation. A `VlcState` owns its loaded libVLC module and destroys player, media, and instance objects before releasing that module. Moving the entire state to cleanup therefore keeps both plugin and libVLC code mapped through stop/release without a process search-path mutation or manual raw-module lifetime.

## Close and Drawable Lifetime

ViewerVLC must never pass a null drawable while a player/video-output path is live. Close follows this order:

1. Mark the generation closed, stop timers/input producers, remove wheel forwarding, and hide the parent, video, HUD, and overlay surfaces.
2. Keep the parent and video HWND identities alive and parented exactly as they were while moving every player-owning `VlcState` to module-pinned asynchronous stop/release.
3. Count already-running load and state-cleanup work. Each completion carries the viewer pointer, HWND, window identity, and request identity and is accepted only by its owning window generation.
4. After all load and cleanup counts reach zero and no state remains, destroy the owning `wil::unique_hwnd` once on the UI thread. `WM_NCDESTROY` releases child-handle observations and invokes `ViewerClosed` once.

Before playback can begin, ViewerVLC creates one persistent `PTP_WORK` cleanup dispatcher and gives it a persistent `ViewerVLC.dll` module pin. If either resource cannot be created, `Open()` fails before a libVLC state exists. Cleanup submission/allocation/module-pin failure never falls back to blocking stop/release on the UI thread: the whole `VlcState` is spliced into an allocation-free intrusive queue and the pre-created work is submitted with `SubmitThreadpoolWork`, which has no per-submit failure return. A self-reference keeps the viewer alive until the dispatcher drains and releases every player/media/instance/module in the background. The normal close path keeps the parent/video drawable HWNDs hidden, parented, and valid until that drain reaches its identity-bound UI completion. Calling `Close()` again is idempotent. An `Open()` while deferred close is pending returns `ERROR_BUSY`; after close completion the same viewer object may open a new window with a new identity.

Cleanup-completion payload-post failure uses the same identity-bound allocation-free completion aggregate and bounded notification. The worker never destroys an HWND or calls the host callback; only the UI fallback decrements cleanup ownership and reaches the normal once-only destroy/notify path.

`ENABLE_TESTS` cleanup control may supply a manual-reset release event instead of relying on a fixed sleep to keep a retiring preview root alive. ViewerVLC duplicates the event into WIL-owned state and shares that lifetime with direct and dispatcher cleanup work; only a worker may wait on it, the wait is bounded, and every selftest that supplies the gate must signal it from `wil::scope_exit`. This deterministic gate protects preview child-coexistence coverage under every supported selftest timeout multiplier without introducing a UI-thread wait.

External forced HWND destruction first clears the dispatcher's notification HWND/identity, drains posted payloads, and retires every residual state. Cleanup itself does not depend on a timer, posted allocation, or surviving HWND. The dispatcher self-reference and module pins keep `ViewerVLC.dll` mapped until delayed stop/release finishes; only then may the last viewer reference and DLL unload complete. The normal host/API close path still uses the retained-drawable sequence above.

## Performance and Diagnostics

ViewerVLC emits:

- `viewer.vlc.queue` for accepted, rejected, stale, and cleanup queue outcomes;
- `viewer.vlc.load_us` for worker-side probe/load duration;
- `viewer.vlc.close_ui_us` for the synchronous UI portion of `Close()`;
- `viewer.vlc.cleanup_us` for worker stop/release duration.

The close contract is responsiveness-sensitive: injected slow stop/release must leave `Close()` non-blocking, with hidden retained HWNDs until cleanup completes. Regression evidence must archive the metrics under `Specs/TestRuns/`.

## Required Regression Coverage

`ViewerPETests` must cover:

- deterministic loader-submit failure reaching a terminal UI state without synchronous loader work;
- deterministic load-result and cleanup-result payload-post failures reaching their identity-bound terminal UI fallbacks without stranding pending counts;
- deterministic cleanup allocation and submit failures draining through the pre-created dispatcher while `Close()` remains non-blocking;
- forced parent destruction with delayed cleanup proving immediate child-HWND teardown, continued plugin mapping while callback code is active, and clean unload after the HWND-independent dispatcher completes;
- delayed superseding loads rejecting stale request/window identities without stale attach/play;
- parent/video HWND retention and hiding during delayed cleanup, followed by UI-thread destruction and one close callback;
- repeated open/close/reopen cycles with a new per-window identity, including HWND-reuse pressure;
- safe option-token acceptance, native-domain unsigned numeric clamping, and separator/control/option/module-loading/logfile/pidfile rejection;
- source-contract proof that discovery uses the registered VideoLAN install and Program Files known folders, never `SearchPathW`, the current directory, process `PATH`, or environment-derived search roots;
- real dependency loading, instance/player creation, WAV playback, unchanged process DLL directory, and clean close when VLC is installed.

The source contract additionally bans a null drawable, process-global DLL-directory APIs, `SearchPathW`, UI-thread install probing, synchronous loader fallback, inline `VlcAsyncLoadResult` player stop/release, and callback-body module-pin destruction.

Focused implementation evidence is archived at `Specs/TestRuns/SINON/Viewers/2026-07-11_124749_viewervlc_fs13_fs16_terminal_post_fallback/`. It contains the passing trace plus `viewer.vlc.*` JSONL rows, including deterministic `load-post-fallback-terminal` and `cleanup-post-fallback-terminal` outcomes. Across seven delayed close cycles the UI close portion remained at or below 15,231 us while worker stop/release reached 309,872 us.
