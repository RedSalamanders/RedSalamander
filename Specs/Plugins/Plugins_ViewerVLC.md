# ViewerVLC Plugin Specification

Published playback examples MUST use generated media in a marked test sandbox,
the native no-activation capture helper and an identified real VLC runtime. The
fixture proves decoded positive duration/time, a video child and actual pause
stability; a debug-injected clock or setup-error frame is insufficient. Media and
runtime hashes, observation rows and complete frames accompany the build receipt.
The deterministic fixture records software decoding and `wingdi` output, a generated
12-second/5-fps clip, an MTA capture worker and continued owner-UI pumping. Pixel
population witnesses must prove the video is present. These images do not qualify
all GPU output modes; positive decoder time alone cannot qualify a black frame.
See [the screenshot scenario contract](../Testing/Testing_SelfTests.md#product-documentation-gallery).

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

The UI transfers sole VlcState ownership to the module-pinned worker while probing and preparing a replacement. No UI timer, input handler or diagnostic may query that player while it is worker-owned. The worker preserves a compatible instance/player, creates the new media and performs libvlc_media_player_set_media there: that API can synchronously stop its old input. A new player is also constructed on the worker without attaching a drawable or playing. Only a current request/window completion installs the prepared state, attaches the persistent child and starts playback on the UI thread. A superseded result never attaches or plays. One worker may own the published state, with one latest pending path/configuration; another selection replaces that slot and cancels obsolete preparation rather than adding a worker. Stop, close and forced destruction cancel the slot. All results, including post/submission failures, retain module/state ownership through asynchronous cleanup and clear pending accounting.

## Media Paths and Error Messages

Local Windows media paths must work beyond `MAX_PATH`, including when the Windows
`LongPathsEnabled` policy is disabled. This applies to standalone F3 and embedded
preview, Unicode names, and names containing spaces, `%`, and `#`. Keep the selected
path unchanged for navigation, request identity, and diagnostics. On the existing
loader worker, normalize separators/dot segments and use
`Common::Paths::ToExtendedWin32Path` for the file attribute probe. Preserve the actual
Win32 failure code and selected path in file-access diagnostics; a directory is a
separate non-file error. Do not collapse these failures into `ERROR_FILE_NOT_FOUND`
or a claim that the source is remote.
The probe follows symbolic links, preserving the existing regular-file semantics:
a link to a regular media file remains playable, while a broken link is a file
access error before install discovery. Use `status(path, error_code)` directly;
`exists(path, error_code)` clears native errors for known not-found results.

For long or already extended Windows media paths, pass libVLC a `file:///` location
whose complete UTF-8 extended Windows path is percent-encoded using the shared URI
encoder. This preserves drive/UNC prefixes, Unicode, and literal URI delimiters
through libVLC 3's URI decoding. Do not pass a raw extended path to
`libvlc_media_new_path`: its path-to-URI conversion treats the prefix as a UNC host.
Ordinary short paths continue through `libvlc_media_new_path`. Never truncate paths,
copy media to a short temporary path, or change system policy to make playback work.

The installation-required heading, installation instructions, and VLC download link
appear only when install discovery/validation finds no usable install directory.
File access/type errors, loader submission/delivery errors, DLL/export/init errors,
and playback errors use the localized neutral heading "Unable to play this file."
and their specific detail, without installation advice or a download link. Error
category is carried in the identity-checked worker result, rather than inferred from
localized text. Each open clears the preceding overlay state. The Debug snapshot
exposes the actual rendered heading/body and download-link availability for tests.

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
- `viewer.vlc.apply_load_ui_us`, `viewer.vlc.start_playback_ui_us` and `viewer.vlc.play_ui_us` separate owner-UI acceptance, drawable setup and play;
- `viewer.vlc.prepare_media_worker_us` and `viewer.vlc.set_media_worker_us` record media preparation and replacement on the sole-owner worker;
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
- the keyboard fixture waits for the optional backend loader to retire, then presents the real HUD
  through the test-only playback-state seam without assigning focus. It requires actual visibility
  and nonempty layout before the original video-to-HUD Tab and internal-cycle focus assertions.
  Installed-backend playback proof remains separate and required when available; a missing backend's
  asynchronous error overlay must not race HUD keyboard coverage.
- `TestViewerVlcLongPathsAndErrorMessages`: exact 259-, 260-, and 275-character Unicode/URI-delimiter WAV paths; standalone and embedded playback; already extended and forward-slash spellings; positive decoded duration; missing long files, non-local paths, directories, and missing installations with category-correct headings/details/download links. Guarded same-sandbox links cover a valid long-path target and a broken target; only missing OS symbolic-link privilege may skip these conditional link cases. Use a marked sandbox and the governed no-activation mode; never exercise the user's media as a fixture or change `LongPathsEnabled`.

The source contract additionally bans a null drawable, process-global DLL-directory APIs, `SearchPathW`, UI-thread install probing, synchronous loader fallback, inline `VlcAsyncLoadResult` player stop/release, and callback-body module-pin destruction. Any performance or closeout claim must cite a fresh archived run containing the focused `ViewerPETests` result and the corresponding `viewer.vlc.*` JSONL rows; an unavailable historical archive is not normative proof.

The configured-preview Commands fixture holds the media worker on a duplicated manual-reset event. It retains the 900 ms source-focus/publication bounds while the worker is held, checks one worker plus one coalesced path and no UI-owned player, then releases it and proves the original instance/player is reused for the latest audio request. Decoder settlement remains a separate bounded predicate. Cancellation, close and injected submission/delivery failures must neither strand pending media state nor play obsolete content.

The real-player embedded long-path regression additionally injects media submission and completion-post failures, a failed post with a superseding coalesced request, keyboard Stop while media work is held, and Close while media work is held. Each observes actual branch counters and terminal ownership; Open/Stop/Close retain the 100 ms UI-return budget. Stop remains visibly available while a media load is pending. Close retains the hidden parent/video HWND chain until the held worker and cleanup drain. No activation or user media is required.

During replacement of a worker-owned player, the HUD remains shown and has a nonempty layout even after the loading overlay delay expires. Stop stays enabled, and the HUD is above the loading panel so Stop remains reachable in small viewports. Initial loading without a published player retains its existing overlay behavior. The held real-player controls force the delayed-overlay path and verify actual WS_VISIBLE, layout height and sibling Z order before keyboard Stop or Close.
