# ViewerPE Specification

Last updated: 2026-07-12

## Purpose

`builtin/viewer-pe` is the built-in Portable Executable viewer for Windows binaries such as `.exe`, `.dll`, `.sys`, and related PE/COFF files.

It extends the shared viewer contract in `Specs/Plugins/Plugins_ViewerPlugins.md`.

## Window And UI Contract

- Standalone ViewerPE opens as a normal top-level viewer window with the shared DxUi menu bar.
- The main surface is the parsed PE text/content viewport.
- In standalone mode, ViewerPE shows the filename dropdown only when `otherFiles` contains more than one peer file; the dropdown uses the shared compact 28 DIP combo height to leave more vertical room for parsed data.
- In embedded preview mode, ViewerPE hides the standalone filename dropdown and menu/title chrome so the parsed content blends into the preview/tab background.

## Keyboard Contract

- Keyboard focus opens on, and returns to, the parsed-content viewport after peer-file selection or menu commands.
- While focus is inside the filename dropdown or menu, that focused control owns arrow/Enter behavior; `Esc` from that chrome returns focus to the parsed-content viewport.
- `Esc` from the parsed-content viewport posts `WM_CLOSE` and closes the idle viewer.
- Peer navigation follows the shared viewer rules: next/previous/first/last commands use the `otherFiles` list and preserve the same viewer instance where possible.

## Data Contract

- The viewer reads the focused file as bytes and presents best-effort PE metadata, including DOS header, machine/subsystem, timestamp, sections, and import/export-related information when parseable.
- User-facing report narration, section headings, column headings, state words, truncation notices, and field labels come from the ViewerPE resource module so both the rendered and exported reports follow the active plugin language. Standardized PE/COFF identifiers and values (for example `PE32+`, `e_lfanew`, machine/subsystem names, and data-directory enum names) remain their canonical technical tokens.
- Parse failures are reported as viewer content/status rather than modal dialogs.
- PE parsing runs through one module-pinned threadpool scheduler per viewer instance. The scheduler permits at most one active request and one replaceable pending request; a newer open/navigation request invalidates the active generation and replaces, rather than queues behind, any older pending request.
- The whole-file parser rejects inputs larger than 256 MiB before allocating its byte buffer; the limit is reported through the normal in-view parse error path.
- After a successful `GetSize`, the reader must seek to byte zero and report position zero. That size is a commitment: premature EOF, bytes beyond the reported size, a seek-position mismatch, or a `Read` result larger than its requested buffer is invalid provider data. File reads are requested in chunks no larger than 1 MiB. The worker checks the request generation and window identity between chunks and at bounded parse-enumeration checkpoints.
- `Close()` and window teardown invalidate pending/active generations without waiting for an in-flight `IFileReader::Read` to return. The worker owns no `ViewerPE*`, keeps the plugin DLL mapped until it exits, and transfers its last pin to `FreeLibraryWhenCallbackReturns(...)` so the callback cannot unmap its own code before returning. It posts only identity-tagged results; a completion targeting a stale or recycled `HWND` is ignored and its payload is reclaimed.
- A current request must reach a terminal UI state even if its result allocation or payload post fails. While parsing, a bounded UI timer polls an allocation-free failure record in the scheduler; successful identity-bound completion or cancellation stops that timer. A destroyed/stale window is not a current request and requires no UI delivery.
- The worker emits `viewer.pe.queue` and `viewer.pe.parse` performance records; `Close()` emits `viewer.pe.close` with the `cancel-without-wait` detail.

## Testing Contract

- `ViewerPETests` covers standalone DxUi combo-host behavior, legacy ComboBox absence, UIA exposure, embedded filename-combo hiding, and clean close behavior.
- `TestViewerPELatestWinsAndCloseDoesNotWaitForBlockedRead` reports a 256 MiB + 1 source and requires rejection before any `Read`; it also rejects a nonzero initial seek position, premature EOF, bytes beyond the committed `GetSize`, and an impossible over-return. With the first valid `Read` blocked, the test proves that the middle request is never opened, the latest pending request runs after release, and every read request stays within `1 MiB`. Result-allocation and payload-post faults must each leave loading through the allocation-free terminal fallback without closing or recycling the current window. Its final blocked-read phase requires sub-500 ms close/window destruction, callback-return-safe module-pin retention after caller/window release, and clean DLL unload only after the blocked callback exits.
