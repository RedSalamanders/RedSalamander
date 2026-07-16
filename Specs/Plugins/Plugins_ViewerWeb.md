# ViewerWeb Plugin Specification

## Overview

**ViewerWeb.dll** is an **optional** viewer plugin DLL that embeds **Microsoft Edge WebView2**.
It exposes three logical viewer plugins (one DLL, three plugin IDs):

- `builtin/viewer-web`: isolated HTML/PDF viewer. HTML is served in memory from a private viewer origin; every PDF is copied from the active provider reader to a constrained staging file before WebView2 sees it.
- `builtin/viewer-json`: JSON/JSON5/JSONL viewer (pretty highlighted text by default; optional tree view; structured JSONL log cards).
- `builtin/viewer-markdown`: Markdown viewer with syntax highlighting (renders an internal HTML page).

## Invocation (Host Integration)

ViewerWeb is invoked via the standard viewer association mechanism:
- **Shortcut**: `F3` (View)
- Host chooses the viewer by file extension using `extensions.openWithViewerByExtension` (see `Specs/Plugins/Plugins_ViewerPlugins.md` and `Specs/Core/Core_SettingsStore.md`).

Deployment:
- The host loads the optional DLL from `<exeDir>\\Plugins\\ViewerWeb.dll` when present.
- ViewerWeb.dll uses `RedSalamanderEnumeratePlugins` together with `RedSalamanderCreate(pluginId)` to expose the three viewer IDs.
- ViewerWeb.dll depends on the WebView2 loader (`WebView2Loader.dll`) which must be deployed next to the executable/plugin (non-static loader).
  - This project deploys it from vcpkg into the plugin output folder as part of the build.
  - Alternative: use the WebView2 **static loader** approach described by Microsoft (not used by default in this repo).

Runtime requirements:
- Microsoft Edge **WebView2 Runtime** must be installed on the target machine.

Module lifetime:
- ViewerWeb owns a DLL-global shared `ICoreWebView2Environment` so multiple viewer instances do not each pay WebView2 environment startup cost.
- ViewerWeb exports `RedSalamanderPluginShutdown()` to reset that shared environment, unregister the viewer window class, release its WIL-owned class brushes, and retry staged-file deletion as the required module quiet point. The hook is idempotent. `DllMain` process detach deliberately performs no COM release, GDI/container destruction, or staged-file I/O because those operations are forbidden under the loader lock.
- WebView2 callbacks are protected by the host's unload protocol rather than by a callback-owned `FreeLibrary` reference. Callback construction records the WebView2 owner STA; every callback object contributes to a DLL-global live count, and invocation/final release must occur on that STA. A thread-affinity violation permanently votes not-unloadable for that DLL image. Environment/controller creation callbacks also capture a shared-environment generation so stale callbacks self-drop after shutdown/rediscovery. Removing the callback-owned module pin eliminates the self-unmap hazard inside `ComCallback::Release()`. After the last callback is deleted on the owner STA, a release epoch forces the first same-STA zero-live `RedSalamanderPluginCanUnloadNow()` poll to return false; only a later same-STA deferred-unload poll may return true, at which point the releasing stack is necessarily gone.
- Load and Save-As threadpool work has a separate active-worker gate. Each work item acquires a module pin before queueing, transfers it at callback entry with `FreeLibraryWhenCallbackReturns`, and decrements the worker gate only at the callback epilogue. `CanUnloadNow()` requires zero WebView2 callbacks, zero active workers, completed class/brush shutdown cleanup, and no retained staged-file cleanup.
- `RedSalamanderPluginRetainModuleUntilProcessExit()` votes true only when the same quiet-state check is still false during process shutdown. Runtime refresh uses `CanUnloadNow()` deferral; after a later clean poll the host may explicitly unload the module.

## Supported Extensions

Intended associations:
- `builtin/viewer-web`: `.html`, `.htm`, `.pdf`
- `builtin/viewer-json`: `.json`, `.json5`, `.jsonl`, `.ndjson`
- `builtin/viewer-markdown`: `.md`

## Configuration

ViewerWeb exposes a per-plugin configuration schema (`GetConfigurationSchema`) and accepts configuration via `SetConfiguration`.

Keys (defaults):
- `maxDocumentMiB` (`32`, `1..64`): maximum accepted source size. It applies to private-origin HTML, JSON/Markdown input, all staged PDF reads, and Save As source copies; the hard schema ceiling remains 64 MiB.
- `viewMode` (`"pretty"`): JSON rendering mode (`"pretty"`, `"tree"`, or `"jsonl"`).
- `allowExternalNavigation` (`false`): allow an explicit, user-initiated `http://` / `https://` link to open in the system browser. It never permits an external document to replace viewer content.
- `devToolsEnabled` (`false`): allow opening WebView2 DevTools.

Notes:
- If ViewerWeb is missing/disabled, the host falls back to `builtin/viewer-text`.
- Settings are per-plugin-ID (`builtin/viewer-web` vs `builtin/viewer-json` vs `builtin/viewer-markdown`).
- Top-level, frame, and new-window requests use the same allow-list policy. Only the exact active private viewer URL or exact active PDF file URL, optionally followed by a fragment, may remain in the WebView. Query/path/scheme changes, frame escapes, redirects, and non-user-initiated external requests are canceled. An explicit user-initiated HTTP(S) request is canceled in-view and opened through the system browser only when `allowExternalNavigation = true`.
- Document isolation is fail-closed. Before the first navigation and on every route transition, ViewerWeb must successfully apply the script, web-message, and DevTools settings. It must also successfully register top-level, frame, new-window, internal-resource filter, and internal-resource response hooks before navigation. Any failure discards/closes the partially configured controller and leaves a terminal initialization error; content is never navigated with inherited settings or missing policy hooks.
- Once a private-origin document is requested, failure to inspect the request, build the in-memory response, or install that response also closes the controller. ViewerWeb never lets its synthetic `.invalid` origin fall through to proxy/DNS/network resolution.
- An allowlisted external request is launched only after `put_Cancel(TRUE)` or `put_Handled(TRUE)` succeeds. If WebView2 cannot suppress the in-view/new-window request, ViewerWeb closes the controller, returns the policy error, and does not launch an external duplicate.
- `.jsonl` / `.ndjson` files auto-prefer the JSONL card view even when `viewMode` is left at the default `"pretty"`.
- If a `.json` / `.json5` document fails as a single JSON value but parses as multiple JSON values separated by newlines, the viewer falls back to the JSONL card view.
- Every HTML/PDF/JSON/Markdown/Save-As provider read requires a successful `GetSize()`, seeks to byte zero and verifies the returned position, treats the reported size as an exact commitment, and bounds every request/returned count. Failed `Seek`/`Read`, premature EOF, and impossible over-return terminate the operation; a one-byte EOF probe rejects trailing data. A filesystem that understates/overstates its size therefore cannot bypass `maxDocumentMiB`, fabricate bytes, or cause an out-of-bounds append/write.
- UTF-16 JSON/Markdown input is converted to UTF-8 incrementally. Conversion stops before the UTF-8 output exceeds the configured byte limit; ViewerWeb never retains a full raw buffer, full UTF-16 copy, and full UTF-8 copy together.
- Generated-page publication has a second deterministic ceiling: 1 MiB fixed renderer allowance plus twice the configured input-byte limit, capped at 128 MiB. If renderer assets/escaping exceed that ceiling, the page is rejected before `_pendingDocumentUtf8` is published.
- JSONL parsing additionally caps retained card count and aggregate retained strings under that publication budget, supplies yyjson with bounded parse/write allocators, and appends escaped fields directly into the final HTML instead of constructing a second full `entriesJs` buffer. Card titles, controls, record counts, and generated value summaries come from localized `IDS_*` resources and are inserted through the JavaScript-string escaper.
- If a JSON/Markdown document exceeds `maxDocumentMiB`, ViewerWeb keeps the error message visible and offers to reopen the same file in `builtin/viewer-text`. Oversized HTML/PDF remains an error in ViewerWeb.

## Content Isolation and Navigation Security

Raw HTML:
- Local and virtual-filesystem HTML bytes are read through `IFileSystemIO`, bounded by advertised and running byte limits, and returned through `WebResourceRequested` at an exact per-load URL under `https://viewer.redsalamander.invalid`.
- Raw HTML is never navigated from its original/staged `file://` path and is never written to a temporary HTML file.
- The raw response carries `Content-Type: text/html` without a forced charset so WebView2 can honor a BOM or document metadata. Its CSP sandbox uses `default-src 'none'`, `script-src 'none'`, `connect-src 'none'`, and explicit frame/object/base/form denial. Inline CSS and `data:` images are the only passive content allowances. WebView2 document scripting, web messages, and DevTools are disabled for this route.

Generated JSON/JSONL/Markdown:
- Generated pages use the same private origin but a separate CSP that permits only the bundled inline renderer and inline styles. Network connections, frames, objects, bases, and forms remain denied.
- WebView2 scripting is enabled only for these generated pages. `devToolsEnabled` has an effect only on this route.

PDF:
- Path syntax is never treated as proof that an `IFileSystem` path identifies the same host-local file. Both Win32-looking and virtual PDF paths are read through `IFileSystemIO` and staged, closing the confused-deputy route where a custom provider could cause WebView2 to open an unrelated host `C:\...` or UNC path.
- ViewerWeb creates `%TEMP%\rsw-{GUID}.pdf` with `CREATE_NEW` and at most 32 collision retries; `StringFromGUID2` supplies the literal braced GUID, and the implementation never falls back to a `.tmp` extension. The copy uses the exact-size reader contract above, writes every byte, flushes and closes before navigation, and remains within `maxDocumentMiB`. ViewerWeb tracks the staged path through completion and active use, deletes it on stale completion/switch/close, and submits it to the OS delayed-delete queue if WebView2 still holds the file at the teardown quiet point. If immediate deletion and delayed scheduling both fail, a module cleanup queue retains the path and keeps `CanUnloadNow()` false while retrying on later loads, shutdown, and unload polls.
- Document scripting, web messages, and DevTools are disabled on the staged PDF route.

Save As:
- The save dialog runs on the UI thread, but provider acquisition, sizing, seek/read, and destination I/O run only in a module-pinned threadpool callback. Close or a newer save increments the atomic request generation and never waits for a provider call that is blocked; the worker owns the viewer/file-system lifetimes it needs and touches only atomic viewer state off-thread.
- Save As enforces the configured and 64 MiB hard ceilings plus the exact-size reader contract. It creates `<destination-directory>\.rsw-save-{GUID}.tmp` with `CREATE_NEW` and at most 32 collision retries, handles partial `WriteFile` completions, flushes and closes the temp, releases the source reader (including for same-path saves), then commits with `MoveFileExW(..., MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)`. Every pre-commit failure closes provider/destination resources and deletes only the temp before publishing completion, leaving the source and any pre-existing destination byte-identical.
- A sharing violation or other atomic-replacement failure against an open/locked destination is a commit failure: ViewerWeb deletes the sibling staging file and leaves every pre-existing destination byte unchanged.
- Completion carries viewer, HWND, and save-generation identity. A stale completion is ignored. If payload posting fails, an allocation-free atomic terminal record plus payload-free message/paint fallback clears the current UI state; closing the viewer never joins the worker. Debug builds expose the same production primitive through `DebugSaveAsRequest` (with write/flush/commit fault bits) so on-disk safety tests do not automate the modal dialog.

Debug control ABI (`ENABLE_TESTS`):
- Registered message `RedSalamander.ViewerWeb.DebugControl.1` takes `DebugControlAction` in `wParam`: `1` arms one load-completion post failure, `2` arms one save-completion post failure, and `3` submits Save As.
- Action `3` requires `lParam = DebugSaveAsRequest*`, `sizeBytes >= sizeof(DebugSaveAsRequest)`, and a non-empty `destinationPath`. `faultMask` accepts only `0x1` (write), `0x2` (flush), and `0x4` (commit), singly or combined; unknown bits fail submission with `E_INVALIDARG`. The handler returns `TRUE` for a structurally valid request and writes the actual queueing result to `submissionHr`.
- Registered message `RedSalamander.ViewerWeb.DebugSnapshot.1` takes `DebugSnapshot*` in `lParam`; its Save-As fields are `saveInProgress` and cumulative `asyncSavePostFailures`. Load-post fallback is exposed as cumulative `asyncLoadPostFailures` plus `loadPostFailureTerminal`.

Diagnostics emit `viewer.web.load_bytes`, `viewer.web.rejected_bytes`, `viewer.web.normalized_bytes`, `viewer.web.output_bytes`, and `viewer.web.save_as_bytes` scenario metrics for accepted/rejected input, conversion/output size, and committed exports. The Debug snapshot also exposes load/save completion-post failure counts and Save-As in-progress state.

Focused regression coverage lives in `ViewerPETests`:
- `TestViewerWebSecurityPolicyAndBounds` covers both CSPs, exact/fragment URL policy, false defaults, running-byte arithmetic, bounded UTF conversion/output, and fault-injected staged-cleanup retry retention.
- `TestViewerWebVirtualHtmlUsesPrivateOriginAndEnforcesByteCaps` loads hostile UTF-16-BOM HTML through a virtual filesystem and observes the actual private WebView2 source and disabled script setting; it also proves exact-size/over-return/output rejection and successful constrained local/virtual PDF staging, collision-safe `.pdf` paths, and quiet-point cleanup.
- `TestViewerWebTransactionalSaveAsAndCloseSafety` exercises same-path and pre-existing-destination success plus byte-for-byte destination preservation for a locked destination, seek mismatch, failed provider `Read`, premature EOF, trailing data, impossible over-return, and injected write/flush/commit failure, with no sibling staging residue. It also covers the allocation-free completion-post fallback, requires a later quiescent `RedSalamanderPluginCanUnloadNow()` observation after the callback release epoch changes, and proves sub-500 ms close, a false unload vote, and continued DLL mapping while a module-pinned Save-As callback is blocked in provider `Read`; after release, the canceled worker must remove its temp, preserve the destination, retire provider state, and only then allow the callback-return pin to unload the module.

## UI / UX

Layout:
- **Header**: in standalone mode only, filename dropdown (combo box) listing `otherFiles` when `otherFileCount > 1` (ViewerText-style); the collapsed combo uses compact chrome so the WebView/content surface keeps as much vertical space as possible. Embedded preview mode hides this standalone header/dropdown and lets the WebView/content surface blend into the preview/tab background.
- **Content**: WebView2 surface.
- Header status text is a DirectWrite/Direct2D-only render path. If the DX path cannot draw the status message, ViewerWeb deliberately shows no fallback GDI status text and logs one error so the failure is visible in diagnostics.
- Internal HTML pages (`json`, `jsonl`, `markdown`) theme their own scrollbars from the current viewer colors, including nested code-block scrollers.
- Raw HTML and generated `json`, `jsonl`, and `markdown` pages are delivered to WebView2 through in-memory `WebResourceRequested` responses on private viewer URLs, so rendered documents do not rely on `NavigateToString()` or a temp `.html` file.

Menu (DxUi-hosted from the hidden native menu model):
- File: Save As, Refresh, Exit, Other Files navigation (Next/Previous/First/Last)
- Search: Find, Find Next, Find Previous
- View: Zoom In/Out/Reset, Toggle DevTools
- Tools: Copy URL, Open in Browser, JSON Expand/Collapse, Toggle Markdown Source

The window detaches its live native `HMENU` after opening and renders the visible top menu bar through the shared `RedSalamander.DxNativeMenuBar` host. `Alt`, `F10`, and menu mnemonics continue to route through that DxUi menu bar.

Keyboard focus opens on the web/content surface. If focus moves to the filename dropdown or menu, that focused control owns its keyboard navigation; `Esc` from that chrome returns focus to the main content surface instead of closing. Accepted peer-file selections or menu commands return focus to the main content surface unless the command intentionally opens an editable find/control surface. `Esc` from the main content surface posts `WM_CLOSE` when the viewer is idle.

JSON viewer modes:
- `pretty`: syntax-highlighted formatted JSON text.
- `tree`: read-only `jsoneditor` tree view with `Expand All` / `Collapse All`.
- `jsonl`: structured log cards for one JSON object per line, with per-entry badges (`level`, `category`, timestamp when present) and lazy syntax-highlighted payload expansion.
- JSON syntax colors are theme-aware but intentionally separated by token kind: keys, string values, numeric values, and boolean/null literals each use a distinct palette.

Oversized document fallback:
- Trigger: JSON/Markdown file exceeds the in-memory size limit (`maxDocumentMiB`).
- UX: ViewerWeb prompts with the size-limit error text plus a Yes/No offer to open the document in Text Viewer.
- On Yes: the host opens `builtin/viewer-text` for the same file-system path and ViewerWeb closes its own window after the handoff succeeds.

## Theme / Rainbow

- Uses `IViewer::SetTheme()` to apply colors (background/text/selection/accent) and DPI.
- For internal pages (JSON/Markdown), the plugin injects theme variables and updates them dynamically via `ExecuteScript`.
- When `theme.rainbowMode` is enabled, the window title bar uses a rainbow accent and the page accent color is derived from the host theme + file identity.

## Third-Party Components (Non-GPL)

ViewerWeb embeds the following JavaScript libraries as resources:
- **jsoneditor 10.4.2** (Apache-2.0) for read-only JSON tree view.
- **markdown-it 14.1.0** (MIT) for Markdown rendering.
- **highlight.js 11.11.1** (BSD-3-Clause) for syntax highlighting in Markdown, JSON pretty view, and expanded JSONL payloads.

License material shipped in-tree:
- `Plugins/ViewerWeb/ThirdParty/jsoneditor.LICENSE.txt`
- `Plugins/ViewerWeb/ThirdParty/jsoneditor.NOTICE.txt`
- `Plugins/ViewerWeb/ThirdParty/markdown-it.LICENSE.txt`
- `Plugins/ViewerWeb/ThirdParty/highlightjs.LICENSE.txt`

JSONL implementation note:
- No new third-party viewer library was added for JSONL. The JSONL view is built in the plugin and reuses the existing permissively licensed `highlight.js` dependency for payload coloration.

## Dependency Maintenance

Vendored web assets:
- Minified assets live under `Plugins/ViewerWeb/Assets/`.
- When updating a vendored asset, replace the minified file, verify the embedded version banner/header, and refresh the matching license/notice text under `Plugins/ViewerWeb/ThirdParty/` if upstream changed it.
- Keep this spec in sync with the vendored asset versions after each update.

vcpkg-managed dependencies:
- Native dependencies (including `webview2`, `yyjson`, `wil`, `curl`, `libraw`, and AWS SDK components) are versioned in `vcpkg.json`.
- To review available updates, refresh the vcpkg baseline (`vcpkg x-update-baseline`) and inspect the resulting version changes before committing.
- After bumping `vcpkg.json`, rebuild with `build.ps1` and verify the viewer plugins still load and the `WebView2Loader.dll` deployment step still succeeds.
