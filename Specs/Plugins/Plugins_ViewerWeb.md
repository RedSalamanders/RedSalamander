# ViewerWeb Plugin Specification

## Overview

**ViewerWeb.dll** is an **optional** viewer plugin DLL that embeds **Microsoft Edge WebView2**.
It exposes three logical viewer plugins (one DLL, three plugin IDs):

- `builtin/viewer-web`: HTML/PDF viewer (navigates to file:// URLs or a temp extract for non-Win32 paths).
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

## Supported Extensions

Intended associations:
- `builtin/viewer-web`: `.html`, `.htm`, `.pdf`
- `builtin/viewer-json`: `.json`, `.json5`, `.jsonl`, `.ndjson`
- `builtin/viewer-markdown`: `.md`

## Configuration

ViewerWeb exposes a per-plugin configuration schema (`GetConfigurationSchema`) and accepts configuration via `SetConfiguration`.

Keys (defaults):
- `maxDocumentMiB` (`32`, `1..512`): maximum size for in-memory loads (JSON/Markdown).
- `viewMode` (`"pretty"`): JSON rendering mode (`"pretty"`, `"tree"`, or `"jsonl"`).
- `allowExternalNavigation` (`true`): allow navigating to `http://` / `https://` links (Web/Markdown).
- `devToolsEnabled` (`false`): allow opening WebView2 DevTools.

Notes:
- If ViewerWeb is missing/disabled, the host falls back to `builtin/viewer-text`.
- Settings are per-plugin-ID (`builtin/viewer-web` vs `builtin/viewer-json` vs `builtin/viewer-markdown`).
- `.jsonl` / `.ndjson` files auto-prefer the JSONL card view even when `viewMode` is left at the default `"pretty"`.
- If a `.json` / `.json5` document fails as a single JSON value but parses as multiple JSON values separated by newlines, the viewer falls back to the JSONL card view.
- If a JSON/Markdown document exceeds `maxDocumentMiB`, ViewerWeb keeps the error message visible and offers to reopen the same file in `builtin/viewer-text`.

## UI / UX

Layout:
- **Header**: in standalone mode only, filename dropdown (combo box) listing `otherFiles` when `otherFileCount > 1` (ViewerText-style); embedded preview mode hides this standalone header/dropdown and lets the WebView/content surface blend into the preview/tab background.
- **Content**: WebView2 surface.
- Header status text is a DirectWrite/Direct2D-only render path. If the DX path cannot draw the status message, ViewerWeb deliberately shows no fallback GDI status text and logs one error so the failure is visible in diagnostics.
- Internal HTML pages (`json`, `jsonl`, `markdown`) theme their own scrollbars from the current viewer colors, including nested code-block scrollers.
- Generated `json`, `jsonl`, and `markdown` pages are delivered to WebView2 through an in-memory `WebResourceRequested` response on a private viewer URL, so large rendered documents do not rely on `NavigateToString()` or a temp `.html` file.

Menu (DxUi-hosted from the hidden native menu model):
- File: Save As, Refresh, Exit, Other Files navigation (Next/Previous/First/Last)
- Search: Find, Find Next, Find Previous
- View: Zoom In/Out/Reset, Toggle DevTools
- Tools: Copy URL, Open in Browser, JSON Expand/Collapse, Toggle Markdown Source

The window detaches its live native `HMENU` after opening and renders the visible top menu bar through the shared `RedSalamander.DxNativeMenuBar` host. `Alt`, `F10`, and menu mnemonics continue to route through that DxUi menu bar.

Keyboard focus opens on the web/content surface. If focus moves to the filename dropdown or menu, that focused control owns its keyboard navigation; accepted peer-file selections or menu commands return focus to the main content surface unless the command intentionally opens an editable find/control surface.

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
