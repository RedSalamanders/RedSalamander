# ViewerWeb Plugin Specification

## Overview

**ViewerWeb.dll** is an **optional** viewer plugin DLL that embeds **Microsoft Edge WebView2**.
It exposes three logical viewer plugins (one DLL, three plugin IDs):

- `builtin/viewer-web`: HTML/PDF viewer (navigates to file:// URLs or a temp extract for non-Win32 paths).
- `builtin/viewer-json`: JSON/JSON5 viewer (pretty highlighted text by default; optional tree view).
- `builtin/viewer-markdown`: Markdown viewer with syntax highlighting (renders an internal HTML page).

## Invocation (Host Integration)

ViewerWeb is invoked via the standard viewer association mechanism:
- **Shortcut**: `F3` (View)
- Host chooses the viewer by file extension using `extensions.openWithViewerByExtension` (see `Specs/PluginsViewer.md` and `Specs/SettingsStoreSpec.md`).

Deployment:
- The host loads the optional DLL from `<exeDir>\\Plugins\\ViewerWeb.dll` when present.
- ViewerWeb.dll uses the optional multi-plugin exports (`RedSalamanderEnumeratePlugins`, `RedSalamanderCreateEx`) to expose the three viewer IDs.
- ViewerWeb.dll depends on the WebView2 loader (`WebView2Loader.dll`) which must be deployed next to the executable/plugin (non-static loader).
  - This project deploys it from vcpkg into the plugin output folder as part of the build.
  - Alternative: use the WebView2 **static loader** approach described by Microsoft (not used by default in this repo).

Runtime requirements:
- Microsoft Edge **WebView2 Runtime** must be installed on the target machine.

## Supported Extensions

Intended associations:
- `builtin/viewer-web`: `.html`, `.htm`, `.pdf`
- `builtin/viewer-json`: `.json`, `.json5`
- `builtin/viewer-markdown`: `.md`

## Configuration

ViewerWeb exposes a per-plugin configuration schema (`GetConfigurationSchema`) and accepts configuration via `SetConfiguration`.

Keys (defaults):
- `maxDocumentMiB` (`32`, `1..512`): maximum size for in-memory loads (JSON/Markdown).
- `viewMode` (`"pretty"`): JSON rendering mode (`"pretty"` or `"tree"`).
- `allowExternalNavigation` (`true`): allow navigating to `http://` / `https://` links (Web/Markdown).
- `devToolsEnabled` (`false`): allow opening WebView2 DevTools.

Notes:
- If ViewerWeb is missing/disabled, the host falls back to `builtin/viewer-text`.
- Settings are per-plugin-ID (`builtin/viewer-web` vs `builtin/viewer-json` vs `builtin/viewer-markdown`).

## UI / UX

Layout:
- **Header**: filename dropdown (combo box) listing `otherFiles` when `otherFileCount > 1` (ViewerText-style).
- **Content**: WebView2 surface.

Menu (themed/owner-drawn):
- File: Save As, Refresh, Exit, Other Files navigation (Next/Previous/First/Last)
- Search: Find, Find Next, Find Previous
- View: Zoom In/Out/Reset, Toggle DevTools
- Tools: Copy URL, Open in Browser, JSON Expand/Collapse, Toggle Markdown Source

## Theme / Rainbow

- Uses `IViewer::SetTheme()` to apply colors (background/text/selection/accent) and DPI.
- For internal pages (JSON/Markdown), the plugin injects theme variables and updates them dynamically via `ExecuteScript`.
- When `theme.rainbowMode` is enabled, the window title bar uses a rainbow accent and the page accent color is derived from the host theme + file identity.

## Third-Party Components (Non-GPL)

ViewerWeb embeds the following JavaScript libraries as resources:
- **jsoneditor** (Apache-2.0) for optional JSON tree view.
- **markdown-it** (MIT) for Markdown rendering.
- **highlight.js** (BSD-3-Clause) for code highlighting (Markdown) and JSON pretty view.
