# ViewerPE Specification

Last updated: 2026-05-15

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
- Parse failures are reported as viewer content/status rather than modal dialogs.

## Testing Contract

- `ViewerPETests` covers standalone DxUi combo-host behavior, legacy ComboBox absence, UIA exposure, embedded filename-combo hiding, and clean close behavior.
