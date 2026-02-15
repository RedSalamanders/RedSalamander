# Localization and Resources Specification

## Goal

All user-facing UI text must be localizable. Static UI structure (menus, context menus) must be defined in `.rc` resources; runtime code should only populate truly dynamic content (theme lists, per-user history entries, etc.).

## Requirements

- **No hardcoded user-facing strings in C++** (except dynamic data such as file paths, folder names, WSL distro names, etc.).
- **Static menus and context menus must be declared in `.rc`** and referenced by resource IDs.
- Runtime code may only extend **dynamic** menu parts:
  - Themes discovered from `Themes\\*.theme.json5` next to the executable
  - Custom themes from settings (`theme.themes[]`)
  - Folder history entries (global MRU)
- Use resource helpers from `Common/Helpers.h`:
  - `LoadStringResource()`
  - `FormatStringResource()`
  - `MessageBoxResource()`

## RedSalamander menus

- Main menu resource lives in `RedSalamander/RedSalamander.rc`.
- `View → Theme` contains fixed built-in theme items in resources and is extended at runtime:
  - Built-ins: `builtin/system`, `builtin/light`, `builtin/dark`, `builtin/rainbow`, `builtin/highContrast` (app-level).
  - A disabled **system high contrast indicator** item may be present; it is not selectable.
  - File themes from `Themes\\*.theme.json5` (separator above and below this section when present).
  - Custom themes from settings (`user/*`) (separated from file themes when both sections exist).
- FolderView context menu is a resource menu (`IDR_FOLDERVIEW_CONTEXT`) defined in `RedSalamander/RedSalamander.rc`.

## RedSalamanderMonitor menus

- Main menu resource lives in `RedSalamanderMonitor/RedSalamanderMonitor.rc`.
- `View → Theme` follows the same rules as RedSalamander: fixed built-ins in resources, dynamic file themes from `Themes\\*.theme.json5`, and dynamic settings themes from `theme.themes[]`.

