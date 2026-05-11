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
- Formatted resource strings use `std::format` syntax and MUST use positional placeholders such as `{0}` and `{1:08X}` so translators can reorder arguments. Bare `{}` and unindexed format specs such as `{:08X}` are forbidden in `.rc` resources.
- Embedded/source resource strings MUST introduce placeholders in argument order (`{0}`, then `{1}`, then `{2}`, etc.) with no skipped indexes. The `FormatStringResource(...)` argument list MUST follow that same source-string order. Translated satellite strings MAY reorder placeholders for grammar, but they MUST use the same placeholder tokens as the source string; translations must not add, drop, duplicate, renumber, or change format specs such as `:L` or `:08X`.
- If a formatted resource string has invalid `std::format` syntax, the helper returns an empty fallback string and logs the failing resource ID plus the format error detail once through diagnostics. `std::bad_alloc` remains fatal and must not be swallowed by the formatting fallback.
- Command labels have two localized forms:
  - Full display names (`IDS_CMD_*`) for menus, Preferences, and shortcut lists.
  - Short display names (`IDS_CMD_SHORT_BASE + IDS_CMD_*`) for compact surfaces such as the function bar; every command must provide one.
  - Resource ids `20000..21999` are reserved for command short labels so the arithmetic mapping cannot collide with unrelated strings.

## Satellite resource DLLs

- English resources remain embedded in the owning executable or plugin DLL.
- Optional translated resources live in per-owner, per-culture resource-only projects under `Lang\<culture>\`.
- Satellite DLLs are named `<OwnerName>-<culture>.dll`, where `<OwnerName>` matches the owning executable or plugin DLL stem, for example `RedSalamander-fr-FR.dll`, `RedSalamanderMonitor-fr-FR.dll`, or `ViewerText-fr-FR.dll`.
- Language resource projects output to `.build\<Platform>\<Configuration>\Lang\`. Normal executable and plugin binary output locations do not change.
- Language resource projects must be resource-only DLLs and must not introduce executable entry points.

## RedConfigure authoring

`RedConfigure.exe` may generate satellite `.rc` files for a selected resource owner and culture. Generated files must keep resource text UTF-16/resource-compiler safe, use deterministic ordering, and preserve positional `std::format` placeholders. RedConfigure must block export when a target translation introduces bare `{}`, unindexed specs such as `{:08X}`, printf-style placeholders, or a placeholder set that does not match the source string exactly.

Resource forms that RedConfigure can parse but cannot safely rewrite yet must remain visible as inventory and must not be written silently.

## Runtime resource lookup

- Each resource owner registers its embedded `HINSTANCE` with the localization manager before localized resources are requested.
- `RedSalamander` and `RedSalamanderMonitor` register their resource owners during startup after settings load and before UI resource lookup.
- Plugins register their resource owner immediately after the plugin DLL loads and unregister it before the module unloads.
- The persisted language setting selects either the Windows preferred UI language chain (`system`) or a concrete BCP-style culture tag such as `fr` or `fr-FR`.
- Lookup tries the selected culture chain in satellite DLLs first, then falls back to the owner module's embedded English resource.
- Missing satellite DLLs, missing satellite resource IDs, and invalid or unavailable culture selections must not block startup or UI creation; embedded English is the fallback.
- App/Common string helper calls route through the localization manager. Plugins that opt in with `REDSAL_USE_COMMON_LOCALIZATION` use the same string helper route.
- Menus, accelerators, dialogs, and other non-string resources that are migrated to satellite support must use localization manager resource-loading helpers instead of direct owner-module loads.
- Non-string loads use the localization helpers that match their resource shape: `LoadMenuResource`, `LoadAcceleratorsResource`, `FindLocalizedResourceHandle`, `LoadResourceBytes`, and `LoadImageResource`.
- Code that needs `SizeofResource`, `LoadResource`, or `LockResource` must use the `HINSTANCE` returned with the localized resource handle. Do not mix an `HRSRC` found in a satellite DLL with the embedded owner module handle.
- Resource handles returned from `FindLocalizedResourceHandle` are transient lookup results. Callers must load or copy the resource immediately and must not cache those handles across language changes or owner unregister.
- Dialog template creation must use the resource-aware Win32 callback helpers so `DialogBoxIndirectParamW` and `CreateDialogIndirectParamW` receive the selected satellite template when one exists.
- Resource-aware dialog helpers must copy template bytes from the selected satellite resource but create the dialog with the embedded owner module `HINSTANCE`, so custom child classes registered by the executable or plugin continue to resolve after switching languages.
- Top-level menus must be loaded explicitly through the localization manager. Window classes must not rely on `lpszMenuName` for localizable menus because class-template loading bypasses the selected satellite.
- Runtime language changes must re-apply the selected language, rebuild already-created top-level menu resources from the current owner, reset cached submenu handles, rebuild dynamic menu sections, and refresh the visible custom menu model. Existing modal dialogs may continue using the language they were opened with until reopened.

## RedSalamander menus

- Main menu resource lives in `RedSalamander/RedSalamander.rc`.
- The main window owns a detached main menu handle for the custom DxUI menu bar. When the language preference changes, the menu handle must be replaced from `LoadMenuResource`, dynamic theme/plugin sections rebuilt, and the DxUI menu model synchronized.
- `View → Theme` contains fixed built-in theme items in resources and is extended at runtime:
  - Built-ins: `builtin/system`, `builtin/light`, `builtin/dark`, `builtin/rainbow`, `builtin/highContrast` (app-level).
  - A disabled **system high contrast indicator** item may be present; it is not selectable.
  - File themes from `Themes\\*.theme.json5` (separator above and below this section when present).
  - Custom themes from settings (`user/*`) (separated from file themes when both sections exist).
- FolderView context menu is a resource menu (`IDR_FOLDERVIEW_CONTEXT`) defined in `RedSalamander/RedSalamander.rc`.

## RedSalamanderMonitor menus

- Main menu resource lives in `RedSalamanderMonitor/RedSalamanderMonitor.rc`.
- The monitor main menu must be attached from `LoadMenuResource` during startup so the persisted language applies before the first window is shown.
- `View → Theme` follows the same rules as RedSalamander: fixed built-ins in resources, dynamic file themes from `Themes\\*.theme.json5`, and dynamic settings themes from `theme.themes[]`.

