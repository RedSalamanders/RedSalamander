---
name: localization
description: Localization and resource management for Windows RC files. Use when adding user-facing strings, creating menus, working with STRINGTABLE, or using LoadStringResource and FormatStringResource helpers.
metadata:
  author: RedSalamander
  version: "1.0"
---

# Localization and Resources

## Mandatory Rules

1. **All user-facing strings** → `.rc` STRINGTABLE
2. **All static menus** → `.rc` MENU resources
3. **Never hardcode UI text** in C++
4. **Always use positional formatting** for variable resource strings to help localization. `.rc` format strings must use `{0}`, `{1:08X}`, etc.; bare `{}` and unindexed specs like `{:08X}` are forbidden.
5. **Keep source placeholders in argument order.** Embedded/source strings introduce placeholders as `{0}`, then `{1}`, then `{2}` with no skipped indexes; pass `FormatStringResource(...)` arguments in that same order.
6. **Keep translations placeholder-equivalent.** Satellite strings may reorder placeholders for grammar, but must not add, drop, duplicate, renumber, or change format specs compared with the source string.
7. **Never treat a resource string as a printf-format string** (`%s`, `%d`, etc.). Use `FormatStringResource(...)` with `std::format`-style positional placeholders and avoid C4774 suppressions.
8. **Run `ResourceLocalizationContracts.Tests.ps1`** when adding or editing resource strings with braces.

## Resource Helpers (Common/Helpers.h)

```cpp
// Load string from resources
std::wstring title = LoadStringResource(IDS_APP_TITLE);

// Format with arguments
std::wstring msg = FormatStringResource(IDS_FILE_COUNT, fileCount);

// Load a documented language-neutral string from the owner module only
std::wstring token = LoadEmbeddedStringResource(g_hInstance, IDS_PLUGIN_PROTOCOL_NAME);

// Format a documented language-neutral skeleton from the owner module only
std::wstring codepage = FormatEmbeddedStringResource(g_hInstance, IDS_CODEPAGE_FORMAT, codepageId);

// Message box from resources
MessageBoxResource(_hWnd, IDS_ERROR_MSG, IDS_ERROR_TITLE, MB_OK | MB_ICONERROR);
```

## Language-Neutral Resource Rule

Stable non-translatable tokens still live in resources, but only in the
embedded `.rc` for the executable or plugin that owns them. Do not duplicate
those IDs in satellite `.rc` files. Load them with `LoadEmbeddedStringResource`
or `FormatEmbeddedStringResource`, passing the owner `HINSTANCE`; plugin code
must pass the plugin DLL handle such as `g_hInstance`, not `nullptr`.

Examples: language autonyms in a language picker, product/protocol/brand names,
file format identifiers, keyboard glyphs, sample technical paths, placeholder-
only layout skeletons, and technical formats such as hex/HRESULT strings.
Ordinary UI labels, tooltips, sentences, command names, and grammar-bearing
formats remain localizable even when translations currently match English.

## STRINGTABLE in .rc

```rc
STRINGTABLE
BEGIN
IDS_APP_TITLE       "RedSalamander"
IDS_FILE_COUNT      "{0} files selected"
IDS_ERROR_MESSAGE   "An error occurred: {0} in {1}"
IDS_HRESULT_DETAILS "0x{0:08X}: {1}"
END
```

## Avoid printf-style resource formats

```rc
// ❌ Wrong (printf-style):
IDS_FATAL_EXCEPTION_FMT "Fatal Exception: %s (0x%08X)"

// ✅ Correct (std::format-style positional):
IDS_FATAL_EXCEPTION_FMT "Fatal Exception: {0} (0x{1:08X})"
```

## Avoid unindexed std::format resource formats

```rc
// ❌ Wrong (cannot be safely reordered by translators):
IDS_SAVE_FAILED "Failed to save {0}: 0x{:08X}"

// ✅ Correct (every argument is indexed):
IDS_SAVE_FAILED "Failed to save {0}: 0x{1:08X}"
```

## Keep source argument order obvious

```rc
// ❌ Wrong (source string uses argument 1 before argument 0):
IDS_ARCHIVE_DONE "Extracted {1} entries from {0}."

// ✅ Correct (code passes entry count first, archive path second):
IDS_ARCHIVE_DONE "Extracted {0} entries from {1}."
```

```cpp
auto msg = FormatStringResource(nullptr, IDS_ARCHIVE_DONE, entryCount, archivePath);
```

Translations may reorder placeholders, but must preserve the exact source
placeholder tokens. For example, a translation of `{0:L} file{1:s}: {2}
selected` can move `{2}`, but it cannot invent `{3}` or change `{2}` to
`{2:s}` unless the source also has that exact token.

## Menu Resources in .rc

```rc
IDR_MAINMENU MENU
BEGIN
    POPUP "&File"
    BEGIN
        MENUITEM "&Open",   ID_FILE_OPEN
        MENUITEM "&Save",   ID_FILE_SAVE
        MENUITEM SEPARATOR
        MENUITEM "E&xit",   ID_FILE_EXIT
    END
END
```

## Dynamic Menu Extensions

Runtime code may **only** extend dynamic parts:
- Theme lists from `Themes\*.theme.json5`
- Custom themes from settings
- Per-user history entries

Base menu structure must remain in `.rc` files.

## Pattern

```cpp
// ✅ Correct - load from resources
auto title = LoadStringResource(IDS_TITLE);

// ❌ Wrong - hardcoded
SetWindowTextW(hWnd, L"My Title");
```
