---
name: localization
description: Localization and resource management for Windows RC files. Use when adding user-facing strings, creating menus, working with STRINGTABLE, or using LoadStringResource and FormatStringResource helpers.
metadata:
  author: DualTail
  version: "1.0"
---

# Localization and Resources

## Mandatory Rules

1. **All user-facing strings** → `.rc` STRINGTABLE
2. **All static menus** → `.rc` MENU resources
3. **Never hardcode UI text** in C++
4. **Always use positional formatting** for variable strings to help localization
5. **Never treat a resource string as a printf-format string** (`%s`, `%d`, etc.). Use `FormatStringResource(...)` with `std::format`-style placeholders and avoid C4774 suppressions.

## Resource Helpers (Common/Helpers.h)

```cpp
// Load string from resources
std::wstring title = LoadStringResource(IDS_APP_TITLE);

// Format with arguments
std::wstring msg = FormatStringResource(IDS_FILE_COUNT, fileCount);

// Message box from resources
MessageBoxResource(_hWnd, IDS_ERROR_MSG, IDS_ERROR_TITLE, MB_OK | MB_ICONERROR);
```

## STRINGTABLE in .rc

```rc
STRINGTABLE
BEGIN
    IDS_APP_TITLE       "Red Salamander"
    IDS_FILE_COUNT      "{0} files selected"
    IDS_ERROR_MESSAGE   "An error occurred: {0} in {1}"
END
```

## Avoid printf-style resource formats

```rc
// ❌ Wrong (printf-style):
IDS_FATAL_EXCEPTION_FMT "Fatal Exception: %s (0x%08X)"

// ✅ Correct (std::format-style positional):
IDS_FATAL_EXCEPTION_FMT "Fatal Exception: {0} (0x{1:08X})"
```

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
