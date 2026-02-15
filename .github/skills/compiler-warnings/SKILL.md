---
name: compiler-warnings
description: Compiler warning policy and suppression patterns for MSVC /W4. Use when fixing warnings, adding pragma suppressions, or dealing with C4365, C5219, C4625, C4626 and other compiler diagnostics.
metadata:
  author: DualTail
  version: "1.0"
---

# Compiler Warnings Policy

**Objective: Zero warnings** (except C4702 unreachable code)

## Critical Rule

**NO GLOBAL `/wd` SUPPRESSIONS** in vcxproj files.
Fix warnings locally with `#pragma warning(push/disable/pop)`.

## Acceptable Global Suppressions

- `/wd5045` - Spectre mitigation
- `/wd4820` - Padding
- `/wd4710` - Not inlined
- `/wd4711` - Auto inline
- `/wd4514` - Unreferenced inline

## Required Fixes (Project Code)

| Warning | Fix |
|---------|-----|
| C4365 | `static_cast<T>()` for signed/unsigned |
| C5219 | Explicit cast for data loss |
| C5264 | Remove unused or add `[[maybe_unused]]` |
| C4464 | Fix include paths, avoid `..` |
| C4774 | Do not use runtime printf-format strings; use `FormatStringResource(...)` (resources) or `std::format` |

## C4774: Format String Not String Literal

**Do not suppress C4774** to keep using `swprintf_s` with a runtime format string (especially from resources).

Preferred fixes:
- Localized/user-facing: change the `.rc` string to `std::format`-style positional placeholders (`{0}`, `{1:08X}`, ...) and use `FormatStringResource(...)`.
- Diagnostics: use `std::format` / `std::format_to_n` instead of `*printf_s`.

## Pragma Patterns

### WIL/System Headers
```cpp
#pragma warning(push)
#pragma warning(disable: 4625 4626 5026 5027) // Deleted operators
#include <wil/resource.h>
#pragma warning(pop)
```

### Windows SDK
```cpp
#pragma warning(push)
#pragma warning(disable: 4820) // Padding
#include <windows.h>
#pragma warning(pop)
```

## Fix Patterns

```cpp
// Signed/unsigned conversion
size_t count = items.size();
int intCount = static_cast<int>(count);

// Unsigned literals
for (size_t i = 0u; i < count; ++i) { }

// Unused variable
[[maybe_unused]] const int kDebugFlag = 1;
```

## Always Comment Suppressions

```cpp
#pragma warning(disable: 4625) // Copy constructor deleted
```

## Additional Fix Patterns

```cpp
// Double-casting for complex type chains
auto val = static_cast<float>(static_cast<int>(source));

// Template safety with static_assert
template<typename T>
void Process(T value) {
    static_assert(std::is_integral_v<T>, "T must be integral");
}
```

## Build Configuration

- Support Debug and Release configurations
- Enable `/W4` for all projects
- Use static analysis tools (PVS-Studio, Clang-Tidy)
- Configure proper optimization flags for Release

## Testing and Quality

- Use static analysis tools regularly
- Follow consistent formatting (.clang-format)
- Write unit tests for complex logic
- Perform regular code reviews
- Profile graphics-intensive operations
- Test with large datasets
- Validate memory usage patterns
- Test on different DPI settings
