# Red Salamander C++ Development Guidelines

## Project Overview

**Red Salamander** is a Windows-based C++ application featuring:
- Advanced text visualization components (ColorTextView with D2D/DirectWrite rendering)
- Real-time debugging and monitoring capabilities
- High-performance graphics rendering using Direct2D, DirectWrite, and DXGI
- Multi-threaded architecture with async operations
- vcpkg-based dependency management

## Project Requirements

- **Platform**: Windows development (Windows 10/11 minimum)
- **Language Standard**: C++23
- **Character Encoding**: Unicode UTF-16
- **Build System**: Visual Studio 2022 with MSBuild
- **Package Manager**: vcpkg
- **Graphics APIs**: Direct2D, DirectWrite, Direct3D 11, DXGI

## Skills Reference

Detailed patterns and guidelines are available as Agent Skills in `.github/skills/`:

| Skill | Description |
|-------|-------------|
| [cpp-build](.github/skills/cpp-build/SKILL.md) | Build system, `build.ps1` usage, project structure |
| [cpp-modern-style](.github/skills/cpp-modern-style/SKILL.md) | C++23 patterns, naming conventions, STL usage |
| [wil-raii](.github/skills/wil-raii/SKILL.md) | WIL RAII wrappers for Windows resources |
| [direct2d-rendering](.github/skills/direct2d-rendering/SKILL.md) | Direct2D/DirectWrite graphics patterns |
| [icon-cache](.github/skills/icon-cache/SKILL.md) | Shell icon management via IconCache |
| [win32-wndproc](.github/skills/win32-wndproc/SKILL.md) | Window procedure and message handling |
| [plugin-callbacks](.github/skills/plugin-callbacks/SKILL.md) | Plugin callback pattern with cookie |
| [localization](.github/skills/localization/SKILL.md) | RC resources, STRINGTABLE, menus |
| [theming](.github/skills/theming/SKILL.md) | Theme color keys and JSON5 themes |
| [compiler-warnings](.github/skills/compiler-warnings/SKILL.md) | MSVC /W4 warning policy |
| [async-threading](.github/skills/async-threading/SKILL.md) | Threading model and async patterns |
| [error-handling](.github/skills/error-handling/SKILL.md) | Error handling, HRESULT, Debug logging |
| [yyjson](.github/skills/yyjson/SKILL.md) | yyjson parsing/serialization patterns, ownership, and cleanup |

## Core Principles

### RAII is Mandatory
All Windows resources MUST use WIL RAII wrappers. Manual cleanup (`DestroyIcon`, `DeleteObject`, `EndPaint`, etc.) is **PROHIBITED**. See [wil-raii skill](.github/skills/wil-raii/SKILL.md).

### Regression Guards (Common Violations)
- **Ban `sprintf_s` / `swprintf_s`** in non-PoC code:
  - Diagnostics: `std::format` / `std::format_to_n` + `OutputDebugStringA/W`
  - User-facing/localized: `.rc` resources + `FormatStringResource(...)` with **positional** placeholders
- **Resource strings must use `std::format`-style positional placeholders** (`{0}`, `{1:08X}`); never treat a resource string as a printf-format string (avoid C4774 suppression).
- **COM ownership:** never store owning raw COM interface pointers (no manual `Release()`); use `wil::com_ptr<T>` for members and locals.
- **COM ref-counting:** never do `obj->AddRef(); ptr.attach(obj);` (two-step hazard); prefer `ptr = obj;` / `wil::com_ptr<T> ptr = obj;`.
- **`wil::unique_hwnd` ownership:** never call `DestroyWindow(_hWnd.get())` on a `wil::unique_hwnd` owner; use `_hWnd.reset()` (or `.release()` only when transferring ownership explicitly).
- **Cross-thread `PostMessageW` payloads:** use `PostMessagePayload(...)` + `TakeMessagePayload<T>(lParam)`; never `PostMessageW(...payload.release())` or raw `new` payload posts. For windows that receive payload messages, call `InitPostedPayloadWindow(hwnd)` during create (`WM_NCCREATE`/`WM_CREATE`) and `DrainPostedPayloadsForWindow(hwnd)` in `WM_NCDESTROY` to prevent leak-on-destroy.
- **IconCache COM contract:** `IconCache::Initialize(...)` stays UI-thread/STA responsibility; any worker thread calling `IconCache::ExtractSystemIcon()` must initialize COM as MTA (`wil::CoInitializeEx(COINIT_MULTITHREADED)`).
- **yyjson mutable builders:** never pass temporary/stack strings to non-copy APIs (`yyjson_mut_obj_add_str`, `yyjson_mut_str`); for dynamic keys use `yyjson_mut_strncpy` + `yyjson_mut_obj_add`, and for string values prefer `*_strcpy`/`*_strncpy` (see `.github/skills/yyjson/SKILL.md`).
- **Thread-safety:** never read/write shared non-atomic state without a lock (or use `std::atomic` with correct memory ordering).
- **Exceptions / `catch (...)`:** `catch (...)` is **FORBIDDEN**. If exception handling is mandatory at an ABI boundary (`noexcept` methods, Win32 callbacks, thread entrypoints), catch only explicitly named exception types and add a short comment explaining why catching is required there and what the fallback is. Prefer non-throwing APIs (e.g. `std::filesystem` overloads with `std::error_code`). If you catch-and-fail, log once with `Debug::Error(...)`; if you catch-and-continue, document the fallback (avoid empty catch without an explanation). Treat `std::bad_alloc` as fatal (`std::terminate()`), and never let it be swallowed by a broader `catch (const std::exception&)`.
- **Detached threads:** avoid `std::thread(...).detach()` in plugin DLLs; prefer `TrySubmitThreadpoolCallback` or `std::jthread` + `stop_token`. If detach is unavoidable, pin module lifetime with `AcquireModuleReferenceFromAddress(...)` and make thread start exception-safe (especially in `noexcept` code).
- **Plugin unload quiet point:** stop producers → request shutdown/cancel → stop posting payload messages → `SetCallback(nullptr, nullptr)` → release instances → unload module only when no callbacks can still run.

### Modern C++
- Smart pointers over raw pointers
- `std::format` over string concatenation
- `std::optional` with `.value()` or `.value_or()` (never `*`)
- Range-based for loops and structured bindings
- See [cpp-modern-style skill](.github/skills/cpp-modern-style/SKILL.md)

### Error Handling
- Use HRESULT for Windows API calls
- `Debug::ErrorWithLastError(...)` for Win32 failures
- `Debug::Error(...)` for unexpected failures
- `Debug::Warning(...)` for recoverable failures
- Don't log normal control-flow (window size 0, cancellation, device recreation)
- See [error-handling skill](.github/skills/error-handling/SKILL.md)

## Patterns to Avoid

- Raw `new`/`delete`
- C-style casts
- `goto` (use early returns + RAII / `wil::scope_exit`)
- Raw Windows handles without WIL
- Manual resource cleanup
- `sprintf_s` / `swprintf_s` in non-PoC code
- `DestroyWindow(...get())` on `wil::unique_hwnd` owners
- `AddRef()` + `attach()` (use `wil::com_ptr` assignment)
- Empty `catch (...) {}` without explaining why it’s safe to ignore
- Global state and singletons
- Blocking UI thread
- Hardcoded UI strings (use `.rc` resources)

## Project Structure

### Key Components
- **RedSalamanderMonitor**: Main monitoring application
- **ColorTextView**: High-performance text editor/viewer
- **Common**: Shared utilities and helpers
- **Plugins**: FileSystem, ViewerText, ViewerSpace
- **PoC**: Proof-of-concept projects

### Output Locations
```text
x64\Debug\*.exe, *.dll       # Debug builds
x64\Release\*.exe, *.dll     # Release builds
```

## Build

Use `build.ps1` for command-line builds. See [cpp-build skill](.github/skills/cpp-build/SKILL.md) for details.

```powershell
.\build.ps1                              # Build all (Debug)
.\build.ps1 -Configuration Release       # Build all (Release)
.\build.ps1 -ProjectName RedSalamander   # Build specific project
```

## Dependencies

- **WIL**: Windows Implementation Library (RAII wrappers)
- **fmt**: Modern C++ formatting
- **DirectX**: Graphics APIs (D2D, D3D11, DXGI)

## LLM Assistant Guidelines

When working with this codebase:
- Always use WIL RAII wrappers for Windows resources
- Prioritize performance and responsiveness
- Consider DPI awareness in UI components
- Respect the Windows-specific architecture
- Suggest modern C++ patterns
- Reference appropriate skills for detailed patterns
