# RedSalamander C++ Development Guidelines

## Project Overview

**RedSalamander** is a Windows-based C++ application featuring:
- Advanced text visualization components (ColorTextView with D2D/DirectWrite rendering)
- Real-time debugging and monitoring capabilities
- High-performance graphics rendering using Direct2D, DirectWrite, and DXGI
- Multi-threaded architecture with async operations
- vcpkg-based dependency management

## Project Requirements

- **Platform**: Windows 11 build 22000.2600 minimum
- **Language Standard**: `stdcpplatest` (latest MSVC mode)
- **Character Encoding**: Unicode UTF-16
- **Build System**: Visual Studio 2026 with MSBuild
- **Package Manager**: vcpkg
- **Graphics APIs**: Direct2D, DirectWrite, Direct3D 11, DXGI

## Skills Reference

Detailed patterns and guidelines are available as Agent Skills in `.github/skills/`:

| Skill | Description |
|-------|-------------|
| [cpp-build](.github/skills/cpp-build/SKILL.md) | Build system, `build.ps1` usage, project structure |
| [cpp-modern-style](.github/skills/cpp-modern-style/SKILL.md) | C++23 patterns, naming conventions, STL usage |
| [perf-validation](.github/skills/perf-validation/SKILL.md) | Mandatory perf-first workflow: instrumentation, selftests, archived runs, and before/after evidence for new features and optimizations |
| [wil-raii](.github/skills/wil-raii/SKILL.md) | WIL RAII wrappers for Windows resources |
| [direct2d-rendering](.github/skills/direct2d-rendering/SKILL.md) | Direct2D/DirectWrite graphics patterns |
| [icon-cache](.github/skills/icon-cache/SKILL.md) | Shell icon management via IconCache |
| [win32-wndproc](.github/skills/win32-wndproc/SKILL.md) | Window procedure and message handling |
| [plugin-callbacks](.github/skills/plugin-callbacks/SKILL.md) | Plugin callback pattern with cookie |
| [localization](.github/skills/localization/SKILL.md) | RC resources, STRINGTABLE, menus |
| [file-actions](.github/skills/file-actions/SKILL.md) | File-action ID semantics, settings validation, direct dispatch, and regression coverage |
| [theming](.github/skills/theming/SKILL.md) | Theme color keys and JSON5 themes |
| [compiler-warnings](.github/skills/compiler-warnings/SKILL.md) | MSVC `/Wall` warning policy with documented PoC exceptions |
| [async-threading](.github/skills/async-threading/SKILL.md) | Threading model and async patterns |
| [error-handling](.github/skills/error-handling/SKILL.md) | Error handling, HRESULT, Debug logging |
| [yyjson](.github/skills/yyjson/SKILL.md) | yyjson parsing/serialization patterns, ownership, and cleanup |

## Core Principles

### RAII is Mandatory
All Windows resources MUST use WIL RAII wrappers. Manual cleanup (`DestroyIcon`, `DeleteObject`, `EndPaint`, etc.) is **PROHIBITED**. See [wil-raii skill](.github/skills/wil-raii/SKILL.md).

### Performance Validation is Mandatory
Any new feature, hot-path change, or optimization that can affect responsiveness, throughput, queueing, rendering, startup, search, Compare Directories, File Operations, plugin I/O, or memory retention MUST integrate:
- scenario definition,
- instrumentation,
- deterministic selftest coverage,
- archived perf evidence under `Specs/TestRuns/`

from the beginning. Do not defer perf validation to a later cleanup pass. See [perf-validation skill](.github/skills/perf-validation/SKILL.md) and `Specs/Testing/Testing_PerformanceValidation.md`.

### Spec Closeout is Mandatory
Completed WIP plans MUST be moved to `Specs/Plans/Done/`. Any durable behavior, UI contract, validation rule, or workflow requirement discovered during implementation MUST be merged into the authoritative domain spec under `Specs/<Domain>/` (or repo-level guidance such as `AGENTS.md` / `Specs/Testing/*` when appropriate) before the work is considered closed. Do not leave normative requirements stranded only in `Specs/Plans/WIP/` or `Specs/Plans/Done/`.

### Regression Guards (Common Violations)
- **Shared-helper reuse is mandatory:** before adding a consumer-local utility, search
  `Specs/Core/Core_SharedHelpers.md`, `Common/`, and (for test code) `Tests/TestSupport/`. Reuse or extend the
  canonical helper when its semantics match; do not reimplement it under a different name. A local variant is
  allowed only when policy, dependency layer, ABI, ownership, error handling, or performance semantics differ.
  Name that difference, document it at the definition, and cover it in the relevant reviewed allowlist or
  source-contract test. Adding a new shared helper also requires updating the shared-helper catalog and the
  authoritative domain spec that owns its behavior.
- **Runtime dependency staging is centralized:** add or change app-local plugin DLLs in
  root `RuntimeDependencies.props`; do not reintroduce plugin-local `PostBuildEvent`/`xcopy` batches. Required
  inputs must fail through the shared MSBuild tasks, obsolete outputs must be removed declaratively, and portable
  packaging must consume the same manifest and pass the clean-extraction smoke contract in
  `Specs/Installer/Installer_PortableZip.md`.
- **Ban `sprintf_s` / `swprintf_s`** in non-PoC code:
  - Diagnostics: `std::format` / `std::format_to_n` + `OutputDebugStringA/W`
  - User-facing/localized: `.rc` resources + `FormatStringResource(...)` with **positional** placeholders
- **Resource strings must use `std::format`-style positional placeholders** (`{0}`, `{1:08X}`); bare `{}` and unindexed format specs like `{:08X}` are forbidden in `.rc` resources because translators must be able to reorder arguments. Embedded/source strings must introduce placeholders in code argument order (`{0}`, then `{1}`, etc.) with no skipped indexes, and `FormatStringResource(...)` calls must pass arguments in that same source-string order. Satellite translations may reorder placeholders for grammar but must preserve the exact source placeholder tokens; they must not add, drop, duplicate, renumber, or change format specs. Never treat a resource string as a printf-format string (avoid C4774 suppression).
- **COM ownership:** never store owning raw COM interface pointers (no manual `Release()`); use `wil::com_ptr<T>` for members and locals.
- **COM ref-counting:** never do `obj->AddRef(); ptr.attach(obj);` (two-step hazard); prefer `ptr = obj;` / `wil::com_ptr<T> ptr = obj;`.
- **`wil::unique_hwnd` ownership:** never call `DestroyWindow(_hWnd.get())` on a `wil::unique_hwnd` owner; use `_hWnd.reset()` (or `.release()` only when transferring ownership explicitly).
- **Cross-thread `PostMessageW` payloads:** use `PostMessagePayload(...)` + `TakeMessagePayload<T>(lParam)`; never `PostMessageW(...payload.release())` or raw `new` payload posts. A helper-owned nonzero `lParam` is an opaque registry token, never a pointer: do not cast or adopt it directly. If a nonzero token yields null, it is stale/drained/type-mismatched and must be ignored before message-specific work. Zero is reserved for an explicitly posted payload-less fallback only when that message defines one. Capture every payload field needed for `msg`/`wParam`/other arguments before the call; never dereference a `unique_ptr` in the same function-call expression that also passes `std::move(payload)`, because argument evaluation may move it first. For windows that receive payload messages, call `InitPostedPayloadWindow(hwnd)` during create (`WM_NCCREATE`/`WM_CREATE`) and `DrainPostedPayloadsForWindow(hwnd)` in `WM_NCDESTROY`; the drain closes posting, invalidates all registered tokens, and deletes the registered storage. A stale queued token can survive numeric HWND reuse but can never be adopted. Every UI-host cross-thread change must review: payload ownership, UI-thread boundary, cancellation path, teardown drain, and a focused teardown stress/selftest when the touched host can queue payloads.
- **IconCache COM contract:** `IconCache::Initialize(...)` stays UI-thread/STA responsibility; any worker thread calling `IconCache::ExtractSystemIcon()` must initialize COM as MTA (`wil::CoInitializeEx(COINIT_MULTITHREADED)`).
- **MTP/WPD worker contract:** portable-device workers must initialize COM as MTA before using WPD, serialize each device session, honor watchdog/cancel/quarantine teardown, and keep `GetCapabilities()` honest for read-only, writable, invalid-profile, and disconnected states.
- **File-action IDs:** treat viewer/editor/user-menu action IDs as stable but case-insensitive identifiers. Uniqueness and lookup are case-insensitive, IDs that differ only by case are invalid duplicates, and parameterized command dispatch or association resolution must never require exact casing.
- **yyjson mutable builders:** never pass temporary/stack strings to non-copy APIs (`yyjson_mut_obj_add_str`, `yyjson_mut_str`); for dynamic keys use `yyjson_mut_strncpy` + `yyjson_mut_obj_add`, and for string values prefer `*_strcpy`/`*_strncpy` (see `.github/skills/yyjson/SKILL.md`).
- **Thread-safety:** never read/write shared non-atomic state without a lock (or use `std::atomic` with correct memory ordering).
- **Exceptions / `catch (...)`:** `catch (...)` is **FORBIDDEN**. If exception handling is mandatory at an ABI boundary (`noexcept` methods, Win32 callbacks, thread entrypoints), catch only explicitly named exception types and add a short comment explaining why catching is required there and what the fallback is. Prefer non-throwing APIs (e.g. `std::filesystem` overloads with `std::error_code`). If you catch-and-fail, log once with `Debug::Error(...)`; if you catch-and-continue, document the fallback (avoid empty catch without an explanation). Treat `std::bad_alloc` as fatal (`std::terminate()`), and never let it be swallowed by a broader `catch (const std::exception&)`.
- **Detached threads:** avoid `std::thread(...).detach()` in plugin DLLs; prefer `TrySubmitThreadpoolCallback` or `std::jthread` + `stop_token`. If detach is unavoidable, pin module lifetime with `AcquireModuleReferenceFromAddress(...)` and make thread start exception-safe (especially in `noexcept` code).
- **Threadpool module pins:** a plugin threadpool callback that owns an `AcquireModuleReferenceFromAddress(...)` handle MUST transfer it at callback entry with `FreeLibraryWhenCallbackReturns(instance, pin.release())`. Destroying the last `wil::unique_hmodule` inside the callback can unmap the callback's own code before it returns.
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
.build\x64\Debug\*.exe, *.dll       # Debug builds
.build\x64\Release\*.exe, *.dll     # Release builds
.build\ARM64\Debug\*.exe, *.dll     # Debug builds
.build\ARM64\Release\*.exe, *.dll   # Release builds
```

## Build

Use `build.ps1` for command-line builds. See [cpp-build skill](.github/skills/cpp-build/SKILL.md) for details.

```powershell
.\build.ps1                              # Build all (Debug)
.\build.ps1 -Configuration Release       # Build all (Release)
.\build.ps1 -ProjectName RedSalamander   # Build specific project
```

To verify a change is green, use `.\Tools\Run-AllTests.ps1 -Suite Full` (builds + full suite) or `.\Tools\Run-AllTests.ps1 -SkipBuild` (in-process selftests only against an existing build). See README "Self-tests" for details.

## Dependencies

- **WIL**: Windows Implementation Library (RAII wrappers)
- **yyjson**: JSON parsing and serialization
- **DirectX**: Graphics APIs (D2D, D3D11, DXGI)

## LLM Assistant Guidelines

When working with this codebase:
- Always use WIL RAII wrappers for Windows resources
- Prioritize performance and responsiveness
- Treat performance validation as part of the feature contract: add or reuse metrics, add deterministic selftests, and archive runs when the scenario is perf-sensitive
- When finishing a plan, move it to `Specs/Plans/Done/` and update the authoritative spec or repo guidance so the lasting contract does not live only in the plan
- Consider DPI awareness in UI components
- Respect the Windows-specific architecture
- Suggest modern C++ patterns
- Reference appropriate skills for detailed patterns
