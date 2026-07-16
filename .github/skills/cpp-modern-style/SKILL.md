---
name: cpp-modern-style
description: Modern C++23 coding style and conventions for RedSalamander. Use when writing new code, reviewing style, using STL containers, smart pointers, std::format, std::optional, or following naming conventions.
metadata:
  author: RedSalamander
  version: "1.0"
---

# Modern C++23 Style Guide

## Project Requirements
- **C++23** standard
- **Unicode UTF-16** encoding
- **Windows 10/11** minimum

## Smart Pointers

```cpp
// ✅ Exclusive ownership
std::unique_ptr<Widget> widget = std::make_unique<Widget>();

// ✅ Shared ownership (only when needed)
std::shared_ptr<Resource> resource = std::make_shared<Resource>();

// ❌ NEVER raw new/delete
Widget* widget = new Widget();
delete widget;
```

## STL Usage

```cpp
// ✅ std::format for strings
auto msg = std::format(L"Found {} files", count);

// ✅ std::optional (NEVER use * to access)
std::optional<int> result = FindValue();
if (result.has_value()) {
    Use(result.value());
}
int val = result.value_or(0);

// ✅ Range-based for
for (const auto& item : items) { }

// ✅ Structured bindings
auto [key, value] = GetPair();

// ❌ NEVER fixed-size buffers for formatting
wchar_t buf[256];  // BAD
```

### Formatting Into Fixed Buffers (When Required)

Prefer `std::format_to_n` over `*printf_s` when you must write into an existing fixed buffer (e.g. `noexcept` hot paths, Win32 APIs).

```cpp
std::array<wchar_t, 32> buf{};
constexpr size_t max = buf.size() - 1;
const auto r         = std::format_to_n(buf.data(), max, L"VK_{:02X}", vk);
buf[(r.size < max) ? r.size : max] = L'\0';
```

## Naming Conventions

| Type | Style | Example |
|------|-------|---------|
| Classes | PascalCase | `ColorTextView` |
| Public methods | PascalCase | `RenderText()` |
| Private methods | camelCase | `calculateLayout()` |
| Variables | camelCase | `textLayout` |
| Members | `_` prefix | `_scrollY` |
| Constants | kName | `kMaxBufferSize` |

## Code Style

```cpp
// ✅ One declaration per line
int width;
int height;

// ❌ Multiple declarations
int width, height;

// ✅ Use auto appropriately
auto iter = container.begin();

// ✅ constexpr for compile-time
constexpr int kBufferSize = 1024;

// ✅ string_view and span for non-owning views
void Process(std::wstring_view text);
void Process(std::span<const int> data);
```

## C++ Core Guidelines

Follow [C++ Core Guidelines](https://isocpp.github.io/CppCoreGuidelines/CppCoreGuidelines):
- Prioritize safety, simplicity, and maintainability
- Avoid undefined behavior
- Use GSL concepts where applicable

## RAII Pattern

Every resource should have an owner:
```cpp
class Resource 
{
    Handle _handle;
public:
    Resource(Args args) : _handle(acquire(args)) {}
    ~Resource() { release(_handle); }
    // Rule of 5: Disable copy, implement move if needed
    Resource(const Resource&) = delete;
    Resource& operator=(const Resource&) = delete;
    Resource(Resource&&) noexcept = default;
    Resource& operator=(Resource&&) noexcept = default;
};
```

## File Organization

- **Headers**: Keep interfaces minimal, use forward declarations
- **Implementation**: Group related functionality together
- **Dependencies**: Minimize header includes
- **Platform**: Isolate platform-specific code when possible

## Shared Helper Reuse

Before defining a local helper, consult `Specs/Core/Core_SharedHelpers.md` and search `Common/` (or
`Tests/TestSupport/` for test infrastructure). If an existing helper has the required semantics, use it directly.
Do not create a differently named copy or a leaf wrapper that only forwards the same policy.

If the existing helper is at the correct dependency layer but lacks a generally useful operation, extend it and
add focused tests. Keep a local implementation only when it intentionally differs in policy, dependency, ABI,
ownership, failure handling, or hot-path constraints; name and document that difference. When adding a canonical
shared helper, update the catalog and its authoritative domain spec, and add a source-contract guard when a
regression could silently recreate prior copies.

## Comments and Documentation

- Write self-documenting code with meaningful names
- Document public APIs with Doxygen-style comments
- Explain "why" rather than "what"
- Include performance notes for critical sections
- For perf-sensitive features or hot-path changes, define the protected scenario and required measurement path in code/docs early rather than after the implementation is complete

## Patterns to Avoid

- `new`/`delete` (use smart pointers)
- C-style casts (use `static_cast`, `reinterpret_cast`)
- `goto` (use early returns + RAII / `wil::scope_exit`)
- Raw Windows handles (use WIL wrappers)
- `sprintf_s` / `swprintf_s` in non-PoC code
- `catch (...)` (FORBIDDEN; catch explicitly named exception types only, and document why catching is mandatory at that boundary)
- Global state and singletons unless absolutely necessary
- Blocking UI thread with synchronous operations
- String concatenation in loops (use `std::format` or reserve capacity)
- Multiple variable declarations on same line
- Perf-sensitive feature work with no instrumentation, no deterministic selftest, or no archived evidence
