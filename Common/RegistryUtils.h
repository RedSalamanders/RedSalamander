#pragma once

#include <cstddef>
#include <optional>
#include <string>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

#ifndef COMMON_API
#ifdef COMMON_EXPORTS
#define COMMON_API __declspec(dllexport)
#else
#define COMMON_API __declspec(dllimport)
#endif
#endif

namespace Common::Registry
{
// This is a bounded policy adapter, not a replacement for wil::reg. WIL owns ordinary key lifetime,
// open/create, iteration, and typed-value operations. Its allocating string getters intentionally do
// not expose this adapter's pre-allocation byte cap, query/read mutation policy, or test observer.
inline constexpr size_t kDefaultMaxStringBytes        = 64u * 1024u;
inline constexpr unsigned int kDefaultMaxReadAttempts = 3u;

using StringReadObserver = void (*)(HKEY key, const wchar_t* valueName, unsigned int attempt, void* cookie) noexcept;

struct StringReadOptions
{
    size_t maxBytes              = kDefaultMaxStringBytes;
    unsigned int maxAttempts     = kDefaultMaxReadAttempts;
    bool allowExpandString       = false;
    bool expandEnvironmentString = false;
    bool trimWhitespace          = false;
    bool allowEmpty              = false;

    // Synchronous seam for deterministic mutation tests. Production callers leave this null.
    StringReadObserver afterSizeReadObserver = nullptr;
    void* observerCookie                     = nullptr;
};

[[nodiscard]] COMMON_API std::optional<std::wstring> ReadBoundedStringValue(HKEY key, const wchar_t* valueName, const StringReadOptions& options = {}) noexcept;
} // namespace Common::Registry
