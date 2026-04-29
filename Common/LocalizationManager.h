#pragma once

#include <cstddef>
#include <string>
#include <string_view>
#include <vector>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#ifdef LoadString
#undef LoadString
#endif

#ifndef COMMON_API
#ifdef COMMON_EXPORTS
#define COMMON_API __declspec(dllexport)
#else
#define COMMON_API __declspec(dllimport)
#endif
#endif

namespace Localization
{
enum class LanguagePreferenceKind
{
    System,
    Culture,
};

struct LanguagePreference final
{
    LanguagePreferenceKind kind{LanguagePreferenceKind::System};
    std::wstring culture;
};

struct ResourceLookupResult final
{
    HINSTANCE instance = nullptr;
    HRSRC resource     = nullptr;

    [[nodiscard]] explicit operator bool() const noexcept
    {
        return instance && resource;
    }
};

COMMON_API HRESULT RegisterResourceOwner(std::wstring_view ownerName, HINSTANCE embeddedInstance) noexcept;
COMMON_API void UnregisterResourceOwner(HINSTANCE embeddedInstance) noexcept;
COMMON_API HRESULT ApplyLanguagePreference(const LanguagePreference& preference) noexcept;
COMMON_API int LoadString(HINSTANCE embeddedInstance, UINT id, std::wstring& result) noexcept;
COMMON_API HMENU LoadMenuResource(HINSTANCE embeddedInstance, UINT menuId) noexcept;
COMMON_API HACCEL LoadAcceleratorsResource(HINSTANCE embeddedInstance, PCWSTR tableName) noexcept;
COMMON_API ResourceLookupResult FindLocalizedResourceHandle(HINSTANCE embeddedInstance, PCWSTR name, PCWSTR type) noexcept;
COMMON_API HRSRC FindLocalizedResource(HINSTANCE embeddedInstance, PCWSTR name, PCWSTR type) noexcept;
COMMON_API HINSTANCE ResolveResourceInstance(HINSTANCE embeddedInstance, PCWSTR name, PCWSTR type) noexcept;
COMMON_API HANDLE LoadImageResource(HINSTANCE embeddedInstance, PCWSTR name, UINT type, int cx, int cy, UINT flags) noexcept;
COMMON_API bool LoadResourceBytes(HINSTANCE embeddedInstance, PCWSTR name, PCWSTR type, std::vector<std::byte>& result) noexcept;
COMMON_API std::vector<std::wstring> DiscoverAvailableCultures() noexcept;
} // namespace Localization
