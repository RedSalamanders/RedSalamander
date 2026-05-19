#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#ifndef COMMON_API
#ifdef COMMON_EXPORTS
#define COMMON_API __declspec(dllexport)
#else
#define COMMON_API __declspec(dllimport)
#endif
#endif

namespace Common::MinimumOsVersion
{
inline constexpr DWORD kMinimumWindowsMajorVersion = 10u;
inline constexpr DWORD kMinimumWindowsMinorVersion = 0u;
inline constexpr DWORD kMinimumWindowsBuildNumber   = 22000u;
inline constexpr DWORD kMinimumWindowsBuildRevision = 2600u;
inline constexpr wchar_t kMinimumWindowsDisplayName[] = L"Windows 11 build 22000.2600";
inline constexpr wchar_t kUnsupportedWindowsMessage[] = L"RedSalamander requires Windows 11 build 22000.2600 or later.";

enum class UnsupportedVersionNotification : unsigned char
{
    None,
    MessageBox,
};

COMMON_API bool IsCurrentWindowsVersionSupported() noexcept;
COMMON_API bool EnsureCurrentWindowsVersionSupported(
    HWND owner,
    UnsupportedVersionNotification notification = UnsupportedVersionNotification::MessageBox) noexcept;
} // namespace Common::MinimumOsVersion
