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
inline constexpr DWORD kMinimumWindowsBuildNumber  = 26100u;
inline constexpr wchar_t kMinimumWindowsDisplayName[] = L"Windows 11 24H2 (build 26100)";
inline constexpr wchar_t kUnsupportedWindowsMessage[] = L"RedSalamander requires Windows 11 24H2 (build 26100) or later.";

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
