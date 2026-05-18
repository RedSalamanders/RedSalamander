#include "MinimumOsVersion.h"

namespace
{
using RtlGetVersionFn = LONG(WINAPI*)(OSVERSIONINFOW*);

[[nodiscard]] bool TryGetWindowsVersion(OSVERSIONINFOW& version) noexcept
{
    version                      = {};
    version.dwOSVersionInfoSize = sizeof(version);

    const HMODULE ntdll = ::GetModuleHandleW(L"ntdll.dll");
    if (! ntdll)
    {
        return false;
    }

    const FARPROC proc = ::GetProcAddress(ntdll, "RtlGetVersion");
#pragma warning(suppress : 4191)
    const auto rtlGetVersion = reinterpret_cast<RtlGetVersionFn>(proc);
    return rtlGetVersion != nullptr && rtlGetVersion(&version) == 0;
}

[[nodiscard]] bool IsAtLeastMinimumVersion(const OSVERSIONINFOW& version) noexcept
{
    using namespace Common::MinimumOsVersion;

    if (version.dwMajorVersion != kMinimumWindowsMajorVersion)
    {
        return version.dwMajorVersion > kMinimumWindowsMajorVersion;
    }
    if (version.dwMinorVersion != kMinimumWindowsMinorVersion)
    {
        return version.dwMinorVersion > kMinimumWindowsMinorVersion;
    }
    return version.dwBuildNumber >= kMinimumWindowsBuildNumber;
}
} // namespace

namespace Common::MinimumOsVersion
{
bool IsCurrentWindowsVersionSupported() noexcept
{
    OSVERSIONINFOW version{};
    return TryGetWindowsVersion(version) && IsAtLeastMinimumVersion(version);
}

bool EnsureCurrentWindowsVersionSupported(HWND owner, UnsupportedVersionNotification notification) noexcept
{
    if (IsCurrentWindowsVersionSupported())
    {
        return true;
    }

    ::OutputDebugStringW(kUnsupportedWindowsMessage);
    if (notification == UnsupportedVersionNotification::MessageBox)
    {
        static_cast<void>(::MessageBoxW(owner, kUnsupportedWindowsMessage, L"RedSalamander", MB_OK | MB_ICONERROR));
    }
    return false;
}
} // namespace Common::MinimumOsVersion
