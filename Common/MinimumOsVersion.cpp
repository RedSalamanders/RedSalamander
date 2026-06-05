#include "MinimumOsVersion.h"

#pragma comment(lib, "Advapi32.lib")

namespace
{
using RtlGetVersionFn = LONG(WINAPI*)(OSVERSIONINFOW*);

constexpr wchar_t kWindowsCurrentVersionSubKey[]     = LR"(SOFTWARE\Microsoft\Windows NT\CurrentVersion)";
constexpr wchar_t kWindowsUpdateBuildRevisionValue[] = L"UBR";

[[nodiscard]] bool TryGetWindowsVersion(OSVERSIONINFOW& version) noexcept
{
    version                     = {};
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

[[nodiscard]] bool TryGetWindowsBuildRevision(DWORD& revision) noexcept
{
    revision = 0u;

    DWORD value          = 0u;
    DWORD size           = sizeof(value);
    const LSTATUS status = ::RegGetValueW(
        HKEY_LOCAL_MACHINE, kWindowsCurrentVersionSubKey, kWindowsUpdateBuildRevisionValue, RRF_RT_REG_DWORD | RRF_SUBKEY_WOW6464KEY, nullptr, &value, &size);
    if (status != ERROR_SUCCESS || size != sizeof(value))
    {
        return false;
    }

    revision = value;
    return true;
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
    if (version.dwBuildNumber != kMinimumWindowsBuildNumber)
    {
        return version.dwBuildNumber > kMinimumWindowsBuildNumber;
    }

    DWORD buildRevision = 0u;
    return TryGetWindowsBuildRevision(buildRevision) && buildRevision >= kMinimumWindowsBuildRevision;
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
