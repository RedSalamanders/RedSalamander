#include "WSLDistro.h"
#include "resource.h"

#include <algorithm>
#include <optional>
#include <shlwapi.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820) // WIL warnings
#include <wil/resource.h>
#pragma warning(pop)

#pragma comment(lib, "shlwapi.lib")

std::vector<WSLDistribution> WSLDistro::EnumerateDistributions() noexcept
{
    std::vector<WSLDistribution> distributions;

    auto wslRootKey = OpenWslRegKey();
    if (! wslRootKey)
    {
        return distributions; // WSL not installed or no access
    }

    // Enumerate all distribution GUIDs
    std::vector<std::wstring> guids;
    if (! EnumerateDistroGuids(wslRootKey, guids))
    {
        return distributions;
    }

    distributions.reserve(guids.size());

    // Process each distribution
    for (const auto& guid : guids)
    {
        auto distroKey = OpenDistroKey(wslRootKey, guid);
        if (! distroKey)
        {
            continue;
        }

        // Read distribution name
        auto name = ReadDistroName(distroKey);
        if (! name.has_value() || name->empty())
        {
            continue;
        }

        // Filter out utility distributions
        if (ShouldFilterDistro(*name))
        {
            continue;
        }

        // Read WSL version
        bool isWsl2 = ReadModernFlag(distroKey);

        // Build distribution info
        WSLDistribution distro;
        distro.name        = *name;
        distro.guid        = guid;
        distro.networkPath = BuildNetworkPath(*name);
        distro.isWsl2      = isWsl2;

        distributions.push_back(std::move(distro));
    }

    // Sort by name for consistent ordering
    std::sort(distributions.begin(),
              distributions.end(),
              [](const WSLDistribution& a, const WSLDistribution& b) { return _wcsicmp(a.name.c_str(), b.name.c_str()) < 0; });

    return distributions;
}

bool WSLDistro::IsWSLInstalled() noexcept
{
    auto wslRootKey = OpenWslRegKey();
    return wslRootKey.is_valid();
}

std::wstring WSLDistro::BuildNetworkPath(const std::wstring& name) noexcept
{
    // Use modern \\wsl.localhost\{name} format
    return L"\\\\wsl.localhost\\" + name;
}

wil::unique_hkey WSLDistro::OpenWslRegKey() noexcept
{
    HKEY hKey = nullptr;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, kLxssRegKey, 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
        return wil::unique_hkey{hKey};
    }
    return wil::unique_hkey{};
}

wil::unique_hkey WSLDistro::OpenDistroKey(const wil::unique_hkey& wslKey, const std::wstring& guid) noexcept
{
    if (! wslKey)
    {
        return wil::unique_hkey{};
    }

    HKEY hKey = nullptr;
    if (RegOpenKeyExW(wslKey.get(), guid.c_str(), 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
        return wil::unique_hkey{hKey};
    }
    return wil::unique_hkey{};
}

bool WSLDistro::EnumerateDistroGuids(const wil::unique_hkey& wslKey, std::vector<std::wstring>& guids) noexcept
{
    if (! wslKey)
    {
        return false;
    }

    guids.clear();

    wchar_t buffer[39]; // GUIDs are 38 chars + null terminator: {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}

    for (DWORD index = 0;; ++index)
    {
        DWORD length = ARRAYSIZE(buffer);
        LONG result  = RegEnumKeyExW(wslKey.get(), index, buffer, &length, nullptr, nullptr, nullptr, nullptr);

        if (result == ERROR_NO_MORE_ITEMS)
        {
            break;
        }

        if (result == ERROR_SUCCESS && length == 38 && buffer[0] == L'{' && buffer[37] == L'}')
        {
            guids.emplace_back(&buffer[0], length);
        }
    }

    return true;
}

std::optional<std::wstring> WSLDistro::ReadDistroName(const wil::unique_hkey& distroKey) noexcept
{
    if (! distroKey)
    {
        return std::nullopt;
    }

    wchar_t buffer[256];
    DWORD bufferSize = sizeof(buffer);
    DWORD type       = 0;

    LONG result = RegQueryValueExW(distroKey.get(), kRegKeyDistroName, nullptr, &type, reinterpret_cast<BYTE*>(buffer), &bufferSize);

    if (result == ERROR_SUCCESS && type == REG_SZ)
    {
        return std::wstring(buffer);
    }

    return std::nullopt;
}

bool WSLDistro::ReadModernFlag(const wil::unique_hkey& distroKey) noexcept
{
    if (! distroKey)
    {
        return false;
    }

    DWORD modernValue = 0;
    DWORD bufferSize  = sizeof(modernValue);
    DWORD type        = 0;

    LONG result = RegQueryValueExW(distroKey.get(), kRegKeyModern, nullptr, &type, reinterpret_cast<BYTE*>(&modernValue), &bufferSize);

    if (result == ERROR_SUCCESS && type == REG_DWORD)
    {
        return modernValue == 1u;
    }

    return false;
}

bool WSLDistro::ShouldFilterDistro(const std::wstring& name) noexcept
{
    // Filter out docker-desktop* and rancher-desktop* utility distributions
    if (name.size() >= kDockerDistroPrefix.size())
    {
        if (_wcsnicmp(name.c_str(), kDockerDistroPrefix.data(), kDockerDistroPrefix.size()) == 0)
        {
            return true;
        }
    }

    if (name.size() >= kRancherDistroPrefix.size())
    {
        if (_wcsnicmp(name.c_str(), kRancherDistroPrefix.data(), kRancherDistroPrefix.size()) == 0)
        {
            return true;
        }
    }

    return false;
}

wil::unique_hicon WSLDistro::LoadDistributionIcon(const std::wstring& distroName, int iconSize) noexcept
{
    // Map distribution names to resource IDs using case-insensitive substring matching
    struct IconMapping
    {
        int resourceId;
        const wchar_t* keyword;
    };

    static const IconMapping iconMap[] = {
        {IDI_WSL_UBUNTU, L"ubuntu"},
        {IDI_WSL_DEBIAN, L"debian"},
        {IDI_WSL_FEDORA, L"fedora"},
        {IDI_WSL_LINUX_GENERIC, L"kali"},
        {IDI_WSL_LINUX_GENERIC, L"opensuse"},
        {IDI_WSL_LINUX_GENERIC, L"suse"},
        {IDI_WSL_LINUX_GENERIC, L"alpine"},
        {IDI_WSL_LINUX_GENERIC, L"arch"},
        {IDI_WSL_LINUX_GENERIC, L"manjaro"},
        {IDI_WSL_LINUX_GENERIC, L"alma"},
        {IDI_WSL_LINUX_GENERIC, L"rocky"},
    };

    // Find matching resource ID via case-insensitive substring search
    int resourceId = 0;
    for (const auto& mapping : iconMap)
    {
        if (StrStrIW(distroName.c_str(), mapping.keyword) != nullptr)
        {
            resourceId = mapping.resourceId;
            break;
        }
    }

    if (resourceId == 0)
    {
        return {}; // No icon for this distribution
    }

    // Load PNG resource from embedded resources
    HINSTANCE hInstance = GetModuleHandleW(nullptr);
    if (! hInstance)
    {
        return nullptr;
    }

    HRSRC hResource = FindResourceW(hInstance, MAKEINTRESOURCEW(resourceId), L"PNG");
    if (! hResource)
    {
        return nullptr;
    }

    DWORD imageSize = SizeofResource(hInstance, hResource);
    if (imageSize == 0)
    {
        return nullptr;
    }

    HGLOBAL hMemory = LoadResource(hInstance, hResource);
    if (! hMemory)
    {
        return nullptr;
    }

    BYTE* pImageData = static_cast<BYTE*>(LockResource(hMemory));
    if (! pImageData)
    {
        return nullptr;
    }

    // CreateIconFromResourceEx automatically handles PNG format and scaling
    // Version 0x00030000 supports PNG format (Windows Vista+)
    // Use iconSize for proper DPI-aware sizing, or 0 for default
    HICON hIcon = CreateIconFromResourceEx(pImageData,
                                           imageSize,
                                           TRUE,             // fIcon = TRUE (not cursor)
                                           0x00030000,       // dwVersion = 3.0 (supports PNG)
                                           iconSize,         // cxDesired (DPI-aware size)
                                           iconSize,         // cyDesired (DPI-aware size)
                                           LR_DEFAULTCOLOR); // Flags

    return wil::unique_hicon{hIcon}; // May be empty if creation failed
}
