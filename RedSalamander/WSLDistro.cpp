#include "WSLDistro.h"
#include "LocalizationManager.h"
#include "WslDistributionCatalog.h"
#include "resource.h"

#include <cstddef>
#include <shlwapi.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820) // WIL warnings
#include <wil/resource.h>
#pragma warning(pop)

#pragma comment(lib, "shlwapi.lib")

std::vector<WSLDistribution> WSLDistro::EnumerateDistributions() noexcept
{
    std::vector<WSLDistribution> distributions;
    const auto records = Common::Wsl::EnumerateDistributions();
    distributions.reserve(records.size());
    for (const auto& record : records)
    {
        WSLDistribution distro;
        distro.name        = record.name;
        distro.guid        = record.guid;
        distro.basePath    = record.basePath;
        distro.networkPath = BuildNetworkPath(record.name);
        distro.isWsl2      = record.isWsl2;
        distributions.push_back(std::move(distro));
    }
    return distributions;
}

bool WSLDistro::IsWSLInstalled() noexcept
{
    return Common::Wsl::IsInstalled();
}

std::wstring WSLDistro::BuildNetworkPath(const std::wstring& name) noexcept
{
    // Use modern \\wsl.localhost\{name} format
    return L"\\\\wsl.localhost\\" + name;
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

    // Load PNG resource from localized resources
    HINSTANCE hInstance = GetModuleHandleW(nullptr);
    if (! hInstance)
    {
        return nullptr;
    }

    std::vector<std::byte> imageBytes;
    if (! Localization::LoadResourceBytes(hInstance, MAKEINTRESOURCEW(resourceId), L"PNG", imageBytes))
    {
        return nullptr;
    }

    // CreateIconFromResourceEx automatically handles PNG format and scaling
    // Version 0x00030000 supports PNG format (Windows Vista+)
    // Use iconSize for proper DPI-aware sizing, or 0 for default
    HICON hIcon = CreateIconFromResourceEx(reinterpret_cast<PBYTE>(imageBytes.data()),
                                           static_cast<DWORD>(imageBytes.size()),
                                           TRUE,             // fIcon = TRUE (not cursor)
                                           0x00030000,       // dwVersion = 3.0 (supports PNG)
                                           iconSize,         // cxDesired (DPI-aware size)
                                           iconSize,         // cyDesired (DPI-aware size)
                                           LR_DEFAULTCOLOR); // Flags

    return wil::unique_hicon{hIcon}; // May be empty if creation failed
}
