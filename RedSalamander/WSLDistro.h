#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX

#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include <Windows.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820) // Deleted operators, inlining, padding
#include <wil/resource.h>
#pragma warning(pop)

// Represents a single WSL distribution
#pragma warning(push)
#pragma warning(disable : 4625 4626) // WIL headers: deleted copy/move
struct WSLDistribution
{
    WSLDistribution() = default;

    std::wstring name;        // Distribution name (e.g., "Ubuntu")
    std::wstring guid;        // Registry GUID
    std::wstring networkPath; // Network path (\\wsl.localhost\{name})
    bool isWsl2 = false;      // True if Modern=1 (WSL2), false for WSL1
    wil::unique_hicon icon;   // Cached custom icon (DPI-aware, lifetime-cached)
};
#pragma warning(pop)

// WSL distribution enumeration and management
class WSLDistro
{
public:
    // Enumerate all registered WSL distributions using registry-based approach
    // Filters out docker-desktop and rancher-desktop utility distros
    // Returns empty vector if WSL is not installed or registry access fails
    static std::vector<WSLDistribution> EnumerateDistributions() noexcept;

    // Check if WSL is installed on the system
    static bool IsWSLInstalled() noexcept;

    // Build network path for a distribution using modern format
    static std::wstring BuildNetworkPath(const std::wstring& name) noexcept;

    // Load distribution icon from PNG resources based on distribution name
    // Returns a RAII icon on success, or an empty icon on failure.
    static wil::unique_hicon LoadDistributionIcon(const std::wstring& distroName, int iconSize) noexcept;

private:
    // Registry key paths
    static constexpr wchar_t kLxssRegKey[]       = L"Software\\Microsoft\\Windows\\CurrentVersion\\Lxss";
    static constexpr wchar_t kRegKeyDistroName[] = L"DistributionName";
    static constexpr wchar_t kRegKeyModern[]     = L"Modern";

    // Utility distro prefixes to filter out
    static constexpr std::wstring_view kDockerDistroPrefix  = L"docker-desktop";
    static constexpr std::wstring_view kRancherDistroPrefix = L"rancher-desktop";

    // Open WSL registry key
    static wil::unique_hkey OpenWslRegKey() noexcept;

    // Open specific distro subkey by GUID
    static wil::unique_hkey OpenDistroKey(const wil::unique_hkey& wslKey, const std::wstring& guid) noexcept;

    // Enumerate all distribution GUIDs
    static bool EnumerateDistroGuids(const wil::unique_hkey& wslKey, std::vector<std::wstring>& guids) noexcept;

    // Read distribution name from registry
    static std::optional<std::wstring> ReadDistroName(const wil::unique_hkey& distroKey) noexcept;

    // Read Modern flag (WSL version)
    static bool ReadModernFlag(const wil::unique_hkey& distroKey) noexcept;

    // Check if distribution name should be filtered out
    static bool ShouldFilterDistro(const std::wstring& name) noexcept;
};
