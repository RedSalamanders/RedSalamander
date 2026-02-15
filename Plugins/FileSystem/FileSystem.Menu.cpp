#include "FileSystem.Internal.h"
#include "resource.h"

#include <algorithm>
#include <filesystem>
#include <optional>
#include <string_view>
#include <vector>

#include <shellapi.h>
#include <shlobj.h>

namespace
{
std::wstring GetDriveLabel(const wchar_t* drive)
{
    wchar_t volumeName[MAX_PATH] = {0};
    GetVolumeInformationW(drive, volumeName, ARRAYSIZE(volumeName), nullptr, nullptr, nullptr, nullptr, 0);
    return volumeName[0] ? volumeName : L"";
}

std::wstring GetDriveFreeSpace(const wchar_t* drive)
{
    ULARGE_INTEGER freeBytes{}, totalBytes{}, availableBytes{};
    if (GetDiskFreeSpaceExW(drive, &availableBytes, &totalBytes, &freeBytes))
    {
        return FormatBytesCompact(freeBytes.QuadPart);
    }
    return L"";
}

struct WslDistributionEntry
{
    std::wstring name;
    std::wstring networkPath;
};

constexpr wchar_t kLxssRegKey[]                  = L"Software\\Microsoft\\Windows\\CurrentVersion\\Lxss";
constexpr std::wstring_view kDockerDistroPrefix  = L"docker-desktop";
constexpr std::wstring_view kRancherDistroPrefix = L"rancher-desktop";

wil::unique_hkey OpenWslRegKey() noexcept
{
    HKEY hKey = nullptr;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, kLxssRegKey, 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
        return wil::unique_hkey{hKey};
    }
    return wil::unique_hkey{};
}

wil::unique_hkey OpenDistroKey(const wil::unique_hkey& wslKey, const std::wstring& guid) noexcept
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

bool EnumerateDistroGuids(const wil::unique_hkey& wslKey, std::vector<std::wstring>& guids) noexcept
{
    if (! wslKey)
    {
        return false;
    }

    guids.clear();

    wchar_t buffer[39];
    for (DWORD index = 0;; ++index)
    {
        DWORD length      = ARRAYSIZE(buffer);
        const LONG result = RegEnumKeyExW(wslKey.get(), index, buffer, &length, nullptr, nullptr, nullptr, nullptr);
        if (result == ERROR_NO_MORE_ITEMS)
        {
            break;
        }

        if (result == ERROR_SUCCESS && length == 38 && buffer[0] == L'{' && buffer[37] == L'}')
        {
            guids.emplace_back(buffer, length);
        }
    }

    return true;
}

std::optional<std::wstring> ReadDistroName(const wil::unique_hkey& distroKey) noexcept
{
    if (! distroKey)
    {
        return std::nullopt;
    }

    wchar_t buffer[256];
    DWORD bufferSize  = sizeof(buffer);
    DWORD type        = 0;
    const LONG result = RegQueryValueExW(distroKey.get(), L"DistributionName", nullptr, &type, reinterpret_cast<BYTE*>(buffer), &bufferSize);
    if (result == ERROR_SUCCESS && type == REG_SZ)
    {
        return std::wstring(buffer);
    }

    return std::nullopt;
}

bool ShouldFilterDistroName(std::wstring_view name) noexcept
{
    if (name.size() >= kDockerDistroPrefix.size())
    {
        if (_wcsnicmp(name.data(), kDockerDistroPrefix.data(), kDockerDistroPrefix.size()) == 0)
        {
            return true;
        }
    }
    if (name.size() >= kRancherDistroPrefix.size())
    {
        if (_wcsnicmp(name.data(), kRancherDistroPrefix.data(), kRancherDistroPrefix.size()) == 0)
        {
            return true;
        }
    }

    return false;
}

std::vector<WslDistributionEntry> EnumerateWslDistributions() noexcept
{
    std::vector<WslDistributionEntry> distributions;

    auto wslKey = OpenWslRegKey();
    if (! wslKey)
    {
        return distributions;
    }

    std::vector<std::wstring> guids;
    if (! EnumerateDistroGuids(wslKey, guids))
    {
        return distributions;
    }

    for (const auto& guid : guids)
    {
        auto distroKey = OpenDistroKey(wslKey, guid);
        if (! distroKey)
        {
            continue;
        }

        const auto name = ReadDistroName(distroKey);
        if (! name.has_value() || name.value().empty())
        {
            continue;
        }

        const std::wstring& nameValue = name.value();
        if (ShouldFilterDistroName(nameValue))
        {
            continue;
        }

        WslDistributionEntry entry;
        entry.name        = nameValue;
        entry.networkPath = L"\\\\wsl.localhost\\" + nameValue;
        distributions.push_back(std::move(entry));
    }

    std::sort(distributions.begin(),
              distributions.end(),
              [](const WslDistributionEntry& a, const WslDistributionEntry& b) { return _wcsicmp(a.name.c_str(), b.name.c_str()) < 0; });

    return distributions;
}
} // namespace

HRESULT STDMETHODCALLTYPE FileSystem::GetMenuItems(const NavigationMenuItem** items, unsigned int* count) noexcept
{
    if (items == nullptr || count == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);

    _menuEntries.clear();
    _menuEntryView.clear();

    const auto addSeparator = [&]
    {
        MenuEntry entry;
        entry.flags = NAV_MENU_ITEM_FLAG_SEPARATOR;
        _menuEntries.push_back(std::move(entry));
    };

    const auto addEntry = [&](UINT labelId, std::wstring path, std::wstring iconPath)
    {
        std::wstring label = LoadStringResource(nullptr, labelId);
        if (label.empty() || path.empty())
        {
            return;
        }

        MenuEntry entry;
        entry.label    = std::move(label);
        entry.path     = std::move(path);
        entry.iconPath = std::move(iconPath);
        _menuEntries.push_back(std::move(entry));
    };

    const auto addRawEntry = [&](std::wstring label, std::wstring path, std::wstring iconPath)
    {
        if (label.empty() || path.empty())
        {
            return;
        }

        MenuEntry entry;
        entry.label    = std::move(label);
        entry.path     = std::move(path);
        entry.iconPath = std::move(iconPath);
        _menuEntries.push_back(std::move(entry));
    };

    const auto tryAddKnownFolder = [&](UINT labelId, const KNOWNFOLDERID& folderId)
    {
        wil::unique_cotaskmem_string folderPath;
        if (FAILED(SHGetKnownFolderPath(folderId, 0, nullptr, &folderPath)) || ! folderPath)
        {
            return;
        }

        std::wstring path(folderPath.get());
        addEntry(labelId, path, std::wstring(folderPath.get()));
    };

    // Quick access items
    tryAddKnownFolder(IDS_MENU_NAV_DESKTOP, FOLDERID_Desktop);
    tryAddKnownFolder(IDS_MENU_NAV_DOCUMENTS, FOLDERID_Documents);
    tryAddKnownFolder(IDS_MENU_NAV_DOWNLOADS, FOLDERID_Downloads);
    tryAddKnownFolder(IDS_MENU_NAV_PICTURES, FOLDERID_Pictures);
    tryAddKnownFolder(IDS_MENU_NAV_MUSIC, FOLDERID_Music);
    tryAddKnownFolder(IDS_MENU_NAV_VIDEOS, FOLDERID_Videos);
    tryAddKnownFolder(IDS_MENU_NAV_ONEDRIVE, FOLDERID_SkyDrive);

    const auto wslDistros = EnumerateWslDistributions();
    if (! wslDistros.empty())
    {
        addSeparator();
        for (const auto& distro : wslDistros)
        {
            addRawEntry(distro.name, distro.networkPath, distro.networkPath);
        }
    }

    // Drives
    const DWORD drives = GetLogicalDrives();
    bool addedDrive    = false;
    for (int i = 0; i < 26; ++i)
    {
        if ((drives & (1 << i)) == 0)
        {
            continue;
        }

        wchar_t drive[4]       = {static_cast<wchar_t>(L'A' + i), L':', L'\\', 0};
        std::wstring label     = GetDriveLabel(drive);
        std::wstring freeSpace = GetDriveFreeSpace(drive);

        std::wstring driveAndLabel = drive;
        if (! label.empty())
        {
            driveAndLabel.append(L" ");
            driveAndLabel.append(label);
        }

        std::wstring text = freeSpace.empty() ? driveAndLabel : std::format(L"{}\t{}", driveAndLabel, freeSpace);

        if (! addedDrive)
        {
            addSeparator();
            addedDrive = true;
        }

        addRawEntry(std::move(text), drive, drive);
    }

    _menuEntryView.reserve(_menuEntries.size());
    for (const auto& entry : _menuEntries)
    {
        NavigationMenuItem item{};
        item.flags     = entry.flags;
        item.label     = entry.label.empty() ? nullptr : entry.label.c_str();
        item.path      = entry.path.empty() ? nullptr : entry.path.c_str();
        item.iconPath  = entry.iconPath.empty() ? nullptr : entry.iconPath.c_str();
        item.commandId = entry.commandId;
        _menuEntryView.push_back(item);
    }

    *items = _menuEntryView.empty() ? nullptr : _menuEntryView.data();
    *count = static_cast<unsigned int>(_menuEntryView.size());
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::ExecuteMenuCommand([[maybe_unused]] unsigned int commandId) noexcept
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE FileSystem::SetCallback(INavigationMenuCallback* callback, void* cookie) noexcept
{
    std::lock_guard lock(_stateMutex);
    _navigationMenuCallback       = callback;
    _navigationMenuCallbackCookie = callback != nullptr ? cookie : nullptr;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::GetDriveInfo(const wchar_t* path, DriveInfo* info) noexcept
{
    if (info == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);

    info->flags       = DRIVE_INFO_FLAG_NONE;
    info->displayName = nullptr;
    info->volumeLabel = nullptr;
    info->fileSystem  = nullptr;
    info->totalBytes  = 0;
    info->freeBytes   = 0;
    info->usedBytes   = 0;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::filesystem::path root = std::filesystem::path(path).root_path();
    if (root.empty())
    {
        return S_FALSE;
    }

    _driveDisplayName = root.wstring();
    if (! _driveDisplayName.empty())
    {
        info->flags       = static_cast<DriveInfoFlags>(info->flags | DRIVE_INFO_FLAG_HAS_DISPLAY_NAME);
        info->displayName = _driveDisplayName.c_str();
    }

    _driveVolumeLabel.clear();
    _driveFileSystem.clear();

    wchar_t volumeName[MAX_PATH]     = {0};
    wchar_t fileSystemName[MAX_PATH] = {0};
    if (GetVolumeInformationW(root.c_str(), volumeName, ARRAYSIZE(volumeName), nullptr, nullptr, nullptr, fileSystemName, ARRAYSIZE(fileSystemName)))
    {
        if (volumeName[0] != L'\0')
        {
            _driveVolumeLabel = volumeName;
            info->flags       = static_cast<DriveInfoFlags>(info->flags | DRIVE_INFO_FLAG_HAS_VOLUME_LABEL);
            info->volumeLabel = _driveVolumeLabel.c_str();
        }

        if (fileSystemName[0] != L'\0')
        {
            _driveFileSystem = fileSystemName;
            info->flags      = static_cast<DriveInfoFlags>(info->flags | DRIVE_INFO_FLAG_HAS_FILE_SYSTEM);
            info->fileSystem = _driveFileSystem.c_str();
        }
    }

    ULARGE_INTEGER freeBytes{};
    ULARGE_INTEGER totalBytes{};
    ULARGE_INTEGER availableBytes{};
    if (GetDiskFreeSpaceExW(root.c_str(), &availableBytes, &totalBytes, &freeBytes))
    {
        const unsigned __int64 total = totalBytes.QuadPart;
        const unsigned __int64 free  = freeBytes.QuadPart;
        if (totalBytes.QuadPart > 0)
        {
            info->flags      = static_cast<DriveInfoFlags>(info->flags | DRIVE_INFO_FLAG_HAS_TOTAL_BYTES);
            info->totalBytes = total;
        }

        info->flags     = static_cast<DriveInfoFlags>(info->flags | DRIVE_INFO_FLAG_HAS_FREE_BYTES);
        info->freeBytes = free;

        if (total >= free)
        {
            info->flags     = static_cast<DriveInfoFlags>(info->flags | DRIVE_INFO_FLAG_HAS_USED_BYTES);
            info->usedBytes = total - free;
        }
    }

    _driveInfo = *info;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::GetDriveMenuItems(const wchar_t* path, const NavigationMenuItem** items, unsigned int* count) noexcept
{
    if (items == nullptr || count == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);

    _driveMenuEntries.clear();
    _driveMenuEntryView.clear();

    if (path == nullptr || path[0] == L'\0')
    {
        *items = nullptr;
        *count = 0;
        return E_INVALIDARG;
    }

    const std::filesystem::path root = std::filesystem::path(path).root_path();
    if (root.empty())
    {
        *items = nullptr;
        *count = 0;
        return S_OK;
    }

    const UINT driveType                = GetDriveTypeW(root.c_str());
    const bool isLocalDisk              = (driveType == DRIVE_FIXED || driveType == DRIVE_REMOVABLE);
    const NavigationMenuItemFlags flags = isLocalDisk ? NAV_MENU_ITEM_FLAG_NONE : NAV_MENU_ITEM_FLAG_DISABLED;

    MenuEntry properties;
    properties.label     = LoadStringResource(nullptr, IDS_MENU_DISK_PROPERTIES);
    properties.flags     = flags;
    properties.commandId = DRIVE_INFO_COMMAND_PROPERTIES;
    _driveMenuEntries.push_back(std::move(properties));

    MenuEntry cleanup;
    cleanup.label     = LoadStringResource(nullptr, IDS_MENU_DISK_CLEANUP);
    cleanup.flags     = flags;
    cleanup.commandId = DRIVE_INFO_COMMAND_CLEANUP;
    _driveMenuEntries.push_back(std::move(cleanup));

    _driveMenuEntryView.reserve(_driveMenuEntries.size());
    for (const auto& entry : _driveMenuEntries)
    {
        NavigationMenuItem item{};
        item.flags     = entry.flags;
        item.label     = entry.label.empty() ? nullptr : entry.label.c_str();
        item.path      = entry.path.empty() ? nullptr : entry.path.c_str();
        item.iconPath  = entry.iconPath.empty() ? nullptr : entry.iconPath.c_str();
        item.commandId = entry.commandId;
        _driveMenuEntryView.push_back(item);
    }

    *items = _driveMenuEntryView.empty() ? nullptr : _driveMenuEntryView.data();
    *count = static_cast<unsigned int>(_driveMenuEntryView.size());
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::ExecuteDriveMenuCommand(unsigned int commandId, const wchar_t* path) noexcept
{
    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::filesystem::path root = std::filesystem::path(path).root_path();
    if (root.empty())
    {
        return E_INVALIDARG;
    }

    if (commandId == DRIVE_INFO_COMMAND_PROPERTIES)
    {
        SHELLEXECUTEINFOW sei{};
        sei.cbSize = sizeof(sei);
        sei.fMask  = SEE_MASK_INVOKEIDLIST;
        sei.lpVerb = L"properties";
        sei.lpFile = root.c_str();
        sei.nShow  = SW_SHOW;
        if (! ShellExecuteExW(&sei))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        return S_OK;
    }

    if (commandId == DRIVE_INFO_COMMAND_CLEANUP)
    {
        std::wstring drive = root.wstring();
        if (drive.size() >= 2)
        {
            drive.resize(2);
        }

        std::wstring args      = std::format(L"/d {}", drive);
        const HINSTANCE result = ShellExecuteW(nullptr, nullptr, L"cleanmgr.exe", args.c_str(), nullptr, SW_SHOW);
        const auto code        = reinterpret_cast<INT_PTR>(result);
        if (code <= 32)
        {
            return HRESULT_FROM_WIN32(static_cast<DWORD>(code));
        }
        return S_OK;
    }

    return E_INVALIDARG;
}
