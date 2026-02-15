#include "FileSystemS3.Internal.h"

namespace FsS3 = FileSystemS3Internal;

HRESULT STDMETHODCALLTYPE FileSystemS3::GetDriveInfo(const wchar_t* path, DriveInfo* info) noexcept
{
    if (info == nullptr)
    {
        return E_POINTER;
    }

    Settings settings;
    wil::com_ptr<IHostConnections> hostConnections;
    FileSystemS3Mode mode = FileSystemS3Mode::S3;
    const wchar_t* scheme = nullptr;
    {
        std::lock_guard lock(_stateMutex);
        settings        = _settings;
        hostConnections = _hostConnections;
        mode            = _mode;
        scheme          = _metaData.shortId ? _metaData.shortId : L"";
    }

    const std::wstring_view pluginPath = (path != nullptr && path[0] != L'\0') ? std::wstring_view(path) : std::wstring_view(L"/");

    FsS3::ResolvedAwsContext ctx{};
    std::wstring canonical;
    const HRESULT hr = FsS3::ResolveAwsContext(mode, settings, pluginPath, hostConnections.get(), false, ctx, canonical);

    const std::wstring normalized = canonical.empty() ? FsS3::NormalizePluginPath(pluginPath) : canonical;

    std::wstring driveDisplayName;
    if (normalized == L"/" || normalized.empty())
    {
        driveDisplayName = std::format(L"{}:/", scheme);
    }
    else
    {
        std::wstring_view rest = normalized;
        while (! rest.empty() && rest.front() == L'/')
        {
            rest.remove_prefix(1);
        }

        const size_t slash                = rest.find(L'/');
        const std::wstring_view authority = (slash == std::wstring_view::npos) ? rest : rest.substr(0, slash);
        const std::wstring_view tail      = (slash == std::wstring_view::npos) ? std::wstring_view{} : rest.substr(slash);
        driveDisplayName                  = std::format(L"{}://{}{}", scheme, authority, tail);
    }

    {
        std::lock_guard lock(_stateMutex);
        _driveDisplayName = std::move(driveDisplayName);

        info->flags       = static_cast<DriveInfoFlags>(DRIVE_INFO_FLAG_HAS_DISPLAY_NAME | DRIVE_INFO_FLAG_HAS_FILE_SYSTEM);
        info->displayName = _driveDisplayName.empty() ? nullptr : _driveDisplayName.c_str();
        info->volumeLabel = nullptr;
        info->fileSystem  = _driveFileSystem.empty() ? nullptr : _driveFileSystem.c_str();
        info->totalBytes  = 0;
        info->freeBytes   = 0;
        info->usedBytes   = 0;
    }

    return SUCCEEDED(hr) ? S_OK : hr;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::GetDriveMenuItems([[maybe_unused]] const wchar_t* path, const NavigationMenuItem** items, unsigned int* count) noexcept
{
    if (items == nullptr || count == nullptr)
    {
        return E_POINTER;
    }

    *items = nullptr;
    *count = 0;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemS3::ExecuteDriveMenuCommand([[maybe_unused]] unsigned int commandId, [[maybe_unused]] const wchar_t* path) noexcept
{
    return E_NOTIMPL;
}
