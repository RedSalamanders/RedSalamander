#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <new>
#include <optional>
#include <string_view>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#define PLUGFACTORY_EXPORTS
#include "PlugInterfaces/Factory.h"

#define REDSAL_DEFINE_TRACE_PROVIDER
#include "Helpers.h"

#include "FileSystemMicrosoftDrive.h"

namespace
{
constexpr wchar_t kPluginIdOneDrivePersonal[]      = L"builtin/file-system-onedrive-personal";
constexpr wchar_t kPluginShortIdOneDrivePersonal[] = L"onedrivep";
constexpr wchar_t kPluginNameOneDrivePersonal[]    = L"OneDrive Personal";
constexpr wchar_t kPluginDescOneDrivePersonal[]    = L"Microsoft OneDrive Personal storage over Microsoft Graph.";

constexpr wchar_t kPluginIdOneDriveBusiness[]      = L"builtin/file-system-onedrive-business";
constexpr wchar_t kPluginShortIdOneDriveBusiness[] = L"onedriveb";
constexpr wchar_t kPluginNameOneDriveBusiness[]    = L"OneDrive Business";
constexpr wchar_t kPluginDescOneDriveBusiness[]    = L"Microsoft OneDrive for Business storage over Microsoft Graph.";

constexpr wchar_t kPluginIdSharePoint[]      = L"builtin/file-system-sharepoint";
constexpr wchar_t kPluginShortIdSharePoint[] = L"sharepoint";
constexpr wchar_t kPluginNameSharePoint[]    = L"SharePoint";
constexpr wchar_t kPluginDescSharePoint[]    = L"Microsoft SharePoint document libraries over Microsoft Graph.";

constexpr PluginMetaData kPlugins[] = {
    {
        .id          = kPluginIdOneDrivePersonal,
        .shortId     = kPluginShortIdOneDrivePersonal,
        .name        = kPluginNameOneDrivePersonal,
        .description = kPluginDescOneDrivePersonal,
        .author      = L"RedSalamander",
        .version     = L"0.1",
    },
    {
        .id          = kPluginIdOneDriveBusiness,
        .shortId     = kPluginShortIdOneDriveBusiness,
        .name        = kPluginNameOneDriveBusiness,
        .description = kPluginDescOneDriveBusiness,
        .author      = L"RedSalamander",
        .version     = L"0.1",
    },
    {
        .id          = kPluginIdSharePoint,
        .shortId     = kPluginShortIdSharePoint,
        .name        = kPluginNameSharePoint,
        .description = kPluginDescSharePoint,
        .author      = L"RedSalamander",
        .version     = L"0.1",
    },
};

[[nodiscard]] std::optional<FileSystemMicrosoftDriveMode> ModeFromPluginId(std::wstring_view pluginId) noexcept
{
    if (pluginId == kPluginIdOneDrivePersonal)
    {
        return FileSystemMicrosoftDriveMode::OneDrivePersonal;
    }
    if (pluginId == kPluginIdOneDriveBusiness)
    {
        return FileSystemMicrosoftDriveMode::OneDriveBusiness;
    }
    if (pluginId == kPluginIdSharePoint)
    {
        return FileSystemMicrosoftDriveMode::SharePoint;
    }
    return std::nullopt;
}
} // namespace

extern "C" HRESULT __stdcall RedSalamanderCreate(REFIID riid, const FactoryOptions* /*factoryOptions*/, IHost* host, void** result)
{
    if (! result)
    {
        return E_POINTER;
    }

    *result = nullptr;

    if (riid != __uuidof(IFileSystem))
    {
        return E_NOINTERFACE;
    }

    auto* instance = new (std::nothrow) FileSystemMicrosoftDrive(FileSystemMicrosoftDriveMode::OneDriveBusiness, host);
    if (! instance)
    {
        return E_OUTOFMEMORY;
    }

    const HRESULT hr = instance->QueryInterface(riid, result);
    instance->Release();
    return hr;
}

extern "C" HRESULT __stdcall RedSalamanderEnumeratePlugins(REFIID riid, const PluginMetaData** metaData, unsigned int* count)
{
    if (! metaData || ! count)
    {
        return E_POINTER;
    }

    *metaData = nullptr;
    *count    = 0;

    if (riid != __uuidof(IFileSystem))
    {
        return E_NOINTERFACE;
    }

    *metaData = kPlugins;
    *count    = static_cast<unsigned int>(std::size(kPlugins));
    return S_OK;
}

extern "C" HRESULT __stdcall RedSalamanderCreateEx(REFIID riid, const FactoryOptions* /*factoryOptions*/, IHost* host, const wchar_t* pluginId, void** result)
{
    if (! result)
    {
        return E_POINTER;
    }

    *result = nullptr;

    if (riid != __uuidof(IFileSystem))
    {
        return E_NOINTERFACE;
    }

    if (! pluginId || pluginId[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const auto mode = ModeFromPluginId(pluginId);
    if (! mode.has_value())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    auto* instance = new (std::nothrow) FileSystemMicrosoftDrive(mode.value(), host);
    if (! instance)
    {
        return E_OUTOFMEMORY;
    }

    const HRESULT hr = instance->QueryInterface(riid, result);
    instance->Release();
    return hr;
}
