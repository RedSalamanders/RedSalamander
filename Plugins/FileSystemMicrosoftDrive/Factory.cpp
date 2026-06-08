#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <array>
#include <new>
#include <optional>
#include <string>
#include <string_view>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#define PLUGFACTORY_EXPORTS
#include "PlugInterfaces/Factory.h"

#define REDSAL_DEFINE_TRACE_PROVIDER
#include "FileSystemMicrosoftDriveResources.h"
#include "Helpers.h"

#include "FileSystemMicrosoftDrive.h"

extern HINSTANCE g_hInstance;

namespace
{
constexpr wchar_t kPluginIdOneDrivePersonal[]      = L"builtin/file-system-onedrive-personal";
constexpr wchar_t kPluginShortIdOneDrivePersonal[] = L"onedrive";

constexpr wchar_t kPluginIdOneDriveBusiness[]      = L"builtin/file-system-onedrive-business";
constexpr wchar_t kPluginShortIdOneDriveBusiness[] = L"onedrive-pro";

constexpr wchar_t kPluginIdSharePoint[]      = L"builtin/file-system-sharepoint";
constexpr wchar_t kPluginShortIdSharePoint[] = L"sharepoint";

struct LocalizedPluginMetaDataSet
{
    std::wstring oneDrivePersonalName;
    std::wstring oneDrivePersonalDescription;
    std::wstring oneDriveBusinessName;
    std::wstring oneDriveBusinessDescription;
    std::wstring sharePointName;
    std::wstring sharePointDescription;
    std::array<PluginMetaData, 3> plugins{};

    LocalizedPluginMetaDataSet()
    {
        LoadStringResource(g_hInstance, IDS_FILESYSTEMMICROSOFTDRIVE_ONEDRIVE_PERSONAL_NAME, oneDrivePersonalName);
        LoadStringResource(g_hInstance, IDS_FILESYSTEMMICROSOFTDRIVE_ONEDRIVE_PERSONAL_DESCRIPTION, oneDrivePersonalDescription);
        LoadStringResource(g_hInstance, IDS_FILESYSTEMMICROSOFTDRIVE_ONEDRIVE_BUSINESS_NAME, oneDriveBusinessName);
        LoadStringResource(g_hInstance, IDS_FILESYSTEMMICROSOFTDRIVE_ONEDRIVE_BUSINESS_DESCRIPTION, oneDriveBusinessDescription);
        LoadEmbeddedStringResource(g_hInstance, IDS_FILESYSTEMMICROSOFTDRIVE_SHAREPOINT_NAME, sharePointName);
        LoadStringResource(g_hInstance, IDS_FILESYSTEMMICROSOFTDRIVE_SHAREPOINT_DESCRIPTION, sharePointDescription);

        // PluginMetaData keeps raw wchar_t* pointers, so bind them only after the
        // backing strings live in their final static storage.
        plugins = {{
            {
                .id          = kPluginIdOneDrivePersonal,
                .shortId     = kPluginShortIdOneDrivePersonal,
                .name        = oneDrivePersonalName.c_str(),
                .description = oneDrivePersonalDescription.c_str(),
                .author      = L"RedSalamander",
                .version     = VERSINFO_PLUGIN_VERSION,
            },
            {
                .id          = kPluginIdOneDriveBusiness,
                .shortId     = kPluginShortIdOneDriveBusiness,
                .name        = oneDriveBusinessName.c_str(),
                .description = oneDriveBusinessDescription.c_str(),
                .author      = L"RedSalamander",
                .version     = VERSINFO_PLUGIN_VERSION,
            },
            {
                .id          = kPluginIdSharePoint,
                .shortId     = kPluginShortIdSharePoint,
                .name        = sharePointName.c_str(),
                .description = sharePointDescription.c_str(),
                .author      = L"RedSalamander",
                .version     = VERSINFO_PLUGIN_VERSION,
            },
        }};
    }
};

[[nodiscard]] const LocalizedPluginMetaDataSet& GetPluginMetaDataSet() noexcept
{
    static const LocalizedPluginMetaDataSet data;

    return data;
}

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

[[nodiscard]] const char* GetPluginSchema(std::wstring_view pluginId) noexcept
{
    const auto mode = ModeFromPluginId(pluginId);
    if (! mode.has_value())
    {
        return nullptr;
    }

    return GetFileSystemMicrosoftDriveStaticConfigurationSchema(mode.value());
}

HRESULT CreatePluginInstance(REFIID riid, IHost* host, std::wstring_view pluginId, void** result)
{
    if (riid != __uuidof(IFileSystem))
    {
        return E_NOINTERFACE;
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
} // namespace

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

    const auto& plugins = GetPluginMetaDataSet().plugins;
    *metaData           = plugins.data();
    *count              = static_cast<unsigned int>(plugins.size());
    return S_OK;
}

extern "C" HRESULT __stdcall RedSalamanderCreate(REFIID riid, const FactoryOptions* /*factoryOptions*/, IHost* host, const wchar_t* pluginId, void** result)
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

    return CreatePluginInstance(riid, host, pluginId, result);
}

extern "C" HRESULT __stdcall RedSalamanderGetConfigurationSchema(REFIID riid, const wchar_t* pluginId, const char** schemaJsonUtf8)
{
    if (! schemaJsonUtf8)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = nullptr;
    if (riid != __uuidof(IFileSystem))
    {
        return E_NOINTERFACE;
    }
    if (! pluginId || pluginId[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const char* schema = GetPluginSchema(pluginId);
    if (! schema)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    *schemaJsonUtf8 = schema;
    return S_OK;
}
