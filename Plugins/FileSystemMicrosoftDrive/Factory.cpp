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

#include "PlugInterfaces/FactoryImpl.h"

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

// Per-entry metadata thunks (return contiguous array elements).
const PluginMetaData* GetMetaDataOneDrivePersonal() noexcept
{
    return &GetPluginMetaDataSet().plugins[0];
}
const PluginMetaData* GetMetaDataOneDriveBusiness() noexcept
{
    return &GetPluginMetaDataSet().plugins[1];
}
const PluginMetaData* GetMetaDataSharePoint() noexcept
{
    return &GetPluginMetaDataSet().plugins[2];
}

// Per-entry schema thunks.
const char* GetSchemaOneDrivePersonal() noexcept
{
    return GetFileSystemMicrosoftDriveStaticConfigurationSchema(FileSystemMicrosoftDriveMode::OneDrivePersonal);
}
const char* GetSchemaOneDriveBusiness() noexcept
{
    return GetFileSystemMicrosoftDriveStaticConfigurationSchema(FileSystemMicrosoftDriveMode::OneDriveBusiness);
}
const char* GetSchemaSharePoint() noexcept
{
    return GetFileSystemMicrosoftDriveStaticConfigurationSchema(FileSystemMicrosoftDriveMode::SharePoint);
}

// Per-entry creation thunks.
HRESULT CreateInstanceOneDrivePersonal(const FactoryOptions* /*factoryOptions*/, IHost* host, void** result) noexcept
{
    auto* instance = new (std::nothrow) FileSystemMicrosoftDrive(FileSystemMicrosoftDriveMode::OneDrivePersonal, host);
    if (! instance)
        return E_OUTOFMEMORY;
    const HRESULT hr = instance->QueryInterface(__uuidof(IFileSystem), result);
    instance->Release();
    return hr;
}
HRESULT CreateInstanceOneDriveBusiness(const FactoryOptions* /*factoryOptions*/, IHost* host, void** result) noexcept
{
    auto* instance = new (std::nothrow) FileSystemMicrosoftDrive(FileSystemMicrosoftDriveMode::OneDriveBusiness, host);
    if (! instance)
        return E_OUTOFMEMORY;
    const HRESULT hr = instance->QueryInterface(__uuidof(IFileSystem), result);
    instance->Release();
    return hr;
}
HRESULT CreateInstanceSharePoint(const FactoryOptions* /*factoryOptions*/, IHost* host, void** result) noexcept
{
    auto* instance = new (std::nothrow) FileSystemMicrosoftDrive(FileSystemMicrosoftDriveMode::SharePoint, host);
    if (! instance)
        return E_OUTOFMEMORY;
    const HRESULT hr = instance->QueryInterface(__uuidof(IFileSystem), result);
    instance->Release();
    return hr;
}

const PluginFactoryEntry kEntries[] = {
    {&GetMetaDataOneDrivePersonal, &GetSchemaOneDrivePersonal, &CreateInstanceOneDrivePersonal},
    {&GetMetaDataOneDriveBusiness, &GetSchemaOneDriveBusiness, &CreateInstanceOneDriveBusiness},
    {&GetMetaDataSharePoint, &GetSchemaSharePoint, &CreateInstanceSharePoint},
};
} // namespace

extern "C" HRESULT __stdcall RedSalamanderEnumeratePlugins(REFIID riid, const PluginMetaData** metaData, unsigned int* count)
{
    return FactoryEnumeratePlugins<IFileSystem>(kEntries, riid, metaData, count);
}

extern "C" HRESULT __stdcall RedSalamanderCreate(REFIID riid, const FactoryOptions* factoryOptions, IHost* host, const wchar_t* pluginId, void** result)
{
    return FactoryCreate<IFileSystem>(kEntries, riid, factoryOptions, host, pluginId, result);
}

extern "C" HRESULT __stdcall RedSalamanderGetConfigurationSchema(REFIID riid, const wchar_t* pluginId, const char** schemaJsonUtf8)
{
    return FactoryGetConfigurationSchema<IFileSystem>(kEntries, riid, pluginId, schemaJsonUtf8);
}
