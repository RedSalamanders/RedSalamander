#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <new>
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
#include "FileSystemGoogleDriveResources.h"
#include "Helpers.h"

#include "FileSystemGoogleDrive.h"

extern HINSTANCE g_hInstance;

namespace
{
[[nodiscard]] const PluginMetaData& GetPluginMetaData() noexcept
{
    static const std::wstring name        = LoadStringResource(g_hInstance, IDS_FILESYSTEMGOOGLEDRIVE_NAME);
    static const std::wstring description = LoadStringResource(g_hInstance, IDS_FILESYSTEMGOOGLEDRIVE_DESCRIPTION);
    static const PluginMetaData metaData  = {
        .id          = L"builtin/file-system-gdrive",
        .shortId     = L"gdrive",
        .name        = name.c_str(),
        .description = description.c_str(),
        .author      = L"RedSalamander",
        .version     = VERSINFO_PLUGIN_VERSION,
    };
    return metaData;
}

[[nodiscard]] const char* GetPluginSchema(std::wstring_view pluginId) noexcept
{
    if (! pluginId.empty() && ! OrdinalString::EqualsNoCase(pluginId, GetPluginMetaData().id))
    {
        return nullptr;
    }

    return GetFileSystemGoogleDriveStaticConfigurationSchema();
}

HRESULT CreatePluginInstance(REFIID riid, IHost* host, void** result)
{
    if (riid != __uuidof(IFileSystem))
    {
        return E_NOINTERFACE;
    }

    auto* instance = new (std::nothrow) FileSystemGoogleDrive(host);
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

    *metaData = &GetPluginMetaData();
    *count    = 1;
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
    if (pluginId && pluginId[0] != L'\0' && ! OrdinalString::EqualsNoCase(pluginId, GetPluginMetaData().id))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    return CreatePluginInstance(riid, host, result);
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

    const std::wstring_view requestedId = pluginId ? std::wstring_view(pluginId) : std::wstring_view{};
    const char* schema                  = GetPluginSchema(requestedId);
    if (! schema)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    *schemaJsonUtf8 = schema;
    return S_OK;
}
