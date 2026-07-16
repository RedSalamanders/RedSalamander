#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <new>
#include <string>
#include <string_view>

#define PLUGFACTORY_EXPORTS
#include "PlugInterfaces/Factory.h"

#include "FileSystemDummy.h"
#include "FileSystemDummyResources.h"
#include "Helpers.h"

#include "PlugInterfaces/FactoryImpl.h"

extern HINSTANCE g_hInstance;

namespace
{
[[nodiscard]] const PluginMetaData* GetMetaData() noexcept
{
    static const std::wstring name        = LoadStringResource(g_hInstance, IDS_FILESYSTEMDUMMY_NAME);
    static const std::wstring description = LoadStringResource(g_hInstance, IDS_FILESYSTEMDUMMY_DESCRIPTION);
    static const PluginMetaData metaData  = {
        .id          = L"builtin/file-system-dummy",
        .shortId     = L"fk",
        .name        = name.c_str(),
        .description = description.c_str(),
        .author      = L"RedSalamander",
        .version     = VERSINFO_PLUGIN_VERSION,
    };
    return &metaData;
}

[[nodiscard]] const char* GetSchema() noexcept
{
    return GetFileSystemDummyStaticConfigurationSchema();
}

HRESULT CreateInstance(const FactoryOptions* /*factoryOptions*/, IHost* host, void** result) noexcept
{
    auto* instance = new (std::nothrow) FileSystemDummy(host);
    if (! instance)
    {
        return E_OUTOFMEMORY;
    }

    const HRESULT hr = instance->QueryInterface(__uuidof(IFileSystem), result);
    instance->Release();
    return hr;
}

const PluginFactoryEntry kEntries[] = {
    {&GetMetaData, &GetSchema, &CreateInstance},
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
