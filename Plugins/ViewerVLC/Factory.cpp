#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <new>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514) // WIL headers: deleted copy/move and unused inline Helpers
#include <wil/com.h>
#pragma warning(pop)

#define PLUGFACTORY_EXPORTS
#include "PlugInterfaces/Factory.h"

// Define the ETW provider for ViewerVLC.dll
#define REDSAL_DEFINE_TRACE_PROVIDER
#include "Helpers.h"

#include "PlugInterfaces/Viewer.h"
#include "ViewerVLC.h"
#include "resource.h"

#include "PlugInterfaces/FactoryImpl.h"

extern HINSTANCE g_hInstance;

namespace
{
[[nodiscard]] const PluginMetaData* GetMetaData() noexcept
{
    static const std::wstring name        = LoadEmbeddedStringResource(g_hInstance, IDS_VIEWERVLC_NAME);
    static const std::wstring description = LoadStringResource(g_hInstance, IDS_VIEWERVLC_DESCRIPTION);
    static const PluginMetaData metaData  = {
        .id          = L"builtin/viewer-vlc",
        .shortId     = L"viewvlc",
        .name        = name.c_str(),
        .description = description.c_str(),
        .author      = nullptr,
        .version     = VERSINFO_PLUGIN_VERSION,
    };
    return &metaData;
}

[[nodiscard]] const char* GetSchema() noexcept
{
    return GetViewerVlcStaticConfigurationSchema();
}

HRESULT CreateInstance(const FactoryOptions* /*factoryOptions*/, IHost* host, void** result) noexcept
{
    auto* instance = new (std::nothrow) ViewerVLC();
    if (instance == nullptr)
    {
        return E_OUTOFMEMORY;
    }

    instance->SetHost(host);

    const HRESULT hr = instance->QueryInterface(__uuidof(IViewer), result);
    instance->Release();
    return hr;
}

const PluginFactoryEntry kEntries[] = {
    {&GetMetaData, &GetSchema, &CreateInstance},
};
} // namespace

extern "C" HRESULT __stdcall RedSalamanderEnumeratePlugins(REFIID riid, const PluginMetaData** metaData, unsigned int* count)
{
    return FactoryEnumeratePlugins<IViewer>(kEntries, riid, metaData, count);
}

extern "C" HRESULT __stdcall RedSalamanderCreate(REFIID riid, const FactoryOptions* factoryOptions, IHost* host, const wchar_t* pluginId, void** result)
{
    return FactoryCreate<IViewer>(kEntries, riid, factoryOptions, host, pluginId, result);
}

extern "C" HRESULT __stdcall RedSalamanderGetConfigurationSchema(REFIID riid, const wchar_t* pluginId, const char** schemaJsonUtf8)
{
    return FactoryGetConfigurationSchema<IViewer>(kEntries, riid, pluginId, schemaJsonUtf8);
}
