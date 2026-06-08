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

extern HINSTANCE g_hInstance;

namespace
{
[[nodiscard]] const PluginMetaData& GetPluginMetaData() noexcept
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
    return metaData;
}

[[nodiscard]] const char* GetPluginSchema(std::wstring_view pluginId) noexcept
{
    if (! pluginId.empty() && ! OrdinalString::EqualsNoCase(pluginId, GetPluginMetaData().id))
    {
        return nullptr;
    }

    return GetViewerVlcStaticConfigurationSchema();
}
} // namespace

HRESULT CreatePluginInstance(REFIID riid, IHost* host, void** result)
{
    if (result == nullptr)
    {
        return E_POINTER;
    }

    *result = nullptr;

    if (riid == __uuidof(IViewer))
    {
        auto* instance = new (std::nothrow) ViewerVLC();
        if (instance == nullptr)
        {
            return E_OUTOFMEMORY;
        }

        instance->SetHost(host);

        HRESULT hr = instance->QueryInterface(riid, result);
        instance->Release();
        return hr;
    }

    return E_NOINTERFACE;
}

extern "C" HRESULT __stdcall RedSalamanderEnumeratePlugins(REFIID riid, const PluginMetaData** metaData, unsigned int* count)
{
    if (! metaData || ! count)
    {
        return E_POINTER;
    }

    *metaData = nullptr;
    *count    = 0;
    if (riid != __uuidof(IViewer))
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
    if (riid != __uuidof(IViewer))
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
    if (riid != __uuidof(IViewer))
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
