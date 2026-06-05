#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <array>
#include <iterator>
#include <new>
#include <optional>
#include <string>
#include <string_view>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514) // WIL headers: deleted copy/move and unused inline Helpers
#include <wil/com.h>
#pragma warning(pop)

#define PLUGFACTORY_EXPORTS
#include "PlugInterfaces/Factory.h"

// Define the ETW provider for ViewerWeb.dll
#define REDSAL_DEFINE_TRACE_PROVIDER
#include "Helpers.h"

#include "ViewerWeb.h"
#include "resource.h"

extern HINSTANCE g_hInstance;
void ResetSharedEnvironment() noexcept;

namespace
{
struct LocalizedPluginMetaDataSet
{
    std::wstring webName;
    std::wstring webDescription;
    std::wstring jsonName;
    std::wstring jsonDescription;
    std::wstring markdownName;
    std::wstring markdownDescription;
    std::array<PluginMetaData, 3> plugins{};

    LocalizedPluginMetaDataSet()
    {
        LoadStringResource(g_hInstance, IDS_VIEWERWEB_NAME, webName);
        LoadStringResource(g_hInstance, IDS_VIEWERWEB_DESCRIPTION, webDescription);
        LoadStringResource(g_hInstance, IDS_VIEWERJSON_NAME, jsonName);
        LoadStringResource(g_hInstance, IDS_VIEWERJSON_DESCRIPTION, jsonDescription);
        LoadStringResource(g_hInstance, IDS_VIEWERMARKDOWN_NAME, markdownName);
        LoadStringResource(g_hInstance, IDS_VIEWERMARKDOWN_DESCRIPTION, markdownDescription);

        // PluginMetaData keeps raw wchar_t* pointers, so bind them only after the
        // backing strings live in their final static storage.
        plugins = {{
            {
                .id          = L"builtin/viewer-web",
                .shortId     = L"web",
                .name        = webName.c_str(),
                .description = webDescription.c_str(),
                .author      = L"RedSalamander",
                .version     = VERSINFO_PLUGIN_VERSION,
            },
            {
                .id          = L"builtin/viewer-json",
                .shortId     = L"json",
                .name        = jsonName.c_str(),
                .description = jsonDescription.c_str(),
                .author      = L"RedSalamander",
                .version     = VERSINFO_PLUGIN_VERSION,
            },
            {
                .id          = L"builtin/viewer-markdown",
                .shortId     = L"md",
                .name        = markdownName.c_str(),
                .description = markdownDescription.c_str(),
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

static std::optional<ViewerWebKind> KindFromPluginId(std::wstring_view pluginId) noexcept
{
    if (pluginId == L"builtin/viewer-web")
    {
        return ViewerWebKind::Web;
    }
    if (pluginId == L"builtin/viewer-json")
    {
        return ViewerWebKind::Json;
    }
    if (pluginId == L"builtin/viewer-markdown")
    {
        return ViewerWebKind::Markdown;
    }
    return std::nullopt;
}

[[nodiscard]] const char* GetPluginSchema(std::wstring_view pluginId) noexcept
{
    const auto kind = KindFromPluginId(pluginId);
    if (! kind.has_value())
    {
        return nullptr;
    }

    return GetViewerWebStaticConfigurationSchema(kind.value());
}
} // namespace

HRESULT CreatePluginInstance(REFIID riid, IHost* host, ViewerWebKind kind, void** result)
{
    if (result == nullptr)
    {
        return E_POINTER;
    }

    *result = nullptr;

    if (riid == __uuidof(IViewer))
    {
        auto* instance = new (std::nothrow) ViewerWeb(kind);
        if (! instance)
        {
            return E_OUTOFMEMORY;
        }

        instance->SetHost(host);

        const HRESULT hr = instance->QueryInterface(riid, result);
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

    if (riid != __uuidof(IViewer))
    {
        return E_NOINTERFACE;
    }

    if (! pluginId || pluginId[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const auto kind = KindFromPluginId(pluginId);
    if (! kind.has_value())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    return CreatePluginInstance(riid, host, kind.value(), result);
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

extern "C" PLUGFACTORY_API void __stdcall RedSalamanderPluginShutdown() noexcept
{
    ResetSharedEnvironment();
}

extern "C" PLUGFACTORY_API BOOL __stdcall RedSalamanderPluginRetainModuleUntilProcessExit() noexcept
{
    return FALSE;
}
