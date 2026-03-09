#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <iterator>
#include <new>
#include <array>
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

#include "resource.h"
#include "ViewerWeb.h"

extern HINSTANCE g_hInstance;

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
};

[[nodiscard]] const LocalizedPluginMetaDataSet& GetPluginMetaDataSet() noexcept
{
    static const LocalizedPluginMetaDataSet data = [] {
        LocalizedPluginMetaDataSet value{};
        value.webName             = LoadStringResource(g_hInstance, IDS_VIEWERWEB_NAME);
        value.webDescription      = LoadStringResource(g_hInstance, IDS_VIEWERWEB_DESCRIPTION);
        value.jsonName            = LoadStringResource(g_hInstance, IDS_VIEWERJSON_NAME);
        value.jsonDescription     = LoadStringResource(g_hInstance, IDS_VIEWERJSON_DESCRIPTION);
        value.markdownName        = LoadStringResource(g_hInstance, IDS_VIEWERMARKDOWN_NAME);
        value.markdownDescription = LoadStringResource(g_hInstance, IDS_VIEWERMARKDOWN_DESCRIPTION);
        value.plugins             = {{
            {
                .id          = L"builtin/viewer-web",
                .shortId     = L"web",
                .name        = value.webName.c_str(),
                .description = value.webDescription.c_str(),
                .author      = L"RedSalamander",
                .version     = L"0.1",
            },
            {
                .id          = L"builtin/viewer-json",
                .shortId     = L"json",
                .name        = value.jsonName.c_str(),
                .description = value.jsonDescription.c_str(),
                .author      = L"RedSalamander",
                .version     = L"0.1",
            },
            {
                .id          = L"builtin/viewer-markdown",
                .shortId     = L"md",
                .name        = value.markdownName.c_str(),
                .description = value.markdownDescription.c_str(),
                .author      = L"RedSalamander",
                .version     = L"0.1",
            },
        }};
        return value;
    }();

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
} // namespace

extern "C" HRESULT __stdcall RedSalamanderCreate(REFIID riid, const FactoryOptions* /*factoryOptions*/, IHost* host, void** result)
{
    if (result == nullptr)
    {
        return E_POINTER;
    }

    *result = nullptr;

    if (riid == __uuidof(IViewer))
    {
        // Backward-compatible single-plugin entry point.
        auto* instance = new (std::nothrow) ViewerWeb(ViewerWebKind::Web);
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

extern "C" HRESULT __stdcall RedSalamanderCreateEx(REFIID riid, const FactoryOptions* /*factoryOptions*/, IHost* host, const wchar_t* pluginId, void** result)
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

    auto* instance = new (std::nothrow) ViewerWeb(kind.value());
    if (! instance)
    {
        return E_OUTOFMEMORY;
    }

    instance->SetHost(host);

    const HRESULT hr = instance->QueryInterface(riid, result);
    instance->Release();
    return hr;
}
