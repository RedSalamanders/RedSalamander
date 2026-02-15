#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <iterator>
#include <new>
#include <optional>
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

namespace
{
static constexpr PluginMetaData kViewerWebPlugins[] = {
    {
        .id          = L"builtin/viewer-web",
        .shortId     = L"web",
        .name        = L"Web Viewer",
        .description = L"WebView2-based viewer for HTML and PDF files.",
        .author      = L"RedSalamander",
        .version     = L"0.1",
    },
    {
        .id          = L"builtin/viewer-json",
        .shortId     = L"json",
        .name        = L"JSON Viewer",
        .description = L"WebView2-based JSON/JSON5 viewer with folding and syntax highlighting.",
        .author      = L"RedSalamander",
        .version     = L"0.1",
    },
    {
        .id          = L"builtin/viewer-markdown",
        .shortId     = L"md",
        .name        = L"Markdown Viewer",
        .description = L"WebView2-based Markdown viewer with syntax highlighting.",
        .author      = L"RedSalamander",
        .version     = L"0.1",
    },
};

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

    *metaData = kViewerWebPlugins;
    *count    = static_cast<unsigned int>(std::size(kViewerWebPlugins));
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
