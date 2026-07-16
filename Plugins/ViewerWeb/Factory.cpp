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

#include "PlugInterfaces/FactoryImpl.h"

extern HINSTANCE g_hInstance;
void ResetSharedEnvironment() noexcept;
[[nodiscard]] bool CanUnloadViewerWebModuleNow() noexcept;

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

// Per-entry metadata thunks (return contiguous array elements).
const PluginMetaData* GetMetaDataWeb() noexcept
{
    return &GetPluginMetaDataSet().plugins[0];
}
const PluginMetaData* GetMetaDataJson() noexcept
{
    return &GetPluginMetaDataSet().plugins[1];
}
const PluginMetaData* GetMetaDataMarkdown() noexcept
{
    return &GetPluginMetaDataSet().plugins[2];
}

// Per-entry schema thunks.
const char* GetSchemaWeb() noexcept
{
    return GetViewerWebStaticConfigurationSchema(ViewerWebKind::Web);
}
const char* GetSchemaJson() noexcept
{
    return GetViewerWebStaticConfigurationSchema(ViewerWebKind::Json);
}
const char* GetSchemaMarkdown() noexcept
{
    return GetViewerWebStaticConfigurationSchema(ViewerWebKind::Markdown);
}

// Per-entry creation thunks.
HRESULT CreateInstanceWeb(const FactoryOptions* /*factoryOptions*/, IHost* host, void** result) noexcept
{
    auto* instance = new (std::nothrow) ViewerWeb(ViewerWebKind::Web);
    if (! instance)
        return E_OUTOFMEMORY;
    instance->SetHost(host);
    const HRESULT hr = instance->QueryInterface(__uuidof(IViewer), result);
    instance->Release();
    return hr;
}
HRESULT CreateInstanceJson(const FactoryOptions* /*factoryOptions*/, IHost* host, void** result) noexcept
{
    auto* instance = new (std::nothrow) ViewerWeb(ViewerWebKind::Json);
    if (! instance)
        return E_OUTOFMEMORY;
    instance->SetHost(host);
    const HRESULT hr = instance->QueryInterface(__uuidof(IViewer), result);
    instance->Release();
    return hr;
}
HRESULT CreateInstanceMarkdown(const FactoryOptions* /*factoryOptions*/, IHost* host, void** result) noexcept
{
    auto* instance = new (std::nothrow) ViewerWeb(ViewerWebKind::Markdown);
    if (! instance)
        return E_OUTOFMEMORY;
    instance->SetHost(host);
    const HRESULT hr = instance->QueryInterface(__uuidof(IViewer), result);
    instance->Release();
    return hr;
}

const PluginFactoryEntry kEntries[] = {
    {&GetMetaDataWeb, &GetSchemaWeb, &CreateInstanceWeb},
    {&GetMetaDataJson, &GetSchemaJson, &CreateInstanceJson},
    {&GetMetaDataMarkdown, &GetSchemaMarkdown, &CreateInstanceMarkdown},
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

extern "C" PLUGFACTORY_API void __stdcall RedSalamanderPluginShutdown() noexcept
{
    ResetSharedEnvironment();
}

extern "C" PLUGFACTORY_API BOOL __stdcall RedSalamanderPluginRetainModuleUntilProcessExit() noexcept
{
    return CanUnloadViewerWebModuleNow() ? FALSE : TRUE;
}

extern "C" PLUGFACTORY_API BOOL __stdcall RedSalamanderPluginCanUnloadNow() noexcept
{
    return CanUnloadViewerWebModuleNow() ? TRUE : FALSE;
}
