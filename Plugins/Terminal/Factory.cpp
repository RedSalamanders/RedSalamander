#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <memory>
#include <mutex>
#include <new>
#include <string>

#define PLUGFACTORY_EXPORTS
#include "PlugInterfaces/Factory.h"
#include "PlugInterfaces/Terminal.h"

#define REDSAL_DEFINE_TRACE_PROVIDER
#include "Helpers.h"
#include "PlugInterfaces/FactoryImpl.h"
#include "Terminal.h"
#include "resource.h"

extern HINSTANCE g_hInstance;

namespace
{
struct PluginMetaDataStorage final
{
    PluginMetaDataStorage()
        : name(LoadStringResource(g_hInstance, IDS_TERMINAL_NAME)),
          description(LoadStringResource(g_hInstance, IDS_TERMINAL_DESCRIPTION)),
          metaData{
              .id          = L"builtin/terminal",
              .shortId     = L"terminal",
              .name        = name.c_str(),
              .description = description.c_str(),
              .author      = L"RedSalamander",
              .version     = VERSINFO_PLUGIN_VERSION,
          }
    {
    }

    std::wstring name;
    std::wstring description;
    PluginMetaData metaData;
};

std::mutex g_metaDataMutex;
std::unique_ptr<PluginMetaDataStorage> g_metaData;

[[nodiscard]] const PluginMetaData* GetMetaData() noexcept
{
    std::scoped_lock lock(g_metaDataMutex);
    if (! g_metaData)
    {
        g_metaData = std::make_unique<PluginMetaDataStorage>();
    }
    return &g_metaData->metaData;
}

[[nodiscard]] const char* GetSchema() noexcept
{
    return GetTerminalStaticConfigurationSchema();
}

HRESULT CreateInstance(const FactoryOptions* /*factoryOptions*/, IHost* /*host*/, void** result) noexcept
{
    auto* instance = new (std::nothrow) Terminal();
    if (instance == nullptr)
    {
        return E_OUTOFMEMORY;
    }
    const HRESULT hr = instance->QueryInterface(__uuidof(ITerminal), result);
    instance->Release();
    return hr;
}

const PluginFactoryEntry kEntries[] = {
    {&GetMetaData, &GetSchema, &CreateInstance},
};
} // namespace

extern "C" HRESULT __stdcall RedSalamanderEnumeratePlugins(REFIID riid, const PluginMetaData** metaData, unsigned int* count)
{
    return FactoryEnumeratePlugins<ITerminal>(kEntries, riid, metaData, count);
}

extern "C" HRESULT __stdcall RedSalamanderCreate(REFIID riid, const FactoryOptions* factoryOptions, IHost* host, const wchar_t* pluginId, void** result)
{
    return FactoryCreate<ITerminal>(kEntries, riid, factoryOptions, host, pluginId, result);
}

extern "C" HRESULT __stdcall RedSalamanderGetConfigurationSchema(REFIID riid, const wchar_t* pluginId, const char** schemaJsonUtf8)
{
    return FactoryGetConfigurationSchema<ITerminal>(kEntries, riid, pluginId, schemaJsonUtf8);
}

extern "C" __declspec(dllexport) void __stdcall RedSalamanderPluginShutdown() noexcept
{
    std::scoped_lock lock(g_metaDataMutex);
    g_metaData.reset();
}

extern "C" __declspec(dllexport) BOOL __stdcall RedSalamanderPluginCanUnloadNow() noexcept
{
    return CanUnloadTerminalModuleNow() ? TRUE : FALSE;
}
