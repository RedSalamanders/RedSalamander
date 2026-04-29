#pragma once

#include <unknwn.h>
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include "Host.h"
#include "Informations.h"

#ifdef PLUGFACTORY_EXPORTS
#define PLUGFACTORY_API __declspec(dllexport)
#else
#define PLUGFACTORY_API __declspec(dllimport)
#endif
extern "C"
{
#pragma warning(push)
// 4865 : the underlying type will change from 'int' to 'unsigned int' when '/Zc:enumTypes' is specified on the command line
// 4820 : padding in data structure
#pragma warning(disable : 4865 4820)
    enum DebugLevel : uint32_t
    {
        DEBUG_LEVEL_NONE        = 0,
        DEBUG_LEVEL_ERROR       = 1,
        DEBUG_LEVEL_WARNING     = 2,
        DEBUG_LEVEL_INFORMATION = 3,
    };

    typedef struct FactoryOptions
    {
        DebugLevel debugLevel;
    } FactoryOptions;
#pragma warning(pop)

    // Required creation entry point:
    //
    // - A single DLL may implement one or more logical plugins for the same interface type.
    // - The host will call RedSalamanderEnumeratePlugins to get the list of PluginMetaData entries.
    // - The host will then call RedSalamanderCreate with the desired plugin id (metaData[i].id).
    // - For single-plugin DLLs, pluginId may be nullptr or empty.
    //
    // Ownership / lifetime:
    // - The returned object follows normal COM ownership rules.
    // - `host` is caller-owned and remains valid for the lifetime of the created plugin instance.
    PLUGFACTORY_API
    HRESULT __stdcall RedSalamanderCreate(REFIID riid, const FactoryOptions* factoryOptions, IHost* host, const wchar_t* pluginId, void** result);

    // Optional multi-plugin discovery support:
    //
    // - A single DLL may implement multiple logical plugins for the same interface type.
    // - The host calls RedSalamanderEnumeratePlugins to discover PluginMetaData entries before creation.
    //
    // Ownership / lifetime:
    // - The returned PluginMetaData array and all strings are owned by the DLL and remain valid
    //   until the DLL is unloaded. Callers MUST NOT free them.
    // - Because PluginMetaData stores raw wchar_t* fields, plugin DLLs MUST bind those fields only
    //   after the backing string storage lives in its final static/object lifetime; temporary or
    //   lambda-local std::wstring storage is not allowed.
    PLUGFACTORY_API HRESULT __stdcall RedSalamanderEnumeratePlugins(REFIID riid, const PluginMetaData** metaData, unsigned int* count);

    // Optional static configuration-schema support:
    //
    // - Lets hosts fetch JSON schema text without constructing a live plugin instance.
    // - Useful for startup/discovery and settings export paths where only static schema
    //   metadata is needed.
    // - For single-plugin DLLs, pluginId may be nullptr or empty.
    // - The returned JSON string is owned by the DLL and remains valid until unload.
    PLUGFACTORY_API HRESULT __stdcall RedSalamanderGetConfigurationSchema(REFIID riid, const wchar_t* pluginId, const char** schemaJsonUtf8);
}
