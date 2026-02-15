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

    PLUGFACTORY_API HRESULT __stdcall RedSalamanderCreate(REFIID riid, const FactoryOptions* factoryOptions, IHost* host, void** result);

    // Optional multi-plugin support:
    //
    // - A single DLL may implement multiple logical plugins for the same interface type.
    // - The host will call RedSalamanderEnumeratePlugins to get the list of PluginMetaData entries.
    // - The host will then call RedSalamanderCreateEx with the desired plugin id (metaData[i].id).
    //
    // If these exports are missing, the host falls back to RedSalamanderCreate.
    //
    // Ownership / lifetime:
    // - The returned PluginMetaData array and all strings are owned by the DLL and remain valid
    //   until the DLL is unloaded. Callers MUST NOT free them.
    PLUGFACTORY_API HRESULT __stdcall RedSalamanderEnumeratePlugins(REFIID riid, const PluginMetaData** metaData, unsigned int* count);
    PLUGFACTORY_API
    HRESULT __stdcall RedSalamanderCreateEx(REFIID riid, const FactoryOptions* factoryOptions, IHost* host, const wchar_t* pluginId, void** result);
}
