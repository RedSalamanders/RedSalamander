#pragma once

#include <cstdint>

#include <unknwn.h>
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include "Host.h"
#include "Informations.h"

// RedSalamanderEnumeratePlugins returns a borrowed contiguous array whose
// physical length cannot be verified independently by the host. Keep the
// accepted count deliberately small so a buggy or ABI-mismatched plugin cannot
// make the host walk an implausibly large metadata range.
inline constexpr unsigned int kMaxEnumeratedPluginsPerModule = 256u;

[[nodiscard]] constexpr bool IsValidEnumeratedPluginCount(unsigned int count) noexcept
{
    return count > 0u && count <= kMaxEnumeratedPluginsPerModule;
}

[[nodiscard]] constexpr bool IsValidEnumeratedPluginRange(const PluginMetaData* metaData, unsigned int count) noexcept
{
    return metaData != nullptr && IsValidEnumeratedPluginCount(count);
}

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

    enum FactoryConnectionBrowseKind : uint32_t
    {
        FACTORY_CONNECTION_BROWSE_DEVICES  = 1,
        FACTORY_CONNECTION_BROWSE_STORAGES = 2,
    };

    typedef struct FactoryConnectionBrowseRequest
    {
        uint32_t sizeBytes; // sizeof(FactoryConnectionBrowseRequest)
        uint32_t kind;      // FactoryConnectionBrowseKind

        // Required for FACTORY_CONNECTION_BROWSE_STORAGES. The identifier is
        // plugin-defined; for the built-in MTP plugin this is the WPD PnP id
        // on live devices and the fake backend's stable device id in selftests.
        const wchar_t* parentDeviceId;

        uint32_t reserved[8];
    } FactoryConnectionBrowseRequest;

    typedef struct FactoryConnectionBrowseResult
    {
        uint32_t sizeBytes; // sizeof(FactoryConnectionBrowseResult)

        // UTF-8 JSON response allocated with CoTaskMemAlloc. The caller owns it
        // and must free it with CoTaskMemFree().
        //
        // Device response:
        //   {"version":1,"devices":[{"pnpId":"...","friendlyName":"...","devicePuid":"..."}]}
        // Storage response:
        //   {"version":1,"storages":[{"name":"...","persistentId":"...","objectId":"...","initialPath":"..."}]}
        char* jsonUtf8;

        uint32_t reserved[8];
    } FactoryConnectionBrowseResult;
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
    // - S_OK requires a non-null metadata pointer and a count in [1, kMaxEnumeratedPluginsPerModule].
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

    // Optional static connection-browse support:
    //
    // - Lets hosts populate connection-profile pickers through the plugin's own
    //   identity and enumeration rules instead of duplicating device APIs.
    // - For single-plugin DLLs, pluginId may be nullptr or empty.
    // - On S_OK, result->jsonUtf8 is allocated with CoTaskMemAlloc and must be
    //   freed by the caller with CoTaskMemFree().
    // - Omitted export means the plugin does not provide connection browse.
    PLUGFACTORY_API HRESULT __stdcall RedSalamanderBrowseConnectionTargets(
        REFIID riid, const wchar_t* pluginId, const FactoryConnectionBrowseRequest* request, FactoryConnectionBrowseResult* result) noexcept;

    // Optional module-level quiet-point support:
    //
    // - Hosts discover these exports with GetProcAddress; plugins that do not
    //   own DLL-global schedulers, caches, window classes, or driver-backed
    //   resources should omit them.
    // - RedSalamanderPluginShutdown must be idempotent and non-throwing. After
    //   it returns, no DLL-global worker may call host callbacks or touch state
    //   that will be released before FreeLibrary.
    // - RedSalamanderPluginCanUnloadNow is called after shutdown and before
    //   runtime resource-owner unregister / FreeLibrary. Returning FALSE means
    //   live DLL-global work is still unwinding; the host keeps the module
    //   mapped, leaves the resource owner registered, skips same-path reload,
    //   and retries later. Omitted export is treated as TRUE.
    // - RedSalamanderPluginRetainModuleUntilProcessExit is honored only during
    //   process shutdown. Returning TRUE lets a plugin run its quiet point while
    //   leaving the DLL mapped for OS process teardown.
    PLUGFACTORY_API void __stdcall RedSalamanderPluginShutdown() noexcept;
    PLUGFACTORY_API BOOL __stdcall RedSalamanderPluginCanUnloadNow() noexcept;
    PLUGFACTORY_API BOOL __stdcall RedSalamanderPluginRetainModuleUntilProcessExit() noexcept;
}
