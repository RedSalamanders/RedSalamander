#pragma once
// FactoryImpl.h — shared plugin factory implementation.
//
// Usage in each Plugins\*\Factory.cpp:
//   1. #define PLUGFACTORY_EXPORTS before #include "PlugInterfaces/Factory.h"
//   2. #include "Helpers.h"   (provides OrdinalString::EqualsNoCase)
//   3. #include "PlugInterfaces/FactoryImpl.h"
//   4. Define per-entry PluginFactoryEntry values (see below).
//   5. Implement the three extern "C" exports as one-line delegations
//      to FactoryEnumeratePlugins, FactoryCreate, FactoryGetConfigurationSchema.
//
// Include-order requirement: OrdinalString::EqualsNoCase must be declared
// before this header is included. Every Factory.cpp already includes Helpers.h
// (to define the ETW trace provider), so that requirement is satisfied
// automatically. Do NOT include this header before Helpers.h.

#include <cstdint>
#include <limits>
#include <span>

#include "PlugInterfaces/Factory.h"

// ---------------------------------------------------------------------------
// PluginFactoryEntry
//
// One entry per logical plugin id within a DLL. Each function pointer is a
// thin per-entry thunk that delegates to the plugin's existing creation /
// schema functions.
//
// Conventions:
//   - getMetaData : zero-arg thunk returning a pointer to this entry's
//                   PluginMetaData (static lifetime). Called lazily so
//                   DLL-global resource handles (g_hInstance, etc.) are
//                   already valid when the metadata strings are loaded.
//   - getSchema   : zero-arg thunk returning this entry's JSON schema (UTF-8).
//                   May be nullptr if the plugin has no schema.
//   - createInstance : called only after riid/null-result/id-match validation.
//                   The shared helper has already validated everything; the thunk
//                   only constructs the object and returns it via QueryInterface.
//                   First param (factoryOptions) is forwarded from the export;
//                   leave it UNNAMED in each thunk today (currently unused,
//                   naming it triggers C4100 under /W4).
//
// For EnumeratePlugins to return a contiguous array of PluginMetaData, multi-
// plugin DLLs MUST ensure their entries are ordered so that getMetaData() for
// entry[i] returns &pluginsArray[i], where pluginsArray is the contiguous
// std::array<PluginMetaData, N> owned by the DLL. Single-plugin DLLs have
// count == 1 so contiguity is trivially satisfied.
// ---------------------------------------------------------------------------
struct PluginFactoryEntry
{
    const PluginMetaData* (*getMetaData)() noexcept;
    const char* (*getSchema)() noexcept;
    HRESULT (*createInstance)(const FactoryOptions* factoryOptions, IHost* host, void** result) noexcept;
};

// ---------------------------------------------------------------------------
// FactoryEnumeratePlugins<TInterface>
//
// Implements RedSalamanderEnumeratePlugins:
//   E_POINTER     — null out-param
//   E_NOINTERFACE — riid mismatch
//   ERROR_TOO_MANY_NAMES — entries exceeds the host's bounded metadata contract
//   S_OK          — fills *metaData (contiguous array base) and *count
// ---------------------------------------------------------------------------
template <typename TInterface>
[[nodiscard]] HRESULT FactoryEnumeratePlugins(std::span<const PluginFactoryEntry> entries,
                                              REFIID riid,
                                              const PluginMetaData** metaData,
                                              unsigned int* count) noexcept
{
    if (! metaData || ! count)
    {
        return E_POINTER;
    }

    *metaData = nullptr;
    *count    = 0;

    if (riid != __uuidof(TInterface))
    {
        return E_NOINTERFACE;
    }

    if (entries.empty())
    {
        return E_NOINTERFACE;
    }

    if (entries.size() > kMaxEnumeratedPluginsPerModule)
    {
        return HRESULT_FROM_WIN32(ERROR_TOO_MANY_NAMES);
    }

    if (! entries[0].getMetaData)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const PluginMetaData* firstMetaData = entries[0].getMetaData();
    if (! firstMetaData)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const std::uintptr_t firstAddress = reinterpret_cast<std::uintptr_t>(firstMetaData);
    for (size_t index = 1u; index < entries.size(); ++index)
    {
        if (! entries[index].getMetaData || index > ((std::numeric_limits<std::uintptr_t>::max)() - firstAddress) / sizeof(PluginMetaData))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        const PluginMetaData* currentMetaData = entries[index].getMetaData();
        const std::uintptr_t expectedAddress  = firstAddress + index * sizeof(PluginMetaData);
        if (! currentMetaData || reinterpret_cast<std::uintptr_t>(currentMetaData) != expectedAddress)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
    }

    *metaData = firstMetaData; // contiguous array base (verified above)
    *count    = static_cast<unsigned int>(entries.size());
    return S_OK;
}

// ---------------------------------------------------------------------------
// FactoryCreate<TInterface>
//
// Implements RedSalamanderCreate:
//   E_POINTER     — null result
//   E_NOINTERFACE — riid mismatch
//   E_INVALIDARG  — null/empty pluginId when entries.size() > 1
//   ERROR_NOT_FOUND — non-empty pluginId matches no entry
//   delegates to entry.createInstance on match
// ---------------------------------------------------------------------------
template <typename TInterface>
[[nodiscard]] HRESULT FactoryCreate(std::span<const PluginFactoryEntry> entries,
                                    REFIID riid,
                                    const FactoryOptions* factoryOptions,
                                    IHost* host,
                                    const wchar_t* pluginId,
                                    void** result) noexcept
{
    if (! result)
    {
        return E_POINTER;
    }

    *result = nullptr;

    if (riid != __uuidof(TInterface))
    {
        return E_NOINTERFACE;
    }

    if (entries.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    const bool emptyId = (! pluginId || pluginId[0] == L'\0');

    if (entries.size() == 1)
    {
        // Single-plugin DLL: empty/null id means "the only plugin".
        if (! emptyId && ! OrdinalString::EqualsNoCase(pluginId, entries[0].getMetaData()->id))
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
        }
        return entries[0].createInstance(factoryOptions, host, result);
    }

    // Multi-plugin DLL: empty/null id is ambiguous.
    if (emptyId)
    {
        return E_INVALIDARG;
    }

    for (const auto& entry : entries)
    {
        if (OrdinalString::EqualsNoCase(pluginId, entry.getMetaData()->id))
        {
            return entry.createInstance(factoryOptions, host, result);
        }
    }

    return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
}

// ---------------------------------------------------------------------------
// FactoryGetConfigurationSchema<TInterface>
//
// Implements RedSalamanderGetConfigurationSchema:
//   E_POINTER     — null schemaJsonUtf8
//   E_NOINTERFACE — riid mismatch
//   E_INVALIDARG  — null/empty pluginId when entries.size() > 1
//   ERROR_NOT_FOUND — id matches no entry, or matched entry has no schema
//   S_OK          — fills *schemaJsonUtf8
// ---------------------------------------------------------------------------
template <typename TInterface>
[[nodiscard]] HRESULT FactoryGetConfigurationSchema(std::span<const PluginFactoryEntry> entries,
                                                    REFIID riid,
                                                    const wchar_t* pluginId,
                                                    const char** schemaJsonUtf8) noexcept
{
    if (! schemaJsonUtf8)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = nullptr;

    if (riid != __uuidof(TInterface))
    {
        return E_NOINTERFACE;
    }

    if (entries.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    const bool emptyId = (! pluginId || pluginId[0] == L'\0');

    if (entries.size() == 1)
    {
        // Single-plugin DLL: empty/null id means "the only plugin".
        if (! emptyId && ! OrdinalString::EqualsNoCase(pluginId, entries[0].getMetaData()->id))
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
        }
        if (! entries[0].getSchema)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
        }
        const char* schema = entries[0].getSchema();
        if (! schema)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
        }
        *schemaJsonUtf8 = schema;
        return S_OK;
    }

    // Multi-plugin DLL: empty/null id is ambiguous.
    if (emptyId)
    {
        return E_INVALIDARG;
    }

    for (const auto& entry : entries)
    {
        if (OrdinalString::EqualsNoCase(pluginId, entry.getMetaData()->id))
        {
            if (! entry.getSchema)
            {
                return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
            }
            const char* schema = entry.getSchema();
            if (! schema)
            {
                return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
            }
            *schemaJsonUtf8 = schema;
            return S_OK;
        }
    }

    return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
}
