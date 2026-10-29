#pragma once

#include <algorithm>
#include <cstdint>
#include <filesystem>
#include <string_view>
#include <utility>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

namespace PluginModuleLifecycle
{
enum class ModuleUnloadMode : uint8_t
{
    FreeLibrary,
    ProcessShutdown,
};

inline constexpr std::wstring_view kDeferredUnloadError = L"Plugin unload is deferred until the plugin reports it can unload.";

[[nodiscard]] bool UnloadModule(wil::unique_hmodule& module, ModuleUnloadMode mode, const std::filesystem::path& path, std::wstring_view pluginKind) noexcept;

[[nodiscard]] bool PathsEqualNoCase(const std::filesystem::path& left, const std::filesystem::path& right) noexcept;

template <typename Entry> void MarkDeferred(Entry& entry)
{
    entry.unloadDeferred = true;
    entry.loadable       = false;
    entry.loadError      = kDeferredUnloadError;
}

template <typename Entry> [[nodiscard]] Entry MakeDeferredPlaceholder(const Entry& entry)
{
    Entry placeholder;
    placeholder.origin = entry.origin;
    if constexpr (requires { placeholder.type = entry.type; })
    {
        placeholder.type = entry.type;
    }
    placeholder.path            = entry.path;
    placeholder.factoryPluginId = entry.factoryPluginId;
    placeholder.id              = entry.id;
    placeholder.shortId         = entry.shortId;
    placeholder.name            = entry.name;
    placeholder.description     = entry.description;
    placeholder.author          = entry.author;
    placeholder.version         = entry.version;
    placeholder.disabled        = entry.disabled;
    MarkDeferred(placeholder);
    return placeholder;
}

template <typename Entry> [[nodiscard]] bool IsPathDeferred(const std::vector<Entry>& deferredEntries, const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    for (const Entry& entry : deferredEntries)
    {
        if (! entry.path.empty() && PathsEqualNoCase(entry.path, path))
        {
            return true;
        }
    }
    return false;
}

template <typename Entry> void RetainDeferredAliases(const std::vector<const Entry*>& unloadedEntries, std::vector<Entry>& deferredEntries)
{
    // Unload results are per owned module reference, but reload exclusion is per
    // DLL path. Keep every known logical alias until the final owner is gone.
    // These are owned metadata copies, never extra module pins or DLL pointers.
    for (const Entry* entry : unloadedEntries)
    {
        if (! IsPathDeferred(deferredEntries, entry->path))
        {
            continue;
        }

        // Rediscovery supplies the same immutable identities in its placeholders.
        // Do not accumulate another copy on every refresh while the DLL is busy.
        const bool alreadyRetained = std::any_of(deferredEntries.begin(), deferredEntries.end(), [&](const Entry& retained) {
            return retained.id == entry->id && retained.factoryPluginId == entry->factoryPluginId && PathsEqualNoCase(retained.path, entry->path);
        });
        if (! alreadyRetained)
        {
            deferredEntries.push_back(MakeDeferredPlaceholder(*entry));
        }
    }
}

template <typename Entry, typename UnloadEntry>
void UnloadAll(std::vector<Entry>& entries, std::vector<Entry>& deferredEntries, ModuleUnloadMode mode, UnloadEntry&& unloadEntry) noexcept
{
    std::vector<const Entry*> unloadedEntries;
    unloadedEntries.reserve(entries.size());
    for (auto it = entries.rbegin(); it != entries.rend(); ++it)
    {
        if (! unloadEntry(*it, mode))
        {
            deferredEntries.push_back(std::move(*it));
        }
        else
        {
            unloadedEntries.push_back(&*it);
        }
    }
    RetainDeferredAliases(unloadedEntries, deferredEntries);
}

template <typename Entry, typename UnloadEntry>
void SweepDeferred(std::vector<Entry>& deferredEntries, ModuleUnloadMode mode, UnloadEntry&& unloadEntry) noexcept
{
    if (deferredEntries.empty())
    {
        return;
    }

    std::vector<Entry> stillDeferred;
    stillDeferred.reserve(deferredEntries.size());
    std::vector<const Entry*> unloadedEntries;
    unloadedEntries.reserve(deferredEntries.size());
    for (Entry& entry : deferredEntries)
    {
        if (! unloadEntry(entry, mode))
        {
            stillDeferred.push_back(std::move(entry));
        }
        else
        {
            unloadedEntries.push_back(&entry);
        }
    }
    RetainDeferredAliases(unloadedEntries, stillDeferred);
    deferredEntries = std::move(stillDeferred);
}
} // namespace PluginModuleLifecycle
