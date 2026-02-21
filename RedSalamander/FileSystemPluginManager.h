#pragma once

#include <cstdint>
#include <filesystem>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "SettingsStore.h"

class FileSystemPluginManager
{
public:
    enum class PluginOrigin : uint8_t
    {
        Embedded,
        Optional,
        Custom,
    };

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted)
#pragma warning(disable : 4625 4626)
    struct PluginEntry
    {
        PluginOrigin origin = PluginOrigin::Embedded;
        std::filesystem::path path;

        // When non-empty, this DLL exposes multiple logical plugins and this is the
        // plugin id to request via RedSalamanderCreateEx().
        std::wstring factoryPluginId;

        std::wstring id;
        std::wstring shortId;
        std::wstring name;
        std::wstring description;
        std::wstring author;
        std::wstring version;

        bool disabled = false;
        bool loadable = false;
        std::wstring loadError;

        wil::unique_hmodule module;
        wil::com_ptr<IFileSystem> fileSystem;
        wil::com_ptr<IInformations> informations;
    };
#pragma warning(pop)

    static FileSystemPluginManager& GetInstance() noexcept;

    HRESULT Initialize(Common::Settings::Settings& settings) noexcept;
    void Shutdown(Common::Settings::Settings& settings) noexcept;

    const std::vector<PluginEntry>& GetPlugins() const noexcept;

    std::wstring_view GetActivePluginId() const noexcept;
    wil::com_ptr<IFileSystem> GetActiveFileSystem() const noexcept;
    std::optional<std::filesystem::path> TryGetPluginPath(std::wstring_view pluginId) const noexcept;
    std::optional<std::filesystem::path> TryGetActivePluginPath() const noexcept;

    HRESULT Refresh(Common::Settings::Settings& settings) noexcept;
    HRESULT SetActivePlugin(std::wstring_view pluginId, Common::Settings::Settings& settings) noexcept;
    HRESULT DisablePlugin(std::wstring_view pluginId, Common::Settings::Settings& settings) noexcept;
    HRESULT EnablePlugin(std::wstring_view pluginId, Common::Settings::Settings& settings) noexcept;
    HRESULT RemoveCustomPlugin(std::wstring_view pluginId, Common::Settings::Settings& settings) noexcept;
    HRESULT AddCustomPluginPath(const std::filesystem::path& path, Common::Settings::Settings& settings) noexcept;

    HRESULT GetConfigurationSchema(std::wstring_view pluginId, Common::Settings::Settings& settings, std::string& outSchemaJsonUtf8) noexcept;
    HRESULT GetConfiguration(std::wstring_view pluginId, Common::Settings::Settings& settings, std::string& outConfigurationJsonUtf8) noexcept;
    HRESULT SetConfiguration(std::wstring_view pluginId, std::string_view configurationJsonUtf8, Common::Settings::Settings& settings) noexcept;

    HRESULT TestPlugin(std::wstring_view pluginId) noexcept;

private:
    FileSystemPluginManager() = default;

    std::filesystem::path GetExecutableDirectory() noexcept;
    std::filesystem::path GetOptionalPluginsDirectory() noexcept;

    std::optional<size_t> FindPluginIndexById(std::wstring_view pluginId) const noexcept;
    PluginEntry* FindPluginById(std::wstring_view pluginId) noexcept;
    const PluginEntry* FindPluginById(std::wstring_view pluginId) const noexcept;

    HRESULT Discover(Common::Settings::Settings& settings) noexcept;
    HRESULT EnsureLoaded(PluginEntry& entry, Common::Settings::Settings& settings) noexcept;
    void Unload(PluginEntry& entry) noexcept;

    HRESULT ApplyConfigurationFromSettings(PluginEntry& entry, const Common::Settings::Settings& settings) noexcept;
    void PersistConfigurationToSettings(const PluginEntry& entry, Common::Settings::Settings& settings) noexcept;

    bool _initialized = false;
    std::filesystem::path _exeDir;
    std::vector<PluginEntry> _plugins;
    std::wstring _activePluginId;
};
