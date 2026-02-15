#pragma once

#include <cstdint>
#include <filesystem>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027) // WIL: deleted copy/move operators
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/Viewer.h"

namespace Common::Settings
{
struct Settings;
}

class ViewerPluginManager final
{
public:
    enum class PluginType : uint8_t
    {
        Viewer,
    };

    enum class PluginOrigin : uint8_t
    {
        Embedded,
        Optional,
        Custom,
    };

    struct PluginEntry
    {
        PluginEntry()                              = default;
        PluginEntry(const PluginEntry&)            = delete;
        PluginEntry& operator=(const PluginEntry&) = delete;
        PluginEntry(PluginEntry&&)                 = default;
        PluginEntry& operator=(PluginEntry&&)      = default;
        ~PluginEntry()                             = default;

        PluginOrigin origin = PluginOrigin::Embedded;
        std::filesystem::path path;

        // When non-empty, this DLL exposes multiple logical plugins and this is the
        // plugin id to request via RedSalamanderCreateEx().
        std::wstring factoryPluginId;

        bool loadable = false;
        bool disabled = false;
        std::wstring loadError;

        std::wstring id;
        std::wstring shortId;
        std::wstring name;
        std::wstring description;
        std::wstring author;
        std::wstring version;

        wil::unique_hmodule module;
        FARPROC createFactory   = nullptr;
        FARPROC createFactoryEx = nullptr;
    };

    static ViewerPluginManager& GetInstance() noexcept;

    HRESULT Initialize(Common::Settings::Settings& settings) noexcept;
    void Shutdown(Common::Settings::Settings& settings) noexcept;

    HRESULT Refresh(Common::Settings::Settings& settings) noexcept;

    const std::vector<PluginEntry>& GetPlugins() const noexcept;

    HRESULT CreateViewerInstance(std::wstring_view pluginId, Common::Settings::Settings& settings, wil::com_ptr<IViewer>& outViewer) noexcept;

    HRESULT DisablePlugin(std::wstring_view pluginId, Common::Settings::Settings& settings) noexcept;
    HRESULT EnablePlugin(std::wstring_view pluginId, Common::Settings::Settings& settings) noexcept;
    HRESULT AddCustomPluginPath(const std::filesystem::path& path, Common::Settings::Settings& settings) noexcept;

    HRESULT GetConfigurationSchema(std::wstring_view pluginId, Common::Settings::Settings& settings, std::string& outSchemaJsonUtf8) noexcept;
    HRESULT GetConfiguration(std::wstring_view pluginId, Common::Settings::Settings& settings, std::string& outConfigurationJsonUtf8) noexcept;
    HRESULT SetConfiguration(std::wstring_view pluginId, std::string_view configurationJsonUtf8, Common::Settings::Settings& settings) noexcept;

    HRESULT TestPlugin(std::wstring_view pluginId) noexcept;

private:
    ViewerPluginManager()  = default;
    ~ViewerPluginManager() = default;

    ViewerPluginManager(const ViewerPluginManager&)            = delete;
    ViewerPluginManager(ViewerPluginManager&&)                 = delete;
    ViewerPluginManager& operator=(const ViewerPluginManager&) = delete;
    ViewerPluginManager& operator=(ViewerPluginManager&&)      = delete;

    HRESULT Discover(Common::Settings::Settings& settings) noexcept;
    HRESULT EnsureLoaded(PluginEntry& entry) noexcept;
    void Unload(PluginEntry& entry) noexcept;

    std::filesystem::path GetExecutableDirectory() noexcept;
    std::filesystem::path GetOptionalPluginsDirectory() noexcept;

    std::optional<size_t> FindPluginIndexById(std::wstring_view pluginId) const noexcept;
    PluginEntry* FindPluginById(std::wstring_view pluginId) noexcept;
    const PluginEntry* FindPluginById(std::wstring_view pluginId) const noexcept;

private:
    bool _initialized = false;
    std::filesystem::path _exeDir;
    std::vector<PluginEntry> _plugins;
};
