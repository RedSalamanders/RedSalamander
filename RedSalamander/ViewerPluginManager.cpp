#include "Framework.h"

#include "ViewerPluginManager.h"

#include <algorithm>
#include <cwctype>
#include <format>
#include <limits>
#include <unordered_set>

#include "Helpers.h"
#include "HostServices.h"
#include "PlugInterfaces/Factory.h"
#include "SettingsStore.h"

namespace
{
using CreateFactoryFunc    = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, void**);
using EnumeratePluginsFunc = HRESULT(__stdcall*)(REFIID, const PluginMetaData** metaData, unsigned int* count);
using CreateFactoryExFunc  = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const wchar_t* pluginId, void**);

bool IsDllPath(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    std::error_code ec;
    if (! std::filesystem::is_regular_file(path, ec))
    {
        return false;
    }

    const std::wstring ext = path.extension().wstring();
    return _wcsicmp(ext.c_str(), L".dll") == 0;
}

std::wstring SafeCoalesce(const wchar_t* value) noexcept
{
    return value ? std::wstring(value) : std::wstring();
}

std::string SafeCoalesce(const char* value) noexcept
{
    return value ? std::string(value) : std::string();
}

bool IsValidShortId(std::wstring_view shortId) noexcept
{
    if (shortId.empty())
    {
        return false;
    }

    for (wchar_t ch : shortId)
    {
        if (std::iswalnum(static_cast<wint_t>(ch)) == 0)
        {
            return false;
        }
    }

    return true;
}

std::wstring ToLowerInvariant(std::wstring_view text)
{
    std::wstring lowered(text);
    std::transform(lowered.begin(), lowered.end(), lowered.begin(), [](wchar_t ch) { return static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch))); });
    return lowered;
}

bool EqualsNoCase(std::wstring_view a, std::wstring_view b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    if (a.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return false;
    }

    const int len = static_cast<int>(a.size());
    return CompareStringOrdinal(a.data(), len, b.data(), len, TRUE) == CSTR_EQUAL;
}

void RemoveStringFromVector(std::vector<std::wstring>& values, std::wstring_view needle)
{
    values.erase(std::remove_if(values.begin(), values.end(), [&](const std::wstring& v) { return EqualsNoCase(v, needle); }), values.end());
}

HRESULT ApplyConfigurationFromSettings(IInformations& infos, std::wstring_view pluginId, const Common::Settings::Settings& settings) noexcept
{
    if (pluginId.empty())
    {
        return S_FALSE;
    }

    const auto it = settings.plugins.configurationByPluginId.find(std::wstring(pluginId));
    if (it == settings.plugins.configurationByPluginId.end())
    {
        return infos.SetConfiguration(nullptr);
    }

    const Common::Settings::JsonValue& configValue = it->second;
    if (std::holds_alternative<std::monostate>(configValue.value))
    {
        return infos.SetConfiguration(nullptr);
    }

    std::string configText;
    const HRESULT serializeHr = Common::Settings::SerializeJsonValue(configValue, configText);
    if (FAILED(serializeHr))
    {
        Debug::Warning(L"Failed to serialize viewer plugin configuration JSON for '{}' (hr=0x{:08X}); configuration will be ignored.",
                       std::wstring(pluginId),
                       static_cast<unsigned long>(serializeHr));
        return infos.SetConfiguration(nullptr);
    }

    return infos.SetConfiguration(configText.empty() ? nullptr : configText.c_str());
}
} // namespace

ViewerPluginManager& ViewerPluginManager::GetInstance() noexcept
{
    static ViewerPluginManager instance;
    return instance;
}

HRESULT ViewerPluginManager::Initialize(Common::Settings::Settings& settings) noexcept
{
    if (_initialized)
    {
        return S_OK;
    }

    _exeDir = GetExecutableDirectory();
    if (_exeDir.empty())
    {
        Debug::ErrorWithLastError(L"Failed to get executable directory.");
        return E_FAIL;
    }

    const HRESULT hr = Refresh(settings);
    if (FAILED(hr))
    {
        Debug::Error(L"Failed to discover viewer plugins (hr=0x{:08X}).", static_cast<unsigned long>(hr));
        return hr;
    }

    _initialized = true;
    return S_OK;
}

void ViewerPluginManager::Shutdown(Common::Settings::Settings& /*settings*/) noexcept
{
    if (! _initialized)
    {
        return;
    }

    for (PluginEntry& entry : _plugins)
    {
        Unload(entry);
    }

    _plugins.clear();
    _exeDir.clear();
    _initialized = false;
}

HRESULT ViewerPluginManager::Refresh(Common::Settings::Settings& settings) noexcept
{
    return Discover(settings);
}

const std::vector<ViewerPluginManager::PluginEntry>& ViewerPluginManager::GetPlugins() const noexcept
{
    return _plugins;
}

HRESULT ViewerPluginManager::CreateViewerInstance(std::wstring_view pluginId, Common::Settings::Settings& settings, wil::com_ptr<IViewer>& outViewer) noexcept
{
    outViewer.reset();

    PluginEntry* entry = FindPluginById(pluginId);
    if (! entry)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    if (entry->disabled || ! entry->loadable || ! entry->module || ! entry->createFactory || (! entry->factoryPluginId.empty() && ! entry->createFactoryEx))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    FactoryOptions options{};
    options.debugLevel = DEBUG_LEVEL_NONE;

#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
    const auto createFactory   = reinterpret_cast<CreateFactoryFunc>(entry->createFactory);
    const auto createFactoryEx = reinterpret_cast<CreateFactoryExFunc>(entry->createFactoryEx);
#pragma warning(pop)

    wil::com_ptr<IViewer> viewer;
    const HRESULT createHr = entry->factoryPluginId.empty()
                                 ? createFactory(__uuidof(IViewer), &options, GetHostServices(), viewer.put_void())
                                 : createFactoryEx(__uuidof(IViewer), &options, GetHostServices(), entry->factoryPluginId.c_str(), viewer.put_void());
    if (FAILED(createHr))
    {
        return createHr;
    }

    wil::com_ptr<IInformations> infos;
    const HRESULT qiHr = viewer->QueryInterface(__uuidof(IInformations), infos.put_void());
    if (SUCCEEDED(qiHr) && infos)
    {
        static_cast<void>(ApplyConfigurationFromSettings(*infos, entry->id, settings));
    }

    outViewer = std::move(viewer);
    return S_OK;
}

HRESULT ViewerPluginManager::DisablePlugin(std::wstring_view pluginId, Common::Settings::Settings& settings) noexcept
{
    PluginEntry* entry = FindPluginById(pluginId);
    if (! entry)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    if (entry->id.empty())
    {
        return E_INVALIDARG;
    }

    if (! entry->disabled)
    {
        entry->disabled = true;
        settings.plugins.disabledPluginIds.push_back(entry->id);
    }

    Unload(*entry);
    return S_OK;
}

HRESULT ViewerPluginManager::EnablePlugin(std::wstring_view pluginId, Common::Settings::Settings& settings) noexcept
{
    PluginEntry* entry = FindPluginById(pluginId);
    if (! entry)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    if (entry->id.empty())
    {
        return E_INVALIDARG;
    }

    if (entry->disabled)
    {
        entry->disabled = false;
        RemoveStringFromVector(settings.plugins.disabledPluginIds, entry->id);
    }

    return EnsureLoaded(*entry);
}

HRESULT ViewerPluginManager::AddCustomPluginPath(const std::filesystem::path& path, Common::Settings::Settings& settings) noexcept
{
    if (path.empty())
    {
        return E_INVALIDARG;
    }

    const auto exists =
        std::find(settings.plugins.customPluginPaths.begin(), settings.plugins.customPluginPaths.end(), path) != settings.plugins.customPluginPaths.end();
    if (exists)
    {
        return Refresh(settings);
    }

    if (! IsDllPath(path))
    {
        return E_INVALIDARG;
    }

    const HRESULT refreshHr = Refresh(settings);
    if (FAILED(refreshHr))
    {
        return refreshHr;
    }

    PluginEntry probe;
    probe.origin = PluginOrigin::Custom;
    probe.path   = path;

    const HRESULT probeHr = EnsureLoaded(probe);
    if (FAILED(probeHr))
    {
        return probeHr;
    }

    if (probe.id.empty())
    {
        return E_INVALIDARG;
    }

    if (FindPluginById(probe.id))
    {
        return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
    }

    settings.plugins.customPluginPaths.push_back(path);
    return Refresh(settings);
}

HRESULT ViewerPluginManager::GetConfigurationSchema(std::wstring_view pluginId, Common::Settings::Settings& settings, std::string& outSchemaJsonUtf8) noexcept
{
    PluginEntry* entry = FindPluginById(pluginId);
    if (! entry)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    const HRESULT hr = EnsureLoaded(*entry);
    if (FAILED(hr))
    {
        return hr;
    }

    if (! entry->module || ! entry->createFactory || (! entry->factoryPluginId.empty() && ! entry->createFactoryEx))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    FactoryOptions options{};
    options.debugLevel = DEBUG_LEVEL_NONE;

#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
    const auto createFactory   = reinterpret_cast<CreateFactoryFunc>(entry->createFactory);
    const auto createFactoryEx = reinterpret_cast<CreateFactoryExFunc>(entry->createFactoryEx);
#pragma warning(pop)

    wil::com_ptr<IViewer> viewer;
    const HRESULT createHr = entry->factoryPluginId.empty()
                                 ? createFactory(__uuidof(IViewer), &options, GetHostServices(), viewer.put_void())
                                 : createFactoryEx(__uuidof(IViewer), &options, GetHostServices(), entry->factoryPluginId.c_str(), viewer.put_void());
    if (FAILED(createHr))
    {
        return createHr;
    }

    wil::com_ptr<IInformations> infos;
    const HRESULT qiHr = viewer->QueryInterface(__uuidof(IInformations), infos.put_void());
    if (FAILED(qiHr) || ! infos)
    {
        return qiHr;
    }

    static_cast<void>(ApplyConfigurationFromSettings(*infos, entry->id, settings));

    const char* schema     = nullptr;
    const HRESULT schemaHr = infos->GetConfigurationSchema(&schema);
    if (FAILED(schemaHr))
    {
        return schemaHr;
    }

    outSchemaJsonUtf8 = SafeCoalesce(schema);
    return S_OK;
}

HRESULT ViewerPluginManager::GetConfiguration(std::wstring_view pluginId, Common::Settings::Settings& settings, std::string& outConfigurationJsonUtf8) noexcept
{
    PluginEntry* entry = FindPluginById(pluginId);
    if (! entry)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    const HRESULT hr = EnsureLoaded(*entry);
    if (FAILED(hr))
    {
        return hr;
    }

    if (! entry->module || ! entry->createFactory || (! entry->factoryPluginId.empty() && ! entry->createFactoryEx))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    FactoryOptions options{};
    options.debugLevel = DEBUG_LEVEL_NONE;

#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
    const auto createFactory   = reinterpret_cast<CreateFactoryFunc>(entry->createFactory);
    const auto createFactoryEx = reinterpret_cast<CreateFactoryExFunc>(entry->createFactoryEx);
#pragma warning(pop)

    wil::com_ptr<IViewer> viewer;
    const HRESULT createHr = entry->factoryPluginId.empty()
                                 ? createFactory(__uuidof(IViewer), &options, GetHostServices(), viewer.put_void())
                                 : createFactoryEx(__uuidof(IViewer), &options, GetHostServices(), entry->factoryPluginId.c_str(), viewer.put_void());
    if (FAILED(createHr))
    {
        return createHr;
    }

    wil::com_ptr<IInformations> infos;
    const HRESULT qiHr = viewer->QueryInterface(__uuidof(IInformations), infos.put_void());
    if (FAILED(qiHr) || ! infos)
    {
        return qiHr;
    }

    static_cast<void>(ApplyConfigurationFromSettings(*infos, entry->id, settings));

    const char* config  = nullptr;
    const HRESULT cfgHr = infos->GetConfiguration(&config);
    if (FAILED(cfgHr))
    {
        return cfgHr;
    }

    outConfigurationJsonUtf8 = SafeCoalesce(config);
    return S_OK;
}

HRESULT ViewerPluginManager::SetConfiguration(std::wstring_view pluginId, std::string_view configurationJsonUtf8, Common::Settings::Settings& settings) noexcept
{
    PluginEntry* entry = FindPluginById(pluginId);
    if (! entry)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    const HRESULT hr = EnsureLoaded(*entry);
    if (FAILED(hr))
    {
        return hr;
    }

    if (! entry->module || ! entry->createFactory || (! entry->factoryPluginId.empty() && ! entry->createFactoryEx))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    FactoryOptions options{};
    options.debugLevel = DEBUG_LEVEL_NONE;

#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
    const auto createFactory   = reinterpret_cast<CreateFactoryFunc>(entry->createFactory);
    const auto createFactoryEx = reinterpret_cast<CreateFactoryExFunc>(entry->createFactoryEx);
#pragma warning(pop)

    wil::com_ptr<IViewer> viewer;
    const HRESULT createHr = entry->factoryPluginId.empty()
                                 ? createFactory(__uuidof(IViewer), &options, GetHostServices(), viewer.put_void())
                                 : createFactoryEx(__uuidof(IViewer), &options, GetHostServices(), entry->factoryPluginId.c_str(), viewer.put_void());
    if (FAILED(createHr))
    {
        return createHr;
    }

    wil::com_ptr<IInformations> infos;
    const HRESULT qiHr = viewer->QueryInterface(__uuidof(IInformations), infos.put_void());
    if (FAILED(qiHr) || ! infos)
    {
        return qiHr;
    }

    const std::string configText(configurationJsonUtf8);
    const HRESULT setHr = infos->SetConfiguration(configText.empty() ? nullptr : configText.c_str());
    if (FAILED(setHr))
    {
        return setHr;
    }

    BOOL something            = FALSE;
    const HRESULT saveCheckHr = infos->SomethingToSave(&something);
    if (SUCCEEDED(saveCheckHr) && ! something)
    {
        settings.plugins.configurationByPluginId.erase(entry->id);
        return S_OK;
    }

    const char* persistedConfig     = nullptr;
    const HRESULT getHr             = infos->GetConfiguration(&persistedConfig);
    const std::string persistedText = SUCCEEDED(getHr) ? SafeCoalesce(persistedConfig) : configText;

    Common::Settings::JsonValue persistedValue;
    HRESULT parseHr = Common::Settings::ParseJsonValue(persistedText, persistedValue);
    if (FAILED(parseHr))
    {
        parseHr = Common::Settings::ParseJsonValue(configText, persistedValue);
        if (FAILED(parseHr))
        {
            Debug::Warning(L"Failed to parse viewer plugin configuration JSON for '{}' (hr=0x{:08X}); configuration will not be persisted.",
                           entry->id,
                           static_cast<unsigned long>(parseHr));
            return S_OK;
        }
    }

    settings.plugins.configurationByPluginId[entry->id] = std::move(persistedValue);
    return S_OK;
}

HRESULT ViewerPluginManager::TestPlugin(std::wstring_view pluginId) noexcept
{
    PluginEntry* entry = FindPluginById(pluginId);
    if (! entry)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    return EnsureLoaded(*entry);
}

std::filesystem::path ViewerPluginManager::GetExecutableDirectory() noexcept
{
    wil::unique_cotaskmem_string modulePath;
    const HRESULT hr = wil::GetModuleFileNameW<wil::unique_cotaskmem_string>(nullptr, modulePath);
    if (FAILED(hr) || ! modulePath)
    {
        return {};
    }
    return std::filesystem::path(modulePath.get()).parent_path();
}

std::filesystem::path ViewerPluginManager::GetOptionalPluginsDirectory() noexcept
{
    if (_exeDir.empty())
    {
        return {};
    }
    return _exeDir / L"Plugins";
}

std::optional<size_t> ViewerPluginManager::FindPluginIndexById(std::wstring_view pluginId) const noexcept
{
    if (pluginId.empty())
    {
        return std::nullopt;
    }

    for (size_t i = 0; i < _plugins.size(); ++i)
    {
        if (EqualsNoCase(_plugins[i].id, pluginId))
        {
            return i;
        }
    }

    return std::nullopt;
}

ViewerPluginManager::PluginEntry* ViewerPluginManager::FindPluginById(std::wstring_view pluginId) noexcept
{
    const std::optional<size_t> index = FindPluginIndexById(pluginId);
    if (! index.has_value())
    {
        return nullptr;
    }

    return &_plugins[index.value()];
}

const ViewerPluginManager::PluginEntry* ViewerPluginManager::FindPluginById(std::wstring_view pluginId) const noexcept
{
    const std::optional<size_t> index = FindPluginIndexById(pluginId);
    if (! index.has_value())
    {
        return nullptr;
    }

    return &_plugins[index.value()];
}

HRESULT ViewerPluginManager::Discover(Common::Settings::Settings& settings) noexcept
{
    _plugins.clear();

    if (_exeDir.empty())
    {
        _exeDir = GetExecutableDirectory();
    }

    if (_exeDir.empty())
    {
        return E_FAIL;
    }

    std::unordered_set<std::wstring> disabledIds;
    disabledIds.reserve(settings.plugins.disabledPluginIds.size());
    for (const auto& id : settings.plugins.disabledPluginIds)
    {
        if (! id.empty())
        {
            disabledIds.insert(ToLowerInvariant(id));
        }
    }

    struct Candidate
    {
        PluginOrigin origin = PluginOrigin::Embedded;
        std::filesystem::path path;
    };

    std::vector<Candidate> candidates;
    std::unordered_set<std::wstring> seenPaths;
    seenPaths.reserve(static_cast<size_t>(8) + settings.plugins.customPluginPaths.size());

    const auto tryAddCandidate = [&](PluginOrigin origin, const std::filesystem::path& path)
    {
        if (path.empty())
        {
            return;
        }

        const std::wstring key = ToLowerInvariant(path.wstring());
        if (! seenPaths.insert(key).second)
        {
            return;
        }

        candidates.push_back({origin, path});
    };

    const std::filesystem::path embeddedDir = _exeDir / L"Plugins";
    tryAddCandidate(PluginOrigin::Embedded, embeddedDir / L"ViewerText.dll");
    tryAddCandidate(PluginOrigin::Embedded, embeddedDir / L"ViewerSpace.dll");
    tryAddCandidate(PluginOrigin::Embedded, embeddedDir / L"ViewerImgRaw.dll");

    const std::filesystem::path optionalDir = GetOptionalPluginsDirectory();
    std::error_code ec;
    if (! optionalDir.empty() && std::filesystem::exists(optionalDir, ec))
    {
        for (const auto& item : std::filesystem::directory_iterator(optionalDir, ec))
        {
            if (ec)
            {
                break;
            }

            const std::filesystem::path p = item.path();
            if (IsDllPath(p))
            {
                tryAddCandidate(PluginOrigin::Optional, p);
            }
        }
    }

    for (const auto& p : settings.plugins.customPluginPaths)
    {
        tryAddCandidate(PluginOrigin::Custom, p);
    }

    std::unordered_set<std::wstring> seenIds;
    std::unordered_set<std::wstring> seenShortIds;

    const auto addLoadedEntry = [&](PluginEntry&& entry)
    {
        const std::wstring idKey    = ToLowerInvariant(entry.id);
        const std::wstring shortKey = ToLowerInvariant(entry.shortId);

        entry.disabled = ! entry.id.empty() && disabledIds.contains(idKey);

        bool conflict = false;

        if (entry.id.empty())
        {
            entry.loadError = L"Plugin id is missing.";
            conflict        = true;
        }
        else if (seenIds.contains(idKey))
        {
            entry.loadError = std::format(L"Duplicate plugin id '{}'.", entry.id);
            conflict        = true;
        }

        if (entry.shortId.empty())
        {
            entry.loadError = L"Short id is missing.";
            conflict        = true;
        }
        else if (seenShortIds.contains(shortKey))
        {
            entry.loadError = std::format(L"Duplicate short id '{}'.", entry.shortId);
            conflict        = true;
        }

        if (conflict)
        {
            entry.loadable = false;
            Unload(entry);
            return;
        }

        seenIds.insert(idKey);
        seenShortIds.insert(shortKey);

        if (entry.disabled)
        {
            Unload(entry);
        }

        _plugins.push_back(std::move(entry));
    };

    const auto tryLoadAndAddEntry = [&](PluginEntry&& entry)
    {
        const HRESULT loadHr = EnsureLoaded(entry);
        if (FAILED(loadHr))
        {
            if (loadHr == E_NOINTERFACE)
            {
                return;
            }

            if (loadHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND) && entry.loadError.find(L"RedSalamanderCreateEx") == std::wstring::npos)
            {
                // Not a viewer plugin (may be another RedSalamander plugin type).
                return;
            }

            return;
        }

        addLoadedEntry(std::move(entry));
    };

    for (const Candidate& candidate : candidates)
    {
        if (! IsDllPath(candidate.path))
        {
            continue;
        }

        bool handledAsMulti = false;
        bool isViewer       = true;

        wil::unique_hmodule probe(LoadLibraryExW(candidate.path.c_str(), nullptr, LOAD_WITH_ALTERED_SEARCH_PATH));
        if (probe)
        {
#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
            const auto enumerate = reinterpret_cast<EnumeratePluginsFunc>(GetProcAddress(probe.get(), "RedSalamanderEnumeratePlugins"));
#pragma warning(pop)
            if (enumerate)
            {
                const PluginMetaData* metaData = nullptr;
                unsigned int count             = 0;
                const HRESULT enumHr           = enumerate(__uuidof(IViewer), &metaData, &count);
                if (enumHr == E_NOINTERFACE)
                {
                    isViewer = false;
                }
                else if (SUCCEEDED(enumHr) && metaData != nullptr && count > 0)
                {
                    handledAsMulti = true;
                    for (unsigned int i = 0; i < count; ++i)
                    {
                        PluginEntry entry;
                        entry.origin          = candidate.origin;
                        entry.path            = candidate.path;
                        entry.factoryPluginId = SafeCoalesce(metaData[i].id);
                        tryLoadAndAddEntry(std::move(entry));
                    }
                }
            }
        }

        if (! isViewer)
        {
            continue;
        }

        if (handledAsMulti)
        {
            continue;
        }

        PluginEntry entry;
        entry.origin = candidate.origin;
        entry.path   = candidate.path;
        tryLoadAndAddEntry(std::move(entry));
    }

    const auto byOriginNameId = [](const PluginEntry& a, const PluginEntry& b)
    {
        if (a.origin != b.origin)
        {
            return static_cast<int>(a.origin) < static_cast<int>(b.origin);
        }

        const std::wstring an = a.name.empty() ? a.path.filename().wstring() : a.name;
        const std::wstring bn = b.name.empty() ? b.path.filename().wstring() : b.name;

        const int cmp = _wcsicmp(an.c_str(), bn.c_str());
        if (cmp != 0)
        {
            return cmp < 0;
        }

        return a.id < b.id;
    };

    std::sort(_plugins.begin(), _plugins.end(), byOriginNameId);
    return S_OK;
}

HRESULT ViewerPluginManager::EnsureLoaded(PluginEntry& entry) noexcept
{
    if (entry.module && entry.createFactory && entry.loadable && (entry.factoryPluginId.empty() || entry.createFactoryEx))
    {
        return S_OK;
    }

    entry.loadable = false;
    entry.loadError.clear();
    entry.module.reset();
    entry.createFactory   = nullptr;
    entry.createFactoryEx = nullptr;

    if (entry.path.empty())
    {
        entry.loadError = L"Plugin path is empty.";
        return E_INVALIDARG;
    }

    wil::unique_hmodule module(LoadLibraryExW(entry.path.c_str(), nullptr, LOAD_WITH_ALTERED_SEARCH_PATH));
    if (! module)
    {
        const DWORD lastError = GetLastError();
        entry.loadError       = std::format(L"LoadLibraryExW failed (0x{:08X}).", lastError);
        return HRESULT_FROM_WIN32(lastError);
    }

#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
    const auto createFactory   = reinterpret_cast<CreateFactoryFunc>(GetProcAddress(module.get(), "RedSalamanderCreate"));
    const auto createFactoryEx = reinterpret_cast<CreateFactoryExFunc>(GetProcAddress(module.get(), "RedSalamanderCreateEx"));
#pragma warning(pop)
    if (! createFactory)
    {
        DWORD lastError = GetLastError();
        if (lastError == ERROR_SUCCESS)
        {
            lastError = ERROR_PROC_NOT_FOUND;
        }
        entry.loadError = L"Missing RedSalamanderCreate export.";
        return HRESULT_FROM_WIN32(lastError);
    }

    FactoryOptions options{};
    options.debugLevel = DEBUG_LEVEL_NONE;

    wil::com_ptr<IViewer> viewer;
    HRESULT createHr = E_FAIL;
    if (! entry.factoryPluginId.empty())
    {
        if (! createFactoryEx)
        {
            entry.loadError = L"Missing export RedSalamanderCreateEx for multi-plugin DLL.";
            return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
        }
        createHr = createFactoryEx(__uuidof(IViewer), &options, GetHostServices(), entry.factoryPluginId.c_str(), viewer.put_void());
    }
    else
    {
        createHr = createFactory(__uuidof(IViewer), &options, GetHostServices(), viewer.put_void());
    }
    if (createHr == E_NOINTERFACE)
    {
        return E_NOINTERFACE;
    }
    if (FAILED(createHr))
    {
        entry.loadError = std::format(L"Factory failed (hr=0x{:08X}).", static_cast<unsigned long>(createHr));
        return createHr;
    }

    wil::com_ptr<IInformations> infos;
    const HRESULT qiHr = viewer->QueryInterface(__uuidof(IInformations), infos.put_void());
    if (FAILED(qiHr))
    {
        entry.loadError = std::format(L"IInformations not supported (hr=0x{:08X}).", static_cast<unsigned long>(qiHr));
        return qiHr;
    }

    const PluginMetaData* meta = nullptr;
    const HRESULT metaHr       = infos->GetMetaData(&meta);
    if (FAILED(metaHr))
    {
        entry.loadError = std::format(L"GetMetaData failed (hr=0x{:08X}).", static_cast<unsigned long>(metaHr));
        return metaHr;
    }

    if (meta)
    {
        entry.id          = SafeCoalesce(meta->id);
        entry.shortId     = SafeCoalesce(meta->shortId);
        entry.name        = SafeCoalesce(meta->name);
        entry.description = SafeCoalesce(meta->description);
        entry.author      = SafeCoalesce(meta->author);
        entry.version     = SafeCoalesce(meta->version);
    }

    if (! entry.factoryPluginId.empty() && ! entry.id.empty() && ! EqualsNoCase(entry.factoryPluginId, entry.id))
    {
        entry.loadError = std::format(L"Plugin id mismatch: requested '{}' but instance reported '{}'.", entry.factoryPluginId, entry.id);
        return E_FAIL;
    }

    if (entry.id.empty())
    {
        entry.loadError = L"Plugin id is missing.";
        return E_INVALIDARG;
    }

    if (! IsValidShortId(entry.shortId))
    {
        entry.loadError = std::format(L"Invalid or missing short id '{}'.", entry.shortId);
        return E_INVALIDARG;
    }

    entry.module = std::move(module);
#pragma warning(push)
#pragma warning(disable : 4191) // C4191: 'reinterpret_cast': unsafe conversion between function pointer and 'FARPROC'
    entry.createFactory   = reinterpret_cast<FARPROC>(createFactory);
    entry.createFactoryEx = reinterpret_cast<FARPROC>(createFactoryEx);
#pragma warning(pop)
    entry.loadable = true;
    return S_OK;
}

void ViewerPluginManager::Unload(PluginEntry& entry) noexcept
{
    entry.module.reset();
    entry.createFactory   = nullptr;
    entry.createFactoryEx = nullptr;
}
