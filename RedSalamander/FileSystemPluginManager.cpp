#include "Framework.h"

#include "FileSystemPluginManager.h"

#include <algorithm>
#include <cwchar>
#include <cwctype>
#include <format>
#include <unordered_set>

#include "Helpers.h"
#include "HostServices.h"
#include "PlugInterfaces/Factory.h"

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/win32_helpers.h>
#pragma warning(pop)

namespace
{
using CreateFactoryFunc    = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, void**);
using CreateFactoryExFunc  = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const wchar_t*, void**);
using EnumeratePluginsFunc = HRESULT(__stdcall*)(REFIID, const PluginMetaData**, unsigned int*);

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

void RemovePathFromVector(std::vector<std::filesystem::path>& values, const std::filesystem::path& needle)
{
    values.erase(std::remove_if(values.begin(), values.end(), [&](const std::filesystem::path& v) { return v == needle; }), values.end());
}
} // namespace

FileSystemPluginManager& FileSystemPluginManager::GetInstance() noexcept
{
    static FileSystemPluginManager instance;
    return instance;
}

HRESULT FileSystemPluginManager::Initialize(Common::Settings::Settings& settings) noexcept
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
        Debug::Error(L"Failed to discover file system plugins (hr=0x{:08X}).", static_cast<unsigned long>(hr));
        return hr;
    }

    _initialized = true;
    return S_OK;
}

void FileSystemPluginManager::Shutdown(Common::Settings::Settings& settings) noexcept
{
    if (! _initialized)
    {
        return;
    }

    for (const PluginEntry& entry : _plugins)
    {
        PersistConfigurationToSettings(entry, settings);
    }

    for (PluginEntry& entry : _plugins)
    {
        Unload(entry);
    }

    _plugins.clear();
    _activePluginId.clear();
    _initialized = false;
}

const std::vector<FileSystemPluginManager::PluginEntry>& FileSystemPluginManager::GetPlugins() const noexcept
{
    return _plugins;
}

std::wstring_view FileSystemPluginManager::GetActivePluginId() const noexcept
{
    return _activePluginId;
}

wil::com_ptr<IFileSystem> FileSystemPluginManager::GetActiveFileSystem() const noexcept
{
    const PluginEntry* entry = FindPluginById(_activePluginId);
    if (! entry)
    {
        return {};
    }

    return entry->fileSystem;
}

std::optional<std::filesystem::path> FileSystemPluginManager::TryGetPluginPath(std::wstring_view pluginId) const noexcept
{
    const PluginEntry* entry = FindPluginById(pluginId);
    if (! entry)
    {
        return std::nullopt;
    }

    if (entry->path.empty())
    {
        return std::nullopt;
    }

    return entry->path;
}

std::optional<std::filesystem::path> FileSystemPluginManager::TryGetActivePluginPath() const noexcept
{
    if (_activePluginId.empty())
    {
        return std::nullopt;
    }

    return TryGetPluginPath(_activePluginId);
}

HRESULT FileSystemPluginManager::Refresh(Common::Settings::Settings& settings) noexcept
{
    const HRESULT hr = Discover(settings);
    if (FAILED(hr))
    {
        return hr;
    }

    std::wstring_view wantedId = settings.plugins.currentFileSystemPluginId;

    const PluginEntry* wanted = FindPluginById(wantedId);
    if (! wanted || wanted->disabled || ! wanted->loadable)
    {
        wantedId = {};
    }

    if (wantedId.empty())
    {
        for (const PluginEntry& entry : _plugins)
        {
            if (entry.loadable && ! entry.disabled && ! entry.id.empty())
            {
                wantedId = entry.id;
                break;
            }
        }
    }

    if (! wantedId.empty())
    {
        return SetActivePlugin(wantedId, settings);
    }

    _activePluginId.clear();
    return S_OK;
}

HRESULT FileSystemPluginManager::SetActivePlugin(std::wstring_view pluginId, Common::Settings::Settings& settings) noexcept
{
    PluginEntry* entry = FindPluginById(pluginId);
    if (! entry)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    if (entry->disabled)
    {
        entry->disabled = false;
        RemoveStringFromVector(settings.plugins.disabledPluginIds, entry->id);
    }

    const HRESULT hr = EnsureLoaded(*entry, settings);
    if (FAILED(hr))
    {
        return hr;
    }

    _activePluginId                            = entry->id;
    settings.plugins.currentFileSystemPluginId = entry->id;
    return S_OK;
}

HRESULT FileSystemPluginManager::DisablePlugin(std::wstring_view pluginId, Common::Settings::Settings& settings) noexcept
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

    if (entry->id == _activePluginId)
    {
        std::wstring_view fallback;
        for (const PluginEntry& candidate : _plugins)
        {
            if (candidate.id.empty() || candidate.id == entry->id)
            {
                continue;
            }
            if (candidate.loadable && ! candidate.disabled)
            {
                fallback = candidate.id;
                break;
            }
        }

        if (fallback.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }

        const HRESULT switchHr = SetActivePlugin(fallback, settings);
        if (FAILED(switchHr))
        {
            return switchHr;
        }
    }

    if (! entry->disabled)
    {
        entry->disabled = true;
        settings.plugins.disabledPluginIds.push_back(entry->id);
    }

    Unload(*entry);
    return S_OK;
}

HRESULT FileSystemPluginManager::EnablePlugin(std::wstring_view pluginId, Common::Settings::Settings& settings) noexcept
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

    return EnsureLoaded(*entry, settings);
}

HRESULT FileSystemPluginManager::RemoveCustomPlugin(std::wstring_view pluginId, Common::Settings::Settings& settings) noexcept
{
    PluginEntry* entry = FindPluginById(pluginId);
    if (! entry)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    if (entry->origin != PluginOrigin::Custom)
    {
        return E_INVALIDARG;
    }

    const std::filesystem::path path = entry->path;

    if (entry->id == _activePluginId)
    {
        std::wstring_view fallback;
        for (const PluginEntry& candidate : _plugins)
        {
            if (candidate.origin == PluginOrigin::Custom && candidate.path == path)
            {
                continue;
            }
            if (candidate.id.empty())
            {
                continue;
            }
            if (candidate.loadable && ! candidate.disabled)
            {
                fallback = candidate.id;
                break;
            }
        }

        if (! fallback.empty())
        {
            const HRESULT switchHr = SetActivePlugin(fallback, settings);
            if (FAILED(switchHr))
            {
                return switchHr;
            }
        }
        else
        {
            _activePluginId.clear();
        }
    }

    RemovePathFromVector(settings.plugins.customPluginPaths, path);
    return Refresh(settings);
}

HRESULT FileSystemPluginManager::AddCustomPluginPath(const std::filesystem::path& path, Common::Settings::Settings& settings) noexcept
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

    Common::Settings::Settings scratch = settings;
    const HRESULT probeHr              = EnsureLoaded(probe, scratch);
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

HRESULT
FileSystemPluginManager::GetConfigurationSchema(std::wstring_view pluginId, Common::Settings::Settings& settings, std::string& outSchemaJsonUtf8) noexcept
{
    PluginEntry* entry = FindPluginById(pluginId);
    if (! entry)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    const HRESULT hr = EnsureLoaded(*entry, settings);
    if (FAILED(hr))
    {
        return hr;
    }

    if (! entry->informations)
    {
        return E_NOINTERFACE;
    }

    const char* schema     = nullptr;
    const HRESULT schemaHr = entry->informations->GetConfigurationSchema(&schema);
    if (FAILED(schemaHr))
    {
        return schemaHr;
    }

    outSchemaJsonUtf8 = SafeCoalesce(schema);
    return S_OK;
}

HRESULT
FileSystemPluginManager::GetConfiguration(std::wstring_view pluginId, Common::Settings::Settings& settings, std::string& outConfigurationJsonUtf8) noexcept
{
    PluginEntry* entry = FindPluginById(pluginId);
    if (! entry)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    const HRESULT hr = EnsureLoaded(*entry, settings);
    if (FAILED(hr))
    {
        return hr;
    }

    if (! entry->informations)
    {
        return E_NOINTERFACE;
    }

    const char* config  = nullptr;
    const HRESULT cfgHr = entry->informations->GetConfiguration(&config);
    if (FAILED(cfgHr))
    {
        return cfgHr;
    }

    outConfigurationJsonUtf8 = SafeCoalesce(config);
    return S_OK;
}

HRESULT
FileSystemPluginManager::SetConfiguration(std::wstring_view pluginId, std::string_view configurationJsonUtf8, Common::Settings::Settings& settings) noexcept
{
    PluginEntry* entry = FindPluginById(pluginId);
    if (! entry)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    const HRESULT hr = EnsureLoaded(*entry, settings);
    if (FAILED(hr))
    {
        return hr;
    }

    if (! entry->informations)
    {
        return E_NOINTERFACE;
    }

    const std::string configText(configurationJsonUtf8);
    const HRESULT setHr = entry->informations->SetConfiguration(configText.empty() ? nullptr : configText.c_str());
    if (FAILED(setHr))
    {
        return setHr;
    }

    BOOL something            = FALSE;
    const HRESULT saveCheckHr = entry->informations->SomethingToSave(&something);
    if (SUCCEEDED(saveCheckHr) && ! something)
    {
        settings.plugins.configurationByPluginId.erase(entry->id);
        return S_OK;
    }

    const char* persistedConfig = nullptr;
    const HRESULT getHr         = entry->informations->GetConfiguration(&persistedConfig);

    const std::string persistedText = SUCCEEDED(getHr) ? SafeCoalesce(persistedConfig) : configText;

    Common::Settings::JsonValue persistedValue;
    HRESULT parseHr = Common::Settings::ParseJsonValue(persistedText, persistedValue);
    if (FAILED(parseHr))
    {
        parseHr = Common::Settings::ParseJsonValue(configText, persistedValue);
        if (FAILED(parseHr))
        {
            Debug::Warning(L"Failed to parse plugin configuration JSON for '{}' (hr=0x{:08X}); configuration will not be persisted.",
                           entry->id,
                           static_cast<unsigned long>(parseHr));
            return S_OK;
        }
    }

    settings.plugins.configurationByPluginId[entry->id] = std::move(persistedValue);
    return S_OK;
}

HRESULT FileSystemPluginManager::TestPlugin(std::wstring_view pluginId) noexcept
{
    PluginEntry* entry = FindPluginById(pluginId);
    if (! entry)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    Common::Settings::Settings dummySettings;
    return EnsureLoaded(*entry, dummySettings);
}

std::filesystem::path FileSystemPluginManager::GetExecutableDirectory() noexcept
{
    wil::unique_cotaskmem_string modulePath;
    const HRESULT hr = wil::GetModuleFileNameW<wil::unique_cotaskmem_string>(nullptr, modulePath);
    if (FAILED(hr) || ! modulePath)
    {
        return {};
    }
    return std::filesystem::path(modulePath.get()).parent_path();
}

std::filesystem::path FileSystemPluginManager::GetOptionalPluginsDirectory() noexcept
{
    if (_exeDir.empty())
    {
        return {};
    }
    return _exeDir / L"Plugins";
}

std::optional<size_t> FileSystemPluginManager::FindPluginIndexById(std::wstring_view pluginId) const noexcept
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

FileSystemPluginManager::PluginEntry* FileSystemPluginManager::FindPluginById(std::wstring_view pluginId) noexcept
{
    const std::optional<size_t> index = FindPluginIndexById(pluginId);
    if (! index.has_value())
    {
        return nullptr;
    }

    return &_plugins[index.value()];
}

const FileSystemPluginManager::PluginEntry* FileSystemPluginManager::FindPluginById(std::wstring_view pluginId) const noexcept
{
    const std::optional<size_t> index = FindPluginIndexById(pluginId);
    if (! index.has_value())
    {
        return nullptr;
    }

    return &_plugins[index.value()];
}

HRESULT FileSystemPluginManager::Discover(Common::Settings::Settings& settings) noexcept
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

    tryAddCandidate(PluginOrigin::Embedded, _exeDir / L"Plugins" / L"FileSystem.dll");
    tryAddCandidate(PluginOrigin::Embedded, _exeDir / L"Plugins" / L"ViewerText.dll");

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
            Debug::Error(L"Plugin '{}' skipped: {}", entry.path.wstring(), entry.loadError);
            Unload(entry);
            _plugins.push_back(std::move(entry));
            return;
        }

        seenIds.insert(idKey);
        seenShortIds.insert(shortKey);

        if (entry.disabled && ! EqualsNoCase(entry.id, settings.plugins.currentFileSystemPluginId))
        {
            Unload(entry);
        }

        _plugins.push_back(std::move(entry));
    };

    const auto tryLoadAndAddEntry = [&](PluginEntry&& entry)
    {
        Common::Settings::Settings scratch = settings;
        const HRESULT loadHr               = EnsureLoaded(entry, scratch);
        if (FAILED(loadHr))
        {
            if (loadHr == E_NOINTERFACE)
            {
                // Not a file system plugin (may be another RedSalamander plugin type).
                return;
            }

            if (loadHr == HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND) && entry.loadError.find(L"RedSalamanderCreateEx") == std::wstring::npos)
            {
                Debug::Warning(L"Plugin '{}' skipped: missing RedSalamanderCreate export.", entry.path.wstring());
                return;
            }

            _plugins.push_back(std::move(entry));
            return;
        }

        addLoadedEntry(std::move(entry));
    };

    for (const Candidate& candidate : candidates)
    {
        if (! IsDllPath(candidate.path))
        {
            PluginEntry entry;
            entry.origin    = candidate.origin;
            entry.path      = candidate.path;
            entry.loadable  = false;
            entry.loadError = L"File is missing or not a DLL.";
            _plugins.push_back(std::move(entry));
            continue;
        }

        bool handledAsMulti = false;
        bool isFileSystem   = true;

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
                const HRESULT enumHr           = enumerate(__uuidof(IFileSystem), &metaData, &count);
                if (enumHr == E_NOINTERFACE)
                {
                    isFileSystem = false;
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

        if (! isFileSystem)
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

HRESULT FileSystemPluginManager::EnsureLoaded(PluginEntry& entry, Common::Settings::Settings& settings) noexcept
{
    if (entry.module && entry.fileSystem && entry.informations)
    {
        return S_OK;
    }

    entry.loadable = false;
    entry.loadError.clear();
    entry.module.reset();
    entry.fileSystem   = nullptr;
    entry.informations = nullptr;

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
        Debug::Error(L"Failed to load plugin '{}': {}", entry.path.wstring(), entry.loadError);
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
        entry.loadError = L"Missing export RedSalamanderCreate.";
        return HRESULT_FROM_WIN32(lastError);
    }

    FactoryOptions options{};
    options.debugLevel = DEBUG_LEVEL_NONE;

    wil::com_ptr<IFileSystem> fileSystem;
    HRESULT createHr = E_FAIL;
    if (! entry.factoryPluginId.empty())
    {
        if (! createFactoryEx)
        {
            entry.loadError = L"Missing export RedSalamanderCreateEx for multi-plugin DLL.";
            return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
        }
        createHr = createFactoryEx(__uuidof(IFileSystem), &options, GetHostServices(), entry.factoryPluginId.c_str(), fileSystem.put_void());
    }
    else
    {
        createHr = createFactory(__uuidof(IFileSystem), &options, GetHostServices(), fileSystem.put_void());
    }
    if (FAILED(createHr))
    {
        entry.loadError = std::format(L"Factory failed (hr=0x{:08X}).", static_cast<unsigned long>(createHr));
        return createHr;
    }

    wil::com_ptr<IInformations> infos;
    const HRESULT qiHr = fileSystem->QueryInterface(__uuidof(IInformations), infos.put_void());
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

    entry.module       = std::move(module);
    entry.fileSystem   = std::move(fileSystem);
    entry.informations = std::move(infos);
    entry.loadable     = true;

    static_cast<void>(ApplyConfigurationFromSettings(entry, settings));
    return S_OK;
}

void FileSystemPluginManager::Unload(PluginEntry& entry) noexcept
{
    entry.informations = nullptr;
    entry.fileSystem   = nullptr;
    entry.module.reset();
}

HRESULT FileSystemPluginManager::ApplyConfigurationFromSettings(PluginEntry& entry, const Common::Settings::Settings& settings) noexcept
{
    if (! entry.informations || entry.id.empty())
    {
        return S_FALSE;
    }

    const auto it = settings.plugins.configurationByPluginId.find(entry.id);
    if (it == settings.plugins.configurationByPluginId.end())
    {
        return entry.informations->SetConfiguration(nullptr);
    }

    const Common::Settings::JsonValue& configValue = it->second;
    if (std::holds_alternative<std::monostate>(configValue.value))
    {
        return entry.informations->SetConfiguration(nullptr);
    }

    std::string configText;
    const HRESULT serializeHr = Common::Settings::SerializeJsonValue(configValue, configText);
    if (FAILED(serializeHr))
    {
        Debug::Warning(L"Failed to serialize plugin configuration JSON for '{}' (hr=0x{:08X}); configuration will be ignored.",
                       entry.id,
                       static_cast<unsigned long>(serializeHr));
        return entry.informations->SetConfiguration(nullptr);
    }

    return entry.informations->SetConfiguration(configText.empty() ? nullptr : configText.c_str());
}

void FileSystemPluginManager::PersistConfigurationToSettings(const PluginEntry& entry, Common::Settings::Settings& settings) noexcept
{
    if (! entry.informations || entry.id.empty())
    {
        return;
    }

    BOOL something            = FALSE;
    const HRESULT saveCheckHr = entry.informations->SomethingToSave(&something);
    if (FAILED(saveCheckHr))
    {
        return;
    }

    if (! something)
    {
        settings.plugins.configurationByPluginId.erase(entry.id);
        return;
    }

    const char* config  = nullptr;
    const HRESULT getHr = entry.informations->GetConfiguration(&config);
    if (FAILED(getHr))
    {
        return;
    }

    Common::Settings::JsonValue persistedValue;
    const HRESULT parseHr = Common::Settings::ParseJsonValue(SafeCoalesce(config), persistedValue);
    if (FAILED(parseHr))
    {
        Debug::Warning(L"Failed to parse plugin configuration JSON for '{}' (hr=0x{:08X}); configuration will not be persisted.",
                       entry.id,
                       static_cast<unsigned long>(parseHr));
        return;
    }

    settings.plugins.configurationByPluginId[entry.id] = std::move(persistedValue);
}
