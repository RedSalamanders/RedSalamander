#include "DirectoryInfoCache.h"

#include <algorithm>
#include <format>
#include <limits>
#include <unordered_set>

#include "FolderWatcher.h"
#include "Helpers.h"
#include "NavigationLocation.h"
#include "PlugInterfaces/Informations.h"

namespace
{
constexpr uint64_t kMiB = 1024ull * 1024ull;
constexpr uint64_t kGiB = 1024ull * 1024ull * 1024ull;

constexpr uint64_t kMinDefaultCacheSize = 256ull * kMiB;
constexpr uint64_t kMaxDefaultCacheSize = 4ull * kGiB;

constexpr uint32_t kMaxWatchersHardCap = 1024u;
constexpr uint32_t kMruWatchedHardCap  = 256u;

std::wstring NormalizePath(std::wstring_view path, bool isFilePlugin) noexcept;

[[nodiscard]] bool EqualsNoCase(std::wstring_view left, std::wstring_view right) noexcept
{
    return CompareStringOrdinal(left.data(), static_cast<int>(left.size()), right.data(), static_cast<int>(right.size()), TRUE) == CSTR_EQUAL;
}

std::wstring MakeCaseInsensitivePathKey(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    std::wstring key(text);
    bool asciiOnly = true;
    for (wchar_t& ch : key)
    {
        if (ch >= 0x80)
        {
            asciiOnly = false;
            break;
        }

        if (ch >= L'A' && ch <= L'Z')
        {
            ch = static_cast<wchar_t>(ch | 0x20);
        }
    }

    if (asciiOnly)
    {
        return key;
    }

    if (text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return key;
    }

    const int srcLength   = static_cast<int>(text.size());
    const int requiredLen = LCMapStringEx(LOCALE_NAME_INVARIANT, LCMAP_LOWERCASE, text.data(), srcLength, nullptr, 0, nullptr, nullptr, 0);
    if (requiredLen <= 0)
    {
        Debug::ErrorWithLastError(L"DirectoryInfoCache: LCMapStringEx() failed to query lowercase size.");
        return key;
    }

    std::wstring mapped(static_cast<size_t>(requiredLen), L'\0');
    const int written = LCMapStringEx(LOCALE_NAME_INVARIANT, LCMAP_LOWERCASE, text.data(), srcLength, mapped.data(), requiredLen, nullptr, nullptr, 0);
    if (written <= 0)
    {
        Debug::ErrorWithLastError(L"DirectoryInfoCache: LCMapStringEx() failed to lowercase.");
        return key;
    }

    mapped.resize(static_cast<size_t>(written));
    if (! mapped.empty() && mapped.back() == L'\0')
    {
        mapped.pop_back();
    }

    return mapped;
}

[[nodiscard]] bool IsFilePlugin(std::wstring_view pluginId, std::wstring_view pluginShortId) noexcept
{
    return EqualsNoCase(pluginShortId, L"file") || EqualsNoCase(pluginId, L"builtin/file-system");
}

[[nodiscard]] std::wstring NormalizeInstanceContext(std::wstring_view instanceContext) noexcept
{
    std::wstring trimmed = StringUtils::TrimWhitespaceCopy(instanceContext);
    if (trimmed.empty())
    {
        return {};
    }

    if (NavigationLocation::LooksLikeWindowsAbsolutePath(trimmed))
    {
        return NormalizePath(trimmed, true);
    }

    if (trimmed.find_first_of(L"\\/") != std::wstring::npos)
    {
        return NavigationLocation::NormalizePluginPathText(trimmed,
                                                           NavigationLocation::EmptyPathPolicy::ReturnEmpty,
                                                           NavigationLocation::LeadingSlashPolicy::Preserve,
                                                           NavigationLocation::TrailingSlashPolicy::Trim);
    }

    return trimmed;
}

std::wstring NormalizePath(std::wstring_view path, bool isFilePlugin) noexcept
{
    if (path.empty())
    {
        return {};
    }

    if (! isFilePlugin)
    {
        return NavigationLocation::NormalizePluginPathText(path,
                                                           NavigationLocation::EmptyPathPolicy::ReturnEmpty,
                                                           NavigationLocation::LeadingSlashPolicy::Preserve,
                                                           NavigationLocation::TrailingSlashPolicy::Trim);
    }

    std::wstring normalized(path);
    std::replace(normalized.begin(), normalized.end(), L'/', L'\\');

    const bool isExtended = normalized.rfind(L"\\\\?\\", 0) == 0;
    if (! isExtended)
    {
        DWORD required = GetFullPathNameW(normalized.c_str(), 0, nullptr, nullptr);
        if (required > 0)
        {
            std::wstring absolute(static_cast<size_t>(required), L'\0');
            const DWORD written = GetFullPathNameW(normalized.c_str(), required, absolute.data(), nullptr);
            if (written > 0 && written < required)
            {
                absolute.resize(static_cast<size_t>(written));
                normalized = std::move(absolute);
            }
        }
    }

    while (normalized.size() > 3 && (normalized.back() == L'\\' || normalized.back() == L'/'))
    {
        normalized.pop_back();
    }

    return normalized;
}

uint64_t ClampCacheBytes(uint64_t value) noexcept
{
    if (value == 0)
    {
        return 0;
    }
    return std::clamp<uint64_t>(value, 8ull * kMiB, 64ull * kGiB);
}

uint32_t ClampWatchers(uint32_t value) noexcept
{
    return std::min(value, kMaxWatchersHardCap);
}

uint32_t ClampMruWatched(uint32_t value) noexcept
{
    return std::min(value, kMruWatchedHardCap);
}

[[nodiscard]] wchar_t GetPathSeparator(bool isFilePlugin) noexcept
{
    return isFilePlugin ? L'\\' : L'/';
}

[[nodiscard]] bool IsSameOrDescendantPath(std::wstring_view candidatePathKey, std::wstring_view rootPathKey, bool isFilePlugin) noexcept
{
    if (candidatePathKey == rootPathKey)
    {
        return true;
    }

    if (candidatePathKey.size() <= rootPathKey.size())
    {
        return false;
    }

    const wchar_t separator = GetPathSeparator(isFilePlugin);
    return candidatePathKey.rfind(rootPathKey, 0) == 0 && candidatePathKey[rootPathKey.size()] == separator;
}

[[nodiscard]] std::wstring GetRelativeSubPath(std::wstring_view fullPath, std::wstring_view rootPath, bool isFilePlugin) noexcept
{
    if (fullPath.size() <= rootPath.size())
    {
        return {};
    }

    size_t offset           = rootPath.size();
    const wchar_t separator = GetPathSeparator(isFilePlugin);
    if (offset < fullPath.size() && fullPath[offset] == separator)
    {
        ++offset;
    }

    return std::wstring(fullPath.substr(offset));
}

[[nodiscard]] std::wstring JoinNormalizedPath(std::wstring_view base, std::wstring_view relative, bool isFilePlugin) noexcept
{
    if (base.empty())
    {
        return NormalizePath(relative, isFilePlugin);
    }

    if (relative.empty())
    {
        return std::wstring(base);
    }

    std::wstring combined(base);
    const wchar_t separator = GetPathSeparator(isFilePlugin);
    if (combined.back() != separator)
    {
        combined.push_back(separator);
    }
    combined.append(relative);
    return NormalizePath(combined, isFilePlugin);
}

[[nodiscard]] std::wstring ParentNormalizedPath(std::wstring_view path, bool isFilePlugin) noexcept
{
    if (path.empty())
    {
        return {};
    }

    std::filesystem::path parent = std::filesystem::path(path).parent_path();
    if (parent.empty())
    {
        return std::wstring(path);
    }

    return NormalizePath(parent.native(), isFilePlugin);
}

[[nodiscard]] std::wstring LeafNameNormalizedPath(std::wstring_view path) noexcept
{
    return std::filesystem::path(path).filename().wstring();
}

[[nodiscard]] std::optional<DirectoryInfoCache::ContextKey> MakeFallbackContext(IFileSystem* fileSystem) noexcept
{
    if (! fileSystem)
    {
        return std::nullopt;
    }

    wil::com_ptr<IInformations> infos;
    const HRESULT qiHr = fileSystem->QueryInterface(__uuidof(IInformations), infos.put_void());
    if (FAILED(qiHr) || ! infos)
    {
        return std::nullopt;
    }

    const PluginMetaData* meta = nullptr;
    const HRESULT metaHr       = infos->GetMetaData(&meta);
    if (FAILED(metaHr) || ! meta)
    {
        return std::nullopt;
    }

    DirectoryInfoCache::ContextKey context{};
    context.pluginId        = meta->id ? meta->id : (meta->shortId ? meta->shortId : L"");
    context.pluginShortId   = meta->shortId ? meta->shortId : L"";
    context.instanceContext = {};
    context.pluginIdKey     = MakeCaseInsensitivePathKey(context.pluginId);
    context.instanceContextKey.clear();
    context.isFilePlugin = IsFilePlugin(context.pluginId, context.pluginShortId);
    return context;
}
} // namespace

DirectoryInfoCache& DirectoryInfoCache::GetInstance()
{
    // Intentionally leaked to avoid shutdown UAF from static destruction order issues.
    static DirectoryInfoCache* instance = new DirectoryInfoCache();
    return *instance;
}

void DirectoryInfoCache::Shutdown() noexcept
{
    _shuttingDown.store(true, std::memory_order_release);

    std::vector<std::unique_ptr<FolderWatcher>> watchersToStop;
    {
        std::lock_guard lock(_mutex);

        for (const auto& pair : _entries)
        {
            const std::shared_ptr<Entry>& entry = pair.second;
            StopWatcherLocked(entry, watchersToStop);
        }

        _entries.clear();
        _lru.clear();
        _contexts.clear();
        _providerContexts.clear();

        _currentBytes = 0;
        _cacheHits    = 0;
        _cacheMisses  = 0;
        _enumerations = 0;
        _evictions    = 0;
        _dirtyMarks   = 0;

        _initialized = false;
        _maxBytes    = 0;
        _maxWatchers = 0;
        _mruWatched  = 0;
    }

    Debug::Info(L"DirectoryInfoCache: Shutdown (stopped {} watcher(s))", watchersToStop.size());
}

size_t DirectoryInfoCache::ContextKeyHash::operator()(const ContextKey& key) const noexcept
{
    const size_t pluginHash  = std::hash<std::wstring_view>{}(key.pluginIdKey);
    const size_t contextHash = std::hash<std::wstring_view>{}(key.instanceContextKey);
    return pluginHash ^ (contextHash + 0x9e3779b97f4a7c15ull + (pluginHash << 6) + (pluginHash >> 2));
}

bool DirectoryInfoCache::ContextKeyEq::operator()(const ContextKey& a, const ContextKey& b) const noexcept
{
    return a.pluginIdKey == b.pluginIdKey && a.instanceContextKey == b.instanceContextKey;
}

size_t DirectoryInfoCache::KeyHash::operator()(const Key& key) const noexcept
{
    const size_t contextHash = ContextKeyHash{}(key.context);
    const size_t pathHash    = std::hash<std::wstring_view>{}(key.pathKey);
    return contextHash ^ (pathHash + 0x9e3779b97f4a7c15ull + (contextHash << 6) + (contextHash >> 2));
}

bool DirectoryInfoCache::KeyEq::operator()(const Key& a, const Key& b) const noexcept
{
    return ContextKeyEq{}(a.context, b.context) && a.pathKey == b.pathKey;
}

DirectoryInfoCache::Borrowed::Borrowed(Borrowed&& other) noexcept
{
    *this = std::move(other);
}

DirectoryInfoCache::Borrowed& DirectoryInfoCache::Borrowed::operator=(Borrowed&& other) noexcept
{
    if (this == &other)
    {
        return *this;
    }

    if (_owner && _entry)
    {
        std::vector<std::unique_ptr<FolderWatcher>> watchersToStop;
        {
            std::lock_guard lock(_owner->_mutex);
            _owner->ReleaseBorrowLocked(_entry);
            _owner->MaybeEvictLocked(watchersToStop);
        }
    }

    _owner  = other._owner;
    _entry  = std::move(other._entry);
    _status = other._status;

    other._owner  = nullptr;
    other._status = E_FAIL;
    return *this;
}

DirectoryInfoCache::Borrowed::~Borrowed()
{
    if (! _owner || ! _entry)
    {
        return;
    }

    std::vector<std::unique_ptr<FolderWatcher>> watchersToStop;
    {
        std::lock_guard lock(_owner->_mutex);
        _owner->ReleaseBorrowLocked(_entry);
        _owner->MaybeEvictLocked(watchersToStop);
    }
}

HRESULT DirectoryInfoCache::Borrowed::Status() const noexcept
{
    return _status;
}

IFilesInformation* DirectoryInfoCache::Borrowed::Get() const noexcept
{
    if (FAILED(_status) || ! _entry)
    {
        return nullptr;
    }
    return _entry->info.get();
}

const std::wstring& DirectoryInfoCache::Borrowed::NormalizedPath() const noexcept
{
    static const std::wstring empty;
    if (! _entry)
    {
        return empty;
    }
    return _entry->key.path;
}

DirectoryInfoCache::Pin::Pin(Pin&& other) noexcept
{
    *this = std::move(other);
}

DirectoryInfoCache::Pin& DirectoryInfoCache::Pin::operator=(Pin&& other) noexcept
{
    if (this == &other)
    {
        return *this;
    }

    if (_owner && _entry)
    {
        std::vector<std::unique_ptr<FolderWatcher>> watchersToStop;
        {
            std::lock_guard lock(_owner->_mutex);
            _owner->RemoveSubscriberLocked(_entry, _hwnd, _message);
            _owner->ReleasePinLocked(_entry);
            _owner->MaybeEvictLocked(watchersToStop);
            _owner->UpdateWatchersLocked(watchersToStop);
        }
    }

    _owner   = other._owner;
    _entry   = std::move(other._entry);
    _hwnd    = other._hwnd;
    _message = other._message;

    other._owner   = nullptr;
    other._hwnd    = nullptr;
    other._message = 0;
    return *this;
}

DirectoryInfoCache::Pin::~Pin()
{
    if (! _owner || ! _entry)
    {
        return;
    }

    std::vector<std::unique_ptr<FolderWatcher>> watchersToStop;
    {
        std::lock_guard lock(_owner->_mutex);
        _owner->RemoveSubscriberLocked(_entry, _hwnd, _message);
        _owner->ReleasePinLocked(_entry);
        _owner->MaybeEvictLocked(watchersToStop);
        _owner->UpdateWatchersLocked(watchersToStop);
    }
}

bool DirectoryInfoCache::Pin::IsValid() const noexcept
{
    return _entry != nullptr;
}

const std::wstring& DirectoryInfoCache::Pin::NormalizedPath() const noexcept
{
    static const std::wstring empty;
    if (! _entry)
    {
        return empty;
    }
    return _entry->key.path;
}

uint64_t DirectoryInfoCache::ComputeDefaultMaxBytes() noexcept
{
    MEMORYSTATUSEX mem{};
    mem.dwLength = sizeof(mem);
    if (! GlobalMemoryStatusEx(&mem))
    {
        return 512ull * kMiB;
    }

    const uint64_t totalPhys = mem.ullTotalPhys;
    uint64_t guess           = totalPhys / 16ull; // ~6.25% of RAM
    guess                    = std::clamp<uint64_t>(guess, kMinDefaultCacheSize, kMaxDefaultCacheSize);
    return guess;
}

void DirectoryInfoCache::ApplySettings(const Common::Settings::Settings& settings) noexcept
{
    uint64_t maxBytes    = 0;
    uint32_t maxWatchers = _maxWatchers;
    uint32_t mruWatched  = _mruWatched;

    if (settings.cache && settings.cache->directoryInfo.maxBytes && *settings.cache->directoryInfo.maxBytes > 0)
    {
        maxBytes = *settings.cache->directoryInfo.maxBytes;
    }

    if (settings.cache && settings.cache->directoryInfo.maxWatchers)
    {
        maxWatchers = *settings.cache->directoryInfo.maxWatchers;
    }

    if (settings.cache && settings.cache->directoryInfo.mruWatched)
    {
        mruWatched = *settings.cache->directoryInfo.mruWatched;
    }

    if (maxBytes == 0)
    {
        maxBytes = ComputeDefaultMaxBytes();
    }

    SetLimits(maxBytes, maxWatchers, mruWatched);
}

void DirectoryInfoCache::SetLimits(uint64_t maxBytes, uint32_t maxWatchers, uint32_t mruWatched) noexcept
{
    std::vector<std::unique_ptr<FolderWatcher>> watchersToStop;
    uint64_t maxBytesLocal    = 0;
    uint32_t maxWatchersLocal = 0;
    uint32_t mruWatchedLocal  = 0;

    {
        std::lock_guard lock(_mutex);

        _maxBytes    = ClampCacheBytes(maxBytes);
        _maxWatchers = ClampWatchers(maxWatchers);
        _mruWatched  = ClampMruWatched(mruWatched);
        _initialized = true;

        MaybeEvictLocked(watchersToStop);
        UpdateWatchersLocked(watchersToStop);

        maxBytesLocal    = _maxBytes;
        maxWatchersLocal = _maxWatchers;
        mruWatchedLocal  = _mruWatched;
    }

    Debug::Info(L"DirectoryInfoCache: configured maxBytes={} MiB, maxWatchers={}, mruWatched={}", maxBytesLocal / kMiB, maxWatchersLocal, mruWatchedLocal);
}

DirectoryInfoCache::Stats DirectoryInfoCache::GetStats() const noexcept
{
    std::lock_guard lock(_mutex);

    uint32_t activeWatchers = 0;
    uint32_t pinnedEntries  = 0;

    for (const auto& entry : _lru)
    {
        if (entry->watcher)
        {
            ++activeWatchers;
        }
        if (entry->pinCount > 0)
        {
            ++pinnedEntries;
        }
    }

    Stats stats{};
    stats.maxBytes       = _maxBytes;
    stats.currentBytes   = _currentBytes;
    stats.cacheHits      = _cacheHits;
    stats.cacheMisses    = _cacheMisses;
    stats.enumerations   = _enumerations;
    stats.evictions      = _evictions;
    stats.dirtyMarks     = _dirtyMarks;
    stats.maxWatchers    = _maxWatchers;
    stats.mruWatched     = _mruWatched;
    stats.activeWatchers = activeWatchers;
    stats.pinnedEntries  = pinnedEntries;
    stats.entryCount     = static_cast<uint32_t>(_entries.size());
    return stats;
}

std::optional<DirectoryInfoCache::ContextKey> DirectoryInfoCache::ResolveContextForFileSystem(IFileSystem* fileSystem) const noexcept
{
    if (! fileSystem)
    {
        return std::nullopt;
    }

    {
        std::lock_guard lock(_mutex);
        const auto it = _providerContexts.find(fileSystem);
        if (it != _providerContexts.end())
        {
            return it->second;
        }
    }

    return MakeFallbackContext(fileSystem);
}

void DirectoryInfoCache::RegisterProvider(IFileSystem* fileSystem,
                                          std::wstring_view pluginId,
                                          std::wstring_view pluginShortId,
                                          std::wstring_view instanceContext) noexcept
{
    if (! fileSystem)
    {
        return;
    }

    ContextKey context{};
    context.pluginId           = pluginId.empty() ? std::wstring(pluginShortId) : std::wstring(pluginId);
    context.pluginShortId      = std::wstring(pluginShortId);
    context.instanceContext    = NormalizeInstanceContext(instanceContext);
    context.pluginIdKey        = MakeCaseInsensitivePathKey(context.pluginId);
    context.instanceContextKey = MakeCaseInsensitivePathKey(context.instanceContext);
    context.isFilePlugin       = IsFilePlugin(context.pluginId, context.pluginShortId);

    Debug::Info(L"DirectoryInfoCache: RegisterProvider fs={} pluginId='{}' pluginShortId='{}' instanceContext='{}' isFilePlugin={}",
                static_cast<const void*>(fileSystem),
                context.pluginId,
                context.pluginShortId,
                context.instanceContext.empty() ? std::wstring_view(L"<none>") : std::wstring_view(context.instanceContext),
                context.isFilePlugin ? L"true" : L"false");

    std::vector<std::unique_ptr<FolderWatcher>> watchersToStop;
    {
        std::lock_guard lock(_mutex);

        const auto current = _providerContexts.find(fileSystem);
        if (current != _providerContexts.end() && ContextKeyEq{}(current->second, context))
        {
            auto ctxIt = _contexts.find(context);
            if (ctxIt != _contexts.end())
            {
                ctxIt->second.providers[fileSystem] = fileSystem;
            }
            else
            {
                ContextState state{};
                state.key                   = context;
                state.providers[fileSystem] = fileSystem;
                _contexts.emplace(state.key, std::move(state));
            }
            return;
        }

        if (current != _providerContexts.end())
        {
            auto ctxIt = _contexts.find(current->second);
            if (ctxIt != _contexts.end())
            {
                ctxIt->second.providers.erase(fileSystem);
                if (ctxIt->second.providers.empty())
                {
                    ClearContextLocked(ctxIt->first, watchersToStop);
                }
                else
                {
                    RestartContextWatchersLocked(ctxIt->first, watchersToStop);
                }
            }
            _providerContexts.erase(current);
        }

        ContextState& state           = _contexts.try_emplace(context).first->second;
        state.key                     = context;
        state.providers[fileSystem]   = fileSystem;
        _providerContexts[fileSystem] = context;

        UpdateWatchersLocked(watchersToStop);
    }

    for (auto& watcher : watchersToStop)
    {
        if (watcher)
        {
            watcher->Stop();
        }
    }
}

void DirectoryInfoCache::UnregisterProvider(IFileSystem* fileSystem) noexcept
{
    if (! fileSystem)
    {
        return;
    }

    std::vector<std::unique_ptr<FolderWatcher>> watchersToStop;
    {
        std::lock_guard lock(_mutex);
        const auto providerIt = _providerContexts.find(fileSystem);
        if (providerIt == _providerContexts.end())
        {
            return;
        }

        const ContextKey context = providerIt->second;
        _providerContexts.erase(providerIt);

        auto ctxIt = _contexts.find(context);
        if (ctxIt != _contexts.end())
        {
            ctxIt->second.providers.erase(fileSystem);
            if (ctxIt->second.providers.empty())
            {
                ClearContextLocked(context, watchersToStop);
            }
            else
            {
                RestartContextWatchersLocked(context, watchersToStop);
            }
        }

        UpdateWatchersLocked(watchersToStop);
    }

    for (auto& watcher : watchersToStop)
    {
        if (watcher)
        {
            watcher->Stop();
        }
    }
}

void DirectoryInfoCache::ClearForFileSystem(IFileSystem* fileSystem) noexcept
{
    if (! fileSystem)
    {
        return;
    }

    const std::optional<ContextKey> contextOpt = ResolveContextForFileSystem(fileSystem);
    if (! contextOpt)
    {
        return;
    }

    std::vector<std::unique_ptr<FolderWatcher>> watchersToStop;
    {
        std::lock_guard lock(_mutex);
        ClearContextLocked(*contextOpt, watchersToStop);
        UpdateWatchersLocked(watchersToStop);
    }

    for (auto& watcher : watchersToStop)
    {
        if (watcher)
        {
            watcher->Stop();
        }
    }
}

void DirectoryInfoCache::InvalidateFolder(IFileSystem* fileSystem, const std::filesystem::path& folder) noexcept
{
    const auto keyOpt = MakeKey(fileSystem, folder);
    if (! keyOpt)
    {
        return;
    }

    std::lock_guard lock(_mutex);
    MarkDirtyLocked(*keyOpt);
}

void DirectoryInfoCache::NotifyFolderContentsChanged(IFileSystem* fileSystem, const std::filesystem::path& folder) noexcept
{
    const auto keyOpt = MakeKey(fileSystem, folder);
    if (! keyOpt)
    {
        return;
    }

    std::lock_guard lock(_mutex);
    NotifyFolderContentsChangedLocked(keyOpt->context, keyOpt->path);
}

void DirectoryInfoCache::NotifyPathCreated(IFileSystem* fileSystem, const std::filesystem::path& path) noexcept
{
    const auto keyOpt = MakeKey(fileSystem, path);
    if (! keyOpt)
    {
        return;
    }

    std::lock_guard lock(_mutex);
    NotifyFolderContentsChangedLocked(keyOpt->context, ParentNormalizedPath(keyOpt->path, keyOpt->context.isFilePlugin));
}

void DirectoryInfoCache::NotifyPathDeleted(IFileSystem* fileSystem, const std::filesystem::path& path) noexcept
{
    const auto keyOpt = MakeKey(fileSystem, path);
    if (! keyOpt)
    {
        return;
    }

    std::lock_guard lock(_mutex);
    NotifyPathDeletedLocked(keyOpt->context, keyOpt->path);
}

void DirectoryInfoCache::NotifyPathMoved(IFileSystem* fileSystem,
                                         const std::filesystem::path& sourcePath,
                                         const std::filesystem::path& destinationPath) noexcept
{
    const auto sourceKeyOpt = MakeKey(fileSystem, sourcePath);
    if (! sourceKeyOpt)
    {
        return;
    }

    std::wstring normalizedDestination = NormalizePath(destinationPath.native(), sourceKeyOpt->context.isFilePlugin);
    if (normalizedDestination.empty())
    {
        return;
    }

    std::lock_guard lock(_mutex);
    NotifyPathMovedLocked(sourceKeyOpt->context, sourceKeyOpt->path, normalizedDestination);
}

bool DirectoryInfoCache::IsFolderWatched(IFileSystem* fileSystem, const std::filesystem::path& folder) const noexcept
{
    const auto keyOpt = MakeKey(fileSystem, folder);
    if (! keyOpt)
    {
        return false;
    }

    std::lock_guard lock(_mutex);
    const auto it = _entries.find(*keyOpt);
    if (it == _entries.end())
    {
        return false;
    }

    const std::shared_ptr<Entry>& entry = it->second;
    return entry && entry->watcher;
}

std::optional<DirectoryInfoCache::Key> DirectoryInfoCache::MakeKey(IFileSystem* fileSystem, const std::filesystem::path& folder) const noexcept
{
    if (! fileSystem)
    {
        return std::nullopt;
    }

    const auto contextOpt = ResolveContextForFileSystem(fileSystem);
    if (! contextOpt)
    {
        return std::nullopt;
    }

    std::wstring normalized = NormalizePath(folder.native(), contextOpt->isFilePlugin);
    if (normalized.empty())
    {
        return std::nullopt;
    }

    Debug::Info(L"DirectoryInfoCache: MakeKey fs={} folder='{}' normalized='{}' pluginId='{}' pluginShortId='{}' instanceContext='{}' isFilePlugin={}",
                static_cast<const void*>(fileSystem),
                folder.native(),
                normalized,
                contextOpt->pluginId,
                contextOpt->pluginShortId,
                contextOpt->instanceContext.empty() ? std::wstring_view(L"<none>") : std::wstring_view(contextOpt->instanceContext),
                contextOpt->isFilePlugin ? L"true" : L"false");

    Key key{};
    key.context = *contextOpt;
    key.path    = std::move(normalized);
    key.pathKey = MakeCaseInsensitivePathKey(key.path);
    return key;
}

std::shared_ptr<DirectoryInfoCache::Entry> DirectoryInfoCache::GetOrCreateEntryLocked(const Key& key) noexcept
{
    auto it = _entries.find(key);
    if (it != _entries.end())
    {
        return it->second;
    }

    auto entry = std::make_shared<Entry>();
    entry->key = key;
    _entries.emplace(entry->key, entry);
    _lru.push_front(entry);
    entry->lruIt      = _lru.begin();
    entry->lruItValid = true;
    return entry;
}

void DirectoryInfoCache::TouchLocked(const std::shared_ptr<Entry>& entry) noexcept
{
    if (! entry)
    {
        return;
    }

    if (! entry->lruItValid)
    {
        _lru.push_front(entry);
        entry->lruIt      = _lru.begin();
        entry->lruItValid = true;
        return;
    }

    if (entry->lruIt != _lru.begin())
    {
        _lru.splice(_lru.begin(), _lru, entry->lruIt);
        entry->lruIt = _lru.begin();
    }
}

void DirectoryInfoCache::AddSubscriberLocked(const std::shared_ptr<Entry>& entry, HWND hwnd, UINT message) noexcept
{
    if (! entry || ! hwnd || message == 0)
    {
        return;
    }

    for (const auto& s : entry->subscribers)
    {
        if (s.hwnd == hwnd && s.message == message)
        {
            return;
        }
    }

    entry->subscribers.push_back({hwnd, message});
}

void DirectoryInfoCache::RemoveSubscriberLocked(const std::shared_ptr<Entry>& entry, HWND hwnd, UINT message) noexcept
{
    if (! entry)
    {
        return;
    }

    std::erase_if(entry->subscribers, [&](const Subscriber& s) { return s.hwnd == hwnd && s.message == message; });
}

void DirectoryInfoCache::AddBorrowLocked(const std::shared_ptr<Entry>& entry) noexcept
{
    if (! entry)
    {
        return;
    }
    ++entry->borrowCount;
}

void DirectoryInfoCache::ReleaseBorrowLocked(const std::shared_ptr<Entry>& entry) noexcept
{
    if (! entry)
    {
        return;
    }
    if (entry->borrowCount > 0)
    {
        --entry->borrowCount;
    }
}

void DirectoryInfoCache::AddPinLocked(const std::shared_ptr<Entry>& entry) noexcept
{
    if (! entry)
    {
        return;
    }
    ++entry->pinCount;
}

void DirectoryInfoCache::ReleasePinLocked(const std::shared_ptr<Entry>& entry) noexcept
{
    if (! entry)
    {
        return;
    }
    if (entry->pinCount > 0)
    {
        --entry->pinCount;
    }
}

void DirectoryInfoCache::PostImpactLocked(const std::shared_ptr<Entry>& entry, const DirectoryImpact& impact) noexcept
{
    if (! entry || entry->subscribers.empty())
    {
        return;
    }

    const bool hasRefreshRenameHint = ! impact.renamedFromDisplayName.empty() && ! impact.renamedToDisplayName.empty();
    if (impact.kind == DirectoryImpact::Kind::RefreshCurrentFolder && entry->refreshPosted && ! hasRefreshRenameHint)
    {
        return;
    }

    if (impact.kind == DirectoryImpact::Kind::RefreshCurrentFolder)
    {
        entry->refreshPosted = true;
    }

    for (const auto& s : entry->subscribers)
    {
        if (s.hwnd && s.message)
        {
            auto payload           = std::make_unique<DirectoryImpact>(impact);
            payload->currentFolder = entry->key.path;
            const bool posted      = PostMessagePayload(s.hwnd, s.message, 0, std::move(payload));
            if (! posted && impact.kind == DirectoryImpact::Kind::RefreshCurrentFolder)
            {
                entry->refreshPosted = false;
            }
        }
    }
}

void DirectoryInfoCache::PostRefreshLocked(const std::shared_ptr<Entry>& entry) noexcept
{
    if (! entry)
    {
        return;
    }

    DirectoryImpact impact{};
    impact.kind = DirectoryImpact::Kind::RefreshCurrentFolder;
    PostImpactLocked(entry, impact);
}

void DirectoryInfoCache::MarkDirtyLocked(const std::shared_ptr<Entry>& entry) noexcept
{
    if (! entry || entry->dirty)
    {
        return;
    }

    entry->dirty = true;
    ++_dirtyMarks;
}

void DirectoryInfoCache::MarkDirtyLocked(const Key& key) noexcept
{
    const auto it = _entries.find(key);
    if (it == _entries.end() || ! it->second)
    {
        return;
    }

    MarkDirtyLocked(it->second);
    PostRefreshLocked(it->second);
}

void DirectoryInfoCache::StartWatcherLocked(const std::shared_ptr<Entry>& entry, std::vector<std::unique_ptr<FolderWatcher>>& watchersToStop) noexcept
{
    if (! entry || entry->watcher)
    {
        return;
    }

    const ContextKey context = entry->key.context;
    const std::wstring path  = entry->key.path;

    auto onNotification = [context, path](const FolderWatcherNotification& notification)
    {
        DirectoryInfoCache& cache = DirectoryInfoCache::GetInstance();
        cache.OnWatcherNotification(context, path, notification);
    };

    wil::com_ptr<IFileSystem> provider = ResolveProviderLocked(entry->key.context);
    if (! provider)
    {
        return;
    }

    wil::com_ptr<IFileSystemDirectoryWatch> dirWatch;
    const HRESULT qiHr = provider->QueryInterface(__uuidof(IFileSystemDirectoryWatch), dirWatch.put_void());
    if (FAILED(qiHr) || ! dirWatch)
    {
        Debug::Info(L"DirectoryInfoCache: watcher unsupported for path='{}' pluginId='{}' pluginShortId='{}' provider={} qiHr=0x{:08X}",
                    path,
                    context.pluginId,
                    context.pluginShortId,
                    static_cast<const void*>(provider.get()),
                    static_cast<unsigned long>(qiHr));
        return;
    }

    Debug::Info(L"DirectoryInfoCache: starting watcher for path='{}' pluginId='{}' pluginShortId='{}' provider={}",
                path,
                context.pluginId,
                context.pluginShortId,
                static_cast<const void*>(provider.get()));
    entry->watcher = std::make_unique<FolderWatcher>(std::move(dirWatch), path, std::move(onNotification));

    const HRESULT hr = entry->watcher->Start();
    if (FAILED(hr))
    {
        Debug::Warning(L"DirectoryInfoCache: Failed to start watcher for '{}' (hr=0x{:08X})", path, static_cast<unsigned long>(hr));
        watchersToStop.emplace_back(std::move(entry->watcher));
    }
}

void DirectoryInfoCache::StopWatcherLocked(const std::shared_ptr<Entry>& entry, std::vector<std::unique_ptr<FolderWatcher>>& watchersToStop) noexcept
{
    if (! entry || ! entry->watcher)
    {
        return;
    }

    watchersToStop.emplace_back(std::move(entry->watcher));
}

void DirectoryInfoCache::RestartContextWatchersLocked(const ContextKey& context, std::vector<std::unique_ptr<FolderWatcher>>& watchersToStop) noexcept
{
    for (const auto& [key, entry] : _entries)
    {
        if (! entry || ! ContextKeyEq{}(key.context, context))
        {
            continue;
        }

        StopWatcherLocked(entry, watchersToStop);
    }
}

void DirectoryInfoCache::ClearContextLocked(const ContextKey& context, std::vector<std::unique_ptr<FolderWatcher>>& watchersToStop) noexcept
{
    auto ctxIt = _contexts.find(context);
    if (ctxIt != _contexts.end())
    {
        for (const auto& [providerPtr, _provider] : ctxIt->second.providers)
        {
            _providerContexts.erase(providerPtr);
        }
    }

    for (auto it = _entries.begin(); it != _entries.end();)
    {
        if (! ContextKeyEq{}(it->first.context, context))
        {
            ++it;
            continue;
        }

        const std::shared_ptr<Entry>& entry = it->second;
        if (entry)
        {
            StopWatcherLocked(entry, watchersToStop);
            const uint64_t bytesFreed = entry->bytes;
            _currentBytes             = (_currentBytes >= bytesFreed) ? (_currentBytes - bytesFreed) : 0;
            if (entry->lruItValid)
            {
                _lru.erase(entry->lruIt);
                entry->lruItValid = false;
            }
        }

        it = _entries.erase(it);
    }

    if (ctxIt != _contexts.end())
    {
        _contexts.erase(ctxIt);
    }
}

wil::com_ptr<IFileSystem> DirectoryInfoCache::ResolveProviderLocked(const ContextKey& context, IFileSystem* preferredFileSystem) const noexcept
{
    if (preferredFileSystem)
    {
        const auto providerIt = _providerContexts.find(preferredFileSystem);
        if (providerIt != _providerContexts.end() && ContextKeyEq{}(providerIt->second, context))
        {
            return preferredFileSystem;
        }
    }

    const auto ctxIt = _contexts.find(context);
    if (ctxIt == _contexts.end() || ctxIt->second.providers.empty())
    {
        return nullptr;
    }

    return ctxIt->second.providers.begin()->second;
}

void DirectoryInfoCache::NotifyFolderContentsChangedLocked(const ContextKey& context,
                                                           const std::wstring& normalizedFolder,
                                                           std::wstring_view renamedFromDisplayName,
                                                           std::wstring_view renamedToDisplayName) noexcept
{
    if (normalizedFolder.empty())
    {
        return;
    }

    Key key{};
    key.context = context;
    key.path    = normalizedFolder;
    key.pathKey = MakeCaseInsensitivePathKey(normalizedFolder);

    const auto it = _entries.find(key);
    if (it == _entries.end() || ! it->second)
    {
        return;
    }

    MarkDirtyLocked(it->second);
    DirectoryImpact impact{};
    impact.kind = DirectoryImpact::Kind::RefreshCurrentFolder;
    impact.renamedFromDisplayName.assign(renamedFromDisplayName);
    impact.renamedToDisplayName.assign(renamedToDisplayName);
    PostImpactLocked(it->second, impact);
}

void DirectoryInfoCache::MarkSubtreeDirtyLocked(const ContextKey& context, const std::wstring& normalizedRootPath) noexcept
{
    if (normalizedRootPath.empty())
    {
        return;
    }

    const std::wstring rootKey = MakeCaseInsensitivePathKey(normalizedRootPath);
    for (const auto& [key, entry] : _entries)
    {
        if (! entry || ! ContextKeyEq{}(key.context, context))
        {
            continue;
        }

        if (! IsSameOrDescendantPath(key.pathKey, rootKey, context.isFilePlugin))
        {
            continue;
        }

        MarkDirtyLocked(entry);
    }
}

void DirectoryInfoCache::NotifyBackedContextPathDeletedLocked(const std::wstring& normalizedPath) noexcept
{
    const std::wstring normalizedPathKey = MakeCaseInsensitivePathKey(normalizedPath);

    for (const auto& [key, entry] : _entries)
    {
        if (! entry || entry->pinCount == 0 || key.context.isFilePlugin || key.context.instanceContext.empty())
        {
            continue;
        }

        if (! NavigationLocation::LooksLikeWindowsAbsolutePath(key.context.instanceContext))
        {
            continue;
        }

        const std::wstring instanceKey = MakeCaseInsensitivePathKey(key.context.instanceContext);
        if (! IsSameOrDescendantPath(instanceKey, normalizedPathKey, true))
        {
            continue;
        }

        DirectoryImpact impact{};
        impact.kind             = DirectoryImpact::Kind::ExitInstanceContext;
        impact.focusDisplayName = LeafNameNormalizedPath(key.context.instanceContext);
        PostImpactLocked(entry, impact);
    }
}

void DirectoryInfoCache::NotifyBackedContextPathMovedLocked(const std::wstring& normalizedSourcePath, const std::wstring& normalizedDestinationPath) noexcept
{
    const std::wstring sourceKey = MakeCaseInsensitivePathKey(normalizedSourcePath);

    for (const auto& [key, entry] : _entries)
    {
        if (! entry || entry->pinCount == 0 || key.context.isFilePlugin || key.context.instanceContext.empty())
        {
            continue;
        }

        if (! NavigationLocation::LooksLikeWindowsAbsolutePath(key.context.instanceContext))
        {
            continue;
        }

        const std::wstring instanceKey = MakeCaseInsensitivePathKey(key.context.instanceContext);
        if (! IsSameOrDescendantPath(instanceKey, sourceKey, true))
        {
            continue;
        }

        DirectoryImpact impact{};
        impact.kind = DirectoryImpact::Kind::RetargetInstanceContext;
        impact.newInstanceContext =
            JoinNormalizedPath(normalizedDestinationPath, GetRelativeSubPath(key.context.instanceContext, normalizedSourcePath, true), true);
        PostImpactLocked(entry, impact);
    }
}

void DirectoryInfoCache::NotifyPathDeletedLocked(const ContextKey& context, const std::wstring& normalizedPath) noexcept
{
    const std::wstring parentPath = ParentNormalizedPath(normalizedPath, context.isFilePlugin);
    NotifyFolderContentsChangedLocked(context, parentPath);
    MarkSubtreeDirtyLocked(context, normalizedPath);

    const std::wstring deletedKey = MakeCaseInsensitivePathKey(normalizedPath);
    for (const auto& [key, entry] : _entries)
    {
        if (! entry || entry->pinCount == 0 || ! ContextKeyEq{}(key.context, context))
        {
            continue;
        }

        if (! IsSameOrDescendantPath(key.pathKey, deletedKey, context.isFilePlugin))
        {
            continue;
        }

        DirectoryImpact impact{};
        impact.kind             = DirectoryImpact::Kind::RelocateCurrentFolder;
        impact.targetFolder     = parentPath.empty() ? normalizedPath : parentPath;
        impact.focusDisplayName = LeafNameNormalizedPath(normalizedPath);
        PostImpactLocked(entry, impact);
    }

    if (context.isFilePlugin)
    {
        NotifyBackedContextPathDeletedLocked(normalizedPath);
    }
}

void DirectoryInfoCache::NotifyPathMovedLocked(const ContextKey& context,
                                               const std::wstring& normalizedSourcePath,
                                               const std::wstring& normalizedDestinationPath) noexcept
{
    const std::wstring sourceParent      = ParentNormalizedPath(normalizedSourcePath, context.isFilePlugin);
    const std::wstring destinationParent = ParentNormalizedPath(normalizedDestinationPath, context.isFilePlugin);
    if (! sourceParent.empty() && ! destinationParent.empty() && EqualsNoCase(sourceParent, destinationParent))
    {
        NotifyFolderContentsChangedLocked(
            context, sourceParent, LeafNameNormalizedPath(normalizedSourcePath), LeafNameNormalizedPath(normalizedDestinationPath));
    }
    else
    {
        NotifyFolderContentsChangedLocked(context, sourceParent);
        NotifyFolderContentsChangedLocked(context, destinationParent);
    }
    MarkSubtreeDirtyLocked(context, normalizedSourcePath);

    const std::wstring sourceKey = MakeCaseInsensitivePathKey(normalizedSourcePath);
    for (const auto& [key, entry] : _entries)
    {
        if (! entry || entry->pinCount == 0 || ! ContextKeyEq{}(key.context, context))
        {
            continue;
        }

        if (! IsSameOrDescendantPath(key.pathKey, sourceKey, context.isFilePlugin))
        {
            continue;
        }

        DirectoryImpact impact{};
        impact.kind = DirectoryImpact::Kind::RelocateCurrentFolder;
        impact.targetFolder =
            JoinNormalizedPath(normalizedDestinationPath, GetRelativeSubPath(key.path, normalizedSourcePath, context.isFilePlugin), context.isFilePlugin);
        PostImpactLocked(entry, impact);
    }

    if (context.isFilePlugin)
    {
        NotifyBackedContextPathMovedLocked(normalizedSourcePath, normalizedDestinationPath);
    }
}

void DirectoryInfoCache::OnWatcherNotification(const ContextKey& context, std::wstring watchedFolder, const FolderWatcherNotification& notification) noexcept
{
    std::lock_guard lock(_mutex);

    if (notification.overflow)
    {
        NotifyFolderContentsChangedLocked(context, watchedFolder);
        return;
    }

    // Track a single in-order rename pair per notification batch. A trailing orphan OLD_NAME degrades to delete below.
    std::optional<std::wstring> pendingRenameOld;
    for (const auto& change : notification.changes)
    {
        const std::wstring fullPath =
            change.relativePath.empty() ? watchedFolder : JoinNormalizedPath(watchedFolder, change.relativePath, context.isFilePlugin);

        switch (change.action)
        {
            case FILESYSTEM_DIR_CHANGE_ADDED:
            case FILESYSTEM_DIR_CHANGE_MODIFIED:
                if (change.relativePath.empty())
                {
                    NotifyFolderContentsChangedLocked(context, watchedFolder);
                }
                else
                {
                    NotifyFolderContentsChangedLocked(context, ParentNormalizedPath(fullPath, context.isFilePlugin));
                }
                pendingRenameOld.reset();
                break;
            case FILESYSTEM_DIR_CHANGE_REMOVED:
                NotifyPathDeletedLocked(context, fullPath);
                pendingRenameOld.reset();
                break;
            case FILESYSTEM_DIR_CHANGE_RENAMED_OLD_NAME: pendingRenameOld = fullPath; break;
            case FILESYSTEM_DIR_CHANGE_RENAMED_NEW_NAME:
                if (pendingRenameOld)
                {
                    NotifyPathMovedLocked(context, *pendingRenameOld, fullPath);
                    pendingRenameOld.reset();
                }
                else
                {
                    if (change.relativePath.empty())
                    {
                        NotifyFolderContentsChangedLocked(context, watchedFolder);
                    }
                    else
                    {
                        NotifyFolderContentsChangedLocked(context, ParentNormalizedPath(fullPath, context.isFilePlugin));
                    }
                }
                break;
            case FILESYSTEM_DIR_CHANGE_UNKNOWN:
                NotifyFolderContentsChangedLocked(context, watchedFolder);
                pendingRenameOld.reset();
                break;
        }
    }

    if (pendingRenameOld)
    {
        NotifyPathDeletedLocked(context, *pendingRenameOld);
    }
}

void DirectoryInfoCache::UpdateWatchersLocked(std::vector<std::unique_ptr<FolderWatcher>>& watchersToStop) noexcept
{
    if (_maxWatchers == 0)
    {
        for (const auto& entry : _lru)
        {
            StopWatcherLocked(entry, watchersToStop);
        }
        return;
    }

    std::unordered_set<const Entry*> wanted;
    wanted.reserve(static_cast<size_t>(_maxWatchers));

    // 1) Pinned folders first (used on screen).
    uint32_t watcherBudget = _maxWatchers;
    for (const auto& entry : _lru)
    {
        if (entry->pinCount == 0)
        {
            continue;
        }

        if (watcherBudget == 0)
        {
            break;
        }

        wanted.insert(entry.get());
        --watcherBudget;
    }

    // 2) Then explicitly opted-in MRU non-pinned entries (best-effort).
    uint32_t watchedMru = 0;
    for (const auto& entry : _lru)
    {
        if (watcherBudget == 0 || watchedMru >= _mruWatched)
        {
            break;
        }
        if (entry->pinCount > 0)
        {
            continue;
        }
        if (! entry->allowOffscreenWatch || ! entry->info || entry->loading)
        {
            continue;
        }

        wanted.insert(entry.get());
        --watcherBudget;
        ++watchedMru;
    }

    // Apply watcher selection.
    for (const auto& entry : _lru)
    {
        if (wanted.contains(entry.get()))
        {
            StartWatcherLocked(entry, watchersToStop);
        }
        else
        {
            StopWatcherLocked(entry, watchersToStop);
        }
    }
}

void DirectoryInfoCache::MaybeEvictLocked(std::vector<std::unique_ptr<FolderWatcher>>& watchersToStop) noexcept
{
    if (_maxBytes == 0)
    {
        return;
    }

    if (_currentBytes <= _maxBytes)
    {
        return;
    }

    const size_t scanLimit = _lru.size();
    size_t scanned         = 0;
    while (_currentBytes > _maxBytes && ! _lru.empty() && scanned < scanLimit)
    {
        ++scanned;

        auto it        = std::prev(_lru.end());
        auto candidate = *it;
        if (! candidate)
        {
            _lru.erase(it);
            continue;
        }

        const bool inUse = (candidate->pinCount > 0) || (candidate->borrowCount > 0) || candidate->loading;
        if (inUse)
        {
            // Cannot evict protected entries. Move on after at most one full pass so an
            // over-budget cache with only pinned/borrowed/loading entries cannot spin.
            _lru.splice(_lru.begin(), _lru, it);
            candidate->lruIt = _lru.begin();
            continue;
        }

        const uint64_t bytesFreed = candidate->bytes;
        _currentBytes             = (_currentBytes >= bytesFreed) ? (_currentBytes - bytesFreed) : 0;
        StopWatcherLocked(candidate, watchersToStop);
        _entries.erase(candidate->key);
        _lru.erase(it);
        candidate->lruItValid = false;
        ++_evictions;

        Debug::Info(L"DirectoryInfoCache: Evicted '{}' ({} MiB), current={} MiB, max={} MiB",
                    candidate->key.path,
                    bytesFreed / kMiB,
                    _currentBytes / kMiB,
                    _maxBytes / kMiB);
    }
}

HRESULT DirectoryInfoCache::EnsureLoaded(const std::shared_ptr<Entry>& entry, BorrowMode mode) noexcept
{
    return EnsureLoaded(entry, mode, std::stop_token{});
}

HRESULT DirectoryInfoCache::EnsureLoaded(const std::shared_ptr<Entry>& entry, BorrowMode mode, std::stop_token stopToken) noexcept
{
    if (! entry)
    {
        return E_INVALIDARG;
    }

    if (_shuttingDown.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (stopToken.stop_requested())
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (! _initialized)
    {
        SetLimits(ComputeDefaultMaxBytes(), _maxWatchers, _mruWatched);
    }

    for (;;)
    {
        std::unique_lock lock(_mutex);

        TouchLocked(entry);

        if (entry->info && ! entry->dirty)
        {
            ++_cacheHits;
            return S_OK;
        }

        if (mode == BorrowMode::CacheOnly)
        {
            if (entry->info)
            {
                ++_cacheHits;
                return S_OK; // Snapshot available (may be stale); caller opted out of re-enumeration.
            }
            return S_FALSE;
        }

        if (stopToken.stop_requested())
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        if (entry->loading)
        {
            std::stop_callback stopWake(stopToken, [&] { entry->cv.notify_all(); });
            entry->cv.wait(lock, [&] { return stopToken.stop_requested() || ! entry->loading; });
            if (stopToken.stop_requested())
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
            continue;
        }

        entry->loading = true;
        break;
    }

    ++_cacheMisses;

    if (stopToken.stop_requested())
    {
        std::vector<std::unique_ptr<FolderWatcher>> watchersToStop;
        {
            std::lock_guard lock(_mutex);
            entry->loading = false;
            TouchLocked(entry);
            MaybeEvictLocked(watchersToStop);
            entry->cv.notify_all();
        }
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    // Perform enumeration outside the cache lock.
    wil::com_ptr<IFileSystem> fileSystem;
    {
        std::lock_guard lock(_mutex);
        fileSystem = ResolveProviderLocked(entry->key.context);
    }

    wil::com_ptr<IFilesInformation> info;
    Debug::Perf::Scope perf(L"DirectoryInfoCache.ReadDirectoryInfo");
    perf.SetDetail(entry->key.path);

    const HRESULT hr = fileSystem ? fileSystem->ReadDirectoryInfo(entry->key.path.c_str(), info.put()) : E_POINTER;
    perf.SetHr(hr);

    uint64_t entryBytes = 0;
    if (SUCCEEDED(hr) && info)
    {
        unsigned long allocated = 0;
        if (SUCCEEDED(info->GetAllocatedSize(&allocated)))
        {
            entryBytes = static_cast<uint64_t>(allocated);
        }
    }
    perf.SetValue0(entryBytes);

    {
        std::vector<std::unique_ptr<FolderWatcher>> watchersToStop;
        std::lock_guard lock(_mutex);
        entry->loading = false;

        if (FAILED(hr))
        {
            // Keep old snapshot if it exists, but mark it dirty so callers can try again later.
            Debug::Warning(L"DirectoryInfoCache: enumeration failed for '{}' (hr=0x{:08X})", entry->key.path, static_cast<unsigned long>(hr));
            entry->refreshPosted = false;
        }
        else
        {
            const uint64_t oldBytes = entry->bytes;
            entry->info             = std::move(info);
            entry->bytes            = entryBytes;
            entry->dirty            = false;
            entry->refreshPosted    = false;
            ++_enumerations;

            if (_currentBytes >= oldBytes)
            {
                _currentBytes -= oldBytes;
            }
            else
            {
                _currentBytes = 0;
            }
            _currentBytes += entry->bytes;
        }

        TouchLocked(entry);
        MaybeEvictLocked(watchersToStop);
        entry->cv.notify_all();
    }

    if (stopToken.stop_requested())
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    return hr;
}

DirectoryInfoCache::Borrowed DirectoryInfoCache::BorrowDirectoryInfo(IFileSystem* fileSystem, const std::filesystem::path& folder, BorrowMode mode) noexcept
{
    return BorrowDirectoryInfo(fileSystem, folder, mode, std::stop_token{});
}

DirectoryInfoCache::Borrowed DirectoryInfoCache::BorrowDirectoryInfo(IFileSystem* fileSystem,
                                                                     const std::filesystem::path& folder,
                                                                     BorrowMode mode,
                                                                     std::stop_token stopToken) noexcept
{
    Borrowed result{};
    result._owner = this;

    if (_shuttingDown.load(std::memory_order_acquire))
    {
        result._status = HRESULT_FROM_WIN32(ERROR_CANCELLED);
        return result;
    }

    const auto keyOpt = MakeKey(fileSystem, folder);
    if (! keyOpt)
    {
        result._status = E_INVALIDARG;
        return result;
    }

    std::shared_ptr<Entry> entry;
    {
        std::lock_guard lock(_mutex);
        entry = GetOrCreateEntryLocked(*keyOpt);
        TouchLocked(entry);
        AddBorrowLocked(entry);
    }

    result._entry  = entry;
    result._status = EnsureLoaded(entry, mode, stopToken);

    if (result._status != S_OK)
    {
        std::lock_guard lock(_mutex);
        ReleaseBorrowLocked(entry);
        result._entry.reset();
    }

    return result;
}

DirectoryInfoCache::Pin DirectoryInfoCache::PinFolder(IFileSystem* fileSystem, const std::filesystem::path& folder, HWND hwnd, UINT message) noexcept
{
    Pin pin{};
    pin._owner   = this;
    pin._hwnd    = hwnd;
    pin._message = message;

    if (_shuttingDown.load(std::memory_order_acquire))
    {
        return pin;
    }

    const auto keyOpt = MakeKey(fileSystem, folder);
    if (! keyOpt)
    {
        return pin;
    }

    {
        std::vector<std::unique_ptr<FolderWatcher>> watchersToStop;
        std::lock_guard lock(_mutex);
        pin._entry = GetOrCreateEntryLocked(*keyOpt);
        AddPinLocked(pin._entry);
        AddSubscriberLocked(pin._entry, hwnd, message);
        TouchLocked(pin._entry);
        UpdateWatchersLocked(watchersToStop);
    }
    return pin;
}
