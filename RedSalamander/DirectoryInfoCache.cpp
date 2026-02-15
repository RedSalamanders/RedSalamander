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

#pragma warning(push)
// C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
struct FileSystemPathSemanticsCache
{
    std::mutex mutex;
    std::unordered_map<IFileSystem*, bool> isFilePluginByPtr;
};
#pragma warning(pop)

FileSystemPathSemanticsCache& GetPathSemanticsCache() noexcept
{
    static FileSystemPathSemanticsCache cache;
    return cache;
}

bool IsFilePlugin(IFileSystem* fileSystem) noexcept
{
    if (! fileSystem)
    {
        return false;
    }

    FileSystemPathSemanticsCache& cache = GetPathSemanticsCache();
    {
        std::lock_guard lock(cache.mutex);
        const auto it = cache.isFilePluginByPtr.find(fileSystem);
        if (it != cache.isFilePluginByPtr.end())
        {
            return it->second;
        }
    }

    wil::com_ptr<IFileSystem> fs = fileSystem;

    wil::com_ptr<IInformations> infos;
    const HRESULT qiHr = fs->QueryInterface(__uuidof(IInformations), infos.put_void());

    bool isFile = false;
    if (SUCCEEDED(qiHr) && infos)
    {
        const PluginMetaData* meta = nullptr;
        const HRESULT metaHr       = infos->GetMetaData(&meta);
        if (SUCCEEDED(metaHr) && meta)
        {
            const wchar_t* idToCheck = meta->shortId ? meta->shortId : meta->id;
            if (idToCheck)
            {
                isFile = CompareStringOrdinal(idToCheck, -1, L"file", -1, TRUE) == CSTR_EQUAL;
            }
        }
    }

    {
        std::lock_guard lock(cache.mutex);
        cache.isFilePluginByPtr.emplace(fileSystem, isFile);
    }

    return isFile;
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
} // namespace

DirectoryInfoCache& DirectoryInfoCache::GetInstance()
{
    static DirectoryInfoCache* instance = new DirectoryInfoCache();
    return *instance;
}

size_t DirectoryInfoCache::KeyHash::operator()(const Key& key) const noexcept
{
    const size_t ptrHash  = std::hash<void*>{}(key.fileSystem.get());
    const size_t pathHash = std::hash<std::wstring_view>{}(key.pathKey);
    return ptrHash ^ (pathHash + 0x9e3779b97f4a7c15ull + (ptrHash << 6) + (ptrHash >> 2));
}

bool DirectoryInfoCache::KeyEq::operator()(const Key& a, const Key& b) const noexcept
{
    if (a.fileSystem.get() != b.fileSystem.get())
    {
        return false;
    }
    return a.pathKey == b.pathKey;
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
            _owner->UpdateWatchersLocked(watchersToStop);
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
        _owner->UpdateWatchersLocked(watchersToStop);
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

void DirectoryInfoCache::ClearForFileSystem(IFileSystem* fileSystem) noexcept
{
    if (! fileSystem)
    {
        return;
    }

    std::vector<std::unique_ptr<FolderWatcher>> watchersToStop;
    {
        std::lock_guard lock(_mutex);

        for (auto it = _entries.begin(); it != _entries.end();)
        {
            const Key& key = it->first;
            if (key.fileSystem.get() != fileSystem)
            {
                ++it;
                continue;
            }

            const std::shared_ptr<Entry>& entry = it->second;
            if (entry)
            {
                StopWatcherLocked(entry, watchersToStop);

                const uint64_t bytesFreed = entry->bytes;
                if (_currentBytes >= bytesFreed)
                {
                    _currentBytes -= bytesFreed;
                }
                else
                {
                    _currentBytes = 0;
                }

                entry->info         = nullptr;
                entry->bytes        = 0;
                entry->dirty        = true;
                entry->notifyPosted = false;

                if (entry->lruItValid)
                {
                    _lru.erase(entry->lruIt);
                    entry->lruItValid = false;
                }
            }

            it = _entries.erase(it);
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

    const bool isFilePlugin = IsFilePlugin(fileSystem);
    std::wstring normalized = NormalizePath(folder.native(), isFilePlugin);
    if (normalized.empty())
    {
        return std::nullopt;
    }

    Key key{};
    key.fileSystem = fileSystem;
    key.path       = std::move(normalized);
    key.pathKey    = MakeCaseInsensitivePathKey(key.path);
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

void DirectoryInfoCache::PostDirtyNotificationsLocked(const std::shared_ptr<Entry>& entry) noexcept
{
    if (! entry || entry->subscribers.empty())
    {
        return;
    }

    if (entry->notifyPosted)
    {
        return;
    }

    entry->notifyPosted = true;
    for (const auto& s : entry->subscribers)
    {
        if (s.hwnd && s.message)
        {
            PostMessageW(s.hwnd, s.message, 0, 0);
        }
    }
}

void DirectoryInfoCache::MarkDirtyLocked(const Key& key) noexcept
{
    auto it = _entries.find(key);
    if (it == _entries.end())
    {
        return;
    }

    const auto& entry = it->second;
    if (! entry)
    {
        return;
    }

    entry->dirty = true;
    ++_dirtyMarks;
    PostDirtyNotificationsLocked(entry);
}

void DirectoryInfoCache::StartWatcherLocked(const std::shared_ptr<Entry>& entry, std::vector<std::unique_ptr<FolderWatcher>>& watchersToStop) noexcept
{
    if (! entry || entry->watcher)
    {
        return;
    }

    const Key key           = entry->key;
    const std::wstring path = entry->key.path;

    auto markDirty = [key]
    {
        DirectoryInfoCache& cache = DirectoryInfoCache::GetInstance();
        std::lock_guard lock(cache._mutex);
        cache.MarkDirtyLocked(key);
    };

    wil::com_ptr<IFileSystemDirectoryWatch> dirWatch;
    const HRESULT qiHr = entry->key.fileSystem ? entry->key.fileSystem->QueryInterface(__uuidof(IFileSystemDirectoryWatch), dirWatch.put_void()) : E_POINTER;
    if (FAILED(qiHr) || ! dirWatch)
    {
        return;
    }

    entry->watcher = std::make_unique<FolderWatcher>(std::move(dirWatch), path, std::move(markDirty));

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

    // 2) Then MRU non-pinned entries (best-effort).
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
        if (! entry->info || entry->loading)
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

    while (_currentBytes > _maxBytes && ! _lru.empty())
    {
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
            // Cannot evict pinned/borrowed entries; rotate to avoid infinite loop.
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
        std::vector<std::unique_ptr<FolderWatcher>> watchersToStop;
        std::unique_lock lock(_mutex);

        TouchLocked(entry);
        UpdateWatchersLocked(watchersToStop);

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
            UpdateWatchersLocked(watchersToStop);
            entry->cv.notify_all();
        }
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    // Perform enumeration outside the cache lock.
    wil::com_ptr<IFileSystem> fileSystem = entry->key.fileSystem;

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
        }
        else
        {
            const uint64_t oldBytes = entry->bytes;
            entry->info             = std::move(info);
            entry->bytes            = entryBytes;
            entry->dirty            = false;
            entry->notifyPosted     = false;
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
        UpdateWatchersLocked(watchersToStop);
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

DirectoryInfoCache::Borrowed
DirectoryInfoCache::BorrowDirectoryInfo(IFileSystem* fileSystem, const std::filesystem::path& folder, BorrowMode mode, std::stop_token stopToken) noexcept
{
    Borrowed result{};
    result._owner = this;

    const auto keyOpt = MakeKey(fileSystem, folder);
    if (! keyOpt)
    {
        result._status = E_INVALIDARG;
        return result;
    }

    std::shared_ptr<Entry> entry;
    {
        std::vector<std::unique_ptr<FolderWatcher>> watchersToStop;
        std::lock_guard lock(_mutex);
        entry = GetOrCreateEntryLocked(*keyOpt);
        TouchLocked(entry);
        AddBorrowLocked(entry);
        UpdateWatchersLocked(watchersToStop);
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
