#pragma once

#include <Windows.h>

#include <atomic>
#include <condition_variable>
#include <cstdint>
#include <filesystem>
#include <list>
#include <memory>
#include <mutex>
#include <optional>
#include <stop_token>
#include <string>
#include <unordered_map>
#include <vector>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/com.h>
#pragma warning(pop)

#include "PlugInterfaces/FileSystem.h"
#include "SettingsStore.h"

class DirectoryInfoCache
{
public:
    static DirectoryInfoCache& GetInstance();

private:
    struct Entry;

public:
    struct Stats
    {
        uint64_t maxBytes       = 0;
        uint64_t currentBytes   = 0;
        uint64_t cacheHits      = 0;
        uint64_t cacheMisses    = 0;
        uint64_t enumerations   = 0;
        uint64_t evictions      = 0;
        uint64_t dirtyMarks     = 0;
        uint32_t maxWatchers    = 0;
        uint32_t mruWatched     = 0;
        uint32_t activeWatchers = 0;
        uint32_t pinnedEntries  = 0;
        uint32_t entryCount     = 0;
    };

    enum class BorrowMode : uint8_t
    {
        CacheOnly,
        AllowEnumerate,
    };

    class Borrowed final
    {
    public:
        Borrowed() = default;
        Borrowed(Borrowed&&) noexcept;
        Borrowed& operator=(Borrowed&&) noexcept;
        ~Borrowed();

        Borrowed(const Borrowed&)            = delete;
        Borrowed& operator=(const Borrowed&) = delete;

        HRESULT Status() const noexcept;
        IFilesInformation* Get() const noexcept;
        const std::wstring& NormalizedPath() const noexcept;

    private:
        friend class DirectoryInfoCache;

        DirectoryInfoCache* _owner = nullptr;
        std::shared_ptr<Entry> _entry;
        HRESULT _status = E_FAIL;
    };

    class Pin final
    {
    public:
        Pin() = default;
        Pin(Pin&&) noexcept;
        Pin& operator=(Pin&&) noexcept;
        ~Pin();

        Pin(const Pin&)            = delete;
        Pin& operator=(const Pin&) = delete;

        bool IsValid() const noexcept;
        const std::wstring& NormalizedPath() const noexcept;

    private:
        friend class DirectoryInfoCache;

        DirectoryInfoCache* _owner = nullptr;
        std::shared_ptr<Entry> _entry;
        HWND _hwnd    = nullptr;
        UINT _message = 0;
    };

    void ApplySettings(const Common::Settings::Settings& settings) noexcept;
    void SetLimits(uint64_t maxBytes, uint32_t maxWatchers, uint32_t mruWatched) noexcept;
    Stats GetStats() const noexcept;
    void ClearForFileSystem(IFileSystem* fileSystem) noexcept;
    void InvalidateFolder(IFileSystem* fileSystem, const std::filesystem::path& folder) noexcept;
    bool IsFolderWatched(IFileSystem* fileSystem, const std::filesystem::path& folder) const noexcept;

    Borrowed BorrowDirectoryInfo(IFileSystem* fileSystem, const std::filesystem::path& folder, BorrowMode mode) noexcept;
    Borrowed BorrowDirectoryInfo(IFileSystem* fileSystem, const std::filesystem::path& folder, BorrowMode mode, std::stop_token stopToken) noexcept;
    Pin PinFolder(IFileSystem* fileSystem, const std::filesystem::path& folder, HWND hwnd, UINT message) noexcept;

private:
    DirectoryInfoCache()                                     = default;
    ~DirectoryInfoCache()                                    = default;
    DirectoryInfoCache(const DirectoryInfoCache&)            = delete;
    DirectoryInfoCache& operator=(const DirectoryInfoCache&) = delete;
    DirectoryInfoCache(DirectoryInfoCache&&)                 = delete;
    DirectoryInfoCache& operator=(DirectoryInfoCache&&)      = delete;

    struct Key
    {
        wil::com_ptr<IFileSystem> fileSystem;
        std::wstring path;
        std::wstring pathKey;
    };

    struct KeyHash
    {
        size_t operator()(const Key& key) const noexcept;
    };

    struct KeyEq
    {
        bool operator()(const Key& a, const Key& b) const noexcept;
    };

    struct Subscriber
    {
        HWND hwnd    = nullptr;
        UINT message = 0;
    };

    struct Entry
    {
        Entry()                        = default;
        Entry(const Entry&)            = delete;
        Entry& operator=(const Entry&) = delete;
        Entry(Entry&&)                 = delete;
        Entry& operator=(Entry&&)      = delete;

        Key key{};
        wil::com_ptr<IFilesInformation> info;
        uint64_t bytes       = 0;
        bool dirty           = true;
        bool notifyPosted    = false;
        bool loading         = false;
        uint32_t pinCount    = 0;
        uint32_t borrowCount = 0;
        std::condition_variable cv;
        std::vector<Subscriber> subscribers;
        std::unique_ptr<class FolderWatcher> watcher;
        std::list<std::shared_ptr<Entry>>::iterator lruIt{};
        bool lruItValid = false;
    };

    static uint64_t ComputeDefaultMaxBytes() noexcept;

    std::optional<Key> MakeKey(IFileSystem* fileSystem, const std::filesystem::path& folder) const noexcept;
    HRESULT EnsureLoaded(const std::shared_ptr<Entry>& entry, BorrowMode mode) noexcept;
    HRESULT EnsureLoaded(const std::shared_ptr<Entry>& entry, BorrowMode mode, std::stop_token stopToken) noexcept;
    void TouchLocked(const std::shared_ptr<Entry>& entry) noexcept;
    void MaybeEvictLocked(std::vector<std::unique_ptr<FolderWatcher>>& watchersToStop) noexcept;
    void UpdateWatchersLocked(std::vector<std::unique_ptr<FolderWatcher>>& watchersToStop) noexcept;
    void StartWatcherLocked(const std::shared_ptr<Entry>& entry, std::vector<std::unique_ptr<FolderWatcher>>& watchersToStop) noexcept;
    void StopWatcherLocked(const std::shared_ptr<Entry>& entry, std::vector<std::unique_ptr<FolderWatcher>>& watchersToStop) noexcept;
    void MarkDirtyLocked(const Key& key) noexcept;
    void PostDirtyNotificationsLocked(const std::shared_ptr<Entry>& entry) noexcept;

    void AddSubscriberLocked(const std::shared_ptr<Entry>& entry, HWND hwnd, UINT message) noexcept;
    void RemoveSubscriberLocked(const std::shared_ptr<Entry>& entry, HWND hwnd, UINT message) noexcept;

    void AddBorrowLocked(const std::shared_ptr<Entry>& entry) noexcept;
    void ReleaseBorrowLocked(const std::shared_ptr<Entry>& entry) noexcept;
    void AddPinLocked(const std::shared_ptr<Entry>& entry) noexcept;
    void ReleasePinLocked(const std::shared_ptr<Entry>& entry) noexcept;

    std::shared_ptr<Entry> GetOrCreateEntryLocked(const Key& key) noexcept;

    mutable std::mutex _mutex;
    uint64_t _maxBytes     = 0;
    uint64_t _currentBytes = 0;
    uint32_t _maxWatchers  = 64;
    uint32_t _mruWatched   = 16;
    bool _initialized      = false;

    uint64_t _cacheHits    = 0;
    uint64_t _cacheMisses  = 0;
    uint64_t _enumerations = 0;
    uint64_t _evictions    = 0;
    uint64_t _dirtyMarks   = 0;

    std::list<std::shared_ptr<Entry>> _lru;
    std::unordered_map<Key, std::shared_ptr<Entry>, KeyHash, KeyEq> _entries;
};
