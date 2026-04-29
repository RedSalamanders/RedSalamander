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
#include <string_view>
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

    struct DirectoryImpact
    {
        enum class Kind : uint8_t
        {
            RefreshCurrentFolder,
            RelocateCurrentFolder,
            RetargetInstanceContext,
            ExitInstanceContext,
        };

        Kind kind = Kind::RefreshCurrentFolder;
        std::filesystem::path currentFolder;
        std::filesystem::path targetFolder;
        std::wstring focusDisplayName;
        std::wstring renamedFromDisplayName;
        std::wstring renamedToDisplayName;
        std::wstring newInstanceContext;
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

    // Releases all cached entries and stops watchers. The singleton object itself is intentionally not destroyed
    // (static destruction order safety).
    void Shutdown() noexcept;

    void RegisterProvider(IFileSystem* fileSystem, std::wstring_view pluginId, std::wstring_view pluginShortId, std::wstring_view instanceContext) noexcept;
    void UnregisterProvider(IFileSystem* fileSystem) noexcept;

    void ClearForFileSystem(IFileSystem* fileSystem) noexcept;
    void InvalidateFolder(IFileSystem* fileSystem, const std::filesystem::path& folder) noexcept;
    bool IsFolderWatched(IFileSystem* fileSystem, const std::filesystem::path& folder) const noexcept;

    void NotifyFolderContentsChanged(IFileSystem* fileSystem, const std::filesystem::path& folder) noexcept;
    void NotifyPathCreated(IFileSystem* fileSystem, const std::filesystem::path& path) noexcept;
    void NotifyPathDeleted(IFileSystem* fileSystem, const std::filesystem::path& path) noexcept;
    void NotifyPathMoved(IFileSystem* fileSystem, const std::filesystem::path& sourcePath, const std::filesystem::path& destinationPath) noexcept;

    struct ContextKey
    {
        std::wstring pluginId;
        std::wstring pluginIdKey;
        std::wstring pluginShortId;
        std::wstring instanceContext;
        std::wstring instanceContextKey;
        bool isFilePlugin = false;
    };

    struct ContextKeyHash
    {
        size_t operator()(const ContextKey& key) const noexcept;
    };

    struct ContextKeyEq
    {
        bool operator()(const ContextKey& a, const ContextKey& b) const noexcept;
    };

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
        ContextKey context;
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

    struct ContextState
    {
        ContextKey key{};
        std::unordered_map<IFileSystem*, wil::com_ptr<IFileSystem>> providers;
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
        uint64_t bytes           = 0;
        bool dirty               = true;
        bool refreshPosted       = false;
        bool loading             = false;
        uint32_t pinCount        = 0;
        uint32_t borrowCount     = 0;
        bool allowOffscreenWatch = false;
        std::condition_variable cv;
        std::vector<Subscriber> subscribers;
        std::unique_ptr<class FolderWatcher> watcher;
        std::list<std::shared_ptr<Entry>>::iterator lruIt{};
        bool lruItValid = false;
    };

    static uint64_t ComputeDefaultMaxBytes() noexcept;

    std::optional<ContextKey> ResolveContextForFileSystem(IFileSystem* fileSystem) const noexcept;
    std::optional<Key> MakeKey(IFileSystem* fileSystem, const std::filesystem::path& folder) const noexcept;
    HRESULT EnsureLoaded(const std::shared_ptr<Entry>& entry, BorrowMode mode) noexcept;
    HRESULT EnsureLoaded(const std::shared_ptr<Entry>& entry, BorrowMode mode, std::stop_token stopToken) noexcept;
    void TouchLocked(const std::shared_ptr<Entry>& entry) noexcept;
    void MaybeEvictLocked(std::vector<std::unique_ptr<FolderWatcher>>& watchersToStop) noexcept;
    void UpdateWatchersLocked(std::vector<std::unique_ptr<FolderWatcher>>& watchersToStop) noexcept;
    void StartWatcherLocked(const std::shared_ptr<Entry>& entry, std::vector<std::unique_ptr<FolderWatcher>>& watchersToStop) noexcept;
    void StopWatcherLocked(const std::shared_ptr<Entry>& entry, std::vector<std::unique_ptr<FolderWatcher>>& watchersToStop) noexcept;
    void RestartContextWatchersLocked(const ContextKey& context, std::vector<std::unique_ptr<FolderWatcher>>& watchersToStop) noexcept;
    void MarkDirtyLocked(const Key& key) noexcept;
    void MarkDirtyLocked(const std::shared_ptr<Entry>& entry) noexcept;
    void PostImpactLocked(const std::shared_ptr<Entry>& entry, const DirectoryImpact& impact) noexcept;
    void PostRefreshLocked(const std::shared_ptr<Entry>& entry) noexcept;

    void AddSubscriberLocked(const std::shared_ptr<Entry>& entry, HWND hwnd, UINT message) noexcept;
    void RemoveSubscriberLocked(const std::shared_ptr<Entry>& entry, HWND hwnd, UINT message) noexcept;

    void AddBorrowLocked(const std::shared_ptr<Entry>& entry) noexcept;
    void ReleaseBorrowLocked(const std::shared_ptr<Entry>& entry) noexcept;
    void AddPinLocked(const std::shared_ptr<Entry>& entry) noexcept;
    void ReleasePinLocked(const std::shared_ptr<Entry>& entry) noexcept;

    void ClearContextLocked(const ContextKey& context, std::vector<std::unique_ptr<FolderWatcher>>& watchersToStop) noexcept;
    std::shared_ptr<Entry> GetOrCreateEntryLocked(const Key& key) noexcept;
    wil::com_ptr<IFileSystem> ResolveProviderLocked(const ContextKey& context, IFileSystem* preferredFileSystem = nullptr) const noexcept;
    void OnWatcherNotification(const ContextKey& context, std::wstring watchedFolder, const struct FolderWatcherNotification& notification) noexcept;
    void NotifyFolderContentsChangedLocked(const ContextKey& context,
                                           const std::wstring& normalizedFolder,
                                           std::wstring_view renamedFromDisplayName = {},
                                           std::wstring_view renamedToDisplayName   = {}) noexcept;
    void MarkSubtreeDirtyLocked(const ContextKey& context, const std::wstring& normalizedRootPath) noexcept;
    void NotifyPathDeletedLocked(const ContextKey& context, const std::wstring& normalizedPath) noexcept;
    void NotifyPathMovedLocked(const ContextKey& context, const std::wstring& normalizedSourcePath, const std::wstring& normalizedDestinationPath) noexcept;
    void NotifyBackedContextPathDeletedLocked(const std::wstring& normalizedPath) noexcept;
    void NotifyBackedContextPathMovedLocked(const std::wstring& normalizedSourcePath, const std::wstring& normalizedDestinationPath) noexcept;

    mutable std::mutex _mutex;
    std::atomic<bool> _shuttingDown{false};
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
    std::unordered_map<ContextKey, ContextState, ContextKeyHash, ContextKeyEq> _contexts;
    std::unordered_map<IFileSystem*, ContextKey> _providerContexts;
    std::unordered_map<Key, std::shared_ptr<Entry>, KeyHash, KeyEq> _entries;
};
