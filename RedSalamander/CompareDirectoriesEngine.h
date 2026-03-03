#pragma once

#include <atomic>
#include <condition_variable>
#include <cstdint>
#include <deque>
#include <filesystem>
#include <functional>
#include <list>
#include <limits>
#include <map>
#include <memory>
#include <mutex>
#include <optional>
#include <set>
#include <stop_token>
#include <string>
#include <string_view>
#include <thread>
#include <unordered_map>
#include <vector>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/com.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#include "DirectoryInfoCache.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "SettingsStore.h"

enum class ComparePane : uint8_t
{
    Left,
    Right,
};

struct WStringViewNoCaseLess
{
    using is_transparent = void;

    bool operator()(std::wstring_view left, std::wstring_view right) const noexcept
    {
        if (left.size() > static_cast<size_t>(std::numeric_limits<int>::max()) || right.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
        {
            return left < right;
        }
        return wil::compare_string_ordinal(left, right, true) == wistd::weak_ordering::less;
    }
};

enum class CompareDirectoriesDiffBit : uint32_t
{
    OnlyInLeft  = 0x01u,
    OnlyInRight = 0x02u,

    TypeMismatch = 0x04u,

    Size           = 0x08u,
    DateTime       = 0x10u,
    Attributes     = 0x20u,
    Content        = 0x40u,
    ContentPending = 0x200u,

    SubdirAttributes = 0x80u,
    SubdirContent    = 0x100u,
    SubdirPending    = 0x400u,
};

[[nodiscard]] inline constexpr uint32_t operator|(CompareDirectoriesDiffBit a, CompareDirectoriesDiffBit b) noexcept
{
    return static_cast<uint32_t>(a) | static_cast<uint32_t>(b);
}

[[nodiscard]] inline constexpr uint32_t operator|(uint32_t a, CompareDirectoriesDiffBit b) noexcept
{
    return a | static_cast<uint32_t>(b);
}

[[nodiscard]] inline constexpr bool HasFlag(uint32_t mask, CompareDirectoriesDiffBit bit) noexcept
{
    return (mask & static_cast<uint32_t>(bit)) != 0u;
}

struct CompareDirectoriesItemDecision
{
    bool isDirectory = false;
    bool existsLeft  = false;
    bool existsRight = false;

    bool isDifferent = false;
    bool selectLeft  = false;
    bool selectRight = false;

    uint32_t differenceMask = 0;

    uint64_t leftSizeBytes    = 0;
    int64_t leftLastWriteTime = 0;
    DWORD leftFileAttributes  = 0;

    uint64_t rightSizeBytes    = 0;
    int64_t rightLastWriteTime = 0;
    DWORD rightFileAttributes  = 0;
};

struct CompareDirectoriesFolderDecision
{
    uint64_t version        = 0;
    HRESULT hr              = S_OK;
    bool leftFolderMissing  = false;
    bool rightFolderMissing = false;
    // Pending content-compare jobs for files in this folder that are intentionally not surfaced
    // as per-item ContentPending entries when `showIdenticalItems` is off (keeps memory bounded).
    uint32_t pendingContentCompareCount = 0;
    // Precomputed aggregates over items (+ pendingContentCompareCount) — avoids O(n) scans in hot paths.
    bool anyDifferent = false;
    bool anyPending   = false;
    std::map<std::wstring, CompareDirectoriesItemDecision, WStringViewNoCaseLess> items;
};

struct CompareDirectoriesPerfStats
{
    uint64_t version   = 0;
    uint64_t uiVersion = 0;

    uint32_t scanActiveScans     = 0;
    uint64_t scanFoldersScanned  = 0;
    uint64_t scanEntriesScanned  = 0;
    size_t scanQueueHighSize     = 0;
    size_t scanQueueLowSize      = 0;
    size_t scanQueueSize         = 0;
    size_t scanScheduledKeys     = 0;
    size_t scanInFlightKeys      = 0;
    size_t pendingSubdirUpdates  = 0;

    size_t scanQueueHighWater        = 0;
    size_t scanQueueHighHighWater    = 0;
    size_t scanQueueLowHighWater     = 0;
    size_t scanScheduledHighWater    = 0;
    size_t scanInFlightHighWater     = 0;
    size_t pendingSubdirHighWater    = 0;

    uint64_t contentPendingCompares   = 0;
    uint64_t contentTotalCompares     = 0;
    uint64_t contentCompletedCompares = 0;
    uint64_t contentTotalBytes        = 0;
    uint64_t contentCompletedBytes    = 0;
    size_t contentQueueHighSize       = 0;
    size_t contentQueueLowSize        = 0;
    size_t contentQueueSize           = 0;
    size_t contentInFlightSize        = 0;
    size_t contentCacheSize           = 0;
    size_t pendingContentUpdates      = 0;

    size_t contentQueueHighWater      = 0;
    size_t contentQueueHighHighWater  = 0;
    size_t contentQueueLowHighWater   = 0;
    size_t contentInFlightHighWater   = 0;
    size_t contentCacheHighWater      = 0;
    size_t pendingContentHighWater    = 0;

    size_t decisionCacheEntries                 = 0;
    size_t decisionCacheEntriesHighWater        = 0;
    uint64_t decisionCacheEstimatedBytes        = 0;
    uint64_t decisionCacheEstimatedBytesHighWater = 0;
    uint64_t decisionCacheBudgetBytes           = 0;

    DirectoryInfoCache::Stats directoryInfoCache{};
};

class CompareDirectoriesSession final : public std::enable_shared_from_this<CompareDirectoriesSession>
{
public:
    using ScanProgressCallback    = std::function<void(const std::filesystem::path& relativeFolder,
                                                    std::wstring_view currentEntryName,
                                                    uint64_t scannedFolders,
                                                    uint64_t scannedEntries,
                                                    uint32_t activeScans,
                                                    uint64_t contentCandidateFileCount,
                                                    uint64_t contentCandidateTotalBytes)>;
    using ContentProgressCallback = std::function<void(uint32_t workerIndex,
                                                       const std::filesystem::path& relativeFolder,
                                                       std::wstring_view entryName,
                                                       uint64_t fileTotalBytes,
                                                       uint64_t fileCompletedBytes,
                                                       uint64_t overallTotalBytes,
                                                       uint64_t overallCompletedBytes,
                                                       uint64_t pendingContentCompares,
                                                       uint64_t totalContentCompares,
                                                       uint64_t completedContentCompares)>;
    using DecisionUpdatedCallback = std::function<void()>;

    CompareDirectoriesSession(wil::com_ptr<IFileSystem> leftFileSystem,
                              wil::com_ptr<IFileSystem> rightFileSystem,
                              std::filesystem::path leftRoot,
                              std::filesystem::path rightRoot,
                              Common::Settings::CompareDirectoriesSettings settings);
    ~CompareDirectoriesSession();

    CompareDirectoriesSession(const CompareDirectoriesSession&)            = delete;
    CompareDirectoriesSession& operator=(const CompareDirectoriesSession&) = delete;
    CompareDirectoriesSession(CompareDirectoriesSession&&)                 = delete;
    CompareDirectoriesSession& operator=(CompareDirectoriesSession&&)      = delete;

    void SetRoots(std::filesystem::path leftRoot, std::filesystem::path rightRoot);
    void SetSettings(Common::Settings::CompareDirectoriesSettings settings);
    void SetCompareEnabled(bool enabled) noexcept;
    [[nodiscard]] bool IsCompareEnabled() const noexcept;
    // Controls whether background work is allowed during compare mode:
    // - When disabled, content compare jobs are canceled/cleared and no new background work is queued.
    // - Used by the Compare Directories UI to implement a responsive "Cancel" action.
    void SetBackgroundWorkEnabled(bool enabled) noexcept;
    [[nodiscard]] bool IsBackgroundWorkEnabled() const noexcept;
    void Invalidate() noexcept;
    void InvalidateForAbsolutePath(const std::filesystem::path& absolutePath, bool includeSubtree) noexcept;

    // Applies any queued content-compare results to cached decisions (and updates ancestor folder
    // subtree status) so the UI can reflect completed comparisons without requiring navigation.
    void FlushPendingContentCompareUpdates() noexcept;
    // Applies queued content-compare results in bounded batches so callers can avoid long UI stalls.
    // Returns true when additional updates remain queued.
    [[nodiscard]] bool FlushPendingContentCompareUpdatesBudgeted(size_t maxFoldersToApply) noexcept;

    // Starts a progressive background scan for the current roots/settings.
    // The scan always includes the root folder so progress transitions reliably even when `compareSubdirectories` is off.
    void StartScan() noexcept;

    // Ensures a background scan job is queued for `relativeFolder` (cache key), so callers can
    // trigger progressive computation without doing synchronous I/O on the calling thread.
    // No-op when background work is disabled.
    void RequestScanForFolder(const std::filesystem::path& relativeFolder) noexcept;

    // Cache-only decision lookup. Never performs I/O or triggers traversal.
    [[nodiscard]] std::shared_ptr<const CompareDirectoriesFolderDecision> TryGetCachedDecision(const std::filesystem::path& relativeFolder) noexcept;

    // Pins the currently-visible folders (and their ancestor chains) in the decision cache so background
    // eviction does not evict UI-critical state.
    void SetPinnedFolders(const std::filesystem::path& leftRelativeFolder, const std::filesystem::path& rightRelativeFolder) noexcept;

    // Applies queued subtree status updates in bounded batches so callers can avoid long UI stalls.
    // Returns true when additional updates remain queued.
    [[nodiscard]] bool FlushPendingSubdirUpdatesBudgeted(size_t maxFoldersToApply) noexcept;

    void SetScanProgressCallback(ScanProgressCallback callback) noexcept;
    void SetContentProgressCallback(ContentProgressCallback callback) noexcept;
    void SetDecisionUpdatedCallback(DecisionUpdatedCallback callback) noexcept;

    [[nodiscard]] Common::Settings::CompareDirectoriesSettings GetSettings() const;
    [[nodiscard]] std::filesystem::path GetRoot(ComparePane pane) const;
    [[nodiscard]] uint64_t GetVersion() const noexcept;
    [[nodiscard]] uint64_t GetUiVersion() const noexcept;
    [[nodiscard]] CompareDirectoriesPerfStats GetPerfStats() const noexcept;

#ifdef _DEBUG
    // Selftest hook: allows exercising eviction logic without allocating hundreds of MB.
    // No production code should rely on this.
    void SetDecisionCacheBudgetBytesForSelfTest(uint64_t budgetBytes) noexcept;
#endif

    [[nodiscard]] wil::com_ptr<IFileSystem> GetFileSystem(ComparePane pane) const noexcept;
    [[nodiscard]] wil::com_ptr<IInformations> GetInformations(ComparePane pane) const noexcept;
    [[nodiscard]] wil::com_ptr<IFileSystemIO> GetFileSystemIO(ComparePane pane) const noexcept;
    [[nodiscard]] bool IsContentCompareSupported() const noexcept;

    [[nodiscard]] std::optional<std::filesystem::path> TryMakeRelative(ComparePane pane, const std::filesystem::path& absoluteFolder) const;
    [[nodiscard]] std::filesystem::path ResolveAbsolute(ComparePane pane, const std::filesystem::path& relativeFolder) const;

    [[nodiscard]] std::shared_ptr<const CompareDirectoriesFolderDecision> GetOrComputeDecision(const std::filesystem::path& relativeFolder);

private:
    enum class ScanPriority : uint8_t
    {
        Low,
        High,
    };

    struct FolderScanJob
    {
        uint64_t version     = 0;
        uint64_t cancelToken = 0;
        std::filesystem::path relativeFolder;
        std::wstring key;
        ScanPriority priority = ScanPriority::Low;
    };

    struct ContentCompareKey
    {
        // Cache key for the compared file within the current compare roots.
        // Uses '/' separators (MakeCacheKey semantics), not Win32-normalized paths.
        std::wstring relativeFileKey;
        uint64_t leftSizeBytes     = 0;
        uint64_t rightSizeBytes    = 0;
        int64_t leftLastWriteTime  = 0;
        int64_t rightLastWriteTime = 0;
        // File attributes are intentionally excluded: they do not affect byte content
        // and their presence caused spurious cache misses when only attributes changed.
    };

    struct ContentCompareKeyHash
    {
        size_t operator()(const ContentCompareKey& key) const noexcept;
    };

    struct ContentCompareKeyEq
    {
        bool operator()(const ContentCompareKey& a, const ContentCompareKey& b) const noexcept;
    };

    struct ContentCompareJob
    {
        uint64_t version     = 0;
        uint64_t cancelToken = 0;
        std::filesystem::path relativeFolder;
        std::wstring entryName;
        ContentCompareKey key;
        ScanPriority priority = ScanPriority::Low;
        // Attributes are not part of the cache key but are needed for the pending-update
        // staleness check in ApplyPendingContentCompareUpdatesLocked.
        DWORD leftFileAttributes  = 0;
        DWORD rightFileAttributes = 0;
    };

    struct PendingContentCompareUpdate
    {
        uint64_t version           = 0;
        uint64_t leftSizeBytes     = 0;
        uint64_t rightSizeBytes    = 0;
        int64_t leftLastWriteTime  = 0;
        int64_t rightLastWriteTime = 0;
        DWORD leftFileAttributes   = 0;
        DWORD rightFileAttributes  = 0;
        bool areEqual              = false;
    };

    std::wstring MakeCacheKey(const std::filesystem::path& relativeFolder) const;
    void InvalidateForRelativePathLocked(const std::filesystem::path& relativePath, bool includeSubtree) noexcept;
    void NotifyScanProgress(const std::filesystem::path& relativeFolder, std::wstring_view currentEntryName, bool force) noexcept;
    void NotifyContentProgress(
        uint32_t workerIndex, const std::filesystem::path& relativeFolder, std::wstring_view entryName, uint64_t totalBytes, uint64_t completedBytes) noexcept;
    void NotifyDecisionUpdated(bool force) noexcept;
    void EnsureScanWorkersLocked() noexcept;
    void EnqueueScanLocked(const std::filesystem::path& relativeFolder, uint64_t version, uint64_t cancelToken, ScanPriority priority) noexcept;
    void EnsureContentCompareWorkersLocked() noexcept;
    struct DecisionCacheLruEntry
    {
        uint64_t estimatedBytes = 0;
        std::list<std::wstring>::iterator lruIt{};
    };
    struct ResetCleanup final
    {
        std::map<std::wstring, std::shared_ptr<const CompareDirectoriesFolderDecision>, WStringViewNoCaseLess> cache;
        std::list<std::wstring> decisionCacheLru;
        std::map<std::wstring, DecisionCacheLruEntry, WStringViewNoCaseLess> decisionCacheMeta;
        std::set<std::wstring, WStringViewNoCaseLess> decisionCachePinnedKeys;
        std::deque<FolderScanJob> scanQueueHigh;
        std::deque<FolderScanJob> scanQueueLow;
        std::set<std::wstring, WStringViewNoCaseLess> scanScheduledKeys;
        std::set<std::wstring, WStringViewNoCaseLess> scanHighQueuedKeys;
        std::set<std::wstring, WStringViewNoCaseLess> scanInFlightKeys;
        std::set<std::wstring, WStringViewNoCaseLess> pendingSubdirUpdates;
        std::unordered_map<ContentCompareKey, uint64_t, ContentCompareKeyHash, ContentCompareKeyEq> contentCompareInFlight;
        std::deque<ContentCompareJob> contentCompareQueueHigh;
        std::deque<ContentCompareJob> contentCompareQueueLow;
        std::map<std::wstring, std::map<std::wstring, PendingContentCompareUpdate, WStringViewNoCaseLess>, WStringViewNoCaseLess> pendingContentCompareUpdates;
    };

    static void ScheduleResetCleanup(std::unique_ptr<ResetCleanup> cleanup) noexcept;
    void ResetCompareStateLocked(ResetCleanup& outCleanup) noexcept;
    void ClearContentCompareStateLocked() noexcept;
    static uint64_t EstimateDecisionBytes(std::wstring_view folderKey, const CompareDirectoriesFolderDecision& decision) noexcept;
    void TouchDecisionCacheKeyLocked(std::wstring_view folderKey) noexcept;
    void TrackDecisionCacheInsertOrUpdateLocked(
        std::wstring_view folderKey, const std::shared_ptr<const CompareDirectoriesFolderDecision>& decision) noexcept;
    void TrackDecisionCacheEraseLocked(std::wstring_view folderKey) noexcept;
    void MaybeEvictDecisionCacheLocked() noexcept;
    [[nodiscard]] bool PropagateChildAggregateToAncestorsLocked(
        std::wstring_view childKey, const Common::Settings::CompareDirectoriesSettings& settings, uint64_t currentVersion) noexcept;
    [[nodiscard]] std::shared_ptr<CompareDirectoriesFolderDecision> ComputeDecisionForFolder(
        const std::filesystem::path& relativeFolder,
        const Common::Settings::CompareDirectoriesSettings& settings,
        const std::vector<std::wstring>& ignoreFilePatterns,
        const std::vector<std::wstring>& ignoreDirectoryPatterns,
        uint64_t version,
        uint64_t cancelToken,
        bool allowBackgroundWork,
        bool reportScanProgress,
        bool forceNotifyFolderStart,
        ScanPriority scanPriority,
        std::stop_token stopToken) noexcept;
    void ApplyPendingContentCompareUpdatesLocked(const std::wstring& folderKey) noexcept;
    void ScanWorker(std::stop_token stopToken, uint32_t workerIndex) noexcept;
    void ContentCompareWorker(std::stop_token stopToken, uint32_t workerIndex) noexcept;

    wil::com_ptr<IFileSystem> _leftFileSystem;
    wil::com_ptr<IFileSystem> _rightFileSystem;
    wil::com_ptr<IInformations> _leftInformations;
    wil::com_ptr<IInformations> _rightInformations;
    wil::com_ptr<IFileSystemIO> _leftFileSystemIo;
    wil::com_ptr<IFileSystemIO> _rightFileSystemIo;

    mutable std::mutex _mutex;
    std::filesystem::path _leftRoot;
    std::filesystem::path _rightRoot;
    Common::Settings::CompareDirectoriesSettings _settings;
    std::atomic_uint64_t _version{1};
    std::atomic_bool _compareEnabled{true};
    std::atomic_bool _backgroundWorkEnabled{true};
    std::atomic_uint64_t _backgroundWorkCancelToken{1};
    uint64_t _uiVersion = 1;

    std::map<std::wstring, std::shared_ptr<const CompareDirectoriesFolderDecision>, WStringViewNoCaseLess> _cache;
    std::list<std::wstring> _decisionCacheLru;
    std::map<std::wstring, DecisionCacheLruEntry, WStringViewNoCaseLess> _decisionCacheMeta;
    std::set<std::wstring, WStringViewNoCaseLess> _decisionCachePinnedKeys;

    static constexpr uint64_t kDecisionCacheBudgetBytes = 300ull * 1024ull * 1024ull;
    uint64_t _decisionCacheBudgetBytes                 = kDecisionCacheBudgetBytes;

    std::atomic_uint32_t _scanActiveScans{0};
    std::atomic_uint64_t _scanFoldersScanned{0};
    std::atomic_uint64_t _scanEntriesScanned{0};
    std::atomic_uint64_t _scanLastNotifyTickMs{0};
    std::atomic<std::shared_ptr<const ScanProgressCallback>> _scanProgressCallback;

    std::deque<FolderScanJob> _scanQueueHigh;
    std::deque<FolderScanJob> _scanQueueLow;
    std::set<std::wstring, WStringViewNoCaseLess> _scanScheduledKeys;
    std::set<std::wstring, WStringViewNoCaseLess> _scanHighQueuedKeys;
    std::set<std::wstring, WStringViewNoCaseLess> _scanInFlightKeys;
    std::set<std::wstring, WStringViewNoCaseLess> _pendingSubdirUpdates;
    std::condition_variable _scanCv;
    std::vector<std::jthread> _scanWorkers;

    std::atomic_uint64_t _contentComparePendingCompares{0};
    std::atomic_uint64_t _contentCompareTotalCompares{0};
    std::atomic_uint64_t _contentCompareCompletedCompares{0};
    std::atomic_uint64_t _contentCompareTotalBytes{0};
    std::atomic_uint64_t _contentCompareCompletedBytes{0};
    std::atomic<std::shared_ptr<const ContentProgressCallback>> _contentProgressCallback;

    std::atomic_uint64_t _decisionUpdatedLastNotifyTickMs{0};
    std::atomic<std::shared_ptr<const DecisionUpdatedCallback>> _decisionUpdatedCallback;

    std::unordered_map<ContentCompareKey, bool, ContentCompareKeyHash, ContentCompareKeyEq> _contentCompareCache;
    std::unordered_map<ContentCompareKey, uint64_t, ContentCompareKeyHash, ContentCompareKeyEq> _contentCompareInFlight;
    std::deque<ContentCompareJob> _contentCompareQueueHigh;
    std::deque<ContentCompareJob> _contentCompareQueueLow;
    std::map<std::wstring, std::map<std::wstring, PendingContentCompareUpdate, WStringViewNoCaseLess>, WStringViewNoCaseLess> _pendingContentCompareUpdates;
    std::condition_variable _contentCompareCv;
    std::condition_variable _contentCompareQueueNotFullCv;
    std::vector<std::jthread> _contentCompareWorkers;

    // Perf stats (best-effort, under _mutex).
    size_t _scanQueueHighWater          = 0;
    size_t _scanQueueHighHighWater      = 0;
    size_t _scanQueueLowHighWater       = 0;
    size_t _scanScheduledHighWater      = 0;
    size_t _scanInFlightHighWater       = 0;
    size_t _pendingSubdirHighWater      = 0;
    size_t _contentQueueHighWater       = 0;
    size_t _contentQueueHighHighWater   = 0;
    size_t _contentQueueLowHighWater    = 0;
    size_t _contentInFlightHighWater    = 0;
    size_t _contentCacheHighWater       = 0;
    size_t _pendingContentHighWater     = 0;
    size_t _decisionCacheEntriesHighWater = 0;
    uint64_t _decisionCacheEstimatedBytes = 0;
    uint64_t _decisionCacheEstimatedBytesHighWater = 0;
};

[[nodiscard]] wil::com_ptr<IFileSystem> CreateCompareDirectoriesFileSystem(ComparePane pane, std::shared_ptr<CompareDirectoriesSession> session) noexcept;
