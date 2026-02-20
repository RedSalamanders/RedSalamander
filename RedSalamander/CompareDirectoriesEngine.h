#pragma once

#include <atomic>
#include <condition_variable>
#include <cstdint>
#include <deque>
#include <filesystem>
#include <functional>
#include <limits>
#include <map>
#include <memory>
#include <mutex>
#include <optional>
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
#pragma warning(pop)

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
        const int leftLen  = static_cast<int>(left.size());
        const int rightLen = static_cast<int>(right.size());
        const int cmp      = CompareStringOrdinal(left.data(), leftLen, right.data(), rightLen, TRUE);
        return cmp == CSTR_LESS_THAN;
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
    // Precomputed aggregates over items — avoids O(n) scans in hot ancestor-propagation paths.
    bool anyDifferent = false;
    bool anyPending   = false;
    std::map<std::wstring, CompareDirectoriesItemDecision, WStringViewNoCaseLess> items;
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

    CompareDirectoriesSession(wil::com_ptr<IFileSystem> baseFileSystem,
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

    void SetScanProgressCallback(ScanProgressCallback callback) noexcept;
    void SetContentProgressCallback(ContentProgressCallback callback) noexcept;
    void SetDecisionUpdatedCallback(DecisionUpdatedCallback callback) noexcept;

    [[nodiscard]] Common::Settings::CompareDirectoriesSettings GetSettings() const;
    [[nodiscard]] std::filesystem::path GetRoot(ComparePane pane) const;
    [[nodiscard]] uint64_t GetVersion() const noexcept;
    [[nodiscard]] uint64_t GetUiVersion() const noexcept;

    [[nodiscard]] wil::com_ptr<IFileSystem> GetBaseFileSystem() const noexcept;
    [[nodiscard]] wil::com_ptr<IInformations> GetBaseInformations() const noexcept;
    [[nodiscard]] wil::com_ptr<IFileSystemIO> GetBaseFileSystemIO() const noexcept;

    [[nodiscard]] std::optional<std::filesystem::path> TryMakeRelative(ComparePane pane, const std::filesystem::path& absoluteFolder) const;
    [[nodiscard]] std::filesystem::path ResolveAbsolute(ComparePane pane, const std::filesystem::path& relativeFolder) const;

    [[nodiscard]] std::shared_ptr<const CompareDirectoriesFolderDecision> GetOrComputeDecision(const std::filesystem::path& relativeFolder);

private:
    struct ContentCompareKey
    {
        std::wstring leftPath;
        std::wstring rightPath;
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
        std::wstring folderKey;
        std::filesystem::path relativeFolder;
        std::wstring entryName;
        ContentCompareKey key;
        std::filesystem::path leftPath;
        std::filesystem::path rightPath;
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
    void EnsureContentCompareWorkersLocked() noexcept;
    struct ResetCleanup final
    {
        std::map<std::wstring, std::shared_ptr<const CompareDirectoriesFolderDecision>, WStringViewNoCaseLess> cache;
        std::unordered_map<ContentCompareKey, uint64_t, ContentCompareKeyHash, ContentCompareKeyEq> contentCompareInFlight;
        std::deque<ContentCompareJob> contentCompareQueue;
        std::map<std::wstring, std::map<std::wstring, PendingContentCompareUpdate, WStringViewNoCaseLess>, WStringViewNoCaseLess> pendingContentCompareUpdates;
    };

    static void ScheduleResetCleanup(std::unique_ptr<ResetCleanup> cleanup) noexcept;
    void ResetCompareStateLocked(ResetCleanup& outCleanup) noexcept;
    void ClearContentCompareStateLocked() noexcept;
    void ApplyPendingContentCompareUpdatesLocked(const std::wstring& folderKey) noexcept;
    void ContentCompareWorker(std::stop_token stopToken, uint32_t workerIndex) noexcept;

    wil::com_ptr<IFileSystem> _baseFileSystem;
    wil::com_ptr<IInformations> _baseInformations;
    wil::com_ptr<IFileSystemIO> _baseFileSystemIo;

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

    std::atomic_uint32_t _scanActiveScans{0};
    std::atomic_uint64_t _scanFoldersScanned{0};
    std::atomic_uint64_t _scanEntriesScanned{0};
    std::atomic_uint64_t _scanLastNotifyTickMs{0};
    std::atomic<std::shared_ptr<const ScanProgressCallback>> _scanProgressCallback;

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
    std::deque<ContentCompareJob> _contentCompareQueue;
    std::map<std::wstring, std::map<std::wstring, PendingContentCompareUpdate, WStringViewNoCaseLess>, WStringViewNoCaseLess> _pendingContentCompareUpdates;
    std::condition_variable _contentCompareCv;
    std::vector<std::jthread> _contentCompareWorkers;
};

[[nodiscard]] wil::com_ptr<IFileSystem> CreateCompareDirectoriesFileSystem(ComparePane pane, std::shared_ptr<CompareDirectoriesSession> session) noexcept;
