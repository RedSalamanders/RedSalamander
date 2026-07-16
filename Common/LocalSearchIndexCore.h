#pragma once

#include <cstdint>
#include <memory>
#include <mutex>
#include <regex>
#include <span>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "PlugInterfaces/FileSystem.h"

namespace LocalSearchIndexCore
{
inline constexpr HRESULT kSkipCandidateHr = MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_ITF, 0x201u);

struct VolumeIndex;

enum class PersistentStoreKind : uint32_t
{
    SnapshotBinary = 1,
    Sqlite         = 2,
};

struct RepositoryOptions final
{
    std::wstring snapshotRootDirectory;
    PersistentStoreKind persistentStoreKind = PersistentStoreKind::SnapshotBinary;
    std::wstring sqliteDatabasePath;
    bool sqliteAuthoritative = false;
};

enum class FileSystemKind : uint32_t
{
    Unsupported = 0,
    Ntfs        = 1,
    Refs        = 2,
};

enum class StoreState : uint32_t
{
    Unknown     = 0,
    Ready       = 1,
    Syncing     = 2,
    Recovering  = 3,
    Invalid     = 4,
    Maintenance = 5,
};

enum class SyncPhase : uint32_t
{
    Idle        = 0,
    Discovering = 1,
    Loading     = 2,
    Replaying   = 3,
    Mirroring   = 4,
    Watching    = 5,
};

enum class QueryExecutionMode : uint32_t
{
    Unknown          = 0,
    Sqlite           = 1,
    InMemoryIndex    = 2,
    LiveScanFallback = 3,
};

enum class FallbackReason : uint32_t
{
    None           = 0,
    StoreMissing   = 1,
    StoreInvalid   = 2,
    StoreStale     = 3,
    CutoverBlocked = 4,
    WarmupRunning  = 5,
    SqliteFailure  = 6,
};

struct SupportInfo final
{
    FileSystemKind fileSystemKind = FileSystemKind::Unsupported;
    bool indexable                = false;
    std::wstring normalizedRootPath;
};

struct PersistentStoreInfo final
{
    PersistentStoreKind kind = PersistentStoreKind::SnapshotBinary;
    std::wstring rootDirectory;
    std::wstring primaryPath;
    uint64_t primaryBytes = 0u;
    std::wstring writeAheadLogPath;
    uint64_t writeAheadLogBytes       = 0u;
    uint64_t pageCount                = 0u;
    uint64_t freelistPageCount        = 0u;
    bool incrementalAutoVacuumEnabled = false;
    std::wstring lastCheckpointUtc;
    std::wstring lastCompactionUtc;
    bool inspectionSucceeded                    = false;
    bool readyForQueryCutover                   = false;
    uint64_t storeGeneration                    = 0u;
    uint64_t indexedVolumeCount                 = 0u;
    uint64_t indexedEntryCount                  = 0u;
    uint64_t legacyImportVolumeCount            = 0u;
    bool autoCheckpointEnabled                  = false;
    uint64_t autoCheckpointTargetBytes          = 0u;
    bool autoCompactionEnabled                  = false;
    uint32_t autoCompactionFragmentationPercent = 0u;
    uint64_t autoCompactionMinBytes             = 0u;
};

struct SqliteMaintenancePolicy final
{
    uint64_t autoCheckpointTargetBytes          = 0u;
    uint32_t autoCompactionFragmentationPercent = 0u;
    uint64_t autoCompactionMinBytes             = 0u;
};

struct QueryPlan final
{
    std::wstring rootPath;
    std::wstring namePattern;
    FileSystemSearchNameMode nameMode    = FILESYSTEM_SEARCH_NAME_DISABLED;
    const std::wregex* compiledNameRegex = nullptr;
    bool matchCaseName                   = false;
    bool recursive                       = false;
    bool includeFiles                    = false;
    bool includeDirectories              = false;
    uint64_t maxResults                  = 0u;
};

enum CandidateMetadataFlags : uint32_t
{
    CANDIDATE_METADATA_NONE             = 0u,
    CANDIDATE_METADATA_END_OF_FILE      = 0x1u,
    CANDIDATE_METADATA_LAST_WRITE_TIME  = 0x2u,
    CANDIDATE_METADATA_CREATION_TIME    = 0x4u,
    CANDIDATE_METADATA_LAST_ACCESS_TIME = 0x8u,
    CANDIDATE_METADATA_CHANGE_TIME      = 0x10u,
    CANDIDATE_METADATA_ALLOCATION_SIZE  = 0x20u,
};

struct Candidate final
{
    std::wstring fullPath;
    std::wstring displayName;
    unsigned long fileAttributes = 0u;
    uint32_t metadataFlags       = CANDIDATE_METADATA_NONE;
    int64_t creationTime100ns    = 0;
    int64_t lastAccessTime100ns  = 0;
    int64_t endOfFile            = 0;
    int64_t lastWriteTime100ns   = 0;
    int64_t changeTime100ns      = 0;
    int64_t allocationSize       = 0;
};

struct QueryStats final
{
    FileSystemKind fileSystemKind         = FileSystemKind::Unsupported;
    bool snapshotLoaded                   = false;
    bool snapshotSaved                    = false;
    bool usedSqliteStore                  = false;
    bool sqliteReadOnlyQuery              = false;
    bool usedLiveScanFallback             = false;
    bool usedNamePrefilter                = false;
    bool sqliteCutoverBlocked             = false;
    bool journalAvailable                 = false;
    bool journalReplayApplied             = false;
    bool rebuiltJournalIdMismatch         = false;
    bool rebuiltJournalRangeInvalid       = false;
    bool rebuiltSnapshotCorruption        = false;
    bool usedNtfsEnumeration              = false;
    bool usedTraversalSeed                = false;
    bool hardlinkAliasCoverageIncomplete  = false;
    uint32_t warningFlags                 = FILESYSTEM_SEARCH_WARNING_NONE;
    uint64_t entryCount                   = 0u;
    uint64_t fileCount                    = 0u;
    uint64_t directoryCount               = 0u;
    uint64_t candidateCount               = 0u;
    uint64_t nextUsn                      = 0u;
    uint64_t journalId                    = 0u;
    uint64_t snapshotFileBytes            = 0u;
    uint64_t estimatedMemoryBytes         = 0u;
    uint32_t ensureReadyDurationMs        = 0u;
    uint32_t executeQueryDurationMs       = 0u;
    QueryExecutionMode queryExecutionMode = QueryExecutionMode::Unknown;
    FallbackReason fallbackReason         = FallbackReason::None;
    std::wstring snapshotPath;
};

using CancelCheckFn       = HRESULT(STDMETHODCALLTYPE*)(void* cookie) noexcept;
using CandidateCallbackFn = HRESULT(STDMETHODCALLTYPE*)(Candidate* candidate, void* cookie) noexcept;

struct ProgressUpdate final
{
    FileSystemSearchPhase phase           = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
    uint32_t warningFlags                 = FILESYSTEM_SEARCH_WARNING_NONE;
    HRESULT statusHint                    = S_OK;
    uint64_t scannedDirectories           = 0u;
    uint64_t scannedFiles                 = 0u;
    uint64_t candidateFiles               = 0u;
    uint64_t matchedEntries               = 0u;
    StoreState storeState                 = StoreState::Unknown;
    SyncPhase syncPhase                   = SyncPhase::Idle;
    QueryExecutionMode queryExecutionMode = QueryExecutionMode::Unknown;
    FallbackReason fallbackReason         = FallbackReason::None;
    uint64_t completedRoots               = 0u;
    uint64_t totalRoots                   = 0u;
    std::wstring activeRoot;
    std::wstring currentPath;
};

using ProgressCallbackFn = HRESULT(STDMETHODCALLTYPE*)(const ProgressUpdate* progress, void* cookie) noexcept;

struct RepositoryStatus final
{
    StoreState storeState                 = StoreState::Unknown;
    SyncPhase syncPhase                   = SyncPhase::Idle;
    QueryExecutionMode queryExecutionMode = QueryExecutionMode::Unknown;
    FallbackReason fallbackReason         = FallbackReason::None;
    uint64_t completedRoots               = 0u;
    uint64_t totalRoots                   = 0u;
    std::wstring activeRoot;
};

[[nodiscard]] std::wstring_view GetPersistentStoreKindText(PersistentStoreKind kind) noexcept;
[[nodiscard]] std::wstring_view GetStoreStateText(StoreState state) noexcept;
[[nodiscard]] std::wstring_view GetSyncPhaseText(SyncPhase phase) noexcept;
[[nodiscard]] std::wstring_view GetQueryExecutionModeText(QueryExecutionMode mode) noexcept;
[[nodiscard]] std::wstring_view GetFallbackReasonText(FallbackReason reason) noexcept;
[[nodiscard]] SqliteMaintenancePolicy GetDefaultSqliteMaintenancePolicy() noexcept;
[[nodiscard]] PersistentStoreInfo GetPersistentStoreInfo(const RepositoryOptions& options) noexcept;
[[nodiscard]] std::vector<std::wstring> DiscoverFixedLocalRoots() noexcept;

#ifdef ENABLE_TESTS
enum class SnapshotCorruptionMode : uint32_t
{
    InvalidMagic,
    JournalIdMismatch,
    NextUsnPastEnd,
    EntryCountTooLarge,
};

struct DirectSqliteFreshnessProbe final
{
    bool storedVolumeReady       = false;
    bool currentJournalKnown     = false;
    bool currentJournalAvailable = false;
    uint64_t storedJournalId     = 0u;
    uint64_t storedNextUsn       = 0u;
    uint64_t currentJournalId    = 0u;
    uint64_t currentFirstUsn     = 0u;
    uint64_t currentNextUsn      = 0u;
};

[[nodiscard]] bool ShouldAllowDirectSqliteQueryForTests(const DirectSqliteFreshnessProbe& probe) noexcept;
[[nodiscard]] bool TryParseUsnRecordForTests(const void* recordBytes, size_t recordBytesSize) noexcept;
[[nodiscard]] bool TryParseFileFullDirectoryInformationForTests(const void* entryBytes, size_t entryBytesSize) noexcept;

struct SyntheticJournalRecordForTests final
{
    std::wstring idPath;
    std::wstring parentPath;
    std::wstring name;
    unsigned long fileAttributes = 0u;
    uint32_t reason              = 0u;
};
#endif

class Repository final
{
public:
    explicit Repository(RepositoryOptions options = {}) noexcept;
    ~Repository()                            = default;
    Repository(const Repository&)            = delete;
    Repository(Repository&&)                 = delete;
    Repository& operator=(const Repository&) = delete;
    Repository& operator=(Repository&&)      = delete;

    HRESULT ProbePath(std::wstring_view rootPath, SupportInfo& outSupport) noexcept;
    HRESULT Query(const QueryPlan& plan,
                  CancelCheckFn cancelCheck,
                  void* cancelCookie,
                  std::vector<Candidate>& outCandidates,
                  QueryStats* outStats,
                  ProgressCallbackFn progressCallback = nullptr,
                  void* progressCookie                = nullptr) noexcept;
    HRESULT Enumerate(const QueryPlan& plan,
                      CancelCheckFn cancelCheck,
                      void* cancelCookie,
                      CandidateCallbackFn candidateCallback,
                      void* candidateCookie,
                      QueryStats* outStats,
                      ProgressCallbackFn progressCallback = nullptr,
                      void* progressCookie                = nullptr) noexcept;
    HRESULT EnumerateNoWait(const QueryPlan& plan,
                            CancelCheckFn cancelCheck,
                            void* cancelCookie,
                            CandidateCallbackFn candidateCallback,
                            void* candidateCookie,
                            QueryStats* outStats,
                            ProgressCallbackFn progressCallback = nullptr,
                            void* progressCookie                = nullptr) noexcept;
    HRESULT PrimeRoot(std::wstring_view rootPath) noexcept;
    void CollectCachedRoots(std::vector<std::wstring>& outRoots) noexcept;
    void GetStatus(RepositoryStatus& outStatus) noexcept;
    bool TryGetCachedPersistentStoreInfo(PersistentStoreInfo& outInfo) noexcept;
    bool TryBuildInMemoryPersistentStoreInfo(PersistentStoreInfo& outInfo) noexcept;
    HRESULT EnsureReadyForRoot(std::wstring_view rootPath,
                               CancelCheckFn cancelCheck,
                               void* cancelCookie,
                               QueryStats* outStats,
                               ProgressCallbackFn progressCallback = nullptr,
                               void* progressCookie                = nullptr) noexcept;
    HRESULT InvalidateRoot(std::wstring_view rootPath, bool deleteSnapshot) noexcept;

#ifdef ENABLE_TESTS
    HRESULT DropCachedVolumeForTests(std::wstring_view rootPath) noexcept;
    HRESULT HydrateRootAndQueryForTests(const QueryPlan& plan, std::vector<Candidate>& outCandidates, QueryStats* outStats) noexcept;
    HRESULT CorruptSnapshotForTests(std::wstring_view rootPath, SnapshotCorruptionMode mode) noexcept;
    HRESULT ApplySyntheticJournalForTests(std::wstring_view rootPath, std::span<const SyntheticJournalRecordForTests> records, QueryStats* outStats) noexcept;
    HRESULT QueryPersistedVolumeForTests(const QueryPlan& plan, std::vector<Candidate>& outCandidates, QueryStats* outStats) noexcept;
    void SetNextJournalStateForTests(uint64_t id, uint64_t firstUsn, uint64_t nextUsn) noexcept;
    void SetNextJournalReplayReadFailureForTests(DWORD error) noexcept;
#endif

private:
    struct PersistentStoreFileStamp final
    {
        uint64_t databaseLastWriteTime = 0u;
        uint64_t databaseBytes         = 0u;
        uint64_t walLastWriteTime      = 0u;
        uint64_t walBytes              = 0u;
        bool databaseExists            = false;
        bool walExists                 = false;

        friend bool operator==(const PersistentStoreFileStamp&, const PersistentStoreFileStamp&) noexcept = default;
    };

    HRESULT AcquireOrCreateVolume(const SupportInfo& support, std::shared_ptr<VolumeIndex>& outVolume) noexcept;
    static bool TryCapturePersistentStoreFileStamp(const PersistentStoreInfo& storeInfo, PersistentStoreFileStamp& outStamp) noexcept;
    PersistentStoreInfo GetCachedPersistentStoreInfo() noexcept;
    PersistentStoreInfo GetValidatedCachedPersistentStoreInfoForQuery() noexcept;
    void RefreshCachedPersistentStoreInfo() noexcept;
    void InvalidateCachedPersistentStoreInfo() noexcept;

    RepositoryOptions _options;
    std::mutex _mutex;
    std::unordered_map<std::wstring, std::shared_ptr<VolumeIndex>> _volumes;
    PersistentStoreInfo _cachedPersistentStoreInfo;
    PersistentStoreFileStamp _cachedPersistentStoreFileStamp;
    bool _cachedPersistentStoreInfoValid      = false;
    bool _cachedPersistentStoreFileStampValid = false;
    std::mutex _statusMutex;
    RepositoryStatus _runtimeStatus;
};
} // namespace LocalSearchIndexCore
