#include "LocalSearchIndexCore.h"

#include "Helpers.h"
#include "SqliteIndexStore.h"

#include <algorithm>
#include <array>
#include <chrono>
#include <cstddef>
#include <cstring>
#include <filesystem>
#include <format>
#include <limits>
#include <optional>
#include <span>
#include <string_view>
#include <tuple>
#include <unordered_set>

#include <winioctl.h>
#include <winternl.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

namespace LocalSearchIndexCore
{
namespace
{
constexpr HRESULT kNotSupportedHr         = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
constexpr uint64_t kProgressIntervalItems = 2048u;
constexpr ULONGLONG kProgressIntervalMs   = 200u;

constexpr uint32_t kSnapshotMagic                            = 0x58444953u; // "SIDX"
constexpr uint32_t kSnapshotVersion                          = 1u;
constexpr uint64_t kSqliteAutoCheckpointTargetBytes          = 128u * 1024u * 1024u;
constexpr uint32_t kSqliteAutoCompactionFragmentationPercent = 15u;
constexpr uint64_t kSqliteAutoCompactionMinBytes             = 64u * 1024u * 1024u;
#ifdef ENABLE_TESTS
constexpr wchar_t kForceNtfsTraversalSeedEnvVar[] = L"REDSALAMANDER_TEST_FORCE_NTFS_TRAVERSAL_SEED";
#endif
constexpr DWORD kJournalReplayReasons = USN_REASON_FILE_CREATE | USN_REASON_FILE_DELETE | USN_REASON_RENAME_OLD_NAME | USN_REASON_RENAME_NEW_NAME |
                                        USN_REASON_BASIC_INFO_CHANGE | USN_REASON_HARD_LINK_CHANGE | USN_REASON_REPARSE_POINT_CHANGE;
constexpr HRESULT kSkipCandidateHr    = MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_ITF, 0x201u);

[[nodiscard]] std::wstring FoldPathKey(std::wstring_view path) noexcept;

inline void EmitPerfCount(std::wstring_view name, uint64_t value = 1u) noexcept
{
    Debug::Perf::Emit(name, L"", 0u, value, 0u, S_OK);
}

[[nodiscard]] std::wstring_view GetFallbackReasonMetricName(const FallbackReason reason) noexcept
{
    switch (reason)
    {
        case FallbackReason::StoreMissing: return L"search.backend.sqlite.fallback_reason.store_missing";
        case FallbackReason::StoreInvalid: return L"search.backend.sqlite.fallback_reason.store_invalid";
        case FallbackReason::StoreStale: return L"search.backend.sqlite.fallback_reason.store_stale";
        case FallbackReason::CutoverBlocked: return L"search.backend.sqlite.fallback_reason.cutover_blocked";
        case FallbackReason::WarmupRunning: return L"search.backend.sqlite.fallback_reason.warmup_running";
        case FallbackReason::SqliteFailure: return L"search.backend.sqlite.fallback_reason.sqlite_failure";
        case FallbackReason::None:
        default: return L"search.backend.sqlite.fallback_reason.none";
    }
}

struct NodeId final
{
    uint64_t low  = 0u;
    uint64_t high = 0u;

    [[nodiscard]] bool IsZero() const noexcept
    {
        return low == 0u && high == 0u;
    }

    friend bool operator==(const NodeId&, const NodeId&) noexcept = default;
};

struct NodeIdHash final
{
    [[nodiscard]] size_t operator()(const NodeId& id) const noexcept
    {
        const uint64_t mixed = id.low ^ (id.high + 0x9E3779B97F4A7C15ull + (id.low << 6u) + (id.low >> 2u));
        return static_cast<size_t>(mixed ^ (mixed >> 32u));
    }
};

struct TrackingCandidateCallbackContext final
{
    CandidateCallbackFn callback                      = nullptr;
    void* cookie                                      = nullptr;
    std::unordered_set<std::wstring>* emittedPathKeys = nullptr;
};

HRESULT STDMETHODCALLTYPE TrackCandidateAndForward(Candidate* candidate, void* cookie) noexcept
{
    if (candidate == nullptr || cookie == nullptr)
    {
        return E_POINTER;
    }

    auto& context = *static_cast<TrackingCandidateCallbackContext*>(cookie);
    if (context.callback == nullptr || context.emittedPathKeys == nullptr)
    {
        return E_POINTER;
    }

    HRESULT hr = context.callback(candidate, context.cookie);
    if (FAILED(hr))
    {
        return hr;
    }

    try
    {
        context.emittedPathKeys->insert(FoldPathKey(candidate->fullPath));
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"LocalSearchIndexCore: failed to track emitted SQLite candidate path.");
        return E_FAIL;
    }

    return hr;
}

struct SnapshotHeader final
{
    uint32_t magic          = kSnapshotMagic;
    uint32_t version        = kSnapshotVersion;
    uint32_t fileSystemKind = static_cast<uint32_t>(FileSystemKind::Unsupported);
    uint32_t reserved       = 0u;
    uint64_t journalId      = 0u;
    uint64_t nextUsn        = 0u;
    uint64_t entryCount     = 0u;
    uint64_t rootIdLow      = 0u;
    uint64_t rootIdHigh     = 0u;
};

struct SnapshotEntryHeader final
{
    uint64_t idLow          = 0u;
    uint64_t idHigh         = 0u;
    uint64_t parentIdLow    = 0u;
    uint64_t parentIdHigh   = 0u;
    uint32_t fileAttributes = 0u;
    uint32_t nameBytes      = 0u;
};

struct JournalState final
{
    bool available    = false;
    uint64_t id       = 0u;
    uint64_t firstUsn = 0u;
    uint64_t nextUsn  = 0u;
};

struct Entry final
{
    NodeId id{};
    NodeId parentId{};
    std::wstring name;
    std::wstring fullPath;
    unsigned long fileAttributes = 0u;
    std::vector<NodeId> children;
};

struct SeedEntry final
{
    NodeId id{};
    NodeId parentId{};
    std::wstring name;
    unsigned long fileAttributes = 0u;
};

struct EnumeratedChild final
{
    std::wstring name;
    std::wstring fullPath;
    unsigned long fileAttributes = 0u;
};

struct UsnRecordData final
{
    NodeId id{};
    NodeId parentId{};
    std::wstring name;
    unsigned long fileAttributes = 0u;
    uint32_t reason              = 0u;
};

struct PersistedEntry final
{
    NodeId id{};
    NodeId parentId{};
    std::wstring fullPath;
    std::wstring name;
    unsigned long fileAttributes = 0u;
};

struct PersistedVolumeState final
{
    FileSystemKind fileSystemKind = FileSystemKind::Unsupported;
    NodeId trackedRootId{};
    uint64_t journalId = 0u;
    uint64_t nextUsn   = 0u;
    uint32_t state     = SqliteIndexStore::kVolumeStateReady;
    std::vector<PersistedEntry> entries;
};

struct JournalDelta final
{
    std::unordered_set<NodeId, NodeIdHash> deletedIds;
    std::unordered_set<NodeId, NodeIdHash> upsertIds;
};

class IIndexStore
{
public:
    virtual ~IIndexStore() = default;

    [[nodiscard]] virtual PersistentStoreKind GetActiveKind() const noexcept = 0;
    [[nodiscard]] virtual std::wstring_view GetPrimaryPath() const noexcept  = 0;
    [[nodiscard]] virtual uint64_t GetPrimaryBytes() const noexcept          = 0;

    virtual HRESULT Load(PersistedVolumeState& outState, QueryStats& stats) noexcept    = 0;
    virtual HRESULT Save(const PersistedVolumeState& state, QueryStats& stats) noexcept = 0;
    virtual HRESULT Delete() noexcept                                                   = 0;

#ifdef ENABLE_TESTS
    virtual HRESULT CorruptForTests(SnapshotCorruptionMode mode) noexcept = 0;
#endif
};

enum class NtFileInformationClass : int
{
    FileDirectoryInformation     = 1,
    FileFullDirectoryInformation = 2,
};

using NtQueryDirectoryFile_t = NTSTATUS(NTAPI*)(HANDLE FileHandle,
                                                HANDLE Event,
                                                PIO_APC_ROUTINE ApcRoutine,
                                                PVOID ApcContext,
                                                PIO_STATUS_BLOCK IoStatusBlock,
                                                PVOID FileInformation,
                                                ULONG Length,
                                                NtFileInformationClass FileInformationClass,
                                                BOOLEAN ReturnSingleEntry,
                                                PUNICODE_STRING FileName,
                                                BOOLEAN RestartScan);

using RtlNtStatusToDosError_t = ULONG(NTAPI*)(NTSTATUS Status);

#ifndef NT_SUCCESS
#define NT_SUCCESS(Status) (((NTSTATUS)(Status)) >= 0)
#endif

#ifndef STATUS_NO_MORE_FILES
#define STATUS_NO_MORE_FILES ((NTSTATUS)0x80000006L)
#endif

} // namespace

struct VolumeIndex final
{
    VolumeIndex()                              = default;
    VolumeIndex(const VolumeIndex&)            = delete;
    VolumeIndex& operator=(const VolumeIndex&) = delete;
    VolumeIndex(VolumeIndex&&)                 = delete;
    VolumeIndex& operator=(VolumeIndex&&)      = delete;

    std::mutex mutex;
    std::wstring normalizedRootPath;
    std::wstring rootKey;
    std::wstring volumeRoot;
    std::wstring volumeDevicePath;
    std::wstring snapshotPath;
    std::unique_ptr<IIndexStore> store;
    FileSystemKind fileSystemKind = FileSystemKind::Unsupported;
    NodeId trackedRootId{};
    bool trackedRootIsDirectory       = false;
    uint64_t journalId                = 0u;
    uint64_t nextUsn                  = 0u;
    bool initialized                  = false;
    uint32_t persistentStoreState     = SqliteIndexStore::kVolumeStateReady;
    bool sqliteMirrorSynchronized     = false;
    uint32_t sqliteMirroredState      = 0u;
    uint64_t sqliteMirroredEntryCount = 0u;
    uint64_t sqliteMirroredJournalId  = 0u;
    uint64_t sqliteMirroredNextUsn    = 0u;
    std::unordered_map<NodeId, Entry, NodeIdHash> entries;
    std::unordered_map<std::wstring, NodeId> pathIndex;
};

namespace
{
struct RepositoryProgressState final
{
    ProgressCallbackFn callback = nullptr;
    void* cookie                = nullptr;
    ProgressUpdate latest{};
    uint64_t workUnits                      = 0u;
    uint64_t lastReportedUnits              = 0u;
    ULONGLONG lastReportedTick              = 0u;
    FileSystemSearchPhase lastReportedPhase = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
    bool hasReported                        = false;
};

void UpdateRepositoryRuntimeStatus(RepositoryStatus* status,
                                   std::mutex* statusMutex,
                                   const StoreState storeState,
                                   const SyncPhase syncPhase,
                                   const QueryExecutionMode queryExecutionMode,
                                   const FallbackReason fallbackReason,
                                   std::wstring_view activeRoot,
                                   const uint64_t completedRoots = 0u,
                                   const uint64_t totalRoots     = 0u) noexcept
{
    if (status == nullptr || statusMutex == nullptr)
    {
        return;
    }

    std::lock_guard guard(*statusMutex);
    status->storeState         = storeState;
    status->syncPhase          = syncPhase;
    status->queryExecutionMode = queryExecutionMode;
    status->fallbackReason     = fallbackReason;
    status->completedRoots     = completedRoots;
    status->totalRoots         = totalRoots;
    status->activeRoot.assign(activeRoot);
    EmitPerfCount(L"search.runtime_status.updates");
}

[[nodiscard]] RepositoryStatus CaptureRepositoryRuntimeStatus(const RepositoryStatus* status, std::mutex* statusMutex) noexcept
{
    if (status == nullptr || statusMutex == nullptr)
    {
        return {};
    }

    std::lock_guard guard(*statusMutex);
    return *status;
}

void ApplyRuntimeStatusToProgress(RepositoryProgressState& progress, const RepositoryStatus& status) noexcept
{
    progress.latest.storeState         = status.storeState;
    progress.latest.syncPhase          = status.syncPhase;
    progress.latest.queryExecutionMode = status.queryExecutionMode;
    progress.latest.fallbackReason     = status.fallbackReason;
    progress.latest.completedRoots     = status.completedRoots;
    progress.latest.totalRoots         = status.totalRoots;
    progress.latest.activeRoot         = status.activeRoot;
}

[[nodiscard]] StoreState GetFallbackStoreState(const FallbackReason reason) noexcept
{
    switch (reason)
    {
        case FallbackReason::StoreMissing:
        case FallbackReason::StoreInvalid: return StoreState::Invalid;
        case FallbackReason::SqliteFailure: return StoreState::Recovering;
        case FallbackReason::StoreStale:
        case FallbackReason::CutoverBlocked:
        case FallbackReason::WarmupRunning: return StoreState::Syncing;
        case FallbackReason::None:
        default: return StoreState::Unknown;
    }
}

[[nodiscard]] FallbackReason ClassifyUninspectableStore(const PersistentStoreInfo& storeInfo) noexcept
{
    if (storeInfo.kind != PersistentStoreKind::Sqlite)
    {
        return FallbackReason::None;
    }

    std::error_code existsEc;
    const bool databaseExists = ! storeInfo.primaryPath.empty() && std::filesystem::exists(storeInfo.primaryPath, existsEc);
    return databaseExists ? FallbackReason::StoreInvalid : FallbackReason::StoreMissing;
}

HRESULT EmitRepositoryProgress(RepositoryProgressState& progress, bool force) noexcept
{
    if (progress.callback == nullptr)
    {
        return S_OK;
    }

    EmitPerfCount(L"search.progress.emit_calls");
    if (force)
    {
        EmitPerfCount(L"search.progress.emit_forced");
    }

    const ULONGLONG now = ::GetTickCount64();
    if (! force && progress.hasReported && progress.latest.phase == progress.lastReportedPhase &&
        (progress.workUnits - progress.lastReportedUnits) < kProgressIntervalItems && (now - progress.lastReportedTick) < kProgressIntervalMs)
    {
        EmitPerfCount(L"search.progress.emit_suppressed");
        return S_OK;
    }

    HRESULT hr = S_OK;
    {
        Debug::Perf::Scope callbackPerf(L"search.progress.callback_ms");
        hr = progress.callback(&progress.latest, progress.cookie);
        callbackPerf.SetHr(hr);
    }
    if (FAILED(hr))
    {
        return hr;
    }

    progress.hasReported       = true;
    progress.lastReportedUnits = progress.workUnits;
    progress.lastReportedTick  = now;
    progress.lastReportedPhase = progress.latest.phase;
    return S_OK;
}

HRESULT ReportRepositoryProgress(
    RepositoryProgressState& progress, FileSystemSearchPhase phase, std::wstring_view currentPath, HRESULT statusHint, bool force) noexcept
{
    progress.latest.phase      = phase;
    progress.latest.statusHint = statusHint;
    progress.latest.currentPath.assign(currentPath);
    return EmitRepositoryProgress(progress, force);
}

HRESULT AdvanceRepositoryProgress(RepositoryProgressState& progress,
                                  FileSystemSearchPhase phase,
                                  std::wstring_view currentPath,
                                  uint64_t scannedDirectoriesDelta,
                                  uint64_t scannedFilesDelta,
                                  uint64_t candidateFilesDelta,
                                  uint64_t matchedEntriesDelta,
                                  uint64_t workUnitsDelta,
                                  bool force = false) noexcept
{
    progress.latest.phase      = phase;
    progress.latest.statusHint = S_OK;
    progress.latest.scannedDirectories += scannedDirectoriesDelta;
    progress.latest.scannedFiles += scannedFilesDelta;
    progress.latest.candidateFiles += candidateFilesDelta;
    progress.latest.matchedEntries += matchedEntriesDelta;
    progress.latest.currentPath.assign(currentPath);
    progress.workUnits += workUnitsDelta;
    return EmitRepositoryProgress(progress, force);
}

void AssignRepositoryProgressCounts(
    RepositoryProgressState& progress, uint64_t scannedDirectories, uint64_t scannedFiles, uint64_t candidateFiles, uint64_t matchedEntries) noexcept
{
    progress.latest.scannedDirectories = scannedDirectories;
    progress.latest.scannedFiles       = scannedFiles;
    progress.latest.candidateFiles     = candidateFiles;
    progress.latest.matchedEntries     = matchedEntries;
}

[[nodiscard]] std::wstring FoldText(std::wstring_view text) noexcept
{
    return OrdinalString::FoldCaseInvariant(text);
}

[[nodiscard]] bool IsDriveRoot(std::wstring_view path) noexcept
{
    return path.size() == 3u && path[1] == L':' && (path[2] == L'\\' || path[2] == L'/');
}

[[nodiscard]] bool IsExtendedDriveRoot(std::wstring_view path) noexcept
{
    return path.size() == 7u && path.rfind(L"\\\\?\\", 0) == 0 && path[5] == L':' && (path[6] == L'\\' || path[6] == L'/');
}

[[nodiscard]] bool IsExtendedUncPath(std::wstring_view path) noexcept
{
    return path.rfind(L"\\\\?\\UNC\\", 0) == 0;
}

[[nodiscard]] bool IsUncPath(std::wstring_view path) noexcept
{
    return path.rfind(L"\\\\", 0) == 0 && ! IsExtendedDriveRoot(path);
}

[[nodiscard]] std::wstring MakeAbsolutePath(std::wstring_view path) noexcept
{
    std::wstring input(path);
    if (input.empty())
    {
        input = L".";
    }

    if (input.rfind(L"\\\\?\\", 0) == 0)
    {
        return input;
    }

    const DWORD required = ::GetFullPathNameW(input.c_str(), 0, nullptr, nullptr);
    if (required == 0)
    {
        return input;
    }

    std::wstring absolute(static_cast<size_t>(required) + 1u, L'\0');
    const DWORD written = ::GetFullPathNameW(input.c_str(), static_cast<DWORD>(absolute.size()), absolute.data(), nullptr);
    if (written == 0)
    {
        return input;
    }

    absolute.resize(static_cast<size_t>(written));
    return absolute;
}

[[nodiscard]] std::wstring NormalizePath(std::wstring_view path) noexcept
{
    std::wstring normalized(path);
    std::replace(normalized.begin(), normalized.end(), L'/', L'\\');
    normalized = MakeAbsolutePath(normalized);
    std::replace(normalized.begin(), normalized.end(), L'/', L'\\');

    while (normalized.size() > 1u && (normalized.back() == L'\\' || normalized.back() == L'/'))
    {
        if (IsDriveRoot(normalized) || IsExtendedDriveRoot(normalized) || normalized == L"\\\\")
        {
            break;
        }

        normalized.pop_back();
    }

    return normalized;
}

[[nodiscard]] std::wstring ToExtendedPath(std::wstring_view path) noexcept
{
    std::wstring normalized(path);
    if (normalized.empty())
    {
        normalized = L".";
    }

    if (normalized.rfind(L"\\\\?\\", 0) != 0)
    {
        normalized = NormalizePath(normalized);
    }

    if (normalized.rfind(L"\\\\?\\", 0) == 0)
    {
        return normalized;
    }

    if (normalized.rfind(L"\\\\", 0) == 0)
    {
        return std::wstring(L"\\\\?\\UNC\\") + normalized.substr(2u);
    }

    return std::wstring(L"\\\\?\\") + normalized;
}

[[nodiscard]] std::wstring AppendPath(std::wstring_view basePath, std::wstring_view leafName) noexcept
{
    if (basePath.empty())
    {
        return std::wstring(leafName);
    }

    std::wstring result(basePath);
    const wchar_t last = result.back();
    if (last != L'\\' && last != L'/')
    {
        result.push_back(L'\\');
    }
    result.append(leafName);
    return result;
}

[[nodiscard]] std::wstring GetPathLeaf(std::wstring_view path) noexcept
{
    const size_t pos = path.find_last_of(L"\\/");
    if (pos == std::wstring_view::npos)
    {
        return std::wstring(path);
    }
    return std::wstring(path.substr(pos + 1u));
}

[[nodiscard]] std::wstring FoldPathKey(std::wstring_view path) noexcept
{
    std::wstring key = NormalizePath(path);
    return OrdinalString::FoldCaseInvariant(key);
}

[[nodiscard]] uint64_t Fnv1a64(std::wstring_view text) noexcept
{
    uint64_t hash = 14695981039346656037ull;
    for (const wchar_t ch : text)
    {
        hash ^= static_cast<uint64_t>(ch);
        hash *= 1099511628211ull;
    }
    return hash;
}

[[nodiscard]] std::wstring GetDefaultSnapshotRootDirectory() noexcept
{
    DWORD required = ::GetEnvironmentVariableW(L"LOCALAPPDATA", nullptr, 0u);
    if (required == 0u)
    {
        return {};
    }

    std::wstring root(static_cast<size_t>(required), L'\0');
    const DWORD written = ::GetEnvironmentVariableW(L"LOCALAPPDATA", root.data(), required);
    if (written == 0u)
    {
        return {};
    }

    if (! root.empty() && root.back() == L'\0')
    {
        root.pop_back();
    }

    if (root.empty())
    {
        return {};
    }

    return AppendPath(AppendPath(root, L"RedSalamander"), L"SearchIndex");
}

[[nodiscard]] std::wstring GetSnapshotRootDirectory(const RepositoryOptions& options) noexcept
{
    if (! options.snapshotRootDirectory.empty())
    {
        return NormalizePath(options.snapshotRootDirectory);
    }

    return GetDefaultSnapshotRootDirectory();
}

[[nodiscard]] std::wstring GetDefaultSqliteDatabasePath(const RepositoryOptions& options) noexcept
{
    if (! options.sqliteDatabasePath.empty())
    {
        return NormalizePath(options.sqliteDatabasePath);
    }

    const std::wstring folder = GetSnapshotRootDirectory(options);
    if (folder.empty())
    {
        return {};
    }

    return AppendPath(folder, L"index-v2.sqlite3");
}

[[nodiscard]] std::wstring BuildSnapshotPath(std::wstring_view normalizedRootPath, FileSystemKind kind, const RepositoryOptions& options) noexcept
{
    const std::wstring folder = GetSnapshotRootDirectory(options);
    if (folder.empty())
    {
        return {};
    }

    const uint64_t hash = Fnv1a64(FoldText(normalizedRootPath));
    return AppendPath(folder, std::format(L"{:08x}-{:016x}.bin", static_cast<uint32_t>(kind), static_cast<unsigned long long>(hash)));
}

[[nodiscard]] bool IsSqliteAuthoritative(const RepositoryOptions& options) noexcept
{
    return options.persistentStoreKind == PersistentStoreKind::Sqlite && options.sqliteAuthoritative;
}

[[nodiscard]] uint32_t ClampDurationMs(const std::chrono::steady_clock::duration& duration) noexcept
{
    const auto ms = std::chrono::duration_cast<std::chrono::milliseconds>(duration).count();
    if (ms <= 0)
    {
        return 0u;
    }

    if (ms >= static_cast<std::chrono::milliseconds::rep>((std::numeric_limits<uint32_t>::max)()))
    {
        return (std::numeric_limits<uint32_t>::max)();
    }

    return static_cast<uint32_t>(ms);
}

[[nodiscard]] uint64_t GetFileBytes(std::wstring_view path) noexcept
{
    if (path.empty())
    {
        return 0u;
    }

    std::error_code ec;
    const auto fileSize = std::filesystem::file_size(std::filesystem::path(path), ec);
    return ec ? 0u : static_cast<uint64_t>(fileSize);
}

[[nodiscard]] uint64_t GetSnapshotFileBytes(std::wstring_view snapshotPath) noexcept
{
    return GetFileBytes(snapshotPath);
}

#ifdef ENABLE_TESTS
[[nodiscard]] bool IsForcedNtfsTraversalSeedEnabled() noexcept
{
    const DWORD required = ::GetEnvironmentVariableW(kForceNtfsTraversalSeedEnvVar, nullptr, 0u);
    if (required <= 1u)
    {
        return false;
    }

    std::wstring value;
    value.resize(required - 1u);
    const DWORD written = ::GetEnvironmentVariableW(kForceNtfsTraversalSeedEnvVar, value.data(), required);
    if (written == 0u || written >= required)
    {
        return false;
    }

    value.resize(written);
    return ! value.empty() && ! OrdinalString::EqualsNoCase(value, L"0") && ! OrdinalString::EqualsNoCase(value, L"false");
}
#endif

[[nodiscard]] PersistedVolumeState CapturePersistedVolumeState(const VolumeIndex& volume)
{
    PersistedVolumeState state{};
    state.fileSystemKind = volume.fileSystemKind;
    state.trackedRootId  = volume.trackedRootId;
    state.journalId      = volume.journalId;
    state.nextUsn        = volume.nextUsn;
    state.state          = volume.persistentStoreState;
    state.entries.reserve(volume.entries.size());

    for (const auto& [id, entry] : volume.entries)
    {
        UNREFERENCED_PARAMETER(id);
        state.entries.push_back({
            .id             = entry.id,
            .parentId       = entry.parentId,
            .fullPath       = entry.fullPath,
            .name           = entry.name,
            .fileAttributes = entry.fileAttributes,
        });
    }

    return state;
}

void ApplyPersistedVolumeState(VolumeIndex& volume, PersistedVolumeState&& state) noexcept
{
    volume.entries.clear();
    volume.trackedRootId        = state.trackedRootId;
    volume.journalId            = state.journalId;
    volume.nextUsn              = state.nextUsn;
    volume.persistentStoreState = state.state;

    for (auto& persistedEntry : state.entries)
    {
        Entry entry{};
        entry.id             = persistedEntry.id;
        entry.parentId       = persistedEntry.parentId;
        entry.fullPath       = std::move(persistedEntry.fullPath);
        entry.name           = std::move(persistedEntry.name);
        entry.fileAttributes = persistedEntry.fileAttributes;
        volume.entries.emplace(entry.id, std::move(entry));
    }
}

[[nodiscard]] uint64_t EstimateVolumeMemoryBytes(const VolumeIndex& volume) noexcept
{
    uint64_t total = sizeof(VolumeIndex);
    total += static_cast<uint64_t>(volume.normalizedRootPath.capacity() * sizeof(wchar_t));
    total += static_cast<uint64_t>(volume.volumeRoot.capacity() * sizeof(wchar_t));
    total += static_cast<uint64_t>(volume.volumeDevicePath.capacity() * sizeof(wchar_t));
    total += static_cast<uint64_t>(volume.snapshotPath.capacity() * sizeof(wchar_t));

    for (const auto& [id, entry] : volume.entries)
    {
        UNREFERENCED_PARAMETER(id);
        total += sizeof(entry);
        total += static_cast<uint64_t>(entry.name.capacity() * sizeof(wchar_t));
        total += static_cast<uint64_t>(entry.fullPath.capacity() * sizeof(wchar_t));
        total += static_cast<uint64_t>(entry.children.capacity() * sizeof(NodeId));
    }

    total += static_cast<uint64_t>(volume.pathIndex.size()) * (sizeof(NodeId) + sizeof(std::wstring));
    for (const auto& [path, id] : volume.pathIndex)
    {
        UNREFERENCED_PARAMETER(id);
        total += static_cast<uint64_t>(path.capacity() * sizeof(wchar_t));
    }

    return total;
}

HRESULT EnsureSnapshotDirectory(std::wstring_view snapshotPath) noexcept
{
    if (snapshotPath.empty())
    {
        return E_INVALIDARG;
    }

    const std::filesystem::path snapshotFile(snapshotPath);
    const std::filesystem::path parent = snapshotFile.parent_path();
    if (parent.empty())
    {
        return E_INVALIDARG;
    }

    std::error_code ec;
    std::filesystem::create_directories(parent, ec);
    if (ec)
    {
        return HRESULT_FROM_WIN32(static_cast<DWORD>(ec.value()));
    }

    return S_OK;
}

class SnapshotVolumeStore final : public IIndexStore
{
public:
    SnapshotVolumeStore(std::wstring snapshotPath, FileSystemKind fileSystemKind) noexcept
        : _snapshotPath(std::move(snapshotPath)),
          _fileSystemKind(fileSystemKind)
    {
    }

    [[nodiscard]] PersistentStoreKind GetActiveKind() const noexcept override
    {
        return PersistentStoreKind::SnapshotBinary;
    }

    [[nodiscard]] std::wstring_view GetPrimaryPath() const noexcept override
    {
        return _snapshotPath;
    }

    [[nodiscard]] uint64_t GetPrimaryBytes() const noexcept override
    {
        return GetSnapshotFileBytes(_snapshotPath);
    }

    HRESULT Load(PersistedVolumeState& outState, QueryStats& stats) noexcept override
    {
        stats.snapshotLoaded    = false;
        stats.snapshotFileBytes = 0u;
        outState                = {};

        wil::unique_handle file(::CreateFileW(_snapshotPath.c_str(),
                                              GENERIC_READ,
                                              FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                              nullptr,
                                              OPEN_EXISTING,
                                              FILE_ATTRIBUTE_NORMAL,
                                              nullptr));
        if (! file)
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        LARGE_INTEGER size{};
        if (::GetFileSizeEx(file.get(), &size) == 0)
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        if (size.QuadPart < static_cast<LONGLONG>(sizeof(SnapshotHeader)))
        {
            stats.rebuiltSnapshotCorruption = true;
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        SnapshotHeader header{};
        DWORD bytesRead = 0u;
        if (::ReadFile(file.get(), &header, sizeof(header), &bytesRead, nullptr) == 0 || bytesRead != sizeof(header))
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        if (header.magic != kSnapshotMagic || header.version != kSnapshotVersion || header.fileSystemKind != static_cast<uint32_t>(_fileSystemKind))
        {
            stats.rebuiltSnapshotCorruption = true;
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        outState.fileSystemKind = _fileSystemKind;
        outState.trackedRootId  = NodeId{header.rootIdLow, header.rootIdHigh};
        outState.journalId      = header.journalId;
        outState.nextUsn        = header.nextUsn;
        outState.state          = SqliteIndexStore::kVolumeStateReady;
        outState.entries.clear();
        outState.entries.reserve(static_cast<size_t>(header.entryCount));

        for (uint64_t index = 0u; index < header.entryCount; ++index)
        {
            SnapshotEntryHeader entryHeader{};
            if (::ReadFile(file.get(), &entryHeader, sizeof(entryHeader), &bytesRead, nullptr) == 0 || bytesRead != sizeof(entryHeader))
            {
                stats.rebuiltSnapshotCorruption = true;
                outState.entries.clear();
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }

            if ((entryHeader.nameBytes % sizeof(wchar_t)) != 0u)
            {
                stats.rebuiltSnapshotCorruption = true;
                outState.entries.clear();
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }

            std::wstring name;
            if (entryHeader.nameBytes != 0u)
            {
                name.resize(entryHeader.nameBytes / sizeof(wchar_t));
                if (::ReadFile(file.get(), name.data(), entryHeader.nameBytes, &bytesRead, nullptr) == 0 || bytesRead != entryHeader.nameBytes)
                {
                    stats.rebuiltSnapshotCorruption = true;
                    outState.entries.clear();
                    return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                }
            }

            outState.entries.push_back({
                .id             = NodeId{entryHeader.idLow, entryHeader.idHigh},
                .parentId       = NodeId{entryHeader.parentIdLow, entryHeader.parentIdHigh},
                .name           = std::move(name),
                .fileAttributes = entryHeader.fileAttributes,
            });
        }

        stats.snapshotLoaded    = true;
        stats.snapshotFileBytes = GetPrimaryBytes();
        return S_OK;
    }

    HRESULT Save(const PersistedVolumeState& state, QueryStats& stats) noexcept override
    {
        HRESULT hr = EnsureSnapshotDirectory(_snapshotPath);
        if (FAILED(hr))
        {
            return hr;
        }

        wil::unique_handle file(::CreateFileW(_snapshotPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (! file)
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        SnapshotHeader header{};
        header.fileSystemKind = static_cast<uint32_t>(state.fileSystemKind);
        header.journalId      = state.journalId;
        header.nextUsn        = state.nextUsn;
        header.entryCount     = static_cast<uint64_t>(state.entries.size());
        header.rootIdLow      = state.trackedRootId.low;
        header.rootIdHigh     = state.trackedRootId.high;

        DWORD written = 0u;
        if (::WriteFile(file.get(), &header, sizeof(header), &written, nullptr) == 0 || written != sizeof(header))
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        for (const PersistedEntry& entry : state.entries)
        {
            SnapshotEntryHeader entryHeader{};
            entryHeader.idLow          = entry.id.low;
            entryHeader.idHigh         = entry.id.high;
            entryHeader.parentIdLow    = entry.parentId.low;
            entryHeader.parentIdHigh   = entry.parentId.high;
            entryHeader.fileAttributes = entry.fileAttributes;
            entryHeader.nameBytes      = static_cast<uint32_t>(entry.name.size() * sizeof(wchar_t));

            if (::WriteFile(file.get(), &entryHeader, sizeof(entryHeader), &written, nullptr) == 0 || written != sizeof(entryHeader))
            {
                return HRESULT_FROM_WIN32(::GetLastError());
            }

            if (! entry.name.empty())
            {
                const DWORD nameBytes = static_cast<DWORD>(entry.name.size() * sizeof(wchar_t));
                if (::WriteFile(file.get(), entry.name.data(), nameBytes, &written, nullptr) == 0 || written != nameBytes)
                {
                    return HRESULT_FROM_WIN32(::GetLastError());
                }
            }
        }

        stats.snapshotSaved     = true;
        stats.snapshotFileBytes = GetPrimaryBytes();
        return S_OK;
    }

    HRESULT Delete() noexcept override
    {
        if (_snapshotPath.empty())
        {
            return E_INVALIDARG;
        }

        std::error_code ec;
        static_cast<void>(std::filesystem::remove(std::filesystem::path(_snapshotPath), ec));
        return S_OK;
    }

#ifdef ENABLE_TESTS
    HRESULT CorruptForTests(SnapshotCorruptionMode mode) noexcept override
    {
        wil::unique_handle file(
            ::CreateFileW(_snapshotPath.c_str(), GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (! file)
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        SnapshotHeader header{};
        DWORD bytesRead = 0u;
        if (::ReadFile(file.get(), &header, sizeof(header), &bytesRead, nullptr) == 0 || bytesRead != sizeof(header))
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        switch (mode)
        {
            case SnapshotCorruptionMode::InvalidMagic: header.magic ^= 0x13579BDFu; break;
            case SnapshotCorruptionMode::JournalIdMismatch: header.journalId = header.journalId + 1u; break;
            case SnapshotCorruptionMode::NextUsnPastEnd: header.nextUsn = (std::numeric_limits<uint64_t>::max)(); break;
        }

        LARGE_INTEGER zero{};
        if (::SetFilePointerEx(file.get(), zero, nullptr, FILE_BEGIN) == 0)
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        DWORD written = 0u;
        if (::WriteFile(file.get(), &header, sizeof(header), &written, nullptr) == 0 || written != sizeof(header))
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        return S_OK;
    }
#endif

private:
    std::wstring _snapshotPath;
    FileSystemKind _fileSystemKind = FileSystemKind::Unsupported;
};

HRESULT OpenPathHandle(const std::wstring& path, wil::unique_handle& outHandle) noexcept;

void UpdateSqliteMirrorState(VolumeIndex& volume, const SqliteIndexStore::ReplaceVolumeRequest& request) noexcept
{
    volume.sqliteMirrorSynchronized = true;
    volume.sqliteMirroredState      = request.state;
    volume.sqliteMirroredEntryCount = request.entries.size();
    volume.sqliteMirroredJournalId  = request.journalId;
    volume.sqliteMirroredNextUsn    = request.nextUsn;
}

[[nodiscard]] bool TryBuildSqliteImportedEntry(const Entry& entry, SqliteIndexStore::ImportedEntry& outEntry) noexcept
{
    outEntry = {
        .fileIdLow    = entry.id.low,
        .fileIdHigh   = entry.id.high,
        .parentIdLow  = entry.parentId.low,
        .parentIdHigh = entry.parentId.high,
        .fullPath     = entry.fullPath,
        .name         = entry.name,
        .attributes   = entry.fileAttributes,
    };

    if (entry.fullPath.empty())
    {
        return false;
    }

    bool metadataComplete = true;
    wil::unique_handle handle;
    if (SUCCEEDED(OpenPathHandle(entry.fullPath, handle)) && handle)
    {
        FILE_BASIC_INFO basic{};
        if (::GetFileInformationByHandleEx(handle.get(), FileBasicInfo, &basic, sizeof(basic)) != 0)
        {
            outEntry.creationTime100ns   = static_cast<uint64_t>(basic.CreationTime.QuadPart);
            outEntry.lastAccessTime100ns = static_cast<uint64_t>(basic.LastAccessTime.QuadPart);
            outEntry.writeTime100ns      = static_cast<uint64_t>(basic.LastWriteTime.QuadPart);
            outEntry.changeTime100ns     = static_cast<uint64_t>(basic.ChangeTime.QuadPart);
            if (basic.FileAttributes != 0u)
            {
                outEntry.attributes = basic.FileAttributes;
            }
        }
        else
        {
            metadataComplete = false;
        }

        FILE_STANDARD_INFO standard{};
        if (::GetFileInformationByHandleEx(handle.get(), FileStandardInfo, &standard, sizeof(standard)) != 0)
        {
            outEntry.sizeBytes      = static_cast<uint64_t>(standard.EndOfFile.QuadPart);
            outEntry.allocationSize = static_cast<uint64_t>(standard.AllocationSize.QuadPart);
        }
        else
        {
            metadataComplete = false;
        }
    }
    else
    {
        metadataComplete = false;
    }

    return metadataComplete;
}

void BuildSqliteReplaceRequest(const VolumeIndex& volume,
                               const uint32_t desiredState,
                               SqliteIndexStore::ReplaceVolumeRequest& outRequest,
                               bool& outMetadataComplete)
{
    outRequest                = {};
    outRequest.rootPath       = volume.normalizedRootPath;
    outRequest.fileSystemKind = volume.fileSystemKind;
    outRequest.journalId      = volume.journalId;
    outRequest.nextUsn        = volume.nextUsn;
    outRequest.entries.reserve(volume.entries.size());
    outMetadataComplete = true;

    for (const auto& [id, entry] : volume.entries)
    {
        UNREFERENCED_PARAMETER(id);
        SqliteIndexStore::ImportedEntry imported{};
        outMetadataComplete = TryBuildSqliteImportedEntry(entry, imported) && outMetadataComplete;
        outRequest.entries.push_back(std::move(imported));
    }

    outRequest.state = outMetadataComplete ? desiredState : SqliteIndexStore::kVolumeStateImportedLegacySnapshot;
}

void BuildSqliteApplyJournalDeltaRequest(const VolumeIndex& volume, const JournalDelta& delta, SqliteIndexStore::ApplyJournalDeltaRequest& outRequest)
{
    outRequest                = {};
    outRequest.rootPath       = volume.normalizedRootPath;
    outRequest.fileSystemKind = volume.fileSystemKind;
    outRequest.journalId      = volume.journalId;
    outRequest.nextUsn        = volume.nextUsn;
    outRequest.state          = volume.persistentStoreState;
    outRequest.deletedEntries.reserve(delta.deletedIds.size());
    outRequest.upsertEntries.reserve(delta.upsertIds.size());
    outRequest.seedEntriesIfMissing.reserve(volume.entries.size());
    bool metadataComplete     = true;
    bool seedMetadataComplete = true;

    for (const auto& [id, entry] : volume.entries)
    {
        UNREFERENCED_PARAMETER(id);
        SqliteIndexStore::ImportedEntry imported{};
        seedMetadataComplete = TryBuildSqliteImportedEntry(entry, imported) && seedMetadataComplete;
        outRequest.seedEntriesIfMissing.push_back(std::move(imported));
    }

    for (const NodeId& deletedId : delta.deletedIds)
    {
        outRequest.deletedEntries.push_back({
            .fileIdLow  = deletedId.low,
            .fileIdHigh = deletedId.high,
        });
    }

    for (const NodeId& upsertId : delta.upsertIds)
    {
        const auto it = volume.entries.find(upsertId);
        if (it == volume.entries.end())
        {
            continue;
        }

        SqliteIndexStore::ImportedEntry imported{};
        metadataComplete = TryBuildSqliteImportedEntry(it->second, imported) && metadataComplete;
        outRequest.upsertEntries.push_back(std::move(imported));
    }

    if (! metadataComplete)
    {
        outRequest.state = SqliteIndexStore::kVolumeStateImportedLegacySnapshot;
    }

    outRequest.seedStateIfMissing = seedMetadataComplete ? outRequest.state : SqliteIndexStore::kVolumeStateImportedLegacySnapshot;
}

class SqliteVolumeStore final : public IIndexStore
{
public:
    SqliteVolumeStore(std::wstring databasePath, std::wstring normalizedRootPath, FileSystemKind fileSystemKind) noexcept
        : _databasePath(std::move(databasePath)),
          _normalizedRootPath(std::move(normalizedRootPath)),
          _fileSystemKind(fileSystemKind)
    {
    }

    [[nodiscard]] PersistentStoreKind GetActiveKind() const noexcept override
    {
        return PersistentStoreKind::Sqlite;
    }

    [[nodiscard]] std::wstring_view GetPrimaryPath() const noexcept override
    {
        return _databasePath;
    }

    [[nodiscard]] uint64_t GetPrimaryBytes() const noexcept override
    {
        return GetFileBytes(_databasePath);
    }

    HRESULT Load(PersistedVolumeState& outState, QueryStats& stats) noexcept override
    {
        stats.snapshotLoaded    = false;
        stats.snapshotSaved     = false;
        stats.snapshotFileBytes = 0u;
        outState                = {};

        SqliteIndexStore::ReplaceVolumeRequest volume{};
        const HRESULT hr = SqliteIndexStore::LoadVolume(_databasePath, _normalizedRootPath, volume);
        if (FAILED(hr))
        {
            return hr;
        }

        outState.fileSystemKind = volume.fileSystemKind;
        outState.journalId      = volume.journalId;
        outState.nextUsn        = volume.nextUsn;
        outState.state          = volume.state;
        outState.entries.reserve(volume.entries.size());

        const std::wstring foldedRootPath = FoldPathKey(_normalizedRootPath);
        bool rootFound                    = false;
        for (const SqliteIndexStore::ImportedEntry& entry : volume.entries)
        {
            PersistedEntry persisted{
                .id             = NodeId{entry.fileIdLow, entry.fileIdHigh},
                .parentId       = NodeId{entry.parentIdLow, entry.parentIdHigh},
                .fullPath       = entry.fullPath,
                .name           = entry.name,
                .fileAttributes = entry.attributes,
            };
            if (! rootFound && FoldPathKey(entry.fullPath) == foldedRootPath)
            {
                outState.trackedRootId = persisted.id;
                rootFound              = true;
            }

            outState.entries.push_back(std::move(persisted));
        }

        if (! rootFound)
        {
            outState = {};
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        return S_OK;
    }

    HRESULT Save(const PersistedVolumeState& state, QueryStats& stats) noexcept override
    {
        SqliteIndexStore::ReplaceVolumeRequest request{};
        request.rootPath       = _normalizedRootPath;
        request.fileSystemKind = state.fileSystemKind;
        request.journalId      = state.journalId;
        request.nextUsn        = state.nextUsn;
        request.entries.reserve(state.entries.size());
        bool metadataComplete = true;

        for (const PersistedEntry& entry : state.entries)
        {
            Entry transient{};
            transient.id             = entry.id;
            transient.parentId       = entry.parentId;
            transient.fullPath       = entry.fullPath;
            transient.name           = entry.name;
            transient.fileAttributes = entry.fileAttributes;

            SqliteIndexStore::ImportedEntry imported{};
            metadataComplete = TryBuildSqliteImportedEntry(transient, imported) && metadataComplete;
            request.entries.push_back(std::move(imported));
        }

        request.state    = metadataComplete ? state.state : SqliteIndexStore::kVolumeStateImportedLegacySnapshot;
        const HRESULT hr = SqliteIndexStore::ReplaceVolume(_databasePath, request, nullptr);
        if (SUCCEEDED(hr))
        {
            stats.snapshotLoaded    = false;
            stats.snapshotSaved     = false;
            stats.snapshotFileBytes = 0u;
            stats.snapshotPath.clear();
        }
        return hr;
    }

    HRESULT Delete() noexcept override
    {
        return SqliteIndexStore::DeleteVolume(_databasePath, _normalizedRootPath);
    }

#ifdef ENABLE_TESTS
    HRESULT CorruptForTests(SnapshotCorruptionMode mode) noexcept override
    {
        UNREFERENCED_PARAMETER(mode);
        return E_NOTIMPL;
    }
#endif

private:
    std::wstring _databasePath;
    std::wstring _normalizedRootPath;
    FileSystemKind _fileSystemKind = FileSystemKind::Unsupported;
};

[[nodiscard]] std::unique_ptr<IIndexStore> CreateVolumeStore(std::wstring_view normalizedRootPath,
                                                             FileSystemKind fileSystemKind,
                                                             const RepositoryOptions& options)
{
    if (IsSqliteAuthoritative(options))
    {
        const std::wstring databasePath = GetDefaultSqliteDatabasePath(options);
        if (databasePath.empty())
        {
            return {};
        }

        return std::make_unique<SqliteVolumeStore>(databasePath, std::wstring(normalizedRootPath), fileSystemKind);
    }

    const std::wstring snapshotPath = BuildSnapshotPath(normalizedRootPath, fileSystemKind, options);
    if (snapshotPath.empty())
    {
        return {};
    }

    return std::make_unique<SnapshotVolumeStore>(snapshotPath, fileSystemKind);
}

HRESULT MirrorVolumeToConfiguredSqliteStore(VolumeIndex& volume, const RepositoryOptions& options, bool& outStoreChanged) noexcept
{
    outStoreChanged = false;
    if (options.persistentStoreKind != PersistentStoreKind::Sqlite)
    {
        return S_OK;
    }

    const std::wstring databasePath = GetDefaultSqliteDatabasePath(options);
    if (databasePath.empty())
    {
        return E_INVALIDARG;
    }

    const bool readyMirrorUpToDate = volume.sqliteMirrorSynchronized && volume.sqliteMirroredState == volume.persistentStoreState &&
                                     volume.sqliteMirroredEntryCount == volume.entries.size() && volume.sqliteMirroredJournalId == volume.journalId &&
                                     volume.sqliteMirroredNextUsn == volume.nextUsn;
    if (readyMirrorUpToDate)
    {
        return S_FALSE;
    }

    SqliteIndexStore::ReplaceVolumeRequest request{};
    bool metadataComplete = true;
    BuildSqliteReplaceRequest(volume, volume.persistentStoreState, request, metadataComplete);
    const HRESULT hr = SqliteIndexStore::ReplaceVolume(databasePath, request, nullptr);
    if (SUCCEEDED(hr))
    {
        UpdateSqliteMirrorState(volume, request);
        outStoreChanged = true;
    }
    return hr;
}

HRESULT SaveSnapshot(const VolumeIndex& volume, QueryStats& stats) noexcept;

HRESULT PersistInitialVolumeSeed(
    VolumeIndex& volume, const RepositoryOptions& options, bool snapshotMissingOnStart, QueryStats& stats, bool& outSqliteStoreChanged) noexcept
{
    outSqliteStoreChanged = false;

    if (IsSqliteAuthoritative(options))
    {
        return SaveSnapshot(volume, stats);
    }

    if (options.persistentStoreKind == PersistentStoreKind::Sqlite && snapshotMissingOnStart)
    {
        const HRESULT sqliteHr = MirrorVolumeToConfiguredSqliteStore(volume, options, outSqliteStoreChanged);
        if (FAILED(sqliteHr))
        {
            Debug::Warning(L"LocalSearchIndexCore: initial SQLite sidecar mirror failed for root='{}'. hr=0x{:08X}",
                           volume.normalizedRootPath,
                           static_cast<unsigned long>(sqliteHr));
        }
    }

    return SaveSnapshot(volume, stats);
}

HRESULT CheckCancelled(CancelCheckFn cancelCheck, void* cookie) noexcept
{
    if (cancelCheck == nullptr)
    {
        return S_OK;
    }

    const HRESULT hr = cancelCheck(cookie);
    return FAILED(hr) ? hr : S_OK;
}

[[nodiscard]] NodeId NodeIdFromFileId128(const FILE_ID_128& fileId) noexcept
{
    NodeId id{};
    static_assert(sizeof(fileId.Identifier) >= sizeof(id), "FILE_ID_128 is smaller than expected.");
    std::memcpy(&id, fileId.Identifier, sizeof(id));
    return id;
}

[[nodiscard]] NodeId NodeIdFromUint64(uint64_t value) noexcept
{
    return NodeId{value, 0u};
}

[[nodiscard]] bool IsDirectoryAttributes(unsigned long attributes) noexcept
{
    return (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u;
}

HRESULT OpenPathHandle(const std::wstring& path, wil::unique_handle& outHandle) noexcept;

NtQueryDirectoryFile_t GetNtQueryDirectoryFile() noexcept
{
    static const NtQueryDirectoryFile_t fn = []() noexcept -> NtQueryDirectoryFile_t
    {
        HMODULE ntdll = ::GetModuleHandleW(L"ntdll.dll");
        if (! ntdll)
        {
            return nullptr;
        }

#pragma warning(push)
#pragma warning(disable : 4191)
        return reinterpret_cast<NtQueryDirectoryFile_t>(::GetProcAddress(ntdll, "NtQueryDirectoryFile"));
#pragma warning(pop)
    }();

    return fn;
}

RtlNtStatusToDosError_t GetRtlNtStatusToDosError() noexcept
{
    static const RtlNtStatusToDosError_t fn = []() noexcept -> RtlNtStatusToDosError_t
    {
        HMODULE ntdll = ::GetModuleHandleW(L"ntdll.dll");
        if (! ntdll)
        {
            return nullptr;
        }

#pragma warning(push)
#pragma warning(disable : 4191)
        return reinterpret_cast<RtlNtStatusToDosError_t>(::GetProcAddress(ntdll, "RtlNtStatusToDosError"));
#pragma warning(pop)
    }();

    return fn;
}

HRESULT OpenVolumeHandle(const std::wstring& volumeDevicePath, wil::unique_handle& outHandle) noexcept
{
    constexpr std::array<DWORD, 3> desiredAccesses = {{GENERIC_READ, FILE_READ_ATTRIBUTES, 0u}};
    for (const DWORD desiredAccess : desiredAccesses)
    {
        outHandle.reset(::CreateFileW(volumeDevicePath.c_str(),
                                      desiredAccess,
                                      FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                      nullptr,
                                      OPEN_EXISTING,
                                      FILE_ATTRIBUTE_NORMAL,
                                      nullptr));
        if (outHandle)
        {
            return S_OK;
        }
    }

    return HRESULT_FROM_WIN32(::GetLastError());
}

HRESULT OpenPathHandle(const std::wstring& path, wil::unique_handle& outHandle) noexcept
{
    outHandle.reset(::CreateFileW(ToExtendedPath(path).c_str(),
                                  FILE_READ_ATTRIBUTES,
                                  FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                  nullptr,
                                  OPEN_EXISTING,
                                  FILE_FLAG_BACKUP_SEMANTICS,
                                  nullptr));
    if (! outHandle)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    return S_OK;
}

HRESULT GetPathNodeId(const std::wstring& path, NodeId& outId) noexcept
{
    outId = {};

    wil::unique_handle handle;
    HRESULT hr = OpenPathHandle(path, handle);
    if (FAILED(hr))
    {
        return hr;
    }

    FILE_ID_INFO info{};
    if (::GetFileInformationByHandleEx(handle.get(), FileIdInfo, &info, sizeof(info)) == 0)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    outId = NodeIdFromFileId128(info.FileId);
    return S_OK;
}

HRESULT GetJournalState(std::wstring_view volumeDevicePath, JournalState& outState) noexcept
{
    outState = {};

    wil::unique_handle handle;
    HRESULT hr = OpenVolumeHandle(std::wstring(volumeDevicePath), handle);
    if (FAILED(hr))
    {
        return hr;
    }

    USN_JOURNAL_DATA_V0 journalData{};
    DWORD bytesReturned = 0u;
    if (! ::DeviceIoControl(handle.get(), FSCTL_QUERY_USN_JOURNAL, nullptr, 0u, &journalData, sizeof(journalData), &bytesReturned, nullptr))
    {
        const DWORD error = ::GetLastError();
        if (error == ERROR_JOURNAL_NOT_ACTIVE || error == ERROR_INVALID_FUNCTION)
        {
            return kNotSupportedHr;
        }

        return HRESULT_FROM_WIN32(error);
    }

    outState.available = true;
    outState.id        = static_cast<uint64_t>(journalData.UsnJournalID);
    outState.firstUsn  = static_cast<uint64_t>(journalData.FirstUsn);
    outState.nextUsn   = static_cast<uint64_t>(journalData.NextUsn);
    return S_OK;
}

HRESULT GetJournalState(const VolumeIndex& volume, JournalState& outState) noexcept
{
    return GetJournalState(volume.volumeDevicePath, outState);
}

[[nodiscard]] bool TryParseUsnRecord(const USN_RECORD_COMMON_HEADER* header, UsnRecordData& out) noexcept
{
    out = {};
    if (header == nullptr)
    {
        return false;
    }

    if (header->MajorVersion == 2u)
    {
        const auto* record     = reinterpret_cast<const USN_RECORD_V2*>(header);
        const size_t nameChars = static_cast<size_t>(record->FileNameLength / sizeof(wchar_t));
        out.id                 = NodeIdFromUint64(static_cast<uint64_t>(record->FileReferenceNumber));
        out.parentId           = NodeIdFromUint64(static_cast<uint64_t>(record->ParentFileReferenceNumber));
        out.name.assign(record->FileName, nameChars);
        out.fileAttributes = record->FileAttributes;
        out.reason         = record->Reason;
        return true;
    }

    if (header->MajorVersion == 3u)
    {
        const auto* record     = reinterpret_cast<const USN_RECORD_V3*>(header);
        const size_t nameChars = static_cast<size_t>(record->FileNameLength / sizeof(wchar_t));
        out.id                 = NodeIdFromFileId128(record->FileReferenceNumber);
        out.parentId           = NodeIdFromFileId128(record->ParentFileReferenceNumber);
        out.name.assign(record->FileName, nameChars);
        out.fileAttributes = record->FileAttributes;
        out.reason         = record->Reason;
        return true;
    }

    return false;
}

[[nodiscard]] std::wstring ExtractVolumeRoot(const std::wstring& normalizedRootPath) noexcept
{
    std::array<wchar_t, MAX_PATH> buffer{};
    if (::GetVolumePathNameW(normalizedRootPath.c_str(), buffer.data(), static_cast<DWORD>(buffer.size())) == 0)
    {
        return {};
    }

    std::wstring volumeRoot(buffer.data());
    std::replace(volumeRoot.begin(), volumeRoot.end(), L'/', L'\\');
    return volumeRoot;
}

[[nodiscard]] std::wstring BuildVolumeDevicePath(std::wstring_view volumeRoot) noexcept
{
    if (volumeRoot.size() < 2u || volumeRoot[1] != L':')
    {
        return {};
    }

    return std::wstring(L"\\\\.\\") + std::wstring(volumeRoot.substr(0u, 2u));
}

[[nodiscard]] bool EqualsCaseInsensitive(std::wstring_view left, std::wstring_view right) noexcept
{
    return FoldText(left) == FoldText(right);
}

[[nodiscard]] std::vector<std::wstring> SplitRelativeComponents(std::wstring_view basePath, std::wstring_view childPath) noexcept
{
    std::vector<std::wstring> components;
    if (childPath.size() < basePath.size())
    {
        return components;
    }

    std::wstring_view remainder = childPath.substr(basePath.size());
    while (! remainder.empty() && (remainder.front() == L'\\' || remainder.front() == L'/'))
    {
        remainder.remove_prefix(1u);
    }

    while (! remainder.empty())
    {
        const size_t separator = remainder.find_first_of(L"\\/");
        if (separator == std::wstring_view::npos)
        {
            components.emplace_back(remainder);
            break;
        }

        components.emplace_back(remainder.substr(0u, separator));
        remainder.remove_prefix(separator + 1u);
        while (! remainder.empty() && (remainder.front() == L'\\' || remainder.front() == L'/'))
        {
            remainder.remove_prefix(1u);
        }
    }

    return components;
}

HRESULT PopulateSupportInfo(std::wstring_view rootPath, SupportInfo& outSupport) noexcept
{
    outSupport = {};

    const std::wstring normalized = NormalizePath(rootPath);
    if (normalized.empty())
    {
        return E_INVALIDARG;
    }

    outSupport.normalizedRootPath = normalized;

    if (IsUncPath(normalized) || IsExtendedUncPath(normalized))
    {
        return S_OK;
    }

    const std::wstring volumeRoot = ExtractVolumeRoot(normalized);
    if (volumeRoot.empty())
    {
        return S_OK;
    }

    std::array<wchar_t, 64> fileSystemName{};
    if (::GetVolumeInformationW(volumeRoot.c_str(), nullptr, 0u, nullptr, nullptr, nullptr, fileSystemName.data(), static_cast<DWORD>(fileSystemName.size())) ==
        0)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    const std::wstring fileSystem = FoldText(fileSystemName.data());
    if (fileSystem == L"ntfs")
    {
        outSupport.fileSystemKind = FileSystemKind::Ntfs;
        outSupport.indexable      = true;
    }
    else if (fileSystem == L"refs")
    {
        outSupport.fileSystemKind = FileSystemKind::Refs;
        outSupport.indexable      = true;
    }

    return S_OK;
}

[[nodiscard]] std::vector<std::wstring> DiscoverFixedLocalRootsImpl() noexcept
{
    std::vector<std::wstring> roots;

    const DWORD requiredChars = ::GetLogicalDriveStringsW(0u, nullptr);
    if (requiredChars == 0u)
    {
        return roots;
    }

    std::wstring driveBuffer(static_cast<size_t>(requiredChars), L'\0');
    const DWORD writtenChars = ::GetLogicalDriveStringsW(requiredChars, driveBuffer.data());
    if (writtenChars == 0u || writtenChars > requiredChars)
    {
        return roots;
    }

    std::unordered_set<std::wstring> seenRoots;
    const wchar_t* cursor = driveBuffer.c_str();
    while (*cursor != L'\0')
    {
        std::wstring driveRoot(cursor);
        cursor += driveRoot.size() + 1u;

        if (::GetDriveTypeW(driveRoot.c_str()) != DRIVE_FIXED)
        {
            continue;
        }

        SupportInfo support{};
        const HRESULT hr = PopulateSupportInfo(driveRoot, support);
        if (FAILED(hr) || ! support.indexable)
        {
            continue;
        }

        std::wstring discoveredRoot = ExtractVolumeRoot(support.normalizedRootPath);
        if (discoveredRoot.empty())
        {
            discoveredRoot = support.normalizedRootPath;
        }

        const std::wstring dedupeKey = FoldPathKey(discoveredRoot);
        if (! seenRoots.insert(dedupeKey).second)
        {
            continue;
        }

        roots.push_back(std::move(discoveredRoot));
    }

    std::sort(roots.begin(), roots.end(), [](const std::wstring& left, const std::wstring& right) noexcept { return FoldPathKey(left) < FoldPathKey(right); });
    return roots;
}

void PopulateStatsFromVolume(const VolumeIndex& volume, QueryStats& stats) noexcept
{
    stats.fileSystemKind           = volume.fileSystemKind;
    stats.nextUsn                  = volume.nextUsn;
    stats.journalId                = volume.journalId;
    const bool snapshotBackedStore = volume.store != nullptr && volume.store->GetActiveKind() == PersistentStoreKind::SnapshotBinary;
    stats.snapshotFileBytes        = snapshotBackedStore ? volume.store->GetPrimaryBytes() : 0u;
    if (stats.snapshotFileBytes == 0u && ! stats.snapshotLoaded && ! stats.snapshotSaved)
    {
        stats.snapshotPath.clear();
    }
    else
    {
        stats.snapshotPath = snapshotBackedStore ? volume.snapshotPath : std::wstring{};
    }
    stats.estimatedMemoryBytes = EstimateVolumeMemoryBytes(volume);
    stats.entryCount           = 0u;
    stats.fileCount            = 0u;
    stats.directoryCount       = 0u;

    for (const auto& [id, entry] : volume.entries)
    {
        if (id == volume.trackedRootId && volume.trackedRootIsDirectory)
        {
            continue;
        }

        ++stats.entryCount;
        if (IsDirectoryAttributes(entry.fileAttributes))
        {
            ++stats.directoryCount;
        }
        else
        {
            ++stats.fileCount;
        }
    }
}

void RemoveSubtree(VolumeIndex& volume, const NodeId& id, std::unordered_set<NodeId, NodeIdHash>* removedIds = nullptr) noexcept
{
    const auto it = volume.entries.find(id);
    if (it == volume.entries.end())
    {
        return;
    }

    const std::vector<NodeId> children = it->second.children;
    for (const NodeId& childId : children)
    {
        RemoveSubtree(volume, childId, removedIds);
    }

    if (removedIds != nullptr)
    {
        removedIds->insert(id);
    }
    volume.entries.erase(id);
}

void CollectSubtreeIds(const VolumeIndex& volume, const NodeId& id, std::unordered_set<NodeId, NodeIdHash>& outIds) noexcept
{
    const auto it = volume.entries.find(id);
    if (it == volume.entries.end())
    {
        return;
    }

    if (! outIds.insert(id).second)
    {
        return;
    }

    for (const NodeId& childId : it->second.children)
    {
        CollectSubtreeIds(volume, childId, outIds);
    }
}

void RebuildDerivedState(VolumeIndex& volume) noexcept
{
    for (auto& [id, entry] : volume.entries)
    {
        entry.children.clear();
        if (id != volume.trackedRootId)
        {
            entry.fullPath.clear();
        }
    }

    for (auto& [id, entry] : volume.entries)
    {
        if (id == volume.trackedRootId)
        {
            continue;
        }

        if (const auto parent = volume.entries.find(entry.parentId); parent != volume.entries.end())
        {
            parent->second.children.push_back(id);
        }
    }

    const auto compareChildren = [&](const NodeId& left, const NodeId& right) noexcept
    {
        const auto leftIt  = volume.entries.find(left);
        const auto rightIt = volume.entries.find(right);
        if (leftIt == volume.entries.end() || rightIt == volume.entries.end())
        {
            return left.low < right.low || (left.low == right.low && left.high < right.high);
        }

        return std::tie(leftIt->second.name, left.low, left.high) < std::tie(rightIt->second.name, right.low, right.high);
    };

    std::unordered_set<NodeId, NodeIdHash> visited;
    std::vector<NodeId> stack;

    const auto rootIt = volume.entries.find(volume.trackedRootId);
    if (rootIt == volume.entries.end())
    {
        volume.pathIndex.clear();
        return;
    }

    rootIt->second.fullPath = volume.normalizedRootPath;
    stack.push_back(volume.trackedRootId);

    while (! stack.empty())
    {
        const NodeId currentId = stack.back();
        stack.pop_back();

        if (! visited.insert(currentId).second)
        {
            continue;
        }

        auto currentIt = volume.entries.find(currentId);
        if (currentIt == volume.entries.end())
        {
            continue;
        }

        auto& current = currentIt->second;
        std::sort(current.children.begin(), current.children.end(), compareChildren);

        for (auto childIt = current.children.rbegin(); childIt != current.children.rend(); ++childIt)
        {
            auto found = volume.entries.find(*childIt);
            if (found == volume.entries.end())
            {
                continue;
            }

            found->second.fullPath = AppendPath(current.fullPath, found->second.name);
            stack.push_back(*childIt);
        }
    }

    std::vector<NodeId> unreachable;
    for (const auto& [id, entry] : volume.entries)
    {
        if (! visited.contains(id))
        {
            unreachable.push_back(id);
        }
    }

    for (const NodeId& id : unreachable)
    {
        volume.entries.erase(id);
    }

    for (auto& [id, entry] : volume.entries)
    {
        entry.children.clear();
    }
    for (auto& [id, entry] : volume.entries)
    {
        if (id == volume.trackedRootId)
        {
            continue;
        }

        if (const auto parent = volume.entries.find(entry.parentId); parent != volume.entries.end())
        {
            parent->second.children.push_back(id);
        }
    }
    for (auto& [id, entry] : volume.entries)
    {
        std::sort(entry.children.begin(), entry.children.end(), compareChildren);
    }

    volume.pathIndex.clear();
    for (const auto& [id, entry] : volume.entries)
    {
        volume.pathIndex.emplace(FoldPathKey(entry.fullPath), id);
    }
}

HRESULT SaveSnapshot(const VolumeIndex& volume, QueryStats& stats) noexcept
{
    if (! volume.store)
    {
        return E_INVALIDARG;
    }

    PersistedVolumeState state = CapturePersistedVolumeState(volume);
    return volume.store->Save(state, stats);
}

HRESULT LoadSnapshot(VolumeIndex& volume, QueryStats& stats) noexcept
{
    if (! volume.store)
    {
        return E_INVALIDARG;
    }

    PersistedVolumeState state{};
    HRESULT hr = volume.store->Load(state, stats);
    if (FAILED(hr))
    {
        volume.entries.clear();
        return hr;
    }

    ApplyPersistedVolumeState(volume, std::move(state));

    if (! volume.entries.contains(volume.trackedRootId))
    {
        stats.rebuiltSnapshotCorruption = true;
        volume.entries.clear();
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    RebuildDerivedState(volume);
    if (const auto rootIt = volume.entries.find(volume.trackedRootId); rootIt != volume.entries.end())
    {
        volume.trackedRootIsDirectory = IsDirectoryAttributes(rootIt->second.fileAttributes);
    }

    return S_OK;
}

[[nodiscard]] bool MatchWildcardCaseSensitive(std::wstring_view text, std::wstring_view pattern) noexcept
{
    size_t textPos    = 0u;
    size_t patternPos = 0u;
    size_t starPos    = std::wstring_view::npos;
    size_t matchPos   = 0u;

    while (textPos < text.size())
    {
        if (patternPos < pattern.size() && (pattern[patternPos] == L'?' || pattern[patternPos] == text[textPos]))
        {
            ++patternPos;
            ++textPos;
            continue;
        }

        if (patternPos < pattern.size() && pattern[patternPos] == L'*')
        {
            starPos  = patternPos++;
            matchPos = textPos;
            continue;
        }

        if (starPos != std::wstring_view::npos)
        {
            patternPos = starPos + 1u;
            textPos    = ++matchPos;
            continue;
        }

        return false;
    }

    while (patternPos < pattern.size() && pattern[patternPos] == L'*')
    {
        ++patternPos;
    }

    return patternPos == pattern.size();
}

[[nodiscard]] bool MatchWildcard(std::wstring_view text, std::wstring_view pattern, bool caseSensitive) noexcept
{
    return caseSensitive ? MatchWildcardCaseSensitive(text, pattern) : MatchWildcardCaseSensitive(FoldText(text), FoldText(pattern));
}

[[nodiscard]] bool MatchLiteral(std::wstring_view text, std::wstring_view pattern, bool caseSensitive) noexcept
{
    if (pattern.empty())
    {
        return true;
    }

    if (caseSensitive)
    {
        return text.find(pattern) != std::wstring_view::npos;
    }

    const std::wstring foldedText    = FoldText(text);
    const std::wstring foldedPattern = FoldText(pattern);
    return foldedText.find(foldedPattern) != std::wstring_view::npos;
}

[[nodiscard]] bool MatchName(const QueryPlan& plan, std::wstring_view name) noexcept
{
    switch (plan.nameMode)
    {
        case FILESYSTEM_SEARCH_NAME_DISABLED: return true;
        case FILESYSTEM_SEARCH_NAME_WILDCARD: return MatchWildcard(name, plan.namePattern, plan.matchCaseName);
        case FILESYSTEM_SEARCH_NAME_LITERAL: return MatchLiteral(name, plan.namePattern, plan.matchCaseName);
        case FILESYSTEM_SEARCH_NAME_REGEX: return plan.compiledNameRegex != nullptr && std::regex_search(std::wstring(name), *plan.compiledNameRegex);
    }

    return false;
}

HRESULT EnumerateDirectory(std::wstring_view directoryPath, std::vector<EnumeratedChild>& outChildren) noexcept
{
    outChildren.clear();

    const auto fallbackEnumerate = [&]() noexcept -> HRESULT
    {
        WIN32_FIND_DATAW data{};
        const std::wstring pattern = AppendPath(ToExtendedPath(directoryPath), L"*");
        wil::unique_hfind findHandle(::FindFirstFileW(pattern.c_str(), &data));
        if (! findHandle)
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        do
        {
            const std::wstring_view name(data.cFileName);
            if (name == L"." || name == L"..")
            {
                continue;
            }

            EnumeratedChild child{};
            child.name           = std::wstring(name);
            child.fullPath       = AppendPath(directoryPath, child.name);
            child.fileAttributes = data.dwFileAttributes;
            outChildren.push_back(std::move(child));
        } while (::FindNextFileW(findHandle.get(), &data) != 0);

        const DWORD error = ::GetLastError();
        return error == ERROR_NO_MORE_FILES ? S_OK : HRESULT_FROM_WIN32(error);
    };

    wil::unique_handle directory(::CreateFileW(ToExtendedPath(directoryPath).c_str(),
                                               FILE_LIST_DIRECTORY,
                                               FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                               nullptr,
                                               OPEN_EXISTING,
                                               FILE_FLAG_BACKUP_SEMANTICS,
                                               nullptr));
    if (! directory)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    const auto NtQueryDirectoryFile = GetNtQueryDirectoryFile();
    if (NtQueryDirectoryFile == nullptr)
    {
        return fallbackEnumerate();
    }

    std::vector<std::byte> buffer(128u * 1024u);
    bool restart = true;
    for (;;)
    {
        IO_STATUS_BLOCK iosb{};
        const NTSTATUS status = NtQueryDirectoryFile(directory.get(),
                                                     nullptr,
                                                     nullptr,
                                                     nullptr,
                                                     &iosb,
                                                     buffer.data(),
                                                     static_cast<ULONG>(buffer.size()),
                                                     NtFileInformationClass::FileFullDirectoryInformation,
                                                     FALSE,
                                                     nullptr,
                                                     restart ? TRUE : FALSE);
        if (status == STATUS_NO_MORE_FILES)
        {
            break;
        }

        if (! NT_SUCCESS(status))
        {
            if (const auto RtlNtStatusToDosError = GetRtlNtStatusToDosError(); RtlNtStatusToDosError != nullptr)
            {
                const DWORD error = RtlNtStatusToDosError(status);
                if (error == ERROR_INVALID_FUNCTION || error == ERROR_NOT_SUPPORTED || error == ERROR_INVALID_PARAMETER)
                {
                    return fallbackEnumerate();
                }
                return HRESULT_FROM_WIN32(error != 0u ? error : ERROR_GEN_FAILURE);
            }

            return HRESULT_FROM_WIN32(ERROR_GEN_FAILURE);
        }

        restart = false;

        const size_t bytesValid = static_cast<size_t>(iosb.Information);
        if (bytesValid == 0u || bytesValid > buffer.size())
        {
            return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
        }

        size_t offset = 0u;
        while (offset < bytesValid)
        {
            if (bytesValid - offset < offsetof(FILE_FULL_DIR_INFO, FileName))
            {
                return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
            }

            const auto* info = reinterpret_cast<const FILE_FULL_DIR_INFO*>(buffer.data() + offset);
            if ((info->FileNameLength % sizeof(wchar_t)) != 0u)
            {
                return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
            }

            const size_t nameChars = info->FileNameLength / sizeof(wchar_t);
            const std::wstring_view name(info->FileName, nameChars);
            if (name != L"." && name != L"..")
            {
                EnumeratedChild child{};
                child.name           = std::wstring(name);
                child.fullPath       = AppendPath(directoryPath, child.name);
                child.fileAttributes = info->FileAttributes;
                outChildren.push_back(std::move(child));
            }

            if (info->NextEntryOffset == 0u)
            {
                break;
            }

            offset += static_cast<size_t>(info->NextEntryOffset);
        }
    }

    return S_OK;
}

HRESULT HydrateDirectorySubtree(VolumeIndex& volume,
                                const NodeId& directoryId,
                                std::wstring_view directoryPath,
                                CancelCheckFn cancelCheck,
                                void* cancelCookie,
                                QueryStats& stats,
                                RepositoryProgressState& progress) noexcept
{
    HRESULT hr = CheckCancelled(cancelCheck, cancelCookie);
    if (FAILED(hr))
    {
        return hr;
    }

    std::vector<EnumeratedChild> children;
    hr = EnumerateDirectory(directoryPath, children);
    if (FAILED(hr))
    {
        return hr;
    }

    for (const EnumeratedChild& child : children)
    {
        hr = CheckCancelled(cancelCheck, cancelCookie);
        if (FAILED(hr))
        {
            return hr;
        }

        NodeId childId{};
        hr = GetPathNodeId(child.fullPath, childId);
        if (FAILED(hr))
        {
            if (hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED))
            {
                continue;
            }

            return hr;
        }

        Entry entry{};
        entry.id                = childId;
        entry.parentId          = directoryId;
        entry.name              = child.name;
        entry.fileAttributes    = child.fileAttributes;
        volume.entries[childId] = std::move(entry);

        const HRESULT progressHr = AdvanceRepositoryProgress(progress,
                                                             FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP,
                                                             child.fullPath,
                                                             IsDirectoryAttributes(child.fileAttributes) ? 1u : 0u,
                                                             IsDirectoryAttributes(child.fileAttributes) ? 0u : 1u,
                                                             0u,
                                                             0u,
                                                             1u);
        if (FAILED(progressHr))
        {
            return progressHr;
        }

        if (IsDirectoryAttributes(child.fileAttributes))
        {
            hr = HydrateDirectorySubtree(volume, childId, child.fullPath, cancelCheck, cancelCookie, stats, progress);
            if (FAILED(hr))
            {
                return hr;
            }
        }
    }

    stats.usedTraversalSeed = true;
    return S_OK;
}

HRESULT SeedTraversalIndex(VolumeIndex& volume, CancelCheckFn cancelCheck, void* cancelCookie, QueryStats& stats, RepositoryProgressState& progress) noexcept
{
    volume.entries.clear();

    const DWORD rootAttributes = ::GetFileAttributesW(volume.normalizedRootPath.c_str());
    if (rootAttributes == INVALID_FILE_ATTRIBUTES)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    NodeId rootId{};
    HRESULT hr = GetPathNodeId(volume.normalizedRootPath, rootId);
    if (FAILED(hr))
    {
        return hr;
    }

    Entry rootEntry{};
    rootEntry.id             = rootId;
    rootEntry.parentId       = {};
    rootEntry.fileAttributes = rootAttributes;
    rootEntry.name           = IsDirectoryAttributes(rootAttributes) ? std::wstring() : GetPathLeaf(volume.normalizedRootPath);

    volume.entries.emplace(rootId, std::move(rootEntry));
    volume.trackedRootId          = rootId;
    volume.trackedRootIsDirectory = IsDirectoryAttributes(rootAttributes);

    if (volume.trackedRootIsDirectory)
    {
        hr = ReportRepositoryProgress(progress, FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP, volume.normalizedRootPath, S_OK, true);
        if (FAILED(hr))
        {
            return hr;
        }

        hr = HydrateDirectorySubtree(volume, rootId, volume.normalizedRootPath, cancelCheck, cancelCookie, stats, progress);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    RebuildDerivedState(volume);
    return S_OK;
}

HRESULT SeedNtfsIndex(VolumeIndex& volume, CancelCheckFn cancelCheck, void* cancelCookie, QueryStats& stats, RepositoryProgressState& progress) noexcept
{
    volume.entries.clear();

#ifdef ENABLE_TESTS
    if (IsForcedNtfsTraversalSeedEnabled())
    {
        return SeedTraversalIndex(volume, cancelCheck, cancelCookie, stats, progress);
    }
#endif

    NodeId volumeRootId{};
    HRESULT hr = GetPathNodeId(volume.volumeRoot, volumeRootId);
    if (FAILED(hr))
    {
        return hr;
    }

    std::unordered_map<NodeId, SeedEntry, NodeIdHash> allEntries;
    std::unordered_map<NodeId, std::vector<NodeId>, NodeIdHash> childrenByParent;

    SeedEntry volumeRootEntry{};
    volumeRootEntry.id             = volumeRootId;
    volumeRootEntry.parentId       = {};
    volumeRootEntry.fileAttributes = FILE_ATTRIBUTE_DIRECTORY;
    allEntries.emplace(volumeRootId, std::move(volumeRootEntry));

    wil::unique_handle volumeHandle;
    hr = OpenVolumeHandle(volume.volumeDevicePath, volumeHandle);
    if (FAILED(hr))
    {
        if (hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) || hr == HRESULT_FROM_WIN32(ERROR_PRIVILEGE_NOT_HELD))
        {
            return SeedTraversalIndex(volume, cancelCheck, cancelCookie, stats, progress);
        }

        return hr;
    }

    MFT_ENUM_DATA_V0 enumData{};
    enumData.StartFileReferenceNumber = 0u;
    enumData.LowUsn                   = 0u;
    enumData.HighUsn                  = MAXLONGLONG;

    std::vector<std::byte> buffer(1024u * 1024u);
    for (;;)
    {
        hr = CheckCancelled(cancelCheck, cancelCookie);
        if (FAILED(hr))
        {
            return hr;
        }

        DWORD bytesReturned = 0u;
        if (! ::DeviceIoControl(volumeHandle.get(),
                                FSCTL_ENUM_USN_DATA,
                                &enumData,
                                sizeof(enumData),
                                buffer.data(),
                                static_cast<DWORD>(buffer.size()),
                                &bytesReturned,
                                nullptr))
        {
            const DWORD error = ::GetLastError();
            if (error == ERROR_HANDLE_EOF || error == ERROR_NO_MORE_FILES)
            {
                break;
            }

            if (error == ERROR_ACCESS_DENIED || error == ERROR_PRIVILEGE_NOT_HELD || error == ERROR_INVALID_FUNCTION || error == ERROR_NOT_SUPPORTED)
            {
                return SeedTraversalIndex(volume, cancelCheck, cancelCookie, stats, progress);
            }

            return HRESULT_FROM_WIN32(error);
        }

        if (bytesReturned <= sizeof(uint64_t))
        {
            break;
        }

        enumData.StartFileReferenceNumber = *reinterpret_cast<const DWORDLONG*>(buffer.data());

        size_t offset = sizeof(uint64_t);
        while (offset + sizeof(USN_RECORD_COMMON_HEADER) <= bytesReturned)
        {
            const auto* header = reinterpret_cast<const USN_RECORD_COMMON_HEADER*>(buffer.data() + offset);
            if (header->RecordLength < sizeof(USN_RECORD_COMMON_HEADER) || offset + header->RecordLength > bytesReturned)
            {
                return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
            }

            UsnRecordData parsed{};
            if (TryParseUsnRecord(header, parsed))
            {
                SeedEntry entry{};
                entry.id             = parsed.id;
                entry.parentId       = parsed.parentId;
                entry.name           = std::move(parsed.name);
                entry.fileAttributes = parsed.fileAttributes;
                allEntries[entry.id] = entry;
                childrenByParent[entry.parentId].push_back(entry.id);

                hr = AdvanceRepositoryProgress(progress,
                                               FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP,
                                               volume.normalizedRootPath,
                                               IsDirectoryAttributes(entry.fileAttributes) ? 1u : 0u,
                                               IsDirectoryAttributes(entry.fileAttributes) ? 0u : 1u,
                                               0u,
                                               0u,
                                               1u);
                if (FAILED(hr))
                {
                    return hr;
                }
            }

            offset += header->RecordLength;
        }
    }

    const std::vector<std::wstring> components = SplitRelativeComponents(volume.volumeRoot, volume.normalizedRootPath);
    NodeId trackedRootId                       = volumeRootId;
    unsigned long trackedRootAttributes        = FILE_ATTRIBUTE_DIRECTORY;

    for (const std::wstring& component : components)
    {
        const auto childrenIt = childrenByParent.find(trackedRootId);
        if (childrenIt == childrenByParent.end())
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        std::optional<NodeId> matchedId;
        for (const NodeId& childId : childrenIt->second)
        {
            const auto entryIt = allEntries.find(childId);
            if (entryIt == allEntries.end())
            {
                continue;
            }

            if (EqualsCaseInsensitive(entryIt->second.name, component))
            {
                matchedId             = childId;
                trackedRootAttributes = entryIt->second.fileAttributes;
                break;
            }
        }

        if (! matchedId.has_value())
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        trackedRootId = matchedId.value();
    }

    std::vector<NodeId> subtreeIds;
    subtreeIds.push_back(trackedRootId);
    for (size_t index = 0u; index < subtreeIds.size(); ++index)
    {
        const NodeId currentId = subtreeIds[index];
        const auto childrenIt  = childrenByParent.find(currentId);
        if (childrenIt == childrenByParent.end())
        {
            continue;
        }

        subtreeIds.insert(subtreeIds.end(), childrenIt->second.begin(), childrenIt->second.end());
    }

    volume.trackedRootId          = trackedRootId;
    volume.trackedRootIsDirectory = IsDirectoryAttributes(trackedRootAttributes);

    for (const NodeId& id : subtreeIds)
    {
        const auto seedIt = allEntries.find(id);
        if (seedIt == allEntries.end())
        {
            continue;
        }

        Entry entry{};
        entry.id             = seedIt->second.id;
        entry.parentId       = (id == trackedRootId) ? NodeId{} : seedIt->second.parentId;
        entry.name           = seedIt->second.name;
        entry.fileAttributes = seedIt->second.fileAttributes;
        if (id == trackedRootId && ! volume.trackedRootIsDirectory && entry.name.empty())
        {
            entry.name = GetPathLeaf(volume.normalizedRootPath);
        }

        volume.entries.emplace(entry.id, std::move(entry));
    }

    if (! volume.entries.contains(trackedRootId))
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }

    RebuildDerivedState(volume);
    stats.usedNtfsEnumeration = true;
    return S_OK;
}

HRESULT BuildIndex(VolumeIndex& volume, CancelCheckFn cancelCheck, void* cancelCookie, QueryStats& stats, RepositoryProgressState& progress) noexcept
{
    if (volume.fileSystemKind == FileSystemKind::Ntfs)
    {
        return SeedNtfsIndex(volume, cancelCheck, cancelCookie, stats, progress);
    }

    if (volume.fileSystemKind == FileSystemKind::Refs)
    {
        return SeedTraversalIndex(volume, cancelCheck, cancelCookie, stats, progress);
    }

    return kNotSupportedHr;
}

HRESULT ReplayJournal(VolumeIndex& volume,
                      const JournalState& journalState,
                      CancelCheckFn cancelCheck,
                      void* cancelCookie,
                      QueryStats& stats,
                      RepositoryProgressState& progress,
                      JournalDelta* outDelta = nullptr) noexcept
{
    wil::unique_handle volumeHandle;
    HRESULT hr = OpenVolumeHandle(volume.volumeDevicePath, volumeHandle);
    if (FAILED(hr))
    {
        return hr;
    }

    if (outDelta != nullptr)
    {
        *outDelta = {};
    }

    READ_USN_JOURNAL_DATA_V0 readData{};
    readData.StartUsn          = static_cast<USN>(volume.nextUsn);
    readData.ReasonMask        = kJournalReplayReasons;
    readData.ReturnOnlyOnClose = FALSE;
    readData.Timeout           = 0u;
    readData.BytesToWaitFor    = 0u;
    readData.UsnJournalID      = journalState.id;

    std::vector<Entry> directoriesToHydrate;
    std::vector<std::byte> buffer(256u * 1024u);
    while (static_cast<uint64_t>(readData.StartUsn) < journalState.nextUsn)
    {
        hr = CheckCancelled(cancelCheck, cancelCookie);
        if (FAILED(hr))
        {
            return hr;
        }

        DWORD bytesReturned = 0u;
        if (! ::DeviceIoControl(volumeHandle.get(),
                                FSCTL_READ_USN_JOURNAL,
                                &readData,
                                sizeof(readData),
                                buffer.data(),
                                static_cast<DWORD>(buffer.size()),
                                &bytesReturned,
                                nullptr))
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }

        if (bytesReturned < sizeof(USN))
        {
            break;
        }

        const auto nextUsn = *reinterpret_cast<const USN*>(buffer.data());
        readData.StartUsn  = nextUsn;

        size_t offset = sizeof(USN);
        while (offset + sizeof(USN_RECORD_COMMON_HEADER) <= bytesReturned)
        {
            const auto* header = reinterpret_cast<const USN_RECORD_COMMON_HEADER*>(buffer.data() + offset);
            if (header->RecordLength < sizeof(USN_RECORD_COMMON_HEADER) || offset + header->RecordLength > bytesReturned)
            {
                return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
            }

            UsnRecordData record{};
            if (TryParseUsnRecord(header, record))
            {
                const HRESULT progressHr =
                    AdvanceRepositoryProgress(progress, FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP, volume.normalizedRootPath, 0u, 0u, 0u, 0u, 1u);
                if (FAILED(progressHr))
                {
                    return progressHr;
                }

                const bool wasTracked    = volume.entries.contains(record.id);
                const bool parentTracked = volume.entries.contains(record.parentId) || record.parentId == volume.trackedRootId;

                if ((record.reason & USN_REASON_FILE_DELETE) != 0u)
                {
                    RemoveSubtree(volume, record.id, outDelta != nullptr ? &outDelta->deletedIds : nullptr);
                }
                else if ((record.reason & (USN_REASON_FILE_CREATE | USN_REASON_RENAME_NEW_NAME | USN_REASON_BASIC_INFO_CHANGE | USN_REASON_HARD_LINK_CHANGE |
                                           USN_REASON_REPARSE_POINT_CHANGE)) != 0u)
                {
                    if (wasTracked || parentTracked || record.id == volume.trackedRootId)
                    {
                        Entry updated{};
                        if (const auto existing = volume.entries.find(record.id); existing != volume.entries.end())
                        {
                            updated = existing->second;
                        }

                        updated.id             = record.id;
                        updated.parentId       = (record.id == volume.trackedRootId) ? NodeId{} : record.parentId;
                        updated.fileAttributes = record.fileAttributes;
                        if (record.id != volume.trackedRootId || ! record.name.empty())
                        {
                            updated.name = record.name;
                        }

                        volume.entries[record.id] = updated;
                        if (outDelta != nullptr)
                        {
                            outDelta->upsertIds.insert(record.id);
                        }
                        if (IsDirectoryAttributes(updated.fileAttributes) && ((record.reason & (USN_REASON_FILE_CREATE | USN_REASON_RENAME_NEW_NAME)) != 0u) &&
                            ! wasTracked)
                        {
                            directoriesToHydrate.push_back(updated);
                        }
                    }
                    else if (wasTracked)
                    {
                        RemoveSubtree(volume, record.id, outDelta != nullptr ? &outDelta->deletedIds : nullptr);
                    }
                }
            }

            offset += header->RecordLength;
        }
    }

    RebuildDerivedState(volume);
    for (const Entry& directory : directoriesToHydrate)
    {
        const auto pathIt = volume.pathIndex.find(FoldPathKey(directory.fullPath));
        if (pathIt == volume.pathIndex.end())
        {
            continue;
        }

        hr = HydrateDirectorySubtree(volume, directory.id, directory.fullPath, cancelCheck, cancelCookie, stats, progress);
        if (FAILED(hr))
        {
            return hr;
        }
        RebuildDerivedState(volume);
        if (outDelta != nullptr)
        {
            CollectSubtreeIds(volume, directory.id, outDelta->upsertIds);
        }
    }

    volume.journalId           = journalState.id;
    volume.nextUsn             = journalState.nextUsn;
    stats.journalReplayApplied = true;
    return S_OK;
}

HRESULT EnsureReady(VolumeIndex& volume,
                    const RepositoryOptions& options,
                    CancelCheckFn cancelCheck,
                    void* cancelCookie,
                    QueryStats& stats,
                    RepositoryProgressState& progress,
                    RepositoryStatus* runtimeStatus,
                    std::mutex* runtimeStatusMutex,
                    bool* sqliteStoreChanged) noexcept
{
    if (sqliteStoreChanged != nullptr)
    {
        *sqliteStoreChanged = false;
    }

    const auto start               = std::chrono::steady_clock::now();
    const bool sqliteAuthoritative = IsSqliteAuthoritative(options);
    bool sqliteStoreDirty          = false;
    stats.fileSystemKind           = volume.fileSystemKind;
    PopulateStatsFromVolume(volume, stats);
    UpdateRepositoryRuntimeStatus(runtimeStatus,
                                  runtimeStatusMutex,
                                  StoreState::Syncing,
                                  SyncPhase::Loading,
                                  QueryExecutionMode::Unknown,
                                  FallbackReason::None,
                                  volume.normalizedRootPath);
    ApplyRuntimeStatusToProgress(progress, CaptureRepositoryRuntimeStatus(runtimeStatus, runtimeStatusMutex));

    HRESULT hr = ReportRepositoryProgress(progress, FILESYSTEM_SEARCH_PHASE_INITIALIZING, volume.normalizedRootPath, S_OK, true);
    if (FAILED(hr))
    {
        return hr;
    }

    if (! volume.initialized)
    {
        const HRESULT loadHr = LoadSnapshot(volume, stats);
        if (SUCCEEDED(loadHr))
        {
            volume.initialized = true;
            PopulateStatsFromVolume(volume, stats);
            AssignRepositoryProgressCounts(progress, stats.directoryCount, stats.fileCount, 0u, 0u);
            hr = ReportRepositoryProgress(progress, FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP, volume.normalizedRootPath, S_OK, true);
            if (FAILED(hr))
            {
                return hr;
            }
        }
        else if (loadHr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) && loadHr != HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND) &&
                 loadHr != HRESULT_FROM_WIN32(ERROR_INVALID_DATA))
        {
            return loadHr;
        }
    }

    const bool snapshotMissingOnStart = ! stats.rebuiltSnapshotCorruption && stats.snapshotFileBytes == 0u;

    JournalState journalState{};
    hr = GetJournalState(volume, journalState);
    if (hr == kNotSupportedHr)
    {
        stats.journalAvailable = false;

        const auto rebuildWithoutJournal = [&]() noexcept -> HRESULT
        {
            volume.entries.clear();
            volume.pathIndex.clear();
            volume.initialized = false;
            AssignRepositoryProgressCounts(progress, 0u, 0u, 0u, 0u);

            HRESULT rebuildProgressHr = ReportRepositoryProgress(progress, FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP, volume.normalizedRootPath, S_OK, true);
            if (FAILED(rebuildProgressHr))
            {
                return rebuildProgressHr;
            }

            HRESULT rebuildHr = BuildIndex(volume, cancelCheck, cancelCookie, stats, progress);
            if (FAILED(rebuildHr))
            {
                return rebuildHr;
            }

            volume.journalId            = 0u;
            volume.nextUsn              = 0u;
            volume.initialized          = true;
            volume.persistentStoreState = SqliteIndexStore::kVolumeStateCurrentnessUnproven;
            RebuildDerivedState(volume);
            PopulateStatsFromVolume(volume, stats);
            AssignRepositoryProgressCounts(progress, stats.directoryCount, stats.fileCount, 0u, 0u);
            bool seededSqliteStore  = false;
            const HRESULT persistHr = PersistInitialVolumeSeed(volume, options, snapshotMissingOnStart, stats, seededSqliteStore);
            sqliteStoreDirty        = sqliteStoreDirty || seededSqliteStore;
            return persistHr;
        };

        hr = rebuildWithoutJournal();
        if (FAILED(hr))
        {
            return hr;
        }

        PopulateStatsFromVolume(volume, stats);
        AssignRepositoryProgressCounts(progress, stats.directoryCount, stats.fileCount, 0u, 0u);
        hr = ReportRepositoryProgress(progress, FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP, volume.normalizedRootPath, S_OK, true);
        if (FAILED(hr))
        {
            return hr;
        }

        if (! sqliteAuthoritative)
        {
            bool storeChanged = false;
            UpdateRepositoryRuntimeStatus(runtimeStatus,
                                          runtimeStatusMutex,
                                          StoreState::Syncing,
                                          SyncPhase::Mirroring,
                                          QueryExecutionMode::Unknown,
                                          FallbackReason::None,
                                          volume.normalizedRootPath);
            const HRESULT sqliteMirrorHr = MirrorVolumeToConfiguredSqliteStore(volume, options, storeChanged);
            if (FAILED(sqliteMirrorHr))
            {
                Debug::Warning(L"LocalSearchIndexCore: SQLite sidecar mirror failed for root='{}'. hr=0x{:08X}",
                               volume.normalizedRootPath,
                               static_cast<unsigned long>(sqliteMirrorHr));
            }
            sqliteStoreDirty = sqliteStoreDirty || storeChanged;
        }
        if (sqliteStoreChanged != nullptr)
        {
            *sqliteStoreChanged = sqliteStoreDirty;
        }

        stats.ensureReadyDurationMs = ClampDurationMs(std::chrono::steady_clock::now() - start);
        UpdateRepositoryRuntimeStatus(runtimeStatus,
                                      runtimeStatusMutex,
                                      StoreState::Ready,
                                      SyncPhase::Watching,
                                      QueryExecutionMode::Unknown,
                                      FallbackReason::None,
                                      volume.normalizedRootPath);
        Debug::Info(L"LocalSearchIndexCore: ready root='{}' kind={} journalAvailable={} replay={} rebuildIdMismatch={} rebuildRange={} rebuildCorruption={} "
                    L"usedNtfs={} usedTraversal={} entries={} files={} dirs={} snapshotBytes={} memoryBytes={} readyMs={}",
                    volume.normalizedRootPath,
                    static_cast<uint32_t>(stats.fileSystemKind),
                    stats.journalAvailable,
                    stats.journalReplayApplied,
                    stats.rebuiltJournalIdMismatch,
                    stats.rebuiltJournalRangeInvalid,
                    stats.rebuiltSnapshotCorruption,
                    stats.usedNtfsEnumeration,
                    stats.usedTraversalSeed,
                    stats.entryCount,
                    stats.fileCount,
                    stats.directoryCount,
                    stats.snapshotFileBytes,
                    stats.estimatedMemoryBytes,
                    stats.ensureReadyDurationMs);
        return S_OK;
    }

    if (FAILED(hr))
    {
        return hr;
    }
    stats.journalAvailable = journalState.available;

    const auto rebuild = [&]() noexcept -> HRESULT
    {
        volume.entries.clear();
        volume.pathIndex.clear();
        volume.initialized = false;
        AssignRepositoryProgressCounts(progress, 0u, 0u, 0u, 0u);

        HRESULT rebuildProgressHr = ReportRepositoryProgress(progress, FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP, volume.normalizedRootPath, S_OK, true);
        if (FAILED(rebuildProgressHr))
        {
            return rebuildProgressHr;
        }

        HRESULT rebuildHr = BuildIndex(volume, cancelCheck, cancelCookie, stats, progress);
        if (FAILED(rebuildHr))
        {
            return rebuildHr;
        }

        volume.journalId            = journalState.id;
        volume.nextUsn              = journalState.nextUsn;
        volume.initialized          = true;
        volume.persistentStoreState = stats.usedTraversalSeed ? SqliteIndexStore::kVolumeStateCurrentnessUnproven : SqliteIndexStore::kVolumeStateReady;
        RebuildDerivedState(volume);
        PopulateStatsFromVolume(volume, stats);
        AssignRepositoryProgressCounts(progress, stats.directoryCount, stats.fileCount, 0u, 0u);
        bool seededSqliteStore  = false;
        const HRESULT persistHr = PersistInitialVolumeSeed(volume, options, snapshotMissingOnStart, stats, seededSqliteStore);
        sqliteStoreDirty        = sqliteStoreDirty || seededSqliteStore;
        return persistHr;
    };

    if (! volume.initialized)
    {
        hr = rebuild();
        if (FAILED(hr))
        {
            return hr;
        }
    }
    else if (volume.persistentStoreState == SqliteIndexStore::kVolumeStateCurrentnessUnproven)
    {
        if (volume.journalId == journalState.id && volume.nextUsn == journalState.nextUsn)
        {
            volume.persistentStoreState = SqliteIndexStore::kVolumeStateReady;
            if (sqliteAuthoritative)
            {
                hr = SaveSnapshot(volume, stats);
                if (FAILED(hr))
                {
                    return hr;
                }

                sqliteStoreDirty = true;
            }
        }
        else
        {
            stats.rebuiltJournalRangeInvalid = true;
            hr                               = rebuild();
            if (FAILED(hr))
            {
                return hr;
            }
        }
    }
    else if (volume.journalId != 0u && volume.journalId != journalState.id)
    {
        stats.rebuiltJournalIdMismatch = true;
        hr                             = rebuild();
        if (FAILED(hr))
        {
            return hr;
        }
    }
    else if (volume.nextUsn < journalState.firstUsn || volume.nextUsn > journalState.nextUsn)
    {
        stats.rebuiltJournalRangeInvalid = true;
        hr                               = rebuild();
        if (FAILED(hr))
        {
            return hr;
        }
    }
    else if (volume.nextUsn < journalState.nextUsn)
    {
        JournalDelta journalDelta{};
        hr = ReportRepositoryProgress(progress, FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP, volume.normalizedRootPath, S_OK, true);
        if (FAILED(hr))
        {
            return hr;
        }

        UpdateRepositoryRuntimeStatus(runtimeStatus,
                                      runtimeStatusMutex,
                                      StoreState::Syncing,
                                      SyncPhase::Replaying,
                                      QueryExecutionMode::Unknown,
                                      FallbackReason::None,
                                      volume.normalizedRootPath);
        hr = ReplayJournal(volume, journalState, cancelCheck, cancelCookie, stats, progress, sqliteAuthoritative ? &journalDelta : nullptr);
        if (FAILED(hr))
        {
            return hr;
        }

        volume.persistentStoreState = SqliteIndexStore::kVolumeStateReady;
        if (sqliteAuthoritative)
        {
            SqliteIndexStore::ApplyJournalDeltaRequest request{};
            SqliteIndexStore::ApplyJournalDeltaResult applyResult{};
            BuildSqliteApplyJournalDeltaRequest(volume, journalDelta, request);
            const HRESULT applyHr = SqliteIndexStore::ApplyJournalDelta(GetDefaultSqliteDatabasePath(options), request, &applyResult);
            if (FAILED(applyHr))
            {
                return applyHr;
            }
            if (applyResult.insertedNewVolume)
            {
                Debug::Warning(L"LocalSearchIndexCore: authoritative SQLite replay recreated missing persisted root for '{}'; reseeding full volume.",
                               volume.normalizedRootPath);
                hr = SaveSnapshot(volume, stats);
                if (FAILED(hr))
                {
                    return hr;
                }
            }

            sqliteStoreDirty = true;
        }
        else
        {
            hr = SaveSnapshot(volume, stats);
            if (FAILED(hr))
            {
                return hr;
            }
        }
    }
    else
    {
        volume.journalId = journalState.id;
        volume.nextUsn   = journalState.nextUsn;
        if (volume.persistentStoreState != SqliteIndexStore::kVolumeStateReady)
        {
            volume.persistentStoreState = SqliteIndexStore::kVolumeStateReady;
            if (sqliteAuthoritative)
            {
                hr = SaveSnapshot(volume, stats);
                if (FAILED(hr))
                {
                    return hr;
                }

                sqliteStoreDirty = true;
            }
        }
    }

    PopulateStatsFromVolume(volume, stats);
    AssignRepositoryProgressCounts(progress, stats.directoryCount, stats.fileCount, 0u, 0u);
    hr = ReportRepositoryProgress(progress, FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP, volume.normalizedRootPath, S_OK, true);
    if (FAILED(hr))
    {
        return hr;
    }

    if (! sqliteAuthoritative)
    {
        bool storeChanged = false;
        UpdateRepositoryRuntimeStatus(runtimeStatus,
                                      runtimeStatusMutex,
                                      StoreState::Syncing,
                                      SyncPhase::Mirroring,
                                      QueryExecutionMode::Unknown,
                                      FallbackReason::None,
                                      volume.normalizedRootPath);
        const HRESULT sqliteMirrorHr = MirrorVolumeToConfiguredSqliteStore(volume, options, storeChanged);
        if (FAILED(sqliteMirrorHr))
        {
            Debug::Warning(L"LocalSearchIndexCore: SQLite sidecar mirror failed for root='{}'. hr=0x{:08X}",
                           volume.normalizedRootPath,
                           static_cast<unsigned long>(sqliteMirrorHr));
        }
        sqliteStoreDirty = sqliteStoreDirty || storeChanged;
    }
    if (sqliteStoreChanged != nullptr)
    {
        *sqliteStoreChanged = sqliteStoreDirty;
    }

    stats.ensureReadyDurationMs = ClampDurationMs(std::chrono::steady_clock::now() - start);
    UpdateRepositoryRuntimeStatus(runtimeStatus,
                                  runtimeStatusMutex,
                                  StoreState::Ready,
                                  SyncPhase::Watching,
                                  QueryExecutionMode::Unknown,
                                  FallbackReason::None,
                                  volume.normalizedRootPath);
    Debug::Info(L"LocalSearchIndexCore: ready root='{}' kind={} journalAvailable={} replay={} rebuildIdMismatch={} rebuildRange={} rebuildCorruption={} "
                L"usedNtfs={} usedTraversal={} entries={} files={} dirs={} snapshotBytes={} memoryBytes={} readyMs={}",
                volume.normalizedRootPath,
                static_cast<uint32_t>(stats.fileSystemKind),
                stats.journalAvailable,
                stats.journalReplayApplied,
                stats.rebuiltJournalIdMismatch,
                stats.rebuiltJournalRangeInvalid,
                stats.rebuiltSnapshotCorruption,
                stats.usedNtfsEnumeration,
                stats.usedTraversalSeed,
                stats.entryCount,
                stats.fileCount,
                stats.directoryCount,
                stats.snapshotFileBytes,
                stats.estimatedMemoryBytes,
                stats.ensureReadyDurationMs);
    return S_OK;
}

constexpr uint64_t kQueryCancelCheckInterval = 256u;

template <typename EmitFn>
HRESULT ExecuteQueryImpl(const VolumeIndex& volume,
                         const QueryPlan& plan,
                         CancelCheckFn cancelCheck,
                         void* cancelCookie,
                         QueryStats& stats,
                         RepositoryProgressState& progress,
                         EmitFn&& emitCandidate) noexcept(noexcept(emitCandidate(std::declval<Candidate&&>())))
{
    stats.candidateCount = 0u;

    const auto rootIt = volume.entries.find(volume.trackedRootId);
    if (rootIt == volume.entries.end())
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }

    const auto emitEntry = [&](const Entry& entry) -> HRESULT
    {
        const bool isDirectory = IsDirectoryAttributes(entry.fileAttributes);
        if ((isDirectory && ! plan.includeDirectories) || (! isDirectory && ! plan.includeFiles))
        {
            return S_OK;
        }

        if (! MatchName(plan, entry.name))
        {
            return S_OK;
        }

        Candidate candidate{};
        candidate.fullPath       = entry.fullPath;
        candidate.displayName    = entry.name.empty() ? GetPathLeaf(entry.fullPath) : entry.name;
        candidate.fileAttributes = entry.fileAttributes;

        HRESULT hr = emitCandidate(std::move(candidate));
        if (hr == S_FALSE)
        {
            ++stats.candidateCount;
            progress.latest.candidateFiles = stats.candidateCount;
            return S_FALSE;
        }
        if (hr == kSkipCandidateHr)
        {
            return S_OK;
        }
        if (FAILED(hr))
        {
            return hr;
        }

        ++stats.candidateCount;
        progress.latest.candidateFiles = stats.candidateCount;
        return S_OK;
    };

    HRESULT hr = CheckCancelled(cancelCheck, cancelCookie);
    if (FAILED(hr))
    {
        return hr;
    }

    if (! volume.trackedRootIsDirectory)
    {
        const HRESULT progressHr = AdvanceRepositoryProgress(progress, FILESYSTEM_SEARCH_PHASE_ENUMERATING, rootIt->second.fullPath, 0u, 0u, 0u, 0u, 1u, true);
        if (FAILED(progressHr))
        {
            return progressHr;
        }

        hr = emitEntry(rootIt->second);
        return hr == S_FALSE ? S_OK : hr;
    }

    uint64_t visitedEntries = 0u;
    std::vector<NodeId> stack(rootIt->second.children.begin(), rootIt->second.children.end());
    while (! stack.empty())
    {
        if ((visitedEntries % kQueryCancelCheckInterval) == 0u)
        {
            hr = CheckCancelled(cancelCheck, cancelCookie);
            if (FAILED(hr))
            {
                return hr;
            }
        }

        const NodeId id = stack.back();
        stack.pop_back();
        ++visitedEntries;

        const auto it = volume.entries.find(id);
        if (it == volume.entries.end())
        {
            continue;
        }

        const Entry& entry       = it->second;
        const HRESULT progressHr = AdvanceRepositoryProgress(progress, FILESYSTEM_SEARCH_PHASE_ENUMERATING, entry.fullPath, 0u, 0u, 0u, 0u, 1u);
        if (FAILED(progressHr))
        {
            return progressHr;
        }

        hr = emitEntry(entry);
        if (hr == S_FALSE)
        {
            break;
        }
        if (FAILED(hr))
        {
            return hr;
        }

        if (plan.maxResults != 0u && stats.candidateCount >= plan.maxResults)
        {
            break;
        }

        if (plan.recursive && IsDirectoryAttributes(entry.fileAttributes))
        {
            for (auto childIt = entry.children.rbegin(); childIt != entry.children.rend(); ++childIt)
            {
                stack.push_back(*childIt);
            }
        }
    }

    return S_OK;
}

struct SqliteEnumerateContext final
{
    const QueryPlan* plan                 = nullptr;
    CandidateCallbackFn candidateCallback = nullptr;
    void* candidateCookie                 = nullptr;
    QueryStats* stats                     = nullptr;
    RepositoryProgressState* progress     = nullptr;
};

HRESULT STDMETHODCALLTYPE EmitSqliteCandidate(Candidate* candidate, void* cookie) noexcept
{
    if (candidate == nullptr || cookie == nullptr)
    {
        return E_POINTER;
    }

    auto& context = *static_cast<SqliteEnumerateContext*>(cookie);
    if (context.plan == nullptr || context.candidateCallback == nullptr || context.stats == nullptr || context.progress == nullptr)
    {
        return E_POINTER;
    }

    if (! MatchName(*context.plan, candidate->displayName))
    {
        return S_OK;
    }

    HRESULT hr = context.candidateCallback(candidate, context.candidateCookie);
    if (hr == S_FALSE)
    {
        ++context.stats->candidateCount;
        context.progress->latest.candidateFiles = context.stats->candidateCount;
        const HRESULT progressHr = AdvanceRepositoryProgress(*context.progress, FILESYSTEM_SEARCH_PHASE_ENUMERATING, candidate->fullPath, 0u, 0u, 0u, 0u, 1u);
        return FAILED(progressHr) ? progressHr : S_FALSE;
    }
    if (FAILED(hr))
    {
        return hr;
    }

    ++context.stats->candidateCount;
    context.progress->latest.candidateFiles = context.stats->candidateCount;
    hr = AdvanceRepositoryProgress(*context.progress, FILESYSTEM_SEARCH_PHASE_ENUMERATING, candidate->fullPath, 0u, 0u, 0u, 0u, 1u);
    if (FAILED(hr))
    {
        return hr;
    }

    if (context.plan->maxResults != 0u && context.stats->candidateCount >= context.plan->maxResults)
    {
        return S_FALSE;
    }

    return S_OK;
}

#ifdef ENABLE_TESTS
[[nodiscard]] bool ShouldAllowDirectSqliteQueryImpl(const DirectSqliteFreshnessProbe& probe) noexcept
{
    if (! probe.storedVolumeReady)
    {
        return false;
    }

    if (! probe.currentJournalKnown || ! probe.currentJournalAvailable)
    {
        return false;
    }

    if (probe.storedJournalId != probe.currentJournalId)
    {
        return false;
    }

    if (probe.storedNextUsn < probe.currentFirstUsn || probe.storedNextUsn > probe.currentNextUsn)
    {
        return false;
    }

    return probe.storedNextUsn == probe.currentNextUsn;
}
#else
[[nodiscard]] bool ShouldAllowDirectSqliteQueryImpl(bool storedVolumeReady,
                                                    bool currentJournalKnown,
                                                    bool currentJournalAvailable,
                                                    uint64_t storedJournalId,
                                                    uint64_t storedNextUsn,
                                                    uint64_t currentJournalId,
                                                    uint64_t currentFirstUsn,
                                                    uint64_t currentNextUsn) noexcept
{
    if (! storedVolumeReady)
    {
        return false;
    }

    if (! currentJournalKnown || ! currentJournalAvailable)
    {
        return false;
    }

    if (storedJournalId != currentJournalId)
    {
        return false;
    }

    if (storedNextUsn < currentFirstUsn || storedNextUsn > currentNextUsn)
    {
        return false;
    }

    return storedNextUsn == currentNextUsn;
}
#endif

HRESULT TryEnumerateFromConfiguredSqliteStore(const PersistentStoreInfo& storeInfo,
                                              const QueryPlan& plan,
                                              CancelCheckFn cancelCheck,
                                              void* cancelCookie,
                                              CandidateCallbackFn candidateCallback,
                                              void* candidateCookie,
                                              QueryStats& stats,
                                              RepositoryProgressState& progress,
                                              FallbackReason* outFallbackReason = nullptr) noexcept
{
    if (storeInfo.kind != PersistentStoreKind::Sqlite)
    {
        return S_FALSE;
    }
    if (! storeInfo.inspectionSucceeded)
    {
        if (outFallbackReason != nullptr)
        {
            *outFallbackReason = ClassifyUninspectableStore(storeInfo);
        }
        return S_FALSE;
    }

    if (! storeInfo.readyForQueryCutover)
    {
        stats.sqliteCutoverBlocked = true;
        if (outFallbackReason != nullptr)
        {
            *outFallbackReason = FallbackReason::CutoverBlocked;
        }
        return S_FALSE;
    }

    SqliteIndexStore::VolumeInfo storedVolume{};
    const HRESULT inspectVolumeHr = SqliteIndexStore::InspectVolume(storeInfo.primaryPath, plan.rootPath, storedVolume);
    if (FAILED(inspectVolumeHr))
    {
        if (inspectVolumeHr != HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND))
        {
            Debug::Info(L"LocalSearchIndexCore: direct SQLite volume inspection fallback root='{}' hr=0x{:08X}",
                        plan.rootPath,
                        static_cast<unsigned long>(inspectVolumeHr));
        }
        if (outFallbackReason != nullptr)
        {
            *outFallbackReason = inspectVolumeHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) ? FallbackReason::StoreMissing : FallbackReason::StoreInvalid;
        }
        return S_FALSE;
    }

    if (storedVolume.state == SqliteIndexStore::kVolumeStateCurrentnessUnproven)
    {
        stats.sqliteCutoverBlocked = true;
        if (outFallbackReason != nullptr)
        {
            *outFallbackReason = FallbackReason::StoreStale;
        }
        return S_FALSE;
    }

    if (storedVolume.state == SqliteIndexStore::kVolumeStateReady)
    {
        const std::wstring volumeRoot       = ExtractVolumeRoot(plan.rootPath);
        const std::wstring volumeDevicePath = BuildVolumeDevicePath(volumeRoot);
        bool currentJournalKnown            = false;
        JournalState journalState{};
        if (! volumeDevicePath.empty())
        {
            const HRESULT journalHr = GetJournalState(volumeDevicePath, journalState);
            if (SUCCEEDED(journalHr))
            {
                currentJournalKnown    = true;
                stats.journalAvailable = journalState.available;
            }
        }

#ifdef ENABLE_TESTS
        const DirectSqliteFreshnessProbe probe{
            .storedVolumeReady       = storedVolume.state == SqliteIndexStore::kVolumeStateReady,
            .currentJournalKnown     = currentJournalKnown,
            .currentJournalAvailable = journalState.available,
            .storedJournalId         = storedVolume.journalId,
            .storedNextUsn           = storedVolume.nextUsn,
            .currentJournalId        = journalState.id,
            .currentFirstUsn         = journalState.firstUsn,
            .currentNextUsn          = journalState.nextUsn,
        };
        const bool allowDirectQuery = ShouldAllowDirectSqliteQueryImpl(probe);
#else
        const bool allowDirectQuery = ShouldAllowDirectSqliteQueryImpl(storedVolume.state == SqliteIndexStore::kVolumeStateReady,
                                                                       currentJournalKnown,
                                                                       journalState.available,
                                                                       storedVolume.journalId,
                                                                       storedVolume.nextUsn,
                                                                       journalState.id,
                                                                       journalState.firstUsn,
                                                                       journalState.nextUsn);
#endif
        if (! allowDirectQuery)
        {
            stats.sqliteCutoverBlocked = true;
            if (outFallbackReason != nullptr)
            {
                *outFallbackReason = FallbackReason::StoreStale;
            }
            Debug::Info(L"LocalSearchIndexCore: direct SQLite freshness fallback root='{}' currentJournalKnown={} currentJournalAvailable={} "
                        L"storedJournalId={} currentJournalId={} storedNextUsn={} currentFirstUsn={} currentNextUsn={}",
                        plan.rootPath,
                        currentJournalKnown,
                        journalState.available,
                        storedVolume.journalId,
                        journalState.id,
                        storedVolume.nextUsn,
                        journalState.firstUsn,
                        journalState.nextUsn);
            return S_FALSE;
        }
    }

    SqliteIndexStore::QueryRequest request{};
    request.rootPath           = plan.rootPath;
    request.namePattern        = plan.namePattern;
    request.nameMode           = plan.nameMode;
    request.matchCaseName      = plan.matchCaseName;
    request.recursive          = plan.recursive;
    request.includeFiles       = plan.includeFiles;
    request.includeDirectories = plan.includeDirectories;
    request.maxResults         = plan.maxResults;

    SqliteEnumerateContext context{
        .plan              = &plan,
        .candidateCallback = candidateCallback,
        .candidateCookie   = candidateCookie,
        .stats             = &stats,
        .progress          = &progress,
    };

    SqliteIndexStore::QueryRuntimeStats sqliteStats{};
    const HRESULT hr =
        SqliteIndexStore::EnumerateVolume(storeInfo.primaryPath, request, cancelCheck, cancelCookie, &EmitSqliteCandidate, &context, &sqliteStats);
    if (FAILED(hr))
    {
        if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
        {
            return hr;
        }

        if (outFallbackReason != nullptr)
        {
            *outFallbackReason = FallbackReason::SqliteFailure;
        }
        Debug::Warning(L"LocalSearchIndexCore: SQLite query fallback root='{}' hr=0x{:08X}", plan.rootPath, static_cast<unsigned long>(hr));
        return S_FALSE;
    }

    if (! sqliteStats.volumeReady)
    {
        stats.sqliteCutoverBlocked = true;
        if (outFallbackReason != nullptr)
        {
            *outFallbackReason = FallbackReason::CutoverBlocked;
        }
        return S_FALSE;
    }

    stats.fileSystemKind      = sqliteStats.fileSystemKind;
    stats.entryCount          = sqliteStats.entryCount;
    stats.fileCount           = sqliteStats.fileCount;
    stats.directoryCount      = sqliteStats.directoryCount;
    stats.journalId           = sqliteStats.journalId;
    stats.nextUsn             = sqliteStats.nextUsn;
    stats.usedSqliteStore     = true;
    stats.sqliteReadOnlyQuery = sqliteStats.readOnlyConnection;
    stats.usedNamePrefilter   = sqliteStats.usedNamePrefilter;
    return S_OK;
}

[[nodiscard]] bool IsVolumeCurrentForNoWaitQuery(const VolumeIndex& volume, QueryStats& stats) noexcept
{
    if (! volume.initialized)
    {
        return false;
    }

    JournalState journalState{};
    const HRESULT journalHr = GetJournalState(volume, journalState);
    if (FAILED(journalHr))
    {
        stats.journalAvailable = false;
        return false;
    }

    stats.journalAvailable = journalState.available;
    if (! journalState.available)
    {
        return false;
    }

    if (volume.journalId != journalState.id)
    {
        return false;
    }

    if (volume.nextUsn < journalState.firstUsn || volume.nextUsn > journalState.nextUsn)
    {
        return false;
    }

    return volume.nextUsn == journalState.nextUsn;
}

template <typename EmitFn>
HRESULT EnumerateLiveFileSystem(const QueryPlan& plan,
                                CancelCheckFn cancelCheck,
                                void* cancelCookie,
                                QueryStats& stats,
                                RepositoryProgressState& progress,
                                EmitFn&& emitCandidate) noexcept(noexcept(emitCandidate(std::declval<Candidate&&>())))
{
    struct DirectoryFrame final
    {
        std::wstring fullPath;
    };

    stats.usedLiveScanFallback = true;
    stats.queryExecutionMode   = QueryExecutionMode::LiveScanFallback;

    const DWORD rootAttributes = ::GetFileAttributesW(plan.rootPath.c_str());
    if (rootAttributes == INVALID_FILE_ATTRIBUTES)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    if ((rootAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0u)
    {
        ++stats.fileCount;
        const HRESULT progressHr = AdvanceRepositoryProgress(progress, FILESYSTEM_SEARCH_PHASE_ENUMERATING, plan.rootPath, 0u, 1u, 0u, 0u, 1u, true);
        if (FAILED(progressHr))
        {
            return progressHr;
        }

        if (! plan.includeFiles || ! MatchName(plan, GetPathLeaf(plan.rootPath)))
        {
            stats.entryCount = stats.fileCount;
            return S_OK;
        }

        Candidate candidate{};
        candidate.fullPath       = plan.rootPath;
        candidate.displayName    = std::wstring(GetPathLeaf(plan.rootPath));
        candidate.fileAttributes = rootAttributes;

        HRESULT hr = emitCandidate(std::move(candidate));
        if (FAILED(hr) && hr != S_FALSE)
        {
            return hr;
        }
        if (hr != kSkipCandidateHr)
        {
            ++stats.candidateCount;
            progress.latest.candidateFiles = stats.candidateCount;
        }

        stats.entryCount = stats.fileCount;
        return S_OK;
    }

    std::vector<DirectoryFrame> stack;
    stack.push_back({plan.rootPath});

    while (! stack.empty())
    {
        const HRESULT cancelHr = CheckCancelled(cancelCheck, cancelCookie);
        if (FAILED(cancelHr))
        {
            return cancelHr;
        }

        DirectoryFrame frame = std::move(stack.back());
        stack.pop_back();

        ++stats.directoryCount;
        HRESULT hr = AdvanceRepositoryProgress(progress, FILESYSTEM_SEARCH_PHASE_ENUMERATING, frame.fullPath, 1u, 0u, 0u, 0u, 1u, true);
        if (FAILED(hr))
        {
            return hr;
        }

        std::vector<EnumeratedChild> children;
        hr = EnumerateDirectory(frame.fullPath, children);
        if (FAILED(hr))
        {
            return hr;
        }

        for (auto childIt = children.begin(); childIt != children.end(); ++childIt)
        {
            const EnumeratedChild& child = *childIt;
            const bool isDirectory       = IsDirectoryAttributes(child.fileAttributes);
            if (isDirectory)
            {
                if (plan.recursive)
                {
                    stack.push_back({child.fullPath});
                }
            }
            else
            {
                ++stats.fileCount;
            }

            bool emittedCandidate = false;
            if ((isDirectory && plan.includeDirectories) || (! isDirectory && plan.includeFiles))
            {
                if (MatchName(plan, child.name))
                {
                    Candidate candidate{};
                    candidate.fullPath       = child.fullPath;
                    candidate.displayName    = child.name;
                    candidate.fileAttributes = child.fileAttributes;

                    hr = emitCandidate(std::move(candidate));
                    if (FAILED(hr) && hr != S_FALSE)
                    {
                        return hr;
                    }
                    if (hr != kSkipCandidateHr)
                    {
                        ++stats.candidateCount;
                        progress.latest.candidateFiles = stats.candidateCount;
                        emittedCandidate               = true;
                    }
                }
            }

            hr = AdvanceRepositoryProgress(
                progress, FILESYSTEM_SEARCH_PHASE_ENUMERATING, child.fullPath, 0u, isDirectory ? 0u : 1u, emittedCandidate ? 1u : 0u, 0u, 1u);
            if (FAILED(hr))
            {
                return hr;
            }

            if (plan.maxResults != 0u && stats.candidateCount >= plan.maxResults)
            {
                stats.entryCount = stats.fileCount + stats.directoryCount;
                return S_OK;
            }
        }
    }

    stats.entryCount = stats.fileCount + stats.directoryCount;
    return S_OK;
}
} // namespace

std::wstring_view GetPersistentStoreKindText(const PersistentStoreKind kind) noexcept
{
    switch (kind)
    {
        case PersistentStoreKind::SnapshotBinary: return L"snapshot-v1";
        case PersistentStoreKind::Sqlite: return L"sqlite-v2";
        default: return L"unknown";
    }
}

std::wstring_view GetStoreStateText(const StoreState state) noexcept
{
    switch (state)
    {
        case StoreState::Ready: return L"ready";
        case StoreState::Syncing: return L"syncing";
        case StoreState::Recovering: return L"recovering";
        case StoreState::Invalid: return L"invalid";
        case StoreState::Maintenance: return L"maintenance";
        case StoreState::Unknown:
        default: return L"unknown";
    }
}

std::wstring_view GetSyncPhaseText(const SyncPhase phase) noexcept
{
    switch (phase)
    {
        case SyncPhase::Discovering: return L"discovering";
        case SyncPhase::Loading: return L"loading";
        case SyncPhase::Replaying: return L"replaying";
        case SyncPhase::Mirroring: return L"mirroring";
        case SyncPhase::Watching: return L"watching";
        case SyncPhase::Idle:
        default: return L"idle";
    }
}

std::wstring_view GetQueryExecutionModeText(const QueryExecutionMode mode) noexcept
{
    switch (mode)
    {
        case QueryExecutionMode::Sqlite: return L"sqlite";
        case QueryExecutionMode::InMemoryIndex: return L"in-memory-index";
        case QueryExecutionMode::LiveScanFallback: return L"live-scan";
        case QueryExecutionMode::Unknown:
        default: return L"unknown";
    }
}

std::wstring_view GetFallbackReasonText(const FallbackReason reason) noexcept
{
    switch (reason)
    {
        case FallbackReason::StoreMissing: return L"store-missing";
        case FallbackReason::StoreInvalid: return L"store-invalid";
        case FallbackReason::StoreStale: return L"store-stale";
        case FallbackReason::CutoverBlocked: return L"cutover-blocked";
        case FallbackReason::WarmupRunning: return L"warmup-running";
        case FallbackReason::SqliteFailure: return L"sqlite-failure";
        case FallbackReason::None:
        default: return L"none";
    }
}

SqliteMaintenancePolicy GetDefaultSqliteMaintenancePolicy() noexcept
{
    return {
        .autoCheckpointTargetBytes          = kSqliteAutoCheckpointTargetBytes,
        .autoCompactionFragmentationPercent = kSqliteAutoCompactionFragmentationPercent,
        .autoCompactionMinBytes             = kSqliteAutoCompactionMinBytes,
    };
}

PersistentStoreInfo GetPersistentStoreInfo(const RepositoryOptions& options) noexcept
{
    PersistentStoreInfo info{};
    info.kind          = options.persistentStoreKind;
    info.rootDirectory = GetSnapshotRootDirectory(options);

    if (info.kind == PersistentStoreKind::Sqlite)
    {
        const SqliteMaintenancePolicy policy    = GetDefaultSqliteMaintenancePolicy();
        info.primaryPath                        = GetDefaultSqliteDatabasePath(options);
        info.primaryBytes                       = GetFileBytes(info.primaryPath);
        info.writeAheadLogPath                  = info.primaryPath.empty() ? std::wstring{} : info.primaryPath + L"-wal";
        info.writeAheadLogBytes                 = GetFileBytes(info.writeAheadLogPath);
        info.autoCheckpointEnabled              = true;
        info.autoCheckpointTargetBytes          = policy.autoCheckpointTargetBytes;
        info.autoCompactionEnabled              = true;
        info.autoCompactionFragmentationPercent = policy.autoCompactionFragmentationPercent;
        info.autoCompactionMinBytes             = policy.autoCompactionMinBytes;

        SqliteIndexStore::StoreInfo sqliteInfo{};
        const HRESULT inspectHr = SqliteIndexStore::InspectStore(info.primaryPath, sqliteInfo);
        if (SUCCEEDED(inspectHr))
        {
            info.inspectionSucceeded     = true;
            info.primaryBytes            = sqliteInfo.databaseBytes;
            info.writeAheadLogBytes      = sqliteInfo.writeAheadLogBytes;
            info.pageCount               = sqliteInfo.pageCount;
            info.freelistPageCount       = sqliteInfo.freelistPageCount;
            info.lastCheckpointUtc       = sqliteInfo.lastCheckpointUtc;
            info.lastCompactionUtc       = sqliteInfo.lastCompactionUtc;
            info.indexedVolumeCount      = sqliteInfo.volumeCount;
            info.indexedEntryCount       = sqliteInfo.entryCount;
            info.legacyImportVolumeCount = sqliteInfo.legacyImportVolumeCount;
            info.readyForQueryCutover    = sqliteInfo.legacyImportVolumeCount == 0u;
        }
    }

    return info;
}

#ifdef ENABLE_TESTS
bool ShouldAllowDirectSqliteQueryForTests(const DirectSqliteFreshnessProbe& probe) noexcept
{
    return ShouldAllowDirectSqliteQueryImpl(probe);
}
#endif

Repository::Repository(RepositoryOptions options) noexcept : _options(std::move(options))
{
    _cachedPersistentStoreInfo.kind   = _options.persistentStoreKind;
    _runtimeStatus.storeState         = StoreState::Unknown;
    _runtimeStatus.syncPhase          = SyncPhase::Idle;
    _runtimeStatus.queryExecutionMode = QueryExecutionMode::Unknown;
    _runtimeStatus.fallbackReason     = FallbackReason::None;
}

std::vector<std::wstring> DiscoverFixedLocalRoots() noexcept
{
    try
    {
        return DiscoverFixedLocalRootsImpl();
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"LocalSearchIndexCore: DiscoverFixedLocalRoots failed with an unexpected std::exception.");
        return {};
    }
}

PersistentStoreInfo Repository::GetCachedPersistentStoreInfo() noexcept
{
    std::lock_guard guard(_mutex);
    return _cachedPersistentStoreInfo;
}

void Repository::GetStatus(RepositoryStatus& outStatus) noexcept
{
    outStatus = CaptureRepositoryRuntimeStatus(&_runtimeStatus, &_statusMutex);
}

HRESULT Repository::AcquireOrCreateVolume(const SupportInfo& support, std::shared_ptr<VolumeIndex>& outVolume) noexcept
{
    outVolume.reset();

    if (! support.indexable)
    {
        return kNotSupportedHr;
    }

    const std::wstring rootKey = FoldPathKey(support.normalizedRootPath);

    std::lock_guard guard(_mutex);
    auto it = _volumes.find(rootKey);
    if (it == _volumes.end())
    {
        auto newVolume                = std::make_shared<VolumeIndex>();
        newVolume->normalizedRootPath = support.normalizedRootPath;
        newVolume->rootKey            = rootKey;
        newVolume->fileSystemKind     = support.fileSystemKind;
        newVolume->volumeRoot         = ExtractVolumeRoot(support.normalizedRootPath);
        newVolume->volumeDevicePath   = BuildVolumeDevicePath(newVolume->volumeRoot);
        newVolume->store              = CreateVolumeStore(newVolume->normalizedRootPath, newVolume->fileSystemKind, _options);
        newVolume->snapshotPath       = newVolume->store ? std::wstring(newVolume->store->GetPrimaryPath()) : std::wstring{};
        if (newVolume->volumeRoot.empty() || newVolume->volumeDevicePath.empty() || ! newVolume->store || newVolume->snapshotPath.empty())
        {
            return kNotSupportedHr;
        }

        it = _volumes.emplace(rootKey, std::move(newVolume)).first;
    }

    outVolume = it->second;
    return S_OK;
}

void Repository::RefreshCachedPersistentStoreInfo() noexcept
{
    const PersistentStoreInfo refreshed = GetPersistentStoreInfo(_options);

    std::lock_guard guard(_mutex);
    _cachedPersistentStoreInfo      = refreshed;
    _cachedPersistentStoreInfoValid = true;
}

void Repository::InvalidateCachedPersistentStoreInfo() noexcept
{
    std::lock_guard guard(_mutex);
    _cachedPersistentStoreInfo      = {};
    _cachedPersistentStoreInfo.kind = _options.persistentStoreKind;
    _cachedPersistentStoreInfoValid = false;
}

HRESULT Repository::ProbePath(std::wstring_view rootPath, SupportInfo& outSupport) noexcept
{
    try
    {
        return PopulateSupportInfo(rootPath, outSupport);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"LocalSearchIndexCore: ProbePath failed with an unexpected std::exception.");
        outSupport = {};
        return E_FAIL;
    }
}

HRESULT Repository::Query(const QueryPlan& plan,
                          CancelCheckFn cancelCheck,
                          void* cancelCookie,
                          std::vector<Candidate>& outCandidates,
                          QueryStats* outStats,
                          ProgressCallbackFn progressCallback,
                          void* progressCookie) noexcept
{
    try
    {
        outCandidates.clear();
        HRESULT hr = Enumerate(plan,
                               cancelCheck,
                               cancelCookie,
                               [](Candidate* candidate, void* cookie) noexcept -> HRESULT
        {
            if (candidate == nullptr || cookie == nullptr)
            {
                return E_POINTER;
            }

            try
            {
                auto* results = static_cast<std::vector<Candidate>*>(cookie);
                results->push_back(std::move(*candidate));
                return S_OK;
            }
            catch (const std::bad_alloc&)
            {
                std::terminate();
            }
            catch (const std::exception&)
            {
                // Mandatory: callback boundary. Report failure instead of unwinding through a noexcept callback.
                return E_FAIL;
            }
        },
                               &outCandidates,
                               outStats,
                               progressCallback,
                               progressCookie);
        if (FAILED(hr))
        {
            return hr;
        }

        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::regex_error&)
    {
        return E_INVALIDARG;
    }
    catch (const std::exception&)
    {
        Debug::Error(L"LocalSearchIndexCore: Query failed with an unexpected std::exception.");
        outCandidates.clear();
        if (outStats != nullptr)
        {
            *outStats = {};
        }
        return E_FAIL;
    }
}

HRESULT Repository::Enumerate(const QueryPlan& plan,
                              CancelCheckFn cancelCheck,
                              void* cancelCookie,
                              CandidateCallbackFn candidateCallback,
                              void* candidateCookie,
                              QueryStats* outStats,
                              ProgressCallbackFn progressCallback,
                              void* progressCookie) noexcept
{
    try
    {
        QueryStats stats{};
        RepositoryProgressState progress{
            .callback = progressCallback,
            .cookie   = progressCookie,
        };

        if (candidateCallback == nullptr)
        {
            return E_POINTER;
        }

        if (plan.rootPath.empty())
        {
            return E_INVALIDARG;
        }

        SupportInfo support{};
        HRESULT hr = PopulateSupportInfo(plan.rootPath, support);
        if (FAILED(hr))
        {
            return hr;
        }

        progress.latest.currentPath = support.normalizedRootPath;

        if (! support.indexable)
        {
            return kNotSupportedHr;
        }

        if (_options.persistentStoreKind == PersistentStoreKind::Sqlite)
        {
            bool refreshStoreInfo = false;
            {
                std::lock_guard guard(_mutex);
                refreshStoreInfo = ! _cachedPersistentStoreInfoValid;
            }

            if (refreshStoreInfo)
            {
                RefreshCachedPersistentStoreInfo();
            }

            QueryPlan sqlitePlan = plan;
            sqlitePlan.rootPath  = support.normalizedRootPath;

            stats.fileSystemKind = support.fileSystemKind;

            AssignRepositoryProgressCounts(progress, 0u, 0u, 0u, 0u);
            hr = ReportRepositoryProgress(progress, FILESYSTEM_SEARCH_PHASE_ENUMERATING, sqlitePlan.rootPath, S_OK, true);
            if (FAILED(hr))
            {
                return hr;
            }

            const auto directSqliteStart              = std::chrono::steady_clock::now();
            const PersistentStoreInfo cachedStoreInfo = GetCachedPersistentStoreInfo();
            hr                                        = TryEnumerateFromConfiguredSqliteStore(
                cachedStoreInfo, sqlitePlan, cancelCheck, cancelCookie, candidateCallback, candidateCookie, stats, progress);
            if (hr == S_OK)
            {
                stats.executeQueryDurationMs = ClampDurationMs(std::chrono::steady_clock::now() - directSqliteStart);

                Debug::Info(L"LocalSearchIndexCore: direct SQLite query root='{}' pattern='{}' mode={} recursive={} includeFiles={} includeDirs={} "
                            L"maxResults={} candidates={} entryCount={} snapshotBytes={} queryMs={} prefilter={} cutoverBlocked={}",
                            sqlitePlan.rootPath,
                            sqlitePlan.namePattern,
                            static_cast<uint32_t>(sqlitePlan.nameMode),
                            sqlitePlan.recursive,
                            sqlitePlan.includeFiles,
                            sqlitePlan.includeDirectories,
                            sqlitePlan.maxResults,
                            stats.candidateCount,
                            stats.entryCount,
                            stats.snapshotFileBytes,
                            stats.executeQueryDurationMs,
                            stats.usedNamePrefilter,
                            stats.sqliteCutoverBlocked);

                if (outStats != nullptr)
                {
                    *outStats = stats;
                }

                return S_OK;
            }
            if (FAILED(hr))
            {
                return hr;
            }
        }

        std::shared_ptr<VolumeIndex> volume;
        hr = AcquireOrCreateVolume(support, volume);
        if (FAILED(hr))
        {
            return hr;
        }

        std::lock_guard volumeGuard(volume->mutex);

        bool sqliteStoreChanged = false;
        hr = EnsureReady(*volume, _options, cancelCheck, cancelCookie, stats, progress, &_runtimeStatus, &_statusMutex, &sqliteStoreChanged);
        if (FAILED(hr))
        {
            return hr;
        }

        if (_options.persistentStoreKind == PersistentStoreKind::Sqlite)
        {
            bool refreshStoreInfo = sqliteStoreChanged;
            {
                std::lock_guard guard(_mutex);
                refreshStoreInfo = refreshStoreInfo || ! _cachedPersistentStoreInfoValid;
            }

            if (refreshStoreInfo)
            {
                RefreshCachedPersistentStoreInfo();
            }
        }

        QueryPlan effectivePlan = plan;
        effectivePlan.rootPath  = volume->normalizedRootPath;

        AssignRepositoryProgressCounts(progress, stats.directoryCount, stats.fileCount, 0u, 0u);
        hr = ReportRepositoryProgress(progress, FILESYSTEM_SEARCH_PHASE_ENUMERATING, effectivePlan.rootPath, S_OK, true);
        if (FAILED(hr))
        {
            return hr;
        }

        const auto executeStart                   = std::chrono::steady_clock::now();
        const PersistentStoreInfo cachedStoreInfo = GetCachedPersistentStoreInfo();
        hr                                        = TryEnumerateFromConfiguredSqliteStore(
            cachedStoreInfo, effectivePlan, cancelCheck, cancelCookie, candidateCallback, candidateCookie, stats, progress);
        if (hr == S_FALSE)
        {
            hr = ExecuteQueryImpl(*volume,
                                  effectivePlan,
                                  cancelCheck,
                                  cancelCookie,
                                  stats,
                                  progress,
                                  [candidateCallback, candidateCookie](Candidate&& candidate) noexcept -> HRESULT
            { return candidateCallback(&candidate, candidateCookie); });
        }
        if (FAILED(hr))
        {
            return hr;
        }
        stats.executeQueryDurationMs = ClampDurationMs(std::chrono::steady_clock::now() - executeStart);

        Debug::Info(L"LocalSearchIndexCore: query root='{}' pattern='{}' mode={} recursive={} includeFiles={} includeDirs={} maxResults={} "
                    L"candidates={} entryCount={} snapshotBytes={} memoryBytes={} readyMs={} queryMs={} sqliteStore={} prefilter={} cutoverBlocked={}",
                    effectivePlan.rootPath,
                    effectivePlan.namePattern,
                    static_cast<uint32_t>(effectivePlan.nameMode),
                    effectivePlan.recursive,
                    effectivePlan.includeFiles,
                    effectivePlan.includeDirectories,
                    effectivePlan.maxResults,
                    stats.candidateCount,
                    stats.entryCount,
                    stats.snapshotFileBytes,
                    stats.estimatedMemoryBytes,
                    stats.ensureReadyDurationMs,
                    stats.executeQueryDurationMs,
                    stats.usedSqliteStore,
                    stats.usedNamePrefilter,
                    stats.sqliteCutoverBlocked);

        if (outStats != nullptr)
        {
            *outStats = stats;
        }

        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::regex_error&)
    {
        return E_INVALIDARG;
    }
    catch (const std::exception&)
    {
        Debug::Error(L"LocalSearchIndexCore: Enumerate failed with an unexpected std::exception.");
        if (outStats != nullptr)
        {
            *outStats = {};
        }
        return E_FAIL;
    }
}

HRESULT Repository::EnumerateNoWait(const QueryPlan& plan,
                                    CancelCheckFn cancelCheck,
                                    void* cancelCookie,
                                    CandidateCallbackFn candidateCallback,
                                    void* candidateCookie,
                                    QueryStats* outStats,
                                    ProgressCallbackFn progressCallback,
                                    void* progressCookie) noexcept
{
    try
    {
        const Debug::Perf::Scope totalPerf(L"search.enumerate_nowait.total_ms");
        EmitPerfCount(L"search.enumerate_nowait.calls");

        if (_options.persistentStoreKind != PersistentStoreKind::Sqlite)
        {
            return Enumerate(plan, cancelCheck, cancelCookie, candidateCallback, candidateCookie, outStats, progressCallback, progressCookie);
        }

        QueryStats stats{};
        RepositoryProgressState progress{
            .callback = progressCallback,
            .cookie   = progressCookie,
        };

        if (candidateCallback == nullptr)
        {
            return E_POINTER;
        }

        if (plan.rootPath.empty())
        {
            return E_INVALIDARG;
        }

        SupportInfo support{};
        HRESULT hr = PopulateSupportInfo(plan.rootPath, support);
        if (FAILED(hr))
        {
            return hr;
        }

        progress.latest.currentPath = support.normalizedRootPath;
        stats.fileSystemKind        = support.fileSystemKind;
        if (! support.indexable)
        {
            return kNotSupportedHr;
        }

        bool refreshStoreInfo = false;
        {
            std::lock_guard guard(_mutex);
            refreshStoreInfo = ! _cachedPersistentStoreInfoValid;
        }
        if (refreshStoreInfo)
        {
            RefreshCachedPersistentStoreInfo();
        }

        QueryPlan effectivePlan        = plan;
        effectivePlan.rootPath         = support.normalizedRootPath;
        const bool sqliteAuthoritative = IsSqliteAuthoritative(_options);

        const PersistentStoreInfo cachedStoreInfo = GetCachedPersistentStoreInfo();

        UpdateRepositoryRuntimeStatus(&_runtimeStatus,
                                      &_statusMutex,
                                      cachedStoreInfo.inspectionSucceeded ? StoreState::Ready
                                                                          : GetFallbackStoreState(ClassifyUninspectableStore(cachedStoreInfo)),
                                      SyncPhase::Idle,
                                      QueryExecutionMode::Unknown,
                                      FallbackReason::None,
                                      effectivePlan.rootPath);
        ApplyRuntimeStatusToProgress(progress, CaptureRepositoryRuntimeStatus(&_runtimeStatus, &_statusMutex));

        hr = ReportRepositoryProgress(progress, FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP, effectivePlan.rootPath, S_OK, true);
        if (FAILED(hr))
        {
            return hr;
        }

        const auto executeStart       = std::chrono::steady_clock::now();
        FallbackReason fallbackReason = FallbackReason::None;
        std::unordered_set<std::wstring> sqliteEmittedPathKeys;
        TrackingCandidateCallbackContext sqliteTrackingContext{
            .callback        = candidateCallback,
            .cookie          = candidateCookie,
            .emittedPathKeys = &sqliteEmittedPathKeys,
        };
        const RepositoryStatus runtimeStatusBeforeSqliteQuery = CaptureRepositoryRuntimeStatus(&_runtimeStatus, &_statusMutex);
        UpdateRepositoryRuntimeStatus(&_runtimeStatus,
                                      &_statusMutex,
                                      cachedStoreInfo.inspectionSucceeded && cachedStoreInfo.readyForQueryCutover ? StoreState::Ready
                                                                                                                  : runtimeStatusBeforeSqliteQuery.storeState,
                                      cachedStoreInfo.inspectionSucceeded && cachedStoreInfo.readyForQueryCutover ? SyncPhase::Watching : SyncPhase::Idle,
                                      QueryExecutionMode::Sqlite,
                                      FallbackReason::None,
                                      effectivePlan.rootPath);
        ApplyRuntimeStatusToProgress(progress, CaptureRepositoryRuntimeStatus(&_runtimeStatus, &_statusMutex));
        EmitPerfCount(L"search.backend.sqlite.attempts");
        {
            const Debug::Perf::Scope sqliteQueryPerf(L"search.backend.sqlite.query_ms");
            hr = TryEnumerateFromConfiguredSqliteStore(
                cachedStoreInfo, effectivePlan, cancelCheck, cancelCookie, &TrackCandidateAndForward, &sqliteTrackingContext, stats, progress, &fallbackReason);
        }
        if (hr == S_OK)
        {
            stats.queryExecutionMode     = QueryExecutionMode::Sqlite;
            stats.fallbackReason         = FallbackReason::None;
            stats.executeQueryDurationMs = ClampDurationMs(std::chrono::steady_clock::now() - executeStart);
            EmitPerfCount(L"search.backend.sqlite.successes");
            UpdateRepositoryRuntimeStatus(&_runtimeStatus,
                                          &_statusMutex,
                                          StoreState::Ready,
                                          SyncPhase::Watching,
                                          QueryExecutionMode::Sqlite,
                                          FallbackReason::None,
                                          effectivePlan.rootPath);
            if (outStats != nullptr)
            {
                *outStats = stats;
            }

            return S_OK;
        }
        if (FAILED(hr))
        {
            return hr;
        }

        EmitPerfCount(L"search.backend.sqlite.soft_fallbacks");
        EmitPerfCount(L"search.sqlite.partial_emit_count", static_cast<uint64_t>(sqliteEmittedPathKeys.size()));
        EmitPerfCount(L"search.sqlite.partial_emit_set_size", static_cast<uint64_t>(sqliteEmittedPathKeys.size()));

        std::shared_ptr<VolumeIndex> volume;
        hr = AcquireOrCreateVolume(support, volume);
        if (FAILED(hr))
        {
            return hr;
        }

        {
            EmitPerfCount(L"search.backend.inmemory.attempts");
            std::lock_guard volumeGuard(volume->mutex);
            if (fallbackReason == FallbackReason::None)
            {
                fallbackReason = volume->initialized ? FallbackReason::StoreStale : FallbackReason::WarmupRunning;
            }

            EmitPerfCount(GetFallbackReasonMetricName(fallbackReason));

            if (! sqliteAuthoritative && IsVolumeCurrentForNoWaitQuery(*volume, stats))
            {
                PopulateStatsFromVolume(*volume, stats);
                AssignRepositoryProgressCounts(progress, stats.directoryCount, stats.fileCount, 0u, 0u);
                UpdateRepositoryRuntimeStatus(&_runtimeStatus,
                                              &_statusMutex,
                                              StoreState::Ready,
                                              SyncPhase::Watching,
                                              QueryExecutionMode::InMemoryIndex,
                                              fallbackReason,
                                              effectivePlan.rootPath);
                ApplyRuntimeStatusToProgress(progress, CaptureRepositoryRuntimeStatus(&_runtimeStatus, &_statusMutex));
                hr = ReportRepositoryProgress(progress, FILESYSTEM_SEARCH_PHASE_ENUMERATING, effectivePlan.rootPath, S_OK, true);
                if (FAILED(hr))
                {
                    return hr;
                }

                {
                    const Debug::Perf::Scope inMemoryQueryPerf(L"search.backend.inmemory.query_ms");
                    hr = ExecuteQueryImpl(*volume,
                                          effectivePlan,
                                          cancelCheck,
                                          cancelCookie,
                                          stats,
                                          progress,
                                          [candidateCallback, candidateCookie](Candidate&& candidate) noexcept -> HRESULT
                    { return candidateCallback(&candidate, candidateCookie); });
                }
                if (FAILED(hr))
                {
                    return hr;
                }

                stats.queryExecutionMode     = QueryExecutionMode::InMemoryIndex;
                stats.fallbackReason         = fallbackReason;
                stats.executeQueryDurationMs = ClampDurationMs(std::chrono::steady_clock::now() - executeStart);
                if (outStats != nullptr)
                {
                    *outStats = stats;
                }

                return S_OK;
            }
        }

        stats.queryExecutionMode   = QueryExecutionMode::LiveScanFallback;
        stats.fallbackReason       = fallbackReason;
        stats.usedLiveScanFallback = true;
        if (fallbackReason == FallbackReason::SqliteFailure && ! sqliteEmittedPathKeys.empty())
        {
            stats.usedSqliteStore     = true;
            stats.sqliteReadOnlyQuery = true;
        }
        progress.latest.warningFlags |= FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX;
        EmitPerfCount(L"search.backend.live_fallback.attempts");
        UpdateRepositoryRuntimeStatus(&_runtimeStatus,
                                      &_statusMutex,
                                      GetFallbackStoreState(fallbackReason),
                                      SyncPhase::Idle,
                                      QueryExecutionMode::LiveScanFallback,
                                      fallbackReason,
                                      effectivePlan.rootPath);
        ApplyRuntimeStatusToProgress(progress, CaptureRepositoryRuntimeStatus(&_runtimeStatus, &_statusMutex));

        hr = ReportRepositoryProgress(progress, FILESYSTEM_SEARCH_PHASE_ENUMERATING, effectivePlan.rootPath, S_FALSE, true);
        if (FAILED(hr))
        {
            return hr;
        }

        uint64_t dedupeLookups  = 0u;
        uint64_t dedupeHits     = 0u;
        uint64_t dedupeLookupUs = 0u;
        {
            Debug::Perf::Scope liveFallbackPerf(L"search.backend.live_fallback.query_ms");
            hr = EnumerateLiveFileSystem(effectivePlan,
                                         cancelCheck,
                                         cancelCookie,
                                         stats,
                                         progress,
                                         [candidateCallback, candidateCookie, &sqliteEmittedPathKeys, &dedupeLookups, &dedupeHits, &dedupeLookupUs](
                                             Candidate&& candidate) noexcept -> HRESULT
            {
                try
                {
                    if (! sqliteEmittedPathKeys.empty())
                    {
                        const auto dedupeStart          = std::chrono::steady_clock::now();
                        const std::wstring candidateKey = FoldPathKey(candidate.fullPath);
                        const bool isDuplicate          = sqliteEmittedPathKeys.contains(candidateKey);
                        const auto dedupeElapsed        = std::chrono::steady_clock::now() - dedupeStart;
                        dedupeLookupUs += static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(dedupeElapsed).count());
                        ++dedupeLookups;
                        if (isDuplicate)
                        {
                            ++dedupeHits;
                            return kSkipCandidateHr;
                        }
                    }

                    return candidateCallback(&candidate, candidateCookie);
                }
                catch (const std::bad_alloc&)
                {
                    std::terminate();
                }
                catch (const std::exception&)
                {
                    Debug::Error(L"LocalSearchIndexCore: live fallback dedupe failed after SQLite query degradation.");
                    return E_FAIL;
                }
            });
            liveFallbackPerf.SetHr(hr);
        }
        EmitPerfCount(L"search.dedupe.lookups", dedupeLookups);
        EmitPerfCount(L"search.dedupe.hits", dedupeHits);
        Debug::Perf::Emit(L"search.dedupe.lookup_ms", L"", dedupeLookupUs, dedupeLookups, dedupeHits, S_OK);
        if (FAILED(hr))
        {
            return hr;
        }

        stats.executeQueryDurationMs = ClampDurationMs(std::chrono::steady_clock::now() - executeStart);
        if (outStats != nullptr)
        {
            *outStats = stats;
        }

        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::regex_error&)
    {
        return E_INVALIDARG;
    }
    catch (const std::exception&)
    {
        Debug::Error(L"LocalSearchIndexCore: EnumerateNoWait failed with an unexpected std::exception.");
        if (outStats != nullptr)
        {
            *outStats = {};
        }
        return E_FAIL;
    }
}

HRESULT Repository::InvalidateRoot(std::wstring_view rootPath, bool deleteSnapshot) noexcept
{
    try
    {
        SupportInfo support{};
        HRESULT hr = PopulateSupportInfo(rootPath, support);
        if (FAILED(hr))
        {
            return hr;
        }

        if (! support.indexable)
        {
            return kNotSupportedHr;
        }

        const std::wstring rootKey = FoldPathKey(support.normalizedRootPath);

        {
            std::lock_guard guard(_mutex);
            _volumes.erase(rootKey);
        }
        InvalidateCachedPersistentStoreInfo();

        if (deleteSnapshot)
        {
            std::unique_ptr<IIndexStore> store = CreateVolumeStore(support.normalizedRootPath, support.fileSystemKind, _options);
            if (! store)
            {
                return E_INVALIDARG;
            }

            hr = store->Delete();
            if (FAILED(hr))
            {
                return hr;
            }
        }

        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"LocalSearchIndexCore: InvalidateRoot failed with an unexpected std::exception.");
        return E_FAIL;
    }
}

HRESULT Repository::PrimeRoot(std::wstring_view rootPath) noexcept
{
    try
    {
        SupportInfo support{};
        HRESULT hr = PopulateSupportInfo(rootPath, support);
        if (FAILED(hr))
        {
            return hr;
        }

        std::shared_ptr<VolumeIndex> volume;
        hr = AcquireOrCreateVolume(support, volume);
        if (FAILED(hr))
        {
            return hr;
        }

        return volume ? S_OK : E_FAIL;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"LocalSearchIndexCore: PrimeRoot failed with an unexpected std::exception.");
        return E_FAIL;
    }
}

void Repository::CollectCachedRoots(std::vector<std::wstring>& outRoots) noexcept
{
    outRoots.clear();

    try
    {
        std::lock_guard guard(_mutex);
        outRoots.reserve(_volumes.size());
        for (const auto& [key, volume] : _volumes)
        {
            if (volume)
            {
                outRoots.push_back(volume->normalizedRootPath);
            }
        }
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"LocalSearchIndexCore: CollectCachedRoots failed with an unexpected std::exception.");
        outRoots.clear();
        return;
    }

    std::sort(
        outRoots.begin(), outRoots.end(), [](const std::wstring& left, const std::wstring& right) noexcept { return FoldPathKey(left) < FoldPathKey(right); });
}

HRESULT Repository::EnsureReadyForRoot(std::wstring_view rootPath,
                                       CancelCheckFn cancelCheck,
                                       void* cancelCookie,
                                       QueryStats* outStats,
                                       ProgressCallbackFn progressCallback,
                                       void* progressCookie) noexcept
{
    try
    {
        if (rootPath.empty())
        {
            return E_INVALIDARG;
        }

        QueryStats stats{};
        RepositoryProgressState progress{
            .callback = progressCallback,
            .cookie   = progressCookie,
        };

        SupportInfo support{};
        HRESULT hr = PopulateSupportInfo(rootPath, support);
        if (FAILED(hr))
        {
            return hr;
        }

        progress.latest.currentPath = support.normalizedRootPath;

        if (! support.indexable)
        {
            return kNotSupportedHr;
        }

        std::shared_ptr<VolumeIndex> volume;
        hr = AcquireOrCreateVolume(support, volume);
        if (FAILED(hr))
        {
            return hr;
        }

        std::lock_guard volumeGuard(volume->mutex);

        bool sqliteStoreChanged = false;
        hr = EnsureReady(*volume, _options, cancelCheck, cancelCookie, stats, progress, &_runtimeStatus, &_statusMutex, &sqliteStoreChanged);
        if (FAILED(hr))
        {
            return hr;
        }

        if (_options.persistentStoreKind == PersistentStoreKind::Sqlite)
        {
            bool refreshStoreInfo = sqliteStoreChanged;
            {
                std::lock_guard guard(_mutex);
                refreshStoreInfo = refreshStoreInfo || ! _cachedPersistentStoreInfoValid;
            }

            if (refreshStoreInfo)
            {
                RefreshCachedPersistentStoreInfo();
            }
        }

        if (outStats != nullptr)
        {
            *outStats = stats;
        }

        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"LocalSearchIndexCore: EnsureReadyForRoot failed with an unexpected std::exception.");
        if (outStats != nullptr)
        {
            *outStats = {};
        }
        return E_FAIL;
    }
}

#ifdef ENABLE_TESTS
HRESULT Repository::DropCachedVolumeForTests(std::wstring_view rootPath) noexcept
{
    return InvalidateRoot(rootPath, false);
}

HRESULT Repository::CorruptSnapshotForTests(std::wstring_view rootPath, SnapshotCorruptionMode mode) noexcept
{
    try
    {
        SupportInfo support{};
        HRESULT hr = PopulateSupportInfo(rootPath, support);
        if (FAILED(hr))
        {
            return hr;
        }

        if (! support.indexable)
        {
            return kNotSupportedHr;
        }

        std::unique_ptr<IIndexStore> store = CreateVolumeStore(support.normalizedRootPath, support.fileSystemKind, _options);
        if (! store)
        {
            return E_INVALIDARG;
        }

        return store->CorruptForTests(mode);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"LocalSearchIndexCore: CorruptSnapshotForTests failed with an unexpected std::exception.");
        return E_FAIL;
    }
}
#endif
} // namespace LocalSearchIndexCore
