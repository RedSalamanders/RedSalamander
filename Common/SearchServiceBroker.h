#pragma once

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "LocalSearchIndexCore.h"
#include "PlugInterfaces/FileSystem.h"

namespace SearchServiceBroker
{
inline constexpr uint32_t kProtocolVersion      = 3u;
inline constexpr wchar_t kPipeNameEnvVar[]      = L"REDSALAMANDER_SEARCH_SERVICE_PIPE";
inline constexpr wchar_t kDiscoverRootsEnvVar[] = L"REDSALAMANDER_SEARCH_SERVICE_DISCOVER_ROOTS";

#ifdef _DEBUG
inline constexpr wchar_t kServiceName[]     = L"RedSalamanderSearchService.Debug";
inline constexpr wchar_t kDefaultPipeName[] = LR"(\\.\pipe\RedSalamander.SearchService.Debug.v3)";
#else
inline constexpr wchar_t kServiceName[]     = L"RedSalamanderSearchService";
inline constexpr wchar_t kDefaultPipeName[] = LR"(\\.\pipe\RedSalamander.SearchService.v3)";
#endif

struct ServiceStatus final
{
    uint32_t protocolVersion = 0u;
    uint32_t processId       = 0u;
    std::wstring pipeName;
    std::wstring storageRootDirectory;
    LocalSearchIndexCore::PersistentStoreKind persistentStoreKind = LocalSearchIndexCore::PersistentStoreKind::Sqlite;
    std::wstring persistentStorePath;
    uint64_t persistentStoreBytes = 0u;
    std::wstring writeAheadLogPath;
    uint64_t writeAheadLogBytes               = 0u;
    uint64_t persistentStorePageCount         = 0u;
    uint64_t persistentStoreFreelistPageCount = 0u;
    std::wstring lastCheckpointUtc;
    std::wstring lastCompactionUtc;
    bool maintenanceQueued                      = false;
    bool maintenanceRunning                     = false;
    bool persistentStoreInspectionSucceeded     = false;
    bool readyForQueryCutover                   = false;
    uint64_t indexedVolumeCount                 = 0u;
    uint64_t indexedEntryCount                  = 0u;
    uint64_t legacyImportVolumeCount            = 0u;
    bool autoCheckpointEnabled                  = false;
    uint64_t autoCheckpointTargetBytes          = 0u;
    bool autoCompactionEnabled                  = false;
    uint32_t autoCompactionFragmentationPercent = 0u;
    uint64_t autoCompactionMinBytes             = 0u;
    uint64_t discoveredRootCount                = 0u;
    std::vector<std::wstring> discoveredRoots;
    bool startupWarmupEnabled            = false;
    bool startupWarmupRunning            = false;
    uint64_t startupWarmupTotalRoots     = 0u;
    uint64_t startupWarmupCompletedRoots = 0u;
    uint64_t startupWarmupFailedRoots    = 0u;
    std::wstring startupWarmupCurrentRoot;
    std::wstring startupWarmupLastFailureRoot;
    HRESULT startupWarmupLastFailureHr                          = S_OK;
    bool startupWarmupHasFailure                                = false;
    LocalSearchIndexCore::StoreState storeState                 = LocalSearchIndexCore::StoreState::Unknown;
    LocalSearchIndexCore::SyncPhase syncPhase                   = LocalSearchIndexCore::SyncPhase::Idle;
    LocalSearchIndexCore::QueryExecutionMode queryExecutionMode = LocalSearchIndexCore::QueryExecutionMode::Unknown;
    LocalSearchIndexCore::FallbackReason fallbackReason         = LocalSearchIndexCore::FallbackReason::None;
    uint64_t completedRoots                                     = 0u;
    uint64_t totalRoots                                         = 0u;
    std::wstring activeRoot;
};

struct QueryRequest final
{
    std::wstring rootPath;
    std::wstring namePattern;
    FileSystemSearchNameMode nameMode = FILESYSTEM_SEARCH_NAME_DISABLED;
    FileSystemSearchFlags flags       = FILESYSTEM_SEARCH_NONE;
    bool recursive                    = false;
    bool includeFiles                 = false;
    bool includeDirectories           = false;
    bool matchCaseName                = false;
    uint64_t maxResults               = 0u;
};

struct QueryProgress final
{
    FileSystemSearchPhase phase                                 = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
    uint32_t warningFlags                                       = FILESYSTEM_SEARCH_WARNING_NONE;
    HRESULT statusHint                                          = S_OK;
    uint64_t scannedDirectories                                 = 0u;
    uint64_t scannedFiles                                       = 0u;
    uint64_t candidateFiles                                     = 0u;
    uint64_t matchedEntries                                     = 0u;
    LocalSearchIndexCore::StoreState storeState                 = LocalSearchIndexCore::StoreState::Unknown;
    LocalSearchIndexCore::SyncPhase syncPhase                   = LocalSearchIndexCore::SyncPhase::Idle;
    LocalSearchIndexCore::QueryExecutionMode queryExecutionMode = LocalSearchIndexCore::QueryExecutionMode::Unknown;
    LocalSearchIndexCore::FallbackReason fallbackReason         = LocalSearchIndexCore::FallbackReason::None;
    uint64_t completedRoots                                     = 0u;
    uint64_t totalRoots                                         = 0u;
    std::wstring activeRoot;
    std::wstring currentPath;
};

using ProgressCallbackFn       = HRESULT(STDMETHODCALLTYPE*)(const QueryProgress* progress, void* cookie) noexcept;
using CandidateBatchCallbackFn = HRESULT(STDMETHODCALLTYPE*)(LocalSearchIndexCore::Candidate* candidates,
                                                             size_t count,
                                                             size_t* consumedCount,
                                                             void* cookie) noexcept;

enum ServerEventType : uint32_t
{
    SEARCH_SERVICE_SERVER_EVENT_STARTING = 0u,
    SEARCH_SERVICE_SERVER_EVENT_WAITING_FOR_CLIENT,
    SEARCH_SERVICE_SERVER_EVENT_CLIENT_CONNECTED,
    SEARCH_SERVICE_SERVER_EVENT_REQUEST_RECEIVED,
    SEARCH_SERVICE_SERVER_EVENT_QUERY_PROGRESS,
    SEARCH_SERVICE_SERVER_EVENT_QUERY_BATCH_SENT,
    SEARCH_SERVICE_SERVER_EVENT_QUERY_COMPLETED,
    SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_PROGRESS,
    SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_FAILED,
    SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_COMPLETED,
    SEARCH_SERVICE_SERVER_EVENT_MAINTENANCE_QUEUED,
    SEARCH_SERVICE_SERVER_EVENT_MAINTENANCE_RUNNING,
    SEARCH_SERVICE_SERVER_EVENT_MAINTENANCE_COMPLETED,
    SEARCH_SERVICE_SERVER_EVENT_REQUEST_HANDLED,
    SEARCH_SERVICE_SERVER_EVENT_STOPPING,
    SEARCH_SERVICE_SERVER_EVENT_STOPPED,
    SEARCH_SERVICE_SERVER_EVENT_ERROR,
};

enum ServerRequestType : uint32_t
{
    SEARCH_SERVICE_SERVER_REQUEST_NONE = 0u,
    SEARCH_SERVICE_SERVER_REQUEST_STATUS,
    SEARCH_SERVICE_SERVER_REQUEST_QUERY,
    SEARCH_SERVICE_SERVER_REQUEST_REBUILD,
    SEARCH_SERVICE_SERVER_REQUEST_COMPACT,
};

struct ServerEvent final
{
    ServerEventType type                        = SEARCH_SERVICE_SERVER_EVENT_STARTING;
    HRESULT hr                                  = S_OK;
    uint32_t handledRequests                    = 0u;
    ServerRequestType requestType               = SEARCH_SERVICE_SERVER_REQUEST_NONE;
    FileSystemSearchPhase phase                 = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
    uint32_t requestFlags                       = FILESYSTEM_SEARCH_NONE;
    uint32_t requestNameMode                    = FILESYSTEM_SEARCH_NAME_DISABLED;
    uint32_t batchesSent                        = 0u;
    uint64_t scannedDirectories                 = 0u;
    uint64_t scannedFiles                       = 0u;
    uint64_t candidateFiles                     = 0u;
    uint64_t matchedEntries                     = 0u;
    uint64_t snapshotFileBytes                  = 0u;
    uint64_t estimatedMemoryBytes               = 0u;
    uint64_t ensureReadyDurationMs              = 0u;
    uint64_t executeQueryDurationMs             = 0u;
    uint32_t startupWarmupTotalRoots            = 0u;
    uint32_t startupWarmupCompletedRoots        = 0u;
    uint32_t startupWarmupFailedRoots           = 0u;
    uint32_t warningFlags                       = FILESYSTEM_SEARCH_WARNING_NONE;
    uint32_t storeState                         = static_cast<uint32_t>(LocalSearchIndexCore::StoreState::Unknown);
    uint32_t syncPhase                          = static_cast<uint32_t>(LocalSearchIndexCore::SyncPhase::Idle);
    uint32_t queryExecutionMode                 = static_cast<uint32_t>(LocalSearchIndexCore::QueryExecutionMode::Unknown);
    uint32_t fallbackReason                     = static_cast<uint32_t>(LocalSearchIndexCore::FallbackReason::None);
    uint64_t completedRoots                     = 0u;
    uint64_t totalRoots                         = 0u;
    const wchar_t* rootPath                     = nullptr;
    const wchar_t* namePattern                  = nullptr;
    const wchar_t* currentPath                  = nullptr;
    const wchar_t* activeRoot                   = nullptr;
    const wchar_t* startupWarmupLastFailureRoot = nullptr;
    int32_t startupWarmupLastFailureHr          = S_OK;
};

using ServerEventCallbackFn = void(STDMETHODCALLTYPE*)(const ServerEvent* event, void* cookie) noexcept;

struct ServerOptions final
{
    std::wstring pipeName;
    std::wstring storageRootDirectory;
    LocalSearchIndexCore::PersistentStoreKind persistentStoreKind = LocalSearchIndexCore::PersistentStoreKind::Sqlite;
    std::wstring sqliteDatabasePath;
    uint32_t protocolVersion            = kProtocolVersion;
    uint32_t maxRequestsBeforeExit      = 0u;
    uint32_t disconnectAfterBatches     = 0u;
    bool allowRebuildRequests           = true;
    ServerEventCallbackFn eventCallback = nullptr;
    void* eventCookie                   = nullptr;
};

struct ServerRunResult final
{
    uint32_t handledRequests = 0u;
};

[[nodiscard]] std::wstring GetDefaultPipeName() noexcept;
[[nodiscard]] std::wstring GetConfiguredPipeName() noexcept;
[[nodiscard]] std::wstring GetProgramDataSearchIndexRoot() noexcept;

HRESULT GetStatus(ServiceStatus& outStatus) noexcept;
HRESULT Query(const QueryRequest& request,
              ProgressCallbackFn progressCallback,
              void* progressCookie,
              LocalSearchIndexCore::CancelCheckFn cancelCheck,
              void* cancelCookie,
              std::vector<LocalSearchIndexCore::Candidate>& outCandidates,
              LocalSearchIndexCore::QueryStats* outStats,
              CandidateBatchCallbackFn candidateBatchCallback = nullptr,
              void* candidateBatchCookie                      = nullptr) noexcept;
HRESULT RequestRebuild(std::wstring_view rootPath) noexcept;
HRESULT RequestCompact() noexcept;

HRESULT RunServer(const ServerOptions& options, HANDLE stopEvent, ServerRunResult* outResult) noexcept;
} // namespace SearchServiceBroker
