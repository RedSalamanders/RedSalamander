#include "SearchServiceBroker.h"

#include "Helpers.h"
#include "SqliteIndexStore.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <filesystem>
#include <iterator>
#include <limits>
#include <span>
#include <string_view>
#include <thread>
#include <unordered_map>
#include <vector>

#include <sddl.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

namespace SearchServiceBroker
{
namespace
{
constexpr uint32_t kMessageMagic                    = 0x53535252u; // "RRSS"
constexpr uint32_t kMaxFrameBytes                   = 16u * 1024u * 1024u;
constexpr DWORD kClientConnectTimeoutMs             = 150u;
constexpr HRESULT kProtocolErrorHr                  = HRESULT_FROM_WIN32(RPC_S_PROTOCOL_ERROR);
constexpr DWORD kClientIoPollMs                     = 50u;
constexpr DWORD kClientFrameTimeoutMs               = 30'000u;
constexpr DWORD kClientControlOperationTimeoutMs    = 30'000u;
constexpr DWORD kClientQueryOperationTimeoutMs      = 10u * 60u * 1'000u;
constexpr size_t kMaxClientBufferedCandidates       = 65'536u;
constexpr uint64_t kMaxClientBufferedCandidateBytes = 64ull * 1024ull * 1024ull;
constexpr DWORD kDefaultMissingPipeRetryWindowMs    = 250u;
constexpr DWORD kMaxMissingPipeRetryWindowMs        = 30'000u;
constexpr uint64_t kIdleMaintenanceGraceMs          = 1'000u;
constexpr wchar_t kStartupWarmupDelayMsEnvVar[]     = L"REDSALAMANDER_SEARCH_SERVICE_STARTUP_WARMUP_DELAY_MS";
constexpr wchar_t kServerFrameTimeoutMsEnvVar[]     = L"REDSALAMANDER_SEARCH_SERVICE_SERVER_FRAME_TIMEOUT_MS";

struct StartupWarmupCancelContext final
{
    HANDLE stopEvent = nullptr;
    std::stop_token stopToken{};
};

struct ClientIoContext final
{
    LocalSearchIndexCore::CancelCheckFn cancelCheck = nullptr;
    void* cancelCookie                              = nullptr;
    ULONGLONG operationDeadline                     = 0u;
    DWORD frameTimeoutMs                            = kClientFrameTimeoutMs;
};

struct ServerIoContext final
{
    HANDLE stopEvent     = nullptr;
    DWORD frameTimeoutMs = kClientFrameTimeoutMs;
};

enum class MessageType : uint32_t
{
    StatusRequest   = 1u,
    StatusResponse  = 2u,
    QueryRequest    = 3u,
    QueryProgress   = 4u,
    QueryBatch      = 5u,
    QueryComplete   = 6u,
    RebuildRequest  = 7u,
    Ack             = 8u,
    Error           = 9u,
    CompactRequest  = 10u,
    ShutdownRequest = 11u,
};

struct FrameHeader final
{
    uint32_t magic           = kMessageMagic;
    uint32_t protocolVersion = kProtocolVersion;
    uint32_t messageType     = 0u;
    uint32_t payloadBytes    = 0u;
};

struct StatusResponsePayload final
{
    uint32_t processId        = 0u;
    uint32_t pipeNameBytes    = 0u;
    uint32_t storageRootBytes = 0u;
    uint32_t reserved         = 0u;
};

struct StatusResponseExtendedPayload final
{
    uint32_t persistentStoreKind                = static_cast<uint32_t>(LocalSearchIndexCore::PersistentStoreKind::Sqlite);
    uint32_t persistentStorePathBytes           = 0u;
    uint32_t writeAheadLogPathBytes             = 0u;
    uint32_t flags                              = 0u;
    uint64_t persistentStoreBytes               = 0u;
    uint64_t writeAheadLogBytes                 = 0u;
    uint64_t autoCheckpointTargetBytes          = 0u;
    uint64_t indexedVolumeCount                 = 0u;
    uint64_t indexedEntryCount                  = 0u;
    uint64_t legacyImportVolumeCount            = 0u;
    uint32_t autoCompactionFragmentationPercent = 0u;
    uint32_t reserved                           = 0u;
    uint64_t autoCompactionMinBytes             = 0u;
};

struct StatusResponseMaintenancePayload final
{
    uint32_t lastCheckpointUtcBytes = 0u;
    uint32_t lastCompactionUtcBytes = 0u;
    uint32_t flags                  = 0u;
    uint32_t reserved               = 0u;
    uint64_t pageCount              = 0u;
    uint64_t freelistPageCount      = 0u;
};

struct StatusResponseDiscoveryPayload final
{
    uint32_t discoveredRootCount  = 0u;
    uint32_t discoveredRootsBytes = 0u;
};

struct StatusResponseStartupWarmupPayload final
{
    uint32_t flags            = 0u;
    uint32_t currentRootBytes = 0u;
    uint32_t totalRoots       = 0u;
    uint32_t completedRoots   = 0u;
    uint32_t failedRoots      = 0u;
};

struct StatusResponseStartupWarmupFailurePayload final
{
    int32_t lastFailureHr         = S_OK;
    uint32_t lastFailureRootBytes = 0u;
    uint32_t flags                = 0u;
    uint32_t reserved             = 0u;
};

struct StatusResponseRuntimePayload final
{
    uint32_t storeState         = static_cast<uint32_t>(LocalSearchIndexCore::StoreState::Unknown);
    uint32_t syncPhase          = static_cast<uint32_t>(LocalSearchIndexCore::SyncPhase::Idle);
    uint32_t queryExecutionMode = static_cast<uint32_t>(LocalSearchIndexCore::QueryExecutionMode::Unknown);
    uint32_t fallbackReason     = static_cast<uint32_t>(LocalSearchIndexCore::FallbackReason::None);
    uint32_t activeRootBytes    = 0u;
    uint32_t reserved           = 0u;
    uint64_t completedRoots     = 0u;
    uint64_t totalRoots         = 0u;
};

struct QueryRequestPayload final
{
    uint32_t nameMode         = 0u;
    uint32_t flags            = 0u;
    uint32_t rootPathBytes    = 0u;
    uint32_t namePatternBytes = 0u;
    uint64_t maxResults       = 0u;
};

struct ProgressPayload final
{
    uint32_t phase              = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
    uint32_t warningFlags       = FILESYSTEM_SEARCH_WARNING_NONE;
    int32_t statusHint          = S_OK;
    uint32_t currentPathBytes   = 0u;
    uint64_t scannedDirectories = 0u;
    uint64_t scannedFiles       = 0u;
    uint64_t candidateFiles     = 0u;
    uint64_t matchedEntries     = 0u;
};

struct ProgressRuntimePayload final
{
    uint32_t storeState         = static_cast<uint32_t>(LocalSearchIndexCore::StoreState::Unknown);
    uint32_t syncPhase          = static_cast<uint32_t>(LocalSearchIndexCore::SyncPhase::Idle);
    uint32_t queryExecutionMode = static_cast<uint32_t>(LocalSearchIndexCore::QueryExecutionMode::Unknown);
    uint32_t fallbackReason     = static_cast<uint32_t>(LocalSearchIndexCore::FallbackReason::None);
    uint32_t activeRootBytes    = 0u;
    uint32_t reserved           = 0u;
    uint64_t completedRoots     = 0u;
    uint64_t totalRoots         = 0u;
};

struct CandidateBatchHeader final
{
    uint32_t count    = 0u;
    uint32_t reserved = 0u;
};

struct CandidateEntryHeader final
{
    uint32_t fileAttributes     = 0u;
    uint32_t metadataFlags      = LocalSearchIndexCore::CANDIDATE_METADATA_NONE;
    uint32_t fullPathBytes      = 0u;
    uint32_t displayNameBytes   = 0u;
    int64_t creationTime100ns   = 0;
    int64_t lastAccessTime100ns = 0;
    int64_t endOfFile           = 0;
    int64_t lastWriteTime100ns  = 0;
    int64_t changeTime100ns     = 0;
    int64_t allocationSize      = 0;
};

struct QueryCompletePayload final
{
    int32_t result                  = S_OK;
    uint32_t fileSystemKind         = 0u;
    uint32_t flags                  = 0u;
    uint64_t entryCount             = 0u;
    uint64_t fileCount              = 0u;
    uint64_t directoryCount         = 0u;
    uint64_t candidateCount         = 0u;
    uint64_t nextUsn                = 0u;
    uint64_t journalId              = 0u;
    uint64_t snapshotFileBytes      = 0u;
    uint64_t estimatedMemoryBytes   = 0u;
    uint32_t ensureReadyDurationMs  = 0u;
    uint32_t executeQueryDurationMs = 0u;
};

struct QueryCompleteRuntimePayload final
{
    uint32_t queryExecutionMode = static_cast<uint32_t>(LocalSearchIndexCore::QueryExecutionMode::Unknown);
    uint32_t fallbackReason     = static_cast<uint32_t>(LocalSearchIndexCore::FallbackReason::None);
    uint32_t warningFlags       = FILESYSTEM_SEARCH_WARNING_NONE;
    uint32_t reserved           = 0u;
};

struct RebuildRequestPayload final
{
    uint32_t rootPathBytes = 0u;
};

struct AckPayload final
{
    int32_t result = S_OK;
};

struct ErrorPayload final
{
    int32_t result        = E_FAIL;
    uint32_t messageBytes = 0u;
};

enum StatusResponseExtendedFlags : uint32_t
{
    STATUS_RESPONSE_FLAG_AUTO_CHECKPOINT         = 0x1u,
    STATUS_RESPONSE_FLAG_AUTO_COMPACTION         = 0x2u,
    STATUS_RESPONSE_FLAG_STORE_INSPECTED         = 0x4u,
    STATUS_RESPONSE_FLAG_READY_FOR_QUERY_CUTOVER = 0x8u,
};

enum StatusResponseMaintenanceFlags : uint32_t
{
    STATUS_RESPONSE_MAINTENANCE_FLAG_QUEUED  = 0x1u,
    STATUS_RESPONSE_MAINTENANCE_FLAG_RUNNING = 0x2u,
};

enum StatusResponseStartupWarmupFlags : uint32_t
{
    STATUS_RESPONSE_STARTUP_WARMUP_FLAG_ENABLED     = 0x1u,
    STATUS_RESPONSE_STARTUP_WARMUP_FLAG_RUNNING     = 0x2u,
    STATUS_RESPONSE_STARTUP_WARMUP_FLAG_HAS_FAILURE = 0x4u,
};

enum QueryCompleteFlags : uint32_t
{
    QUERY_COMPLETE_FLAG_SNAPSHOT_LOADED               = 0x1u,
    QUERY_COMPLETE_FLAG_SNAPSHOT_SAVED                = 0x2u,
    QUERY_COMPLETE_FLAG_JOURNAL_AVAILABLE             = 0x4u,
    QUERY_COMPLETE_FLAG_JOURNAL_REPLAY_APPLIED        = 0x8u,
    QUERY_COMPLETE_FLAG_REBUILT_JOURNAL_ID_MISMATCH   = 0x10u,
    QUERY_COMPLETE_FLAG_REBUILT_JOURNAL_RANGE_INVALID = 0x20u,
    QUERY_COMPLETE_FLAG_REBUILT_SNAPSHOT_CORRUPTION   = 0x40u,
    QUERY_COMPLETE_FLAG_USED_NTFS_ENUMERATION         = 0x80u,
    QUERY_COMPLETE_FLAG_USED_TRAVERSAL_SEED           = 0x100u,
    QUERY_COMPLETE_FLAG_USED_SQLITE_STORE             = 0x200u,
    QUERY_COMPLETE_FLAG_SQLITE_CUTOVER_BLOCKED        = 0x400u,
    QUERY_COMPLETE_FLAG_USED_NAME_PREFILTER           = 0x800u,
    QUERY_COMPLETE_FLAG_SQLITE_QUERY_READ_ONLY        = 0x1000u,
    QUERY_COMPLETE_FLAG_HARDLINK_ALIAS_INCOMPLETE     = 0x2000u,
};

struct SessionContext final
{
    SessionContext()                                 = default;
    SessionContext(const SessionContext&)            = delete;
    SessionContext& operator=(const SessionContext&) = delete;
    SessionContext(SessionContext&&)                 = delete;
    SessionContext& operator=(SessionContext&&)      = delete;

    wil::unique_handle pipe;
    LocalSearchIndexCore::Repository* repository        = nullptr;
    struct ServerMaintenanceState* maintenanceState     = nullptr;
    struct ServerStartupWarmupState* startupWarmupState = nullptr;
    ServerOptions options;
    ServerIoContext ioContext;
    uint32_t handledRequests = 0u;
    bool shutdownRequested   = false;
};

struct ServerMaintenanceState final
{
    bool queued                = false;
    bool running               = false;
    uint64_t queuedSinceTickMs = 0u;
};

struct ServerStartupWarmupState final
{
    ServerStartupWarmupState()                                           = default;
    ServerStartupWarmupState(const ServerStartupWarmupState&)            = delete;
    ServerStartupWarmupState& operator=(const ServerStartupWarmupState&) = delete;
    ServerStartupWarmupState(ServerStartupWarmupState&&)                 = delete;
    ServerStartupWarmupState& operator=(ServerStartupWarmupState&&)      = delete;

    mutable std::mutex mutex;
    bool enabled            = false;
    bool running            = false;
    uint32_t totalRoots     = 0u;
    uint32_t completedRoots = 0u;
    uint32_t failedRoots    = 0u;
    std::wstring currentRoot;
    bool hasFailure       = false;
    HRESULT lastFailureHr = S_OK;
    std::wstring lastFailureRoot;
};

struct ServerStartupWarmupSnapshot final
{
    bool enabled            = false;
    bool running            = false;
    uint32_t totalRoots     = 0u;
    uint32_t completedRoots = 0u;
    uint32_t failedRoots    = 0u;
    std::wstring currentRoot;
    bool hasFailure       = false;
    HRESULT lastFailureHr = S_OK;
    std::wstring lastFailureRoot;
};

[[nodiscard]] std::wstring ReadEnvironmentVariableTrimmed(std::wstring_view name) noexcept
{
    return StringUtils::TrimWhitespaceCopy(EnvironmentVariables::Read(name).value_or(std::wstring{}));
}

[[nodiscard]] std::vector<std::wstring> ParseConfiguredRootList(std::wstring_view text)
{
    std::vector<std::wstring> roots;

    size_t tokenStart = 0u;
    auto appendToken  = [&](std::wstring_view token)
    {
        const std::wstring trimmed = StringUtils::TrimWhitespaceCopy(token);
        if (! trimmed.empty())
        {
            roots.push_back(trimmed);
        }
    };

    for (size_t index = 0u; index < text.size(); ++index)
    {
        const wchar_t ch = text[index];
        if (ch != L';' && ch != L'\r' && ch != L'\n')
        {
            continue;
        }

        appendToken(text.substr(tokenStart, index - tokenStart));
        tokenStart = index + 1u;
    }

    appendToken(text.substr(tokenStart));

    std::sort(roots.begin(), roots.end(), [](const std::wstring& left, const std::wstring& right) noexcept { return OrdinalString::LessNoCase(left, right); });
    roots.erase(std::unique(roots.begin(),
                            roots.end(),
                            [](const std::wstring& left, const std::wstring& right) noexcept { return OrdinalString::EqualsNoCase(left, right); }),
                roots.end());
    return roots;
}

struct StartupDiscoveryRoots final
{
    std::vector<std::wstring> roots;
    bool usedOverride = false;
};

[[nodiscard]] StartupDiscoveryRoots ResolveStartupDiscoveryRoots() noexcept
{
    StartupDiscoveryRoots resolved{};
    const std::wstring configuredOverride = ReadEnvironmentVariableTrimmed(kDiscoverRootsEnvVar);
    if (configuredOverride.empty())
    {
        resolved.roots = LocalSearchIndexCore::DiscoverFixedLocalRoots();
        return resolved;
    }

    resolved.roots        = ParseConfiguredRootList(configuredOverride);
    resolved.usedOverride = true;
    return resolved;
}

[[nodiscard]] uint32_t GetStartupWarmupDelayMs() noexcept
{
    const std::wstring configuredDelay = ReadEnvironmentVariableTrimmed(kStartupWarmupDelayMsEnvVar);
    if (configuredDelay.empty())
    {
        return 0u;
    }

    wchar_t* end               = nullptr;
    const unsigned long parsed = ::wcstoul(configuredDelay.c_str(), &end, 10);
    if (end == configuredDelay.c_str() || (end != nullptr && *end != L'\0'))
    {
        return 0u;
    }

    return static_cast<uint32_t>(parsed);
}

[[nodiscard]] DWORD GetServerFrameTimeoutMs() noexcept
{
    const std::wstring configuredTimeout = ReadEnvironmentVariableTrimmed(kServerFrameTimeoutMsEnvVar);
    if (configuredTimeout.empty())
    {
        return kClientFrameTimeoutMs;
    }

    wchar_t* end               = nullptr;
    const unsigned long parsed = ::wcstoul(configuredTimeout.c_str(), &end, 10);
    if (end == configuredTimeout.c_str() || (end != nullptr && *end != L'\0') || parsed == 0ul)
    {
        return kClientFrameTimeoutMs;
    }

    return static_cast<DWORD>((std::min<unsigned long>)(parsed, (std::numeric_limits<DWORD>::max)()));
}

[[nodiscard]] ServerStartupWarmupSnapshot CaptureStartupWarmupSnapshot(const ServerStartupWarmupState* state) noexcept
{
    if (state == nullptr)
    {
        return {};
    }

    std::lock_guard guard(state->mutex);
    return {
        .enabled         = state->enabled,
        .running         = state->running,
        .totalRoots      = state->totalRoots,
        .completedRoots  = state->completedRoots,
        .failedRoots     = state->failedRoots,
        .currentRoot     = state->currentRoot,
        .hasFailure      = state->hasFailure,
        .lastFailureHr   = state->lastFailureHr,
        .lastFailureRoot = state->lastFailureRoot,
    };
}

void UpdateStartupWarmupState(ServerStartupWarmupState& state,
                              const bool enabled,
                              const bool running,
                              const uint32_t totalRoots,
                              const uint32_t completedRoots,
                              const uint32_t failedRoots,
                              std::wstring_view currentRoot) noexcept
{
    std::lock_guard guard(state.mutex);
    state.enabled        = enabled;
    state.running        = running;
    state.totalRoots     = totalRoots;
    state.completedRoots = completedRoots;
    state.failedRoots    = failedRoots;
    state.currentRoot.assign(currentRoot);
}

void SetStartupWarmupFailure(ServerStartupWarmupState& state, const HRESULT hr, std::wstring_view rootPath) noexcept
{
    std::lock_guard guard(state.mutex);
    state.hasFailure    = true;
    state.lastFailureHr = hr;
    state.lastFailureRoot.assign(rootPath);
}

struct ServerQueryContext final
{
    HANDLE pipe                                   = nullptr;
    uint32_t protocolVersion                      = kProtocolVersion;
    const ServerIoContext* ioContext              = nullptr;
    uint32_t handledRequests                      = 0u;
    const ServerOptions* options                  = nullptr;
    const LocalSearchIndexCore::QueryStats* stats = nullptr;
    uint32_t requestFlags                         = FILESYSTEM_SEARCH_NONE;
    uint32_t requestNameMode                      = FILESYSTEM_SEARCH_NAME_DISABLED;
    const std::wstring* rootPath                  = nullptr;
    const std::wstring* namePattern               = nullptr;
};

struct ServerCandidateBatchState final
{
    HANDLE pipe                             = nullptr;
    uint32_t protocolVersion                = kProtocolVersion;
    const ServerIoContext* ioContext        = nullptr;
    uint32_t disconnectAfterBatches         = 0u;
    uint32_t batchesSent                    = 0u;
    uint32_t handledRequests                = 0u;
    const ServerOptions* options            = nullptr;
    size_t payloadBytes                     = sizeof(CandidateBatchHeader);
    bool sentIndexedProgress                = false;
    LocalSearchIndexCore::QueryStats* stats = nullptr;
    uint32_t warningFlags                   = FILESYSTEM_SEARCH_WARNING_NONE;
    uint32_t requestFlags                   = FILESYSTEM_SEARCH_NONE;
    uint32_t requestNameMode                = FILESYSTEM_SEARCH_NAME_DISABLED;
    std::wstring rootPath;
    std::wstring namePattern;
#if defined(RS_SEARCH_TEST_HOOKS)
    ServerTestHook testHook = ServerTestHook::None;
    bool testHookConsumed   = false;
#endif
    std::unordered_map<std::wstring, bool> clientDirectoryAccessCache;
    std::vector<LocalSearchIndexCore::Candidate> bufferedCandidates;
};

[[nodiscard]] HRESULT AuthorizeCandidateForClient(ServerCandidateBatchState& state, const LocalSearchIndexCore::Candidate& candidate);

[[nodiscard]] uint32_t QueryStatsWarningFlags(const LocalSearchIndexCore::QueryStats& stats) noexcept
{
    uint32_t warningFlags = stats.warningFlags;
    if (stats.usedLiveScanFallback)
    {
        warningFlags |= FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX;
    }

    return warningFlags;
}

[[nodiscard]] uint32_t QueryStateWarningFlags(const ServerCandidateBatchState& state) noexcept
{
    return (state.stats != nullptr ? QueryStatsWarningFlags(*state.stats) : FILESYSTEM_SEARCH_WARNING_NONE) | state.warningFlags;
}

struct ServerEventDetails final
{
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
    int32_t startupWarmupLastFailureHr          = S_OK;
    const wchar_t* rootPath                     = nullptr;
    const wchar_t* namePattern                  = nullptr;
    const wchar_t* currentPath                  = nullptr;
    const wchar_t* activeRoot                   = nullptr;
    const wchar_t* startupWarmupLastFailureRoot = nullptr;
};

[[nodiscard]] std::wstring Utf16FromBytes(const std::byte* bytes, uint32_t byteCount) noexcept
{
    if (bytes == nullptr || byteCount == 0u)
    {
        return {};
    }

    if ((byteCount % sizeof(wchar_t)) != 0u)
    {
        return {};
    }

    const wchar_t* text = reinterpret_cast<const wchar_t*>(bytes);
    return std::wstring(text, text + (byteCount / sizeof(wchar_t)));
}

void AppendBytes(std::vector<std::byte>& buffer, const void* data, size_t byteCount)
{
    const std::byte* begin = static_cast<const std::byte*>(data);
    buffer.insert(buffer.end(), begin, begin + byteCount);
}

void AppendUtf16(std::vector<std::byte>& buffer, std::wstring_view text)
{
    if (! text.empty())
    {
        AppendBytes(buffer, text.data(), text.size() * sizeof(wchar_t));
    }
}

template <typename T> [[nodiscard]] bool ReadPod(std::span<const std::byte>& remaining, T& out) noexcept
{
    if (remaining.size_bytes() < sizeof(T))
    {
        return false;
    }

    std::memcpy(&out, remaining.data(), sizeof(T));
    remaining = remaining.subspan(sizeof(T));
    return true;
}

[[nodiscard]] bool ReadUtf16(std::span<const std::byte>& remaining, uint32_t byteCount, std::wstring& out) noexcept
{
    if (remaining.size_bytes() < byteCount || (byteCount % sizeof(wchar_t)) != 0u)
    {
        out.clear();
        return false;
    }

    out       = Utf16FromBytes(remaining.data(), byteCount);
    remaining = remaining.subspan(byteCount);
    return true;
}

HRESULT DecodeQueryBatchPayload(std::span<const std::byte> payloadBytes, std::vector<LocalSearchIndexCore::Candidate>& outCandidates) noexcept
{
    outCandidates.clear();

    CandidateBatchHeader batchHeader{};
    std::span<const std::byte> remaining = payloadBytes;
    if (! ReadPod(remaining, batchHeader) || batchHeader.reserved != 0u)
    {
        return kProtocolErrorHr;
    }

    const size_t maxEntriesInPayload = remaining.size_bytes() / sizeof(CandidateEntryHeader);
    if (static_cast<size_t>(batchHeader.count) > maxEntriesInPayload)
    {
        return kProtocolErrorHr;
    }

    std::vector<LocalSearchIndexCore::Candidate> batchCandidates;
    batchCandidates.reserve(batchHeader.count);
    for (uint32_t index = 0u; index < batchHeader.count; ++index)
    {
        CandidateEntryHeader entryHeader{};
        LocalSearchIndexCore::Candidate candidate{};
        if (! ReadPod(remaining, entryHeader) || ! ReadUtf16(remaining, entryHeader.fullPathBytes, candidate.fullPath) ||
            ! ReadUtf16(remaining, entryHeader.displayNameBytes, candidate.displayName))
        {
            return kProtocolErrorHr;
        }

        candidate.fileAttributes      = entryHeader.fileAttributes;
        candidate.metadataFlags       = entryHeader.metadataFlags;
        candidate.creationTime100ns   = entryHeader.creationTime100ns;
        candidate.lastAccessTime100ns = entryHeader.lastAccessTime100ns;
        candidate.endOfFile           = entryHeader.endOfFile;
        candidate.lastWriteTime100ns  = entryHeader.lastWriteTime100ns;
        candidate.changeTime100ns     = entryHeader.changeTime100ns;
        candidate.allocationSize      = entryHeader.allocationSize;
        batchCandidates.push_back(std::move(candidate));
    }

    outCandidates = std::move(batchCandidates);
    return S_OK;
}

[[nodiscard]] size_t ResolveClientBufferedCandidateLimit(const uint64_t requestMaxResults) noexcept
{
    if (requestMaxResults == 0u)
    {
        return kMaxClientBufferedCandidates;
    }

    const uint64_t capped = (std::min)(requestMaxResults, static_cast<uint64_t>(kMaxClientBufferedCandidates));
    return static_cast<size_t>(capped);
}

[[nodiscard]] uint64_t EstimateClientBufferedCandidateBytes(const LocalSearchIndexCore::Candidate& candidate) noexcept
{
    return static_cast<uint64_t>(sizeof(LocalSearchIndexCore::Candidate)) + (static_cast<uint64_t>(candidate.fullPath.size()) * sizeof(wchar_t)) +
           (static_cast<uint64_t>(candidate.displayName.size()) * sizeof(wchar_t));
}

[[nodiscard]] bool CanAppendClientBufferedCandidates(const size_t currentCount,
                                                     const uint64_t currentBytes,
                                                     std::span<const LocalSearchIndexCore::Candidate> batch,
                                                     const size_t maxCount,
                                                     uint64_t& outBatchBytes) noexcept
{
    outBatchBytes = 0u;
    if (batch.size() > maxCount || currentCount > (maxCount - batch.size()))
    {
        return false;
    }

    for (const auto& candidate : batch)
    {
        const uint64_t candidateBytes = EstimateClientBufferedCandidateBytes(candidate);
        if (candidateBytes > kMaxClientBufferedCandidateBytes || outBatchBytes > (kMaxClientBufferedCandidateBytes - candidateBytes))
        {
            return false;
        }

        outBatchBytes += candidateBytes;
    }

    return currentBytes <= kMaxClientBufferedCandidateBytes && outBatchBytes <= (kMaxClientBufferedCandidateBytes - currentBytes);
}

[[nodiscard]] LocalSearchIndexCore::StoreState ResolveFallbackStoreState(const LocalSearchIndexCore::FallbackReason reason) noexcept
{
    switch (reason)
    {
        case LocalSearchIndexCore::FallbackReason::StoreMissing:
        case LocalSearchIndexCore::FallbackReason::StoreInvalid: return LocalSearchIndexCore::StoreState::Invalid;
        case LocalSearchIndexCore::FallbackReason::SqliteFailure: return LocalSearchIndexCore::StoreState::Recovering;
        case LocalSearchIndexCore::FallbackReason::StoreStale:
        case LocalSearchIndexCore::FallbackReason::CutoverBlocked:
        case LocalSearchIndexCore::FallbackReason::WarmupRunning: return LocalSearchIndexCore::StoreState::Syncing;
        case LocalSearchIndexCore::FallbackReason::None:
        default: return LocalSearchIndexCore::StoreState::Unknown;
    }
}

[[nodiscard]] LocalSearchIndexCore::FallbackReason ClassifyPersistentStoreFallbackReason(const LocalSearchIndexCore::PersistentStoreInfo& storeInfo) noexcept
{
    if (storeInfo.kind != LocalSearchIndexCore::PersistentStoreKind::Sqlite)
    {
        return LocalSearchIndexCore::FallbackReason::None;
    }

    if (! storeInfo.inspectionSucceeded)
    {
        std::error_code existsEc;
        const bool databaseExists = ! storeInfo.primaryPath.empty() && std::filesystem::exists(storeInfo.primaryPath, existsEc);
        return databaseExists ? LocalSearchIndexCore::FallbackReason::StoreInvalid : LocalSearchIndexCore::FallbackReason::StoreMissing;
    }

    if (! storeInfo.readyForQueryCutover)
    {
        return LocalSearchIndexCore::FallbackReason::CutoverBlocked;
    }

    return LocalSearchIndexCore::FallbackReason::None;
}

[[nodiscard]] bool HasRepositoryRuntimeStatus(const LocalSearchIndexCore::RepositoryStatus& status) noexcept
{
    return status.storeState != LocalSearchIndexCore::StoreState::Unknown || status.syncPhase != LocalSearchIndexCore::SyncPhase::Idle ||
           status.queryExecutionMode != LocalSearchIndexCore::QueryExecutionMode::Unknown ||
           status.fallbackReason != LocalSearchIndexCore::FallbackReason::None || status.completedRoots != 0u || status.totalRoots != 0u ||
           ! status.activeRoot.empty();
}

[[nodiscard]] LocalSearchIndexCore::StoreState DeriveRuntimeStoreState(const LocalSearchIndexCore::RepositoryStatus& repositoryStatus,
                                                                       const LocalSearchIndexCore::PersistentStoreInfo& storeInfo,
                                                                       const ServerMaintenanceState* maintenanceState) noexcept
{
    const LocalSearchIndexCore::FallbackReason persistentFallbackReason = ClassifyPersistentStoreFallbackReason(storeInfo);
    const bool staleReadyRuntime =
        repositoryStatus.storeState == LocalSearchIndexCore::StoreState::Ready && persistentFallbackReason != LocalSearchIndexCore::FallbackReason::None;
    if (repositoryStatus.storeState != LocalSearchIndexCore::StoreState::Unknown && ! staleReadyRuntime)
    {
        return repositoryStatus.storeState;
    }

    if (maintenanceState != nullptr && maintenanceState->running)
    {
        return LocalSearchIndexCore::StoreState::Maintenance;
    }

    if (persistentFallbackReason != LocalSearchIndexCore::FallbackReason::None)
    {
        return ResolveFallbackStoreState(persistentFallbackReason);
    }

    return storeInfo.inspectionSucceeded ? LocalSearchIndexCore::StoreState::Ready : LocalSearchIndexCore::StoreState::Unknown;
}

[[nodiscard]] LocalSearchIndexCore::SyncPhase DeriveRuntimeSyncPhase(const LocalSearchIndexCore::RepositoryStatus& repositoryStatus,
                                                                     const LocalSearchIndexCore::PersistentStoreInfo& storeInfo,
                                                                     const ServerStartupWarmupSnapshot& warmupSnapshot) noexcept
{
    const bool staleReadyRuntime = repositoryStatus.storeState == LocalSearchIndexCore::StoreState::Ready && ! storeInfo.readyForQueryCutover;
    if (HasRepositoryRuntimeStatus(repositoryStatus) && ! staleReadyRuntime)
    {
        return repositoryStatus.syncPhase;
    }

    if (warmupSnapshot.running)
    {
        if (! warmupSnapshot.currentRoot.empty())
        {
            return LocalSearchIndexCore::SyncPhase::Loading;
        }

        if (warmupSnapshot.totalRoots != 0u)
        {
            return LocalSearchIndexCore::SyncPhase::Discovering;
        }
    }

    if (storeInfo.inspectionSucceeded && storeInfo.readyForQueryCutover)
    {
        return LocalSearchIndexCore::SyncPhase::Watching;
    }

    return LocalSearchIndexCore::SyncPhase::Idle;
}

[[nodiscard]] bool IsDeadlineExpired(ULONGLONG deadline, ULONGLONG now) noexcept;
[[nodiscard]] DWORD RemainingToDeadlineMs(ULONGLONG deadline, ULONGLONG now) noexcept;

HRESULT CancelServerIoAndReturn(HANDLE handle, OVERLAPPED& overlapped, HRESULT hr) noexcept
{
    static_cast<void>(::CancelIoEx(handle, &overlapped));
    DWORD ignoredBytes = 0u;
    static_cast<void>(::GetOverlappedResult(handle, &overlapped, &ignoredBytes, TRUE));
    return hr;
}

HRESULT WaitForServerIo(HANDLE handle, OVERLAPPED& overlapped, const ServerIoContext* context, DWORD& outBytesTransferred) noexcept
{
    outBytesTransferred           = 0u;
    const ULONGLONG frameStart    = ::GetTickCount64();
    const DWORD frameTimeoutMs    = context != nullptr ? context->frameTimeoutMs : kClientFrameTimeoutMs;
    const ULONGLONG frameDeadline = frameTimeoutMs == INFINITE ? (std::numeric_limits<ULONGLONG>::max)() : frameStart + frameTimeoutMs;

    for (;;)
    {
        const ULONGLONG now = ::GetTickCount64();
        if (IsDeadlineExpired(frameDeadline, now))
        {
            return CancelServerIoAndReturn(handle, overlapped, HRESULT_FROM_WIN32(ERROR_SEM_TIMEOUT));
        }

        DWORD waitMs = (std::min)(kClientIoPollMs, RemainingToDeadlineMs(frameDeadline, now));
        if (waitMs == 0u)
        {
            return CancelServerIoAndReturn(handle, overlapped, HRESULT_FROM_WIN32(ERROR_SEM_TIMEOUT));
        }

        HANDLE waits[2]        = {overlapped.hEvent, context != nullptr ? context->stopEvent : nullptr};
        const DWORD waitCount  = (waits[1] != nullptr) ? 2u : 1u;
        const DWORD waitResult = ::WaitForMultipleObjects(waitCount, waits, FALSE, waitMs);
        if (waitResult == WAIT_TIMEOUT)
        {
            continue;
        }
        if (waitResult == WAIT_FAILED)
        {
            return CancelServerIoAndReturn(handle, overlapped, HRESULT_FROM_WIN32(::GetLastError()));
        }
        if (waitResult == WAIT_OBJECT_0 + 1u)
        {
            return CancelServerIoAndReturn(handle, overlapped, HRESULT_FROM_WIN32(ERROR_CANCELLED));
        }
        if (waitResult != WAIT_OBJECT_0)
        {
            return CancelServerIoAndReturn(handle, overlapped, E_FAIL);
        }

        if (::GetOverlappedResult(handle, &overlapped, &outBytesTransferred, FALSE) == 0)
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }
        return S_OK;
    }
}

HRESULT ServerIoExact(HANDLE handle, void* buffer, uint32_t byteCount, bool write, const ServerIoContext* context) noexcept
{
    auto* bytes    = static_cast<std::byte*>(buffer);
    uint32_t total = 0u;
    wil::unique_event_nothrow ioEvent;
    ioEvent.reset(::CreateEventW(nullptr, TRUE, FALSE, nullptr));
    if (! ioEvent)
    {
        const DWORD lastError = ::GetLastError();
        return lastError != 0u ? HRESULT_FROM_WIN32(lastError) : E_OUTOFMEMORY;
    }

    while (total < byteCount)
    {
        if (context != nullptr && context->stopEvent != nullptr && ::WaitForSingleObject(context->stopEvent, 0u) == WAIT_OBJECT_0)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        static_cast<void>(::ResetEvent(ioEvent.get()));
        OVERLAPPED overlapped{};
        overlapped.hEvent = ioEvent.get();

        DWORD transferred     = 0u;
        const DWORD remaining = byteCount - total;
        const BOOL started    = write ? ::WriteFile(handle, bytes + total, remaining, &transferred, &overlapped)
                                      : ::ReadFile(handle, bytes + total, remaining, &transferred, &overlapped);
        if (started == 0)
        {
            const DWORD error = ::GetLastError();
            if (error != ERROR_IO_PENDING)
            {
                return HRESULT_FROM_WIN32(error);
            }

            HRESULT hr = WaitForServerIo(handle, overlapped, context, transferred);
            if (FAILED(hr))
            {
                return hr;
            }
        }

        if (transferred == 0u)
        {
            return HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE);
        }

        total += transferred;
    }

    return S_OK;
}

HRESULT ReadExact(HANDLE handle, void* buffer, uint32_t byteCount, const ServerIoContext* context) noexcept
{
    return ServerIoExact(handle, buffer, byteCount, false, context);
}

HRESULT WriteExact(HANDLE handle, const void* buffer, uint32_t byteCount, const ServerIoContext* context) noexcept
{
    return ServerIoExact(handle, const_cast<void*>(buffer), byteCount, true, context);
}

HRESULT SendFrame(
    HANDLE handle, MessageType messageType, uint32_t protocolVersion, const std::vector<std::byte>& payload, const ServerIoContext* context) noexcept
{
    if (payload.size() > kMaxFrameBytes)
    {
        return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
    }

    FrameHeader header{};
    header.protocolVersion = protocolVersion;
    header.messageType     = static_cast<uint32_t>(messageType);
    header.payloadBytes    = static_cast<uint32_t>(payload.size());

    HRESULT hr = WriteExact(handle, &header, static_cast<uint32_t>(sizeof(header)), context);
    if (FAILED(hr))
    {
        return hr;
    }

    if (! payload.empty())
    {
        hr = WriteExact(handle, payload.data(), static_cast<uint32_t>(payload.size()), context);
    }
    return hr;
}

HRESULT ReceiveFrame(HANDLE handle, FrameHeader& outHeader, std::vector<std::byte>& outPayload, const ServerIoContext* context) noexcept
{
    outHeader = {};
    outPayload.clear();

    HRESULT hr = ReadExact(handle, &outHeader, static_cast<uint32_t>(sizeof(outHeader)), context);
    if (FAILED(hr))
    {
        return hr;
    }

    if (outHeader.magic != kMessageMagic)
    {
        return kProtocolErrorHr;
    }

    if (outHeader.payloadBytes > kMaxFrameBytes)
    {
        return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
    }

    outPayload.resize(outHeader.payloadBytes);
    if (outHeader.payloadBytes != 0u)
    {
        hr = ReadExact(handle, outPayload.data(), outHeader.payloadBytes, context);
    }
    return hr;
}

[[nodiscard]] ULONGLONG MakeDeadline(DWORD timeoutMs) noexcept
{
    return timeoutMs == INFINITE ? (std::numeric_limits<ULONGLONG>::max)() : ::GetTickCount64() + timeoutMs;
}

[[nodiscard]] bool IsDeadlineExpired(ULONGLONG deadline, ULONGLONG now) noexcept
{
    return deadline != (std::numeric_limits<ULONGLONG>::max)() && now >= deadline;
}

[[nodiscard]] DWORD RemainingToDeadlineMs(ULONGLONG deadline, ULONGLONG now) noexcept
{
    if (deadline == (std::numeric_limits<ULONGLONG>::max)())
    {
        return INFINITE;
    }
    if (now >= deadline)
    {
        return 0u;
    }

    const ULONGLONG remaining = deadline - now;
    return remaining > (std::numeric_limits<DWORD>::max)() ? (std::numeric_limits<DWORD>::max)() : static_cast<DWORD>(remaining);
}

[[nodiscard]] HRESULT CheckClientIoCancelled(const ClientIoContext* context) noexcept
{
    if (context == nullptr || context->cancelCheck == nullptr)
    {
        return S_OK;
    }

    return context->cancelCheck(context->cancelCookie);
}

HRESULT CancelClientIoAndReturn(HANDLE handle, OVERLAPPED& overlapped, HRESULT hr) noexcept
{
    static_cast<void>(::CancelIoEx(handle, &overlapped));
    DWORD ignoredBytes = 0u;
    static_cast<void>(::GetOverlappedResult(handle, &overlapped, &ignoredBytes, TRUE));
    return hr;
}

HRESULT WaitForClientIo(HANDLE handle, OVERLAPPED& overlapped, ClientIoContext* context, DWORD& outBytesTransferred) noexcept
{
    outBytesTransferred           = 0u;
    const ULONGLONG frameStart    = ::GetTickCount64();
    const ULONGLONG frameDeadline = context != nullptr ? frameStart + context->frameTimeoutMs : frameStart + kClientFrameTimeoutMs;

    for (;;)
    {
        HRESULT cancelHr = CheckClientIoCancelled(context);
        if (FAILED(cancelHr))
        {
            return CancelClientIoAndReturn(handle, overlapped, cancelHr);
        }

        const ULONGLONG now = ::GetTickCount64();
        if (IsDeadlineExpired(frameDeadline, now) || (context != nullptr && IsDeadlineExpired(context->operationDeadline, now)))
        {
            return CancelClientIoAndReturn(handle, overlapped, HRESULT_FROM_WIN32(ERROR_SEM_TIMEOUT));
        }

        DWORD waitMs = kClientIoPollMs;
        waitMs       = (std::min)(waitMs, RemainingToDeadlineMs(frameDeadline, now));
        if (context != nullptr)
        {
            waitMs = (std::min)(waitMs, RemainingToDeadlineMs(context->operationDeadline, now));
        }
        if (waitMs == 0u)
        {
            return CancelClientIoAndReturn(handle, overlapped, HRESULT_FROM_WIN32(ERROR_SEM_TIMEOUT));
        }

        const DWORD waitResult = ::WaitForSingleObject(overlapped.hEvent, waitMs);
        if (waitResult == WAIT_TIMEOUT)
        {
            continue;
        }
        if (waitResult == WAIT_FAILED)
        {
            return CancelClientIoAndReturn(handle, overlapped, HRESULT_FROM_WIN32(::GetLastError()));
        }
        if (waitResult != WAIT_OBJECT_0)
        {
            return CancelClientIoAndReturn(handle, overlapped, E_FAIL);
        }

        if (::GetOverlappedResult(handle, &overlapped, &outBytesTransferred, FALSE) == 0)
        {
            return HRESULT_FROM_WIN32(::GetLastError());
        }
        return S_OK;
    }
}

HRESULT ClientIoExact(HANDLE handle, void* buffer, uint32_t byteCount, bool write, ClientIoContext* context) noexcept
{
    auto* bytes    = static_cast<std::byte*>(buffer);
    uint32_t total = 0u;
    wil::unique_event_nothrow ioEvent;
    ioEvent.reset(::CreateEventW(nullptr, TRUE, FALSE, nullptr));
    if (! ioEvent)
    {
        const DWORD lastError = ::GetLastError();
        return lastError != 0u ? HRESULT_FROM_WIN32(lastError) : E_OUTOFMEMORY;
    }

    while (total < byteCount)
    {
        HRESULT cancelHr = CheckClientIoCancelled(context);
        if (FAILED(cancelHr))
        {
            return cancelHr;
        }

        static_cast<void>(::ResetEvent(ioEvent.get()));
        OVERLAPPED overlapped{};
        overlapped.hEvent = ioEvent.get();

        DWORD transferred     = 0u;
        const DWORD remaining = byteCount - total;
        const BOOL started    = write ? ::WriteFile(handle, bytes + total, remaining, &transferred, &overlapped)
                                      : ::ReadFile(handle, bytes + total, remaining, &transferred, &overlapped);
        if (started == 0)
        {
            const DWORD error = ::GetLastError();
            if (error != ERROR_IO_PENDING)
            {
                return HRESULT_FROM_WIN32(error);
            }

            HRESULT hr = WaitForClientIo(handle, overlapped, context, transferred);
            if (FAILED(hr))
            {
                return hr;
            }
        }

        if (transferred == 0u)
        {
            return HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE);
        }

        total += transferred;
    }

    return S_OK;
}

HRESULT ClientReadExact(HANDLE handle, void* buffer, uint32_t byteCount, ClientIoContext* context) noexcept
{
    return ClientIoExact(handle, buffer, byteCount, false, context);
}

HRESULT ClientWriteExact(HANDLE handle, const void* buffer, uint32_t byteCount, ClientIoContext* context) noexcept
{
    return ClientIoExact(handle, const_cast<void*>(buffer), byteCount, true, context);
}

HRESULT ClientSendFrame(
    HANDLE handle, MessageType messageType, uint32_t protocolVersion, const std::vector<std::byte>& payload, ClientIoContext* context) noexcept
{
    if (payload.size() > kMaxFrameBytes)
    {
        return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
    }

    FrameHeader header{};
    header.protocolVersion = protocolVersion;
    header.messageType     = static_cast<uint32_t>(messageType);
    header.payloadBytes    = static_cast<uint32_t>(payload.size());

    HRESULT hr = ClientWriteExact(handle, &header, static_cast<uint32_t>(sizeof(header)), context);
    if (FAILED(hr))
    {
        return hr;
    }

    return payload.empty() ? S_OK : ClientWriteExact(handle, payload.data(), static_cast<uint32_t>(payload.size()), context);
}

HRESULT ClientReceiveFrame(HANDLE handle, FrameHeader& outHeader, std::vector<std::byte>& outPayload, ClientIoContext* context) noexcept
{
    outHeader = {};
    outPayload.clear();

    HRESULT hr = ClientReadExact(handle, &outHeader, static_cast<uint32_t>(sizeof(outHeader)), context);
    if (FAILED(hr))
    {
        return hr;
    }

    if (outHeader.magic != kMessageMagic)
    {
        return kProtocolErrorHr;
    }

    if (outHeader.payloadBytes > kMaxFrameBytes)
    {
        return HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW);
    }

    outPayload.resize(outHeader.payloadBytes);
    return outHeader.payloadBytes == 0u ? S_OK : ClientReadExact(handle, outPayload.data(), outHeader.payloadBytes, context);
}

[[nodiscard]] std::wstring BuildPipeName(std::wstring_view configuredName) noexcept
{
    if (! configuredName.empty())
    {
        return std::wstring(configuredName);
    }

    return std::wstring(kDefaultPipeName);
}

[[nodiscard]] LocalSearchIndexCore::RepositoryOptions BuildRepositoryOptions(const ServerOptions& options) noexcept
{
    LocalSearchIndexCore::RepositoryOptions repositoryOptions{};
    repositoryOptions.snapshotRootDirectory = options.storageRootDirectory;
    repositoryOptions.persistentStoreKind   = options.persistentStoreKind;
    repositoryOptions.sqliteDatabasePath    = options.sqliteDatabasePath;
    repositoryOptions.sqliteAuthoritative   = options.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite;
    return repositoryOptions;
}

void NotifyServerEvent(
    const ServerOptions& options, ServerEventType type, HRESULT hr, uint32_t handledRequests, const ServerEventDetails* details = nullptr) noexcept
{
    if (options.eventCallback == nullptr)
    {
        return;
    }

    ServerEvent event{};
    event.type            = type;
    event.hr              = hr;
    event.handledRequests = handledRequests;
    if (details != nullptr)
    {
        event.requestType                  = details->requestType;
        event.phase                        = details->phase;
        event.requestFlags                 = details->requestFlags;
        event.requestNameMode              = details->requestNameMode;
        event.batchesSent                  = details->batchesSent;
        event.scannedDirectories           = details->scannedDirectories;
        event.scannedFiles                 = details->scannedFiles;
        event.candidateFiles               = details->candidateFiles;
        event.matchedEntries               = details->matchedEntries;
        event.snapshotFileBytes            = details->snapshotFileBytes;
        event.estimatedMemoryBytes         = details->estimatedMemoryBytes;
        event.ensureReadyDurationMs        = details->ensureReadyDurationMs;
        event.executeQueryDurationMs       = details->executeQueryDurationMs;
        event.startupWarmupTotalRoots      = details->startupWarmupTotalRoots;
        event.startupWarmupCompletedRoots  = details->startupWarmupCompletedRoots;
        event.startupWarmupFailedRoots     = details->startupWarmupFailedRoots;
        event.warningFlags                 = details->warningFlags;
        event.storeState                   = details->storeState;
        event.syncPhase                    = details->syncPhase;
        event.queryExecutionMode           = details->queryExecutionMode;
        event.fallbackReason               = details->fallbackReason;
        event.completedRoots               = details->completedRoots;
        event.totalRoots                   = details->totalRoots;
        event.startupWarmupLastFailureHr   = details->startupWarmupLastFailureHr;
        event.rootPath                     = details->rootPath;
        event.namePattern                  = details->namePattern;
        event.currentPath                  = details->currentPath;
        event.activeRoot                   = details->activeRoot;
        event.startupWarmupLastFailureRoot = details->startupWarmupLastFailureRoot;
    }
    options.eventCallback(&event, options.eventCookie);
}

void SetQueuedMaintenanceState(ServerMaintenanceState& state, const bool queued) noexcept
{
    state.queued = queued;
    if (! queued)
    {
        state.queuedSinceTickMs = 0u;
    }
}

[[nodiscard]] uint64_t EstimateSqlitePageBytes(const LocalSearchIndexCore::PersistentStoreInfo& storeInfo) noexcept
{
    if (storeInfo.pageCount == 0u || storeInfo.primaryBytes == 0u)
    {
        return 0u;
    }

    return std::max<uint64_t>(1u, storeInfo.primaryBytes / storeInfo.pageCount);
}

[[nodiscard]] bool ShouldQueueAutomaticMaintenance(const LocalSearchIndexCore::PersistentStoreInfo& storeInfo) noexcept
{
    if (storeInfo.kind != LocalSearchIndexCore::PersistentStoreKind::Sqlite || ! storeInfo.inspectionSucceeded)
    {
        return false;
    }

    const bool shouldCheckpoint =
        storeInfo.autoCheckpointEnabled && storeInfo.autoCheckpointTargetBytes != 0u && storeInfo.writeAheadLogBytes >= storeInfo.autoCheckpointTargetBytes;
    if (shouldCheckpoint)
    {
        return true;
    }

    if (! storeInfo.incrementalAutoVacuumEnabled)
    {
        return true;
    }

    if (! storeInfo.autoCompactionEnabled || storeInfo.pageCount == 0u || storeInfo.freelistPageCount == 0u)
    {
        return false;
    }

    const uint64_t fragmentationPercent = (storeInfo.freelistPageCount * 100u) / storeInfo.pageCount;
    const uint64_t reclaimableBytes     = EstimateSqlitePageBytes(storeInfo) * storeInfo.freelistPageCount;
    return fragmentationPercent >= storeInfo.autoCompactionFragmentationPercent || reclaimableBytes >= storeInfo.autoCompactionMinBytes;
}

enum class RepositoryStoreOverlayResult : uint8_t
{
    Compatible,
    Uninspectable,
    GenerationMismatch,
};

[[nodiscard]] RepositoryStoreOverlayResult OverlayFresherRepositoryStoreInfo(LocalSearchIndexCore::PersistentStoreInfo& storeInfo,
                                                                             LocalSearchIndexCore::Repository* repository) noexcept
{
    if (repository == nullptr || storeInfo.kind != LocalSearchIndexCore::PersistentStoreKind::Sqlite)
    {
        return RepositoryStoreOverlayResult::Compatible;
    }
    if (! storeInfo.inspectionSucceeded)
    {
        return RepositoryStoreOverlayResult::Uninspectable;
    }

    LocalSearchIndexCore::PersistentStoreInfo repositoryStoreInfo{};
    bool hasRepositoryStoreInfo = repository->TryGetCachedPersistentStoreInfo(repositoryStoreInfo) && repositoryStoreInfo.inspectionSucceeded;
    if (! hasRepositoryStoreInfo)
    {
        hasRepositoryStoreInfo = repository->TryBuildInMemoryPersistentStoreInfo(repositoryStoreInfo);
    }
    if (! hasRepositoryStoreInfo || repositoryStoreInfo.kind != storeInfo.kind ||
        ! OrdinalString::EqualsNoCase(repositoryStoreInfo.primaryPath, storeInfo.primaryPath))
    {
        return RepositoryStoreOverlayResult::Compatible;
    }

    if (repositoryStoreInfo.storeGeneration != storeInfo.storeGeneration)
    {
        // The on-disk inspection is the coherent snapshot for this response. A generation mismatch can mean
        // external rotation, so cached or in-memory runtime counts must not overwrite it.
        return RepositoryStoreOverlayResult::GenerationMismatch;
    }

    if (repositoryStoreInfo.indexedVolumeCount <= storeInfo.indexedVolumeCount && repositoryStoreInfo.indexedEntryCount <= storeInfo.indexedEntryCount)
    {
        return RepositoryStoreOverlayResult::Compatible;
    }

    storeInfo.inspectionSucceeded     = true;
    storeInfo.readyForQueryCutover    = repositoryStoreInfo.readyForQueryCutover;
    storeInfo.indexedVolumeCount      = repositoryStoreInfo.indexedVolumeCount;
    storeInfo.indexedEntryCount       = repositoryStoreInfo.indexedEntryCount;
    storeInfo.legacyImportVolumeCount = repositoryStoreInfo.legacyImportVolumeCount;
    return RepositoryStoreOverlayResult::Compatible;
}

void RefreshAutomaticMaintenanceQueue(const ServerOptions& options, ServerMaintenanceState& maintenanceState, uint32_t handledRequests) noexcept
{
    if (options.persistentStoreKind != LocalSearchIndexCore::PersistentStoreKind::Sqlite || maintenanceState.running)
    {
        return;
    }

    const LocalSearchIndexCore::PersistentStoreInfo storeInfo = LocalSearchIndexCore::GetPersistentStoreInfo(BuildRepositoryOptions(options));
    const bool shouldQueue                                    = ShouldQueueAutomaticMaintenance(storeInfo);
    if (shouldQueue)
    {
        if (! maintenanceState.queued)
        {
            maintenanceState.queued            = true;
            maintenanceState.queuedSinceTickMs = ::GetTickCount64();
            const ServerEventDetails details{
                .currentPath = storeInfo.primaryPath.c_str(),
            };
            NotifyServerEvent(options, SEARCH_SERVICE_SERVER_EVENT_MAINTENANCE_QUEUED, S_OK, handledRequests, &details);
        }
        return;
    }

    SetQueuedMaintenanceState(maintenanceState, false);
}

[[nodiscard]] DWORD GetMaintenanceWaitTimeoutMs(const ServerMaintenanceState& maintenanceState) noexcept
{
    if (! maintenanceState.queued || maintenanceState.queuedSinceTickMs == 0u)
    {
        return INFINITE;
    }

    const uint64_t now     = ::GetTickCount64();
    const uint64_t dueTick = maintenanceState.queuedSinceTickMs + kIdleMaintenanceGraceMs;
    if (now >= dueTick)
    {
        return 0u;
    }

    return static_cast<DWORD>(std::min<uint64_t>(dueTick - now, std::numeric_limits<DWORD>::max()));
}

[[nodiscard]] bool ShouldRunAutomaticMaintenanceNow(const ServerMaintenanceState& maintenanceState) noexcept
{
    if (! maintenanceState.queued || maintenanceState.running || maintenanceState.queuedSinceTickMs == 0u)
    {
        return false;
    }

    return ::GetTickCount64() >= (maintenanceState.queuedSinceTickMs + kIdleMaintenanceGraceMs);
}

HRESULT RunQueuedAutomaticMaintenance(const ServerOptions& options, ServerMaintenanceState& maintenanceState, uint32_t handledRequests) noexcept
{
    if (options.persistentStoreKind != LocalSearchIndexCore::PersistentStoreKind::Sqlite)
    {
        SetQueuedMaintenanceState(maintenanceState, false);
        maintenanceState.running = false;
        return S_FALSE;
    }

    const LocalSearchIndexCore::PersistentStoreInfo storeInfo = LocalSearchIndexCore::GetPersistentStoreInfo(BuildRepositoryOptions(options));
    if (! ShouldQueueAutomaticMaintenance(storeInfo))
    {
        SetQueuedMaintenanceState(maintenanceState, false);
        maintenanceState.running = false;
        return S_FALSE;
    }

    maintenanceState.queued  = false;
    maintenanceState.running = true;
    const ServerEventDetails startDetails{
        .currentPath = storeInfo.primaryPath.c_str(),
    };
    NotifyServerEvent(options, SEARCH_SERVICE_SERVER_EVENT_MAINTENANCE_RUNNING, S_OK, handledRequests, &startDetails);

    SqliteIndexStore::AutomaticMaintenanceResult maintenanceResult{};
    const HRESULT maintenanceHr =
        SqliteIndexStore::RunAutomaticMaintenance(storeInfo.primaryPath, LocalSearchIndexCore::GetDefaultSqliteMaintenancePolicy(), &maintenanceResult);

    maintenanceState.running           = false;
    maintenanceState.queuedSinceTickMs = 0u;
    const ServerEventDetails completeDetails{
        .currentPath = storeInfo.primaryPath.c_str(),
    };
    NotifyServerEvent(options, SEARCH_SERVICE_SERVER_EVENT_MAINTENANCE_COMPLETED, maintenanceHr, handledRequests, &completeDetails);

    if (FAILED(maintenanceHr))
    {
        Debug::Warning(L"SearchServiceBroker: automatic SQLite maintenance failed for '{}'. hr=0x{:08X}",
                       storeInfo.primaryPath,
                       static_cast<unsigned long>(maintenanceHr));
        return maintenanceHr;
    }

    if (maintenanceResult.ranVacuum)
    {
        Debug::Info(L"SearchServiceBroker: automatic SQLite maintenance ran full VACUUM rewrite for '{}'.", storeInfo.primaryPath);
    }

    RefreshAutomaticMaintenanceQueue(options, maintenanceState, handledRequests);
    return S_OK;
}

[[nodiscard]] std::wstring GetEnvironmentValue(std::wstring_view name) noexcept
{
    return EnvironmentVariables::Read(name).value_or(std::wstring{});
}

[[nodiscard]] DWORD GetClientMissingPipeRetryWindowMs() noexcept
{
    const std::wstring value = GetEnvironmentValue(kClientMissingPipeRetryMsEnvVar);
    if (value.empty())
    {
        return kDefaultMissingPipeRetryWindowMs;
    }

    uint64_t parsed = 0u;
    for (const wchar_t ch : value)
    {
        if (ch < L'0' || ch > L'9')
        {
            return kDefaultMissingPipeRetryWindowMs;
        }

        parsed = (parsed * 10u) + static_cast<uint64_t>(ch - L'0');
        if (parsed > kMaxMissingPipeRetryWindowMs)
        {
            return kMaxMissingPipeRetryWindowMs;
        }
    }

    return static_cast<DWORD>(parsed);
}

[[nodiscard]] HRESULT WaitForClientConnectRetry(DWORD delayMs, ClientIoContext* context) noexcept
{
    const ULONGLONG deadline = ::GetTickCount64() + delayMs;
    for (;;)
    {
        const HRESULT cancelHr = CheckClientIoCancelled(context);
        if (FAILED(cancelHr))
        {
            return cancelHr;
        }

        const ULONGLONG now = ::GetTickCount64();
        if (now >= deadline)
        {
            return S_OK;
        }

        ::Sleep(static_cast<DWORD>((std::min<ULONGLONG>)(deadline - now, 1u)));
    }
}

HRESULT ConnectClientPipeForName(std::wstring_view requestedPipeName,
                                 DWORD missingPipeRetryWindowMs,
                                 wil::unique_handle& outPipe,
                                 ClientIoContext* context = nullptr) noexcept
{
    const std::wstring pipeName         = BuildPipeName(requestedPipeName);
    const ULONGLONG deadline            = ::GetTickCount64() + 750u;
    const ULONGLONG missingPipeDeadline = ::GetTickCount64() + missingPipeRetryWindowMs;

    for (;;)
    {
        HRESULT hr = CheckClientIoCancelled(context);
        if (FAILED(hr))
        {
            return hr;
        }

        if (::WaitNamedPipeW(pipeName.c_str(), kClientConnectTimeoutMs) == 0)
        {
            const DWORD waitError = ::GetLastError();
            if (waitError == ERROR_FILE_NOT_FOUND)
            {
                if (::GetTickCount64() >= missingPipeDeadline)
                {
                    return HRESULT_FROM_WIN32(waitError);
                }

                hr = WaitForClientConnectRetry(10u, context);
                if (FAILED(hr))
                {
                    return hr;
                }
                continue;
            }
            if (waitError != ERROR_SEM_TIMEOUT && waitError != ERROR_FILE_NOT_FOUND && waitError != ERROR_PIPE_BUSY)
            {
                return HRESULT_FROM_WIN32(waitError);
            }
        }

        outPipe.reset(
            ::CreateFileW(pipeName.c_str(), GENERIC_READ | GENERIC_WRITE, 0u, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_OVERLAPPED, nullptr));
        if (outPipe)
        {
            return S_OK;
        }

        const DWORD error = ::GetLastError();
        if (error == ERROR_FILE_NOT_FOUND)
        {
            if (::GetTickCount64() >= missingPipeDeadline)
            {
                return HRESULT_FROM_WIN32(error);
            }

            hr = WaitForClientConnectRetry(10u, context);
            if (FAILED(hr))
            {
                return hr;
            }
            continue;
        }
        if (error != ERROR_FILE_NOT_FOUND && error != ERROR_PIPE_BUSY && error != ERROR_SEM_TIMEOUT)
        {
            return HRESULT_FROM_WIN32(error);
        }

        if (::GetTickCount64() >= deadline)
        {
            return HRESULT_FROM_WIN32(error);
        }

        hr = WaitForClientConnectRetry(25u, context);
        if (FAILED(hr))
        {
            return hr;
        }
    }
}

HRESULT SendProtocolError(HANDLE handle, uint32_t protocolVersion, HRESULT hr, std::wstring_view message, const ServerIoContext* context) noexcept
{
    ErrorPayload payload{};
    payload.result       = hr;
    payload.messageBytes = static_cast<uint32_t>(message.size() * sizeof(wchar_t));

    std::vector<std::byte> buffer;
    buffer.reserve(sizeof(payload) + payload.messageBytes);
    AppendBytes(buffer, &payload, sizeof(payload));
    AppendUtf16(buffer, message);
    return SendFrame(handle, MessageType::Error, protocolVersion, buffer, context);
}

HRESULT SendStatusResponse(HANDLE handle,
                           uint32_t protocolVersion,
                           const ServerOptions& options,
                           const ServerMaintenanceState* maintenanceState,
                           const ServerStartupWarmupState* startupWarmupState,
                           LocalSearchIndexCore::Repository* repository,
                           const ServerIoContext* context) noexcept
{
    const std::wstring pipeName                         = BuildPipeName(options.pipeName);
    const std::wstring storageRoot                      = options.storageRootDirectory;
    LocalSearchIndexCore::PersistentStoreInfo storeInfo = LocalSearchIndexCore::GetPersistentStoreInfo(BuildRepositoryOptions(options));
    const ServerStartupWarmupSnapshot warmupSnapshot    = CaptureStartupWarmupSnapshot(startupWarmupState);
    LocalSearchIndexCore::RepositoryStatus repositoryStatus{};
    std::vector<std::wstring> discoveredRoots;
    if (repository != nullptr)
    {
        repository->GetStatus(repositoryStatus);
        repository->CollectCachedRoots(discoveredRoots);
        const RepositoryStoreOverlayResult overlayResult = OverlayFresherRepositoryStoreInfo(storeInfo, repository);
        if (overlayResult != RepositoryStoreOverlayResult::Compatible)
        {
            // These fields are tied to the repository generation and must be re-derived from the coherent
            // on-disk inspection plus the independent startup-warmup snapshot. The last request execution
            // mode remains meaningful when this same configured store is merely uninspectable, but not after
            // a detected external generation rotation.
            repositoryStatus.storeState     = LocalSearchIndexCore::StoreState::Unknown;
            repositoryStatus.syncPhase      = LocalSearchIndexCore::SyncPhase::Idle;
            repositoryStatus.fallbackReason = LocalSearchIndexCore::FallbackReason::None;
            repositoryStatus.completedRoots = 0u;
            repositoryStatus.totalRoots     = 0u;
            repositoryStatus.activeRoot.clear();
            if (overlayResult == RepositoryStoreOverlayResult::GenerationMismatch)
            {
                repositoryStatus.queryExecutionMode = LocalSearchIndexCore::QueryExecutionMode::Unknown;
            }
        }
    }
    std::wstring discoveredRootsText;
    for (size_t index = 0; index < discoveredRoots.size(); ++index)
    {
        if (index != 0u)
        {
            discoveredRootsText.append(L"\n");
        }
        discoveredRootsText.append(discoveredRoots[index]);
    }

    StatusResponsePayload payload{};
    payload.processId        = ::GetCurrentProcessId();
    payload.pipeNameBytes    = static_cast<uint32_t>(pipeName.size() * sizeof(wchar_t));
    payload.storageRootBytes = static_cast<uint32_t>(storageRoot.size() * sizeof(wchar_t));

    StatusResponseExtendedPayload extended{};
    extended.persistentStoreKind                = static_cast<uint32_t>(storeInfo.kind);
    extended.persistentStorePathBytes           = static_cast<uint32_t>(storeInfo.primaryPath.size() * sizeof(wchar_t));
    extended.writeAheadLogPathBytes             = static_cast<uint32_t>(storeInfo.writeAheadLogPath.size() * sizeof(wchar_t));
    extended.persistentStoreBytes               = storeInfo.primaryBytes;
    extended.writeAheadLogBytes                 = storeInfo.writeAheadLogBytes;
    extended.autoCheckpointTargetBytes          = storeInfo.autoCheckpointTargetBytes;
    extended.indexedVolumeCount                 = storeInfo.indexedVolumeCount;
    extended.indexedEntryCount                  = storeInfo.indexedEntryCount;
    extended.legacyImportVolumeCount            = storeInfo.legacyImportVolumeCount;
    extended.autoCompactionFragmentationPercent = storeInfo.autoCompactionFragmentationPercent;
    extended.autoCompactionMinBytes             = storeInfo.autoCompactionMinBytes;
    if (storeInfo.autoCheckpointEnabled)
    {
        extended.flags |= STATUS_RESPONSE_FLAG_AUTO_CHECKPOINT;
    }
    if (storeInfo.autoCompactionEnabled)
    {
        extended.flags |= STATUS_RESPONSE_FLAG_AUTO_COMPACTION;
    }
    if (storeInfo.inspectionSucceeded)
    {
        extended.flags |= STATUS_RESPONSE_FLAG_STORE_INSPECTED;
    }
    if (storeInfo.readyForQueryCutover)
    {
        extended.flags |= STATUS_RESPONSE_FLAG_READY_FOR_QUERY_CUTOVER;
    }

    StatusResponseMaintenancePayload maintenance{};
    maintenance.lastCheckpointUtcBytes = static_cast<uint32_t>(storeInfo.lastCheckpointUtc.size() * sizeof(wchar_t));
    maintenance.lastCompactionUtcBytes = static_cast<uint32_t>(storeInfo.lastCompactionUtc.size() * sizeof(wchar_t));
    maintenance.pageCount              = storeInfo.pageCount;
    maintenance.freelistPageCount      = storeInfo.freelistPageCount;
    if (maintenanceState != nullptr)
    {
        if (maintenanceState->queued)
        {
            maintenance.flags |= STATUS_RESPONSE_MAINTENANCE_FLAG_QUEUED;
        }
        if (maintenanceState->running)
        {
            maintenance.flags |= STATUS_RESPONSE_MAINTENANCE_FLAG_RUNNING;
        }
    }

    StatusResponseDiscoveryPayload discovery{};
    discovery.discoveredRootCount  = static_cast<uint32_t>(discoveredRoots.size());
    discovery.discoveredRootsBytes = static_cast<uint32_t>(discoveredRootsText.size() * sizeof(wchar_t));

    StatusResponseStartupWarmupPayload warmup{};
    warmup.currentRootBytes = static_cast<uint32_t>(warmupSnapshot.currentRoot.size() * sizeof(wchar_t));
    warmup.totalRoots       = warmupSnapshot.totalRoots;
    warmup.completedRoots   = warmupSnapshot.completedRoots;
    warmup.failedRoots      = warmupSnapshot.failedRoots;
    if (warmupSnapshot.enabled)
    {
        warmup.flags |= STATUS_RESPONSE_STARTUP_WARMUP_FLAG_ENABLED;
    }
    if (warmupSnapshot.running)
    {
        warmup.flags |= STATUS_RESPONSE_STARTUP_WARMUP_FLAG_RUNNING;
    }
    if (warmupSnapshot.hasFailure)
    {
        warmup.flags |= STATUS_RESPONSE_STARTUP_WARMUP_FLAG_HAS_FAILURE;
    }

    StatusResponseStartupWarmupFailurePayload warmupFailure{};
    warmupFailure.lastFailureHr        = warmupSnapshot.lastFailureHr;
    warmupFailure.lastFailureRootBytes = static_cast<uint32_t>(warmupSnapshot.lastFailureRoot.size() * sizeof(wchar_t));
    if (warmupSnapshot.hasFailure)
    {
        warmupFailure.flags |= STATUS_RESPONSE_STARTUP_WARMUP_FLAG_HAS_FAILURE;
    }

    const LocalSearchIndexCore::FallbackReason derivedFallbackReason = repositoryStatus.fallbackReason != LocalSearchIndexCore::FallbackReason::None
                                                                           ? repositoryStatus.fallbackReason
                                                                           : ClassifyPersistentStoreFallbackReason(storeInfo);
    const LocalSearchIndexCore::StoreState runtimeStoreState         = DeriveRuntimeStoreState(repositoryStatus, storeInfo, maintenanceState);
    const LocalSearchIndexCore::SyncPhase runtimeSyncPhase           = DeriveRuntimeSyncPhase(repositoryStatus, storeInfo, warmupSnapshot);
    const uint64_t runtimeCompletedRoots =
        (repositoryStatus.completedRoots != 0u || repositoryStatus.totalRoots != 0u) ? repositoryStatus.completedRoots : warmupSnapshot.completedRoots;
    const uint64_t runtimeTotalRoots =
        (repositoryStatus.completedRoots != 0u || repositoryStatus.totalRoots != 0u) ? repositoryStatus.totalRoots : warmupSnapshot.totalRoots;
    const std::wstring runtimeActiveRoot = ! repositoryStatus.activeRoot.empty() ? repositoryStatus.activeRoot : warmupSnapshot.currentRoot;

    StatusResponseRuntimePayload runtime{};
    runtime.storeState         = static_cast<uint32_t>(runtimeStoreState);
    runtime.syncPhase          = static_cast<uint32_t>(runtimeSyncPhase);
    runtime.queryExecutionMode = static_cast<uint32_t>(repositoryStatus.queryExecutionMode);
    runtime.fallbackReason     = static_cast<uint32_t>(derivedFallbackReason);
    runtime.completedRoots     = runtimeCompletedRoots;
    runtime.totalRoots         = runtimeTotalRoots;
    runtime.activeRootBytes    = static_cast<uint32_t>(runtimeActiveRoot.size() * sizeof(wchar_t));

    std::vector<std::byte> buffer;
    buffer.reserve(sizeof(payload) + payload.pipeNameBytes + payload.storageRootBytes + sizeof(extended) + extended.persistentStorePathBytes +
                   extended.writeAheadLogPathBytes + sizeof(maintenance) + maintenance.lastCheckpointUtcBytes + maintenance.lastCompactionUtcBytes +
                   sizeof(discovery) + discovery.discoveredRootsBytes + sizeof(warmup) + warmup.currentRootBytes + sizeof(warmupFailure) +
                   warmupFailure.lastFailureRootBytes + sizeof(runtime) + runtime.activeRootBytes);
    AppendBytes(buffer, &payload, sizeof(payload));
    AppendUtf16(buffer, pipeName);
    AppendUtf16(buffer, storageRoot);
    AppendBytes(buffer, &extended, sizeof(extended));
    AppendUtf16(buffer, storeInfo.primaryPath);
    AppendUtf16(buffer, storeInfo.writeAheadLogPath);
    AppendBytes(buffer, &maintenance, sizeof(maintenance));
    AppendUtf16(buffer, storeInfo.lastCheckpointUtc);
    AppendUtf16(buffer, storeInfo.lastCompactionUtc);
    AppendBytes(buffer, &discovery, sizeof(discovery));
    AppendUtf16(buffer, discoveredRootsText);
    AppendBytes(buffer, &warmup, sizeof(warmup));
    AppendUtf16(buffer, warmupSnapshot.currentRoot);
    AppendBytes(buffer, &warmupFailure, sizeof(warmupFailure));
    AppendUtf16(buffer, warmupSnapshot.lastFailureRoot);
    AppendBytes(buffer, &runtime, sizeof(runtime));
    AppendUtf16(buffer, runtimeActiveRoot);
    return SendFrame(handle, MessageType::StatusResponse, protocolVersion, buffer, context);
}

HRESULT SendProgress(HANDLE handle, uint32_t protocolVersion, const QueryProgress& progress, const ServerIoContext* context) noexcept
{
    ProgressPayload payload{};
    payload.phase              = static_cast<uint32_t>(progress.phase);
    payload.warningFlags       = progress.warningFlags;
    payload.statusHint         = progress.statusHint;
    payload.currentPathBytes   = static_cast<uint32_t>(progress.currentPath.size() * sizeof(wchar_t));
    payload.scannedDirectories = progress.scannedDirectories;
    payload.scannedFiles       = progress.scannedFiles;
    payload.candidateFiles     = progress.candidateFiles;
    payload.matchedEntries     = progress.matchedEntries;

    ProgressRuntimePayload runtime{};
    runtime.storeState         = static_cast<uint32_t>(progress.storeState);
    runtime.syncPhase          = static_cast<uint32_t>(progress.syncPhase);
    runtime.queryExecutionMode = static_cast<uint32_t>(progress.queryExecutionMode);
    runtime.fallbackReason     = static_cast<uint32_t>(progress.fallbackReason);
    runtime.completedRoots     = progress.completedRoots;
    runtime.totalRoots         = progress.totalRoots;
    runtime.activeRootBytes    = static_cast<uint32_t>(progress.activeRoot.size() * sizeof(wchar_t));

    std::vector<std::byte> buffer;
    buffer.reserve(sizeof(payload) + payload.currentPathBytes + sizeof(runtime) + runtime.activeRootBytes);
    AppendBytes(buffer, &payload, sizeof(payload));
    AppendUtf16(buffer, progress.currentPath);
    AppendBytes(buffer, &runtime, sizeof(runtime));
    AppendUtf16(buffer, progress.activeRoot);
    return SendFrame(handle, MessageType::QueryProgress, protocolVersion, buffer, context);
}

HRESULT SendCandidates(HANDLE handle,
                       uint32_t protocolVersion,
                       std::span<const LocalSearchIndexCore::Candidate> candidates,
                       uint32_t disconnectAfterBatches,
                       uint32_t& batchesSent,
                       const ServerIoContext* context) noexcept
{
    constexpr size_t kBatchItems = 256u;

    for (size_t offset = 0; offset < candidates.size(); offset += kBatchItems)
    {
        const size_t count = std::min(kBatchItems, candidates.size() - offset);

        CandidateBatchHeader batchHeader{};
        batchHeader.count = static_cast<uint32_t>(count);

        std::vector<std::byte> buffer;
        buffer.reserve(sizeof(batchHeader) + (count * sizeof(CandidateEntryHeader)));
        AppendBytes(buffer, &batchHeader, sizeof(batchHeader));

        for (size_t index = 0; index < count; ++index)
        {
            const auto& candidate = candidates[offset + index];
            CandidateEntryHeader entry{};
            entry.fileAttributes      = candidate.fileAttributes;
            entry.metadataFlags       = candidate.metadataFlags;
            entry.fullPathBytes       = static_cast<uint32_t>(candidate.fullPath.size() * sizeof(wchar_t));
            entry.displayNameBytes    = static_cast<uint32_t>(candidate.displayName.size() * sizeof(wchar_t));
            entry.creationTime100ns   = candidate.creationTime100ns;
            entry.lastAccessTime100ns = candidate.lastAccessTime100ns;
            entry.endOfFile           = candidate.endOfFile;
            entry.lastWriteTime100ns  = candidate.lastWriteTime100ns;
            entry.changeTime100ns     = candidate.changeTime100ns;
            entry.allocationSize      = candidate.allocationSize;
            AppendBytes(buffer, &entry, sizeof(entry));
            AppendUtf16(buffer, candidate.fullPath);
            AppendUtf16(buffer, candidate.displayName);
        }

        HRESULT hr = SendFrame(handle, MessageType::QueryBatch, protocolVersion, buffer, context);
        if (FAILED(hr))
        {
            return hr;
        }

        ++batchesSent;
        if (disconnectAfterBatches != 0u && batchesSent >= disconnectAfterBatches)
        {
            return HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE);
        }
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE StreamServerQueryProgress(const LocalSearchIndexCore::ProgressUpdate* progress, void* cookie) noexcept
{
    if (progress == nullptr || cookie == nullptr)
    {
        return E_POINTER;
    }

    const auto& context = *static_cast<const ServerQueryContext*>(cookie);

    QueryProgress payload{};
    payload.phase              = progress->phase;
    payload.warningFlags       = progress->warningFlags;
    payload.statusHint         = progress->statusHint;
    payload.scannedDirectories = progress->scannedDirectories;
    payload.scannedFiles       = progress->scannedFiles;
    payload.candidateFiles     = progress->candidateFiles;
    payload.matchedEntries     = progress->matchedEntries;
    payload.storeState         = progress->storeState;
    payload.syncPhase          = progress->syncPhase;
    payload.queryExecutionMode = progress->queryExecutionMode;
    payload.fallbackReason     = progress->fallbackReason;
    payload.completedRoots     = progress->completedRoots;
    payload.totalRoots         = progress->totalRoots;
    payload.activeRoot         = progress->activeRoot;
    payload.currentPath        = progress->currentPath;

    HRESULT hr = SendProgress(context.pipe, context.protocolVersion, payload, context.ioContext);
    if (FAILED(hr))
    {
        return hr;
    }

    if (context.options != nullptr && context.rootPath != nullptr && context.namePattern != nullptr)
    {
        const wchar_t* currentPath            = progress->currentPath.empty() ? context.rootPath->c_str() : progress->currentPath.c_str();
        const uint64_t snapshotFileBytes      = (context.stats != nullptr) ? context.stats->snapshotFileBytes : 0u;
        const uint64_t estimatedMemoryBytes   = (context.stats != nullptr) ? context.stats->estimatedMemoryBytes : 0u;
        const uint64_t ensureReadyDurationMs  = (context.stats != nullptr) ? context.stats->ensureReadyDurationMs : 0u;
        const uint64_t executeQueryDurationMs = (context.stats != nullptr) ? context.stats->executeQueryDurationMs : 0u;
        const ServerEventDetails details{
            .requestType            = SEARCH_SERVICE_SERVER_REQUEST_QUERY,
            .phase                  = progress->phase,
            .requestFlags           = context.requestFlags,
            .requestNameMode        = context.requestNameMode,
            .batchesSent            = 0u,
            .scannedDirectories     = progress->scannedDirectories,
            .scannedFiles           = progress->scannedFiles,
            .candidateFiles         = progress->candidateFiles,
            .matchedEntries         = progress->matchedEntries,
            .snapshotFileBytes      = snapshotFileBytes,
            .estimatedMemoryBytes   = estimatedMemoryBytes,
            .ensureReadyDurationMs  = ensureReadyDurationMs,
            .executeQueryDurationMs = executeQueryDurationMs,
            .warningFlags           = progress->warningFlags,
            .storeState             = static_cast<uint32_t>(progress->storeState),
            .syncPhase              = static_cast<uint32_t>(progress->syncPhase),
            .queryExecutionMode     = static_cast<uint32_t>(progress->queryExecutionMode),
            .fallbackReason         = static_cast<uint32_t>(progress->fallbackReason),
            .completedRoots         = progress->completedRoots,
            .totalRoots             = progress->totalRoots,
            .rootPath               = context.rootPath->c_str(),
            .namePattern            = context.namePattern->c_str(),
            .currentPath            = currentPath,
            .activeRoot             = progress->activeRoot.empty() ? nullptr : progress->activeRoot.c_str(),
        };
        NotifyServerEvent(*context.options, SEARCH_SERVICE_SERVER_EVENT_QUERY_PROGRESS, progress->statusHint, context.handledRequests, &details);
    }

    return S_OK;
}

HRESULT FlushServerCandidates(ServerCandidateBatchState& state) noexcept
{
    if (state.bufferedCandidates.empty())
    {
        return S_OK;
    }

    if (! state.sentIndexedProgress && state.stats != nullptr)
    {
        QueryProgress progress{};
        progress.phase              = FILESYSTEM_SEARCH_PHASE_ENUMERATING;
        progress.currentPath        = state.rootPath;
        progress.scannedDirectories = state.stats->directoryCount;
        progress.scannedFiles       = state.stats->fileCount;
        progress.candidateFiles     = state.stats->candidateCount;

        const HRESULT progressHr = SendProgress(state.pipe, state.protocolVersion, progress, state.ioContext);
        if (FAILED(progressHr))
        {
            return progressHr;
        }

        state.sentIndexedProgress = true;
    }

    const HRESULT hr = SendCandidates(state.pipe,
                                      state.protocolVersion,
                                      std::span<const LocalSearchIndexCore::Candidate>(state.bufferedCandidates.data(), state.bufferedCandidates.size()),
                                      state.disconnectAfterBatches,
                                      state.batchesSent,
                                      state.ioContext);
    if (FAILED(hr))
    {
        return hr;
    }

    if (state.options != nullptr && state.stats != nullptr)
    {
        const ServerEventDetails details{
            .requestType            = SEARCH_SERVICE_SERVER_REQUEST_QUERY,
            .phase                  = FILESYSTEM_SEARCH_PHASE_ENUMERATING,
            .requestFlags           = state.requestFlags,
            .requestNameMode        = state.requestNameMode,
            .batchesSent            = state.batchesSent,
            .scannedDirectories     = state.stats->directoryCount,
            .scannedFiles           = state.stats->fileCount,
            .candidateFiles         = state.stats->candidateCount,
            .matchedEntries         = 0u,
            .snapshotFileBytes      = state.stats->snapshotFileBytes,
            .estimatedMemoryBytes   = state.stats->estimatedMemoryBytes,
            .ensureReadyDurationMs  = state.stats->ensureReadyDurationMs,
            .executeQueryDurationMs = state.stats->executeQueryDurationMs,
            .warningFlags           = QueryStateWarningFlags(state),
            .storeState             = static_cast<uint32_t>(state.stats->usedLiveScanFallback ? ResolveFallbackStoreState(state.stats->fallbackReason)
                                                                                              : LocalSearchIndexCore::StoreState::Ready),
            .syncPhase =
                static_cast<uint32_t>(state.stats->usedLiveScanFallback ? LocalSearchIndexCore::SyncPhase::Idle : LocalSearchIndexCore::SyncPhase::Watching),
            .queryExecutionMode = static_cast<uint32_t>(state.stats->queryExecutionMode),
            .fallbackReason     = static_cast<uint32_t>(state.stats->fallbackReason),
            .rootPath           = state.rootPath.c_str(),
            .namePattern        = state.namePattern.c_str(),
            .currentPath        = state.rootPath.c_str(),
            .activeRoot         = state.rootPath.c_str(),
        };
        NotifyServerEvent(*state.options, SEARCH_SERVICE_SERVER_EVENT_QUERY_BATCH_SENT, S_OK, state.handledRequests, &details);
    }

    state.bufferedCandidates.clear();
    state.payloadBytes = sizeof(CandidateBatchHeader);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE StreamServerCandidate(LocalSearchIndexCore::Candidate* candidate, void* cookie) noexcept
{
    if (candidate == nullptr || cookie == nullptr)
    {
        return E_POINTER;
    }

    auto& state = *static_cast<ServerCandidateBatchState*>(cookie);

    HRESULT authorizationHr = S_OK;
    try
    {
        authorizationHr = AuthorizeCandidateForClient(state, *candidate);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory: service callback boundary. Return failure instead of unwinding through a noexcept callback.
        return E_FAIL;
    }
    if (authorizationHr == S_FALSE)
    {
        return LocalSearchIndexCore::kSkipCandidateHr;
    }
    if (FAILED(authorizationHr))
    {
        state.warningFlags |= FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED;
        return LocalSearchIndexCore::kSkipCandidateHr;
    }

    const size_t candidatePayloadBytes =
        sizeof(CandidateEntryHeader) + (candidate->fullPath.size() * sizeof(wchar_t)) + (candidate->displayName.size() * sizeof(wchar_t));
    if (! state.bufferedCandidates.empty() && (state.payloadBytes + candidatePayloadBytes) > kMaxFrameBytes)
    {
        const HRESULT flushHr = FlushServerCandidates(state);
        if (FAILED(flushHr))
        {
            return flushHr;
        }
    }

    try
    {
        state.bufferedCandidates.push_back(std::move(*candidate));
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory: service callback boundary. Return failure instead of unwinding through a noexcept callback.
        return E_FAIL;
    }

    state.payloadBytes += candidatePayloadBytes;
    if (state.bufferedCandidates.size() >= 256u || state.payloadBytes >= (256u * 1024u))
    {
        return FlushServerCandidates(state);
    }

    return S_OK;
}

uint32_t PackQueryStatsFlags(const LocalSearchIndexCore::QueryStats& stats) noexcept
{
    uint32_t flags = 0u;
    if (stats.snapshotLoaded)
    {
        flags |= QUERY_COMPLETE_FLAG_SNAPSHOT_LOADED;
    }
    if (stats.snapshotSaved)
    {
        flags |= QUERY_COMPLETE_FLAG_SNAPSHOT_SAVED;
    }
    if (stats.journalAvailable)
    {
        flags |= QUERY_COMPLETE_FLAG_JOURNAL_AVAILABLE;
    }
    if (stats.journalReplayApplied)
    {
        flags |= QUERY_COMPLETE_FLAG_JOURNAL_REPLAY_APPLIED;
    }
    if (stats.rebuiltJournalIdMismatch)
    {
        flags |= QUERY_COMPLETE_FLAG_REBUILT_JOURNAL_ID_MISMATCH;
    }
    if (stats.rebuiltJournalRangeInvalid)
    {
        flags |= QUERY_COMPLETE_FLAG_REBUILT_JOURNAL_RANGE_INVALID;
    }
    if (stats.rebuiltSnapshotCorruption)
    {
        flags |= QUERY_COMPLETE_FLAG_REBUILT_SNAPSHOT_CORRUPTION;
    }
    if (stats.usedNtfsEnumeration)
    {
        flags |= QUERY_COMPLETE_FLAG_USED_NTFS_ENUMERATION;
    }
    if (stats.usedTraversalSeed)
    {
        flags |= QUERY_COMPLETE_FLAG_USED_TRAVERSAL_SEED;
    }
    if (stats.usedSqliteStore)
    {
        flags |= QUERY_COMPLETE_FLAG_USED_SQLITE_STORE;
    }
    if (stats.sqliteCutoverBlocked)
    {
        flags |= QUERY_COMPLETE_FLAG_SQLITE_CUTOVER_BLOCKED;
    }
    if (stats.usedNamePrefilter)
    {
        flags |= QUERY_COMPLETE_FLAG_USED_NAME_PREFILTER;
    }
    if (stats.sqliteReadOnlyQuery)
    {
        flags |= QUERY_COMPLETE_FLAG_SQLITE_QUERY_READ_ONLY;
    }
    if (stats.hardlinkAliasCoverageIncomplete)
    {
        flags |= QUERY_COMPLETE_FLAG_HARDLINK_ALIAS_INCOMPLETE;
    }
    return flags;
}

void UnpackQueryStatsFlags(uint32_t flags, LocalSearchIndexCore::QueryStats& stats) noexcept
{
    stats.snapshotLoaded                  = (flags & QUERY_COMPLETE_FLAG_SNAPSHOT_LOADED) != 0u;
    stats.snapshotSaved                   = (flags & QUERY_COMPLETE_FLAG_SNAPSHOT_SAVED) != 0u;
    stats.journalAvailable                = (flags & QUERY_COMPLETE_FLAG_JOURNAL_AVAILABLE) != 0u;
    stats.journalReplayApplied            = (flags & QUERY_COMPLETE_FLAG_JOURNAL_REPLAY_APPLIED) != 0u;
    stats.rebuiltJournalIdMismatch        = (flags & QUERY_COMPLETE_FLAG_REBUILT_JOURNAL_ID_MISMATCH) != 0u;
    stats.rebuiltJournalRangeInvalid      = (flags & QUERY_COMPLETE_FLAG_REBUILT_JOURNAL_RANGE_INVALID) != 0u;
    stats.rebuiltSnapshotCorruption       = (flags & QUERY_COMPLETE_FLAG_REBUILT_SNAPSHOT_CORRUPTION) != 0u;
    stats.usedNtfsEnumeration             = (flags & QUERY_COMPLETE_FLAG_USED_NTFS_ENUMERATION) != 0u;
    stats.usedTraversalSeed               = (flags & QUERY_COMPLETE_FLAG_USED_TRAVERSAL_SEED) != 0u;
    stats.usedSqliteStore                 = (flags & QUERY_COMPLETE_FLAG_USED_SQLITE_STORE) != 0u;
    stats.sqliteCutoverBlocked            = (flags & QUERY_COMPLETE_FLAG_SQLITE_CUTOVER_BLOCKED) != 0u;
    stats.usedNamePrefilter               = (flags & QUERY_COMPLETE_FLAG_USED_NAME_PREFILTER) != 0u;
    stats.sqliteReadOnlyQuery             = (flags & QUERY_COMPLETE_FLAG_SQLITE_QUERY_READ_ONLY) != 0u;
    stats.hardlinkAliasCoverageIncomplete = (flags & QUERY_COMPLETE_FLAG_HARDLINK_ALIAS_INCOMPLETE) != 0u;
}

HRESULT SendQueryComplete(
    HANDLE handle, uint32_t protocolVersion, HRESULT result, const LocalSearchIndexCore::QueryStats& stats, const ServerIoContext* context) noexcept
{
    QueryCompletePayload payload{};
    payload.result                 = result;
    payload.fileSystemKind         = static_cast<uint32_t>(stats.fileSystemKind);
    payload.flags                  = PackQueryStatsFlags(stats);
    payload.entryCount             = stats.entryCount;
    payload.fileCount              = stats.fileCount;
    payload.directoryCount         = stats.directoryCount;
    payload.candidateCount         = stats.candidateCount;
    payload.nextUsn                = stats.nextUsn;
    payload.journalId              = stats.journalId;
    payload.snapshotFileBytes      = stats.snapshotFileBytes;
    payload.estimatedMemoryBytes   = stats.estimatedMemoryBytes;
    payload.ensureReadyDurationMs  = stats.ensureReadyDurationMs;
    payload.executeQueryDurationMs = stats.executeQueryDurationMs;

    QueryCompleteRuntimePayload runtime{};
    runtime.queryExecutionMode = static_cast<uint32_t>(stats.queryExecutionMode);
    runtime.fallbackReason     = static_cast<uint32_t>(stats.fallbackReason);
    runtime.warningFlags       = QueryStatsWarningFlags(stats);

    std::vector<std::byte> buffer;
    buffer.reserve(sizeof(payload) + sizeof(runtime));
    AppendBytes(buffer, &payload, sizeof(payload));
    AppendBytes(buffer, &runtime, sizeof(runtime));
    return SendFrame(handle, MessageType::QueryComplete, protocolVersion, buffer, context);
}

HRESULT SendAck(HANDLE handle, uint32_t protocolVersion, HRESULT hr, const ServerIoContext* context) noexcept
{
    AckPayload payload{};
    payload.result = hr;

    std::vector<std::byte> buffer;
    buffer.reserve(sizeof(payload));
    AppendBytes(buffer, &payload, sizeof(payload));
    return SendFrame(handle, MessageType::Ack, protocolVersion, buffer, context);
}

HRESULT STDMETHODCALLTYPE ServerCancelCheck(void* cookie) noexcept
{
    const auto* context = static_cast<const ServerQueryContext*>(cookie);
    if (context == nullptr || context->pipe == nullptr || context->pipe == INVALID_HANDLE_VALUE)
    {
        return E_POINTER;
    }

    DWORD availableBytes = 0u;
    if (::PeekNamedPipe(context->pipe, nullptr, 0u, nullptr, &availableBytes, nullptr) == 0)
    {
        const DWORD error = ::GetLastError();
        if (error == ERROR_BROKEN_PIPE || error == ERROR_PIPE_NOT_CONNECTED || error == ERROR_NO_DATA)
        {
            return HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE);
        }

        return HRESULT_FROM_WIN32(error);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE StartupWarmupCancelCheck(void* cookie) noexcept
{
    const auto* context = static_cast<const StartupWarmupCancelContext*>(cookie);
    if (context == nullptr)
    {
        return E_POINTER;
    }

    if (context->stopEvent != nullptr && ::WaitForSingleObject(context->stopEvent, 0u) == WAIT_OBJECT_0)
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    if (context->stopToken.stop_requested())
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    return S_OK;
}

[[nodiscard]] bool IsAsciiDriveLetter(wchar_t ch) noexcept
{
    return (ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z');
}

[[nodiscard]] bool IsDriveRootedPath(std::wstring_view path) noexcept
{
    return path.size() >= 3u && IsAsciiDriveLetter(path[0]) && path[1] == L':' && (path[2] == L'\\' || path[2] == L'/');
}

[[nodiscard]] bool IsDriveRoot(std::wstring_view path) noexcept
{
    return path.size() == 3u && IsDriveRootedPath(path);
}

[[nodiscard]] bool IsPathSeparator(wchar_t ch) noexcept
{
    return ch == L'\\' || ch == L'/';
}

HRESULT ConnectClientPipe(wil::unique_handle& outPipe, ClientIoContext* context = nullptr) noexcept
{
    return ConnectClientPipeForName(GetConfiguredPipeName(), GetClientMissingPipeRetryWindowMs(), outPipe, context);
}

[[nodiscard]] bool IsMissingPathError(DWORD error) noexcept
{
    return error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND;
}

[[nodiscard]] std::wstring ParentPathForAuthorization(std::wstring_view path)
{
    if (path.empty() || IsDriveRoot(path))
    {
        return {};
    }

    const size_t separator = path.find_last_of(L"\\/");
    if (separator == std::wstring_view::npos)
    {
        return {};
    }

    const size_t parentLength = separator == 2u && IsDriveRootedPath(path) ? 3u : separator;
    return std::wstring(path.substr(0u, parentLength));
}

[[nodiscard]] std::wstring NormalizeResultPathForClientAuthorization(std::wstring_view path)
{
    std::wstring normalized(path);
    std::replace(normalized.begin(), normalized.end(), L'/', L'\\');
    while (normalized.size() > 3u && IsPathSeparator(normalized.back()) && ! IsDriveRoot(normalized))
    {
        normalized.pop_back();
    }

    return normalized;
}

[[nodiscard]] bool IsPathAtOrUnderRoot(std::wstring_view path, std::wstring_view root) noexcept
{
    if (OrdinalString::EqualsNoCase(path, root))
    {
        return true;
    }

    if (path.size() <= root.size() || ! OrdinalString::StartsWithNoCase(path, root))
    {
        return false;
    }

    return IsPathSeparator(root.back()) || IsPathSeparator(path[root.size()]);
}

[[nodiscard]] bool IsDurableClientDirectoryAuthorizationFailure(const DWORD lastError) noexcept
{
    switch (lastError)
    {
        case ERROR_ACCESS_DENIED:
        case ERROR_FILE_NOT_FOUND:
        case ERROR_PATH_NOT_FOUND: return true;
        default: return false;
    }
}

#if defined(RS_SEARCH_TEST_HOOKS)
[[nodiscard]] HRESULT RunClientAuthorizationTestHook(ServerCandidateBatchState& state) noexcept
{
    if (state.testHookConsumed)
    {
        return S_OK;
    }

    switch (state.testHook)
    {
        case ServerTestHook::None: return S_OK;
        case ServerTestHook::FailClientAuthImpersonationOnce: state.testHookConsumed = true; return HRESULT_FROM_WIN32(ERROR_CANNOT_IMPERSONATE);
        case ServerTestHook::RejectMarkedRoot:
        case ServerTestHook::FailMarkedClientDirectoryOpenBadNetPath: return S_OK;
    }

    return E_INVALIDARG;
}
#endif

[[nodiscard]] HRESULT CheckClientCanListDirectory(ServerCandidateBatchState& state, std::wstring_view directoryPath, bool& outAllowed)
{
    outAllowed = false;

    const std::wstring cacheKey = OrdinalString::FoldCaseInvariant(directoryPath);
    if (const auto cached = state.clientDirectoryAccessCache.find(cacheKey); cached != state.clientDirectoryAccessCache.end())
    {
        outAllowed = cached->second;
        return S_OK;
    }

    const std::wstring path(directoryPath);
#if defined(RS_SEARCH_TEST_HOOKS)
    const HRESULT testHookHr = RunClientAuthorizationTestHook(state);
    if (FAILED(testHookHr))
    {
        return testHookHr;
    }

    if (state.testHook == ServerTestHook::FailMarkedClientDirectoryOpenBadNetPath && path.find(L"transient-auth") != std::wstring::npos)
    {
        state.clientDirectoryAccessCache.emplace(cacheKey, false);
        Debug::Perf::Emit(
            L"search.service.authorization.transient_parent_failures", L"shape=ERROR_BAD_NETPATH", 0u, 1u, 0u, HRESULT_FROM_WIN32(ERROR_BAD_NETPATH));
        ::SetLastError(ERROR_BAD_NETPATH);
        const DWORD loggedError = Debug::ErrorWithLastError(L"SearchServiceBroker: transient client directory authorization open failed for '{}'", path);
        return HRESULT_FROM_WIN32(loggedError != ERROR_SUCCESS ? loggedError : ERROR_BAD_NETPATH);
    }
#endif

    if (::ImpersonateNamedPipeClient(state.pipe) == 0)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    const auto revertToSelf = wil::scope_exit([]() noexcept { static_cast<void>(::RevertToSelf()); });

    wil::unique_handle directoryHandle(::CreateFileW(path.c_str(),
                                                     FILE_LIST_DIRECTORY | FILE_READ_ATTRIBUTES,
                                                     FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                     nullptr,
                                                     OPEN_EXISTING,
                                                     FILE_FLAG_BACKUP_SEMANTICS,
                                                     nullptr));
    if (! directoryHandle)
    {
        const DWORD lastError = ::GetLastError();
        if (IsDurableClientDirectoryAuthorizationFailure(lastError))
        {
            state.clientDirectoryAccessCache.emplace(cacheKey, false);
            return S_OK;
        }

        state.clientDirectoryAccessCache.emplace(cacheKey, false);
        Debug::Perf::Emit(L"search.service.authorization.transient_parent_failures", L"", 0u, 1u, 0u, HRESULT_FROM_WIN32(lastError));
        ::SetLastError(lastError);
        const DWORD loggedError = Debug::ErrorWithLastError(L"SearchServiceBroker: transient client directory authorization open failed for '{}'", path);
        return HRESULT_FROM_WIN32(loggedError != ERROR_SUCCESS ? loggedError : ERROR_ACCESS_DENIED);
    }

    outAllowed = true;
    state.clientDirectoryAccessCache.emplace(cacheKey, outAllowed);
    return S_OK;
}

[[nodiscard]] HRESULT AuthorizeCandidateForClient(ServerCandidateBatchState& state, const LocalSearchIndexCore::Candidate& candidate)
{
    const std::wstring_view rootPath = state.rootPath;
    const std::wstring candidatePath = NormalizeResultPathForClientAuthorization(candidate.fullPath);
    if (rootPath.empty() || candidatePath.empty() || ! IsPathAtOrUnderRoot(candidatePath, rootPath))
    {
        return S_FALSE;
    }

    if (OrdinalString::EqualsNoCase(candidatePath, rootPath))
    {
        return S_OK;
    }

    const size_t parentSeparator = candidatePath.find_last_of(L'\\');
    if (parentSeparator == std::wstring::npos)
    {
        return S_FALSE;
    }

    const std::wstring parentPath = candidatePath.substr(0u, parentSeparator == 2u && IsDriveRootedPath(candidatePath) ? 3u : parentSeparator);
    if (! IsPathAtOrUnderRoot(parentPath, rootPath))
    {
        return S_FALSE;
    }
    if (OrdinalString::EqualsNoCase(parentPath, rootPath))
    {
        return S_OK;
    }

    size_t segmentStart = rootPath.size();
    if (segmentStart < parentPath.size() && IsPathSeparator(parentPath[segmentStart]))
    {
        ++segmentStart;
    }

    for (;;)
    {
        const size_t separator = parentPath.find(L'\\', segmentStart);
        const std::wstring_view directory(parentPath.data(), separator == std::wstring::npos ? parentPath.size() : separator);
        if (directory.size() > rootPath.size())
        {
            bool allowed           = false;
            const HRESULT accessHr = CheckClientCanListDirectory(state, directory, allowed);
            if (FAILED(accessHr))
            {
                return accessHr;
            }
            if (! allowed)
            {
                return S_FALSE;
            }
        }

        if (separator == std::wstring::npos)
        {
            break;
        }
        segmentStart = separator + 1u;
    }

    return S_OK;
}

[[nodiscard]] HRESULT OpenExistingClientPathForAuthorization(const std::wstring& path, DWORD attributes) noexcept
{
    const bool isDirectory    = (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u;
    const DWORD desiredAccess = FILE_READ_ATTRIBUTES | (isDirectory ? static_cast<DWORD>(FILE_LIST_DIRECTORY) : static_cast<DWORD>(FILE_READ_DATA));
    const DWORD flags         = isDirectory ? FILE_FLAG_BACKUP_SEMANTICS : FILE_ATTRIBUTE_NORMAL;
    wil::unique_handle pathHandle(
        ::CreateFileW(path.c_str(), desiredAccess, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, flags, nullptr));
    if (! pathHandle)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    return S_OK;
}

[[nodiscard]] HRESULT AuthorizeExistingClientPathAccess(const std::wstring& path) noexcept
{
    const DWORD attributes = ::GetFileAttributesW(path.c_str());
    if (attributes == INVALID_FILE_ATTRIBUTES)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    return OpenExistingClientPathForAuthorization(path, attributes);
}

[[nodiscard]] HRESULT NormalizeClientRequestedRoot(std::wstring_view rootPath, std::wstring& outRoot) noexcept
{
    outRoot.clear();
    if (rootPath.empty())
    {
        return E_INVALIDARG;
    }

    std::wstring path(rootPath);
    if (path.find(L'\0') != std::wstring::npos)
    {
        return E_INVALIDARG;
    }

    std::replace(path.begin(), path.end(), L'/', L'\\');
    if (OrdinalString::StartsWithNoCase(path, L"\\\\?\\UNC\\") || OrdinalString::StartsWithNoCase(path, L"\\\\?\\GLOBALROOT\\") ||
        OrdinalString::StartsWithNoCase(path, L"\\\\.\\") || OrdinalString::StartsWithNoCase(path, L"\\??\\"))
    {
        return HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME);
    }

    if (OrdinalString::StartsWithNoCase(path, L"\\\\?\\"))
    {
        const std::wstring_view withoutPrefix(path.data() + 4u, path.size() - 4u);
        if (! IsDriveRootedPath(withoutPrefix))
        {
            return HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME);
        }
        path.assign(withoutPrefix);
    }

    if (path.rfind(L"\\\\", 0u) == 0 || ! IsDriveRootedPath(path))
    {
        return HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME);
    }

    const DWORD required = ::GetFullPathNameW(path.c_str(), 0u, nullptr, nullptr);
    if (required == 0u)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    std::wstring absolute(static_cast<size_t>(required) + 1u, L'\0');
    const DWORD written = ::GetFullPathNameW(path.c_str(), static_cast<DWORD>(absolute.size()), absolute.data(), nullptr);
    if (written == 0u)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }
    if (written >= absolute.size())
    {
        return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
    }

    absolute.resize(written);
    std::replace(absolute.begin(), absolute.end(), L'/', L'\\');
    if (! IsDriveRootedPath(absolute))
    {
        return HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME);
    }

    while (absolute.size() > 3u && (absolute.back() == L'\\' || absolute.back() == L'/') && ! IsDriveRoot(absolute))
    {
        absolute.pop_back();
    }

    outRoot = std::move(absolute);
    return S_OK;
}

[[nodiscard]] HRESULT AuthorizeClientRootAccess(HANDLE pipe, const std::wstring& normalizedRoot) noexcept
{
    if (pipe == nullptr || pipe == INVALID_HANDLE_VALUE || normalizedRoot.empty())
    {
        return E_INVALIDARG;
    }

    if (::ImpersonateNamedPipeClient(pipe) == 0)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    const auto revertToSelf = wil::scope_exit([]() noexcept { static_cast<void>(::RevertToSelf()); });

    return AuthorizeExistingClientPathAccess(normalizedRoot);
}

[[nodiscard]] HRESULT AuthorizeClientRootRebuildAccess(HANDLE pipe, const std::wstring& normalizedRoot) noexcept
{
    if (pipe == nullptr || pipe == INVALID_HANDLE_VALUE || normalizedRoot.empty())
    {
        return E_INVALIDARG;
    }

    if (::ImpersonateNamedPipeClient(pipe) == 0)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    const auto revertToSelf = wil::scope_exit([]() noexcept { static_cast<void>(::RevertToSelf()); });

    std::wstring authorizationPath = normalizedRoot;
    DWORD lastMissingError         = ERROR_PATH_NOT_FOUND;
    for (;;)
    {
        const DWORD attributes = ::GetFileAttributesW(authorizationPath.c_str());
        if (attributes != INVALID_FILE_ATTRIBUTES)
        {
            return OpenExistingClientPathForAuthorization(authorizationPath, attributes);
        }

        const DWORD error = ::GetLastError();
        if (! IsMissingPathError(error))
        {
            return HRESULT_FROM_WIN32(error);
        }

        lastMissingError        = error;
        std::wstring parentPath = ParentPathForAuthorization(authorizationPath);
        if (parentPath.empty() || OrdinalString::EqualsNoCase(parentPath, authorizationPath))
        {
            return HRESULT_FROM_WIN32(lastMissingError);
        }

        authorizationPath = std::move(parentPath);
    }
}

[[nodiscard]] HRESULT ValidateAndAuthorizeClientRoot(SessionContext& session, std::wstring_view rootPath, std::wstring& outRoot) noexcept
{
    HRESULT hr = NormalizeClientRequestedRoot(rootPath, outRoot);
    if (FAILED(hr))
    {
        return hr;
    }

#if defined(RS_SEARCH_TEST_HOOKS)
    if (session.options.testHook == ServerTestHook::RejectMarkedRoot && outRoot.find(L"service-reject-root") != std::wstring::npos)
    {
        return E_INVALIDARG;
    }
#endif

    return AuthorizeClientRootAccess(session.pipe.get(), outRoot);
}

[[nodiscard]] HRESULT ValidateAndAuthorizeRebuildRoot(SessionContext& session, std::wstring_view rootPath, std::wstring& outRoot) noexcept
{
    HRESULT hr = NormalizeClientRequestedRoot(rootPath, outRoot);
    if (FAILED(hr))
    {
        return hr;
    }

    return AuthorizeClientRootRebuildAccess(session.pipe.get(), outRoot);
}

HRESULT HandleStatusRequest(SessionContext& session) noexcept
{
    const ServerEventDetails details{
        .requestType = SEARCH_SERVICE_SERVER_REQUEST_STATUS,
        .phase       = FILESYSTEM_SEARCH_PHASE_INITIALIZING,
    };
    NotifyServerEvent(session.options, SEARCH_SERVICE_SERVER_EVENT_REQUEST_RECEIVED, S_OK, session.handledRequests, &details);
    return SendStatusResponse(session.pipe.get(),
                              session.options.protocolVersion,
                              session.options,
                              session.maintenanceState,
                              session.startupWarmupState,
                              session.repository,
                              &session.ioContext);
}

HRESULT HandleRebuildRequest(SessionContext& session, std::span<const std::byte> payloadBytes) noexcept
{
    if (! session.options.allowRebuildRequests)
    {
        return SendAck(session.pipe.get(), session.options.protocolVersion, E_ACCESSDENIED, &session.ioContext);
    }

    RebuildRequestPayload payload{};
    if (! ReadPod(payloadBytes, payload))
    {
        return SendProtocolError(session.pipe.get(), session.options.protocolVersion, kProtocolErrorHr, L"Invalid rebuild request.", &session.ioContext);
    }

    std::wstring rootPath;
    if (! ReadUtf16(payloadBytes, payload.rootPathBytes, rootPath))
    {
        return SendProtocolError(session.pipe.get(), session.options.protocolVersion, kProtocolErrorHr, L"Invalid rebuild path.", &session.ioContext);
    }

    std::wstring authorizedRoot;
    HRESULT rootHr = ValidateAndAuthorizeRebuildRoot(session, rootPath, authorizedRoot);
    if (FAILED(rootHr))
    {
        Debug::Warning(L"SearchServiceBroker: rebuild request rejected root='{}'. hr=0x{:08X}", rootPath, static_cast<unsigned long>(rootHr));
        return SendProtocolError(session.pipe.get(), session.options.protocolVersion, rootHr, L"Rejected rebuild root.", &session.ioContext);
    }
    rootPath = std::move(authorizedRoot);

    const ServerEventDetails details{
        .requestType = SEARCH_SERVICE_SERVER_REQUEST_REBUILD,
        .phase       = FILESYSTEM_SEARCH_PHASE_INITIALIZING,
        .rootPath    = rootPath.c_str(),
        .currentPath = rootPath.c_str(),
    };
    NotifyServerEvent(session.options, SEARCH_SERVICE_SERVER_EVENT_REQUEST_RECEIVED, S_OK, session.handledRequests, &details);

    const HRESULT hr = session.repository->InvalidateRoot(rootPath, true);
    Debug::Info(L"SearchServiceBroker: rebuild request root='{}' hr=0x{:08X}", rootPath, static_cast<unsigned long>(hr));
    return SendAck(session.pipe.get(), session.options.protocolVersion, hr, &session.ioContext);
}

HRESULT HandleCompactRequest(SessionContext& session) noexcept
{
    if (session.options.persistentStoreKind != LocalSearchIndexCore::PersistentStoreKind::Sqlite)
    {
        Debug::Warning(L"SearchServiceBroker: compact request rejected because the service backend is not sqlite.");
        return SendAck(session.pipe.get(), session.options.protocolVersion, HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), &session.ioContext);
    }

    const LocalSearchIndexCore::PersistentStoreInfo storeInfo = LocalSearchIndexCore::GetPersistentStoreInfo(BuildRepositoryOptions(session.options));
    const ServerEventDetails requestDetails{
        .requestType = SEARCH_SERVICE_SERVER_REQUEST_COMPACT,
        .phase       = FILESYSTEM_SEARCH_PHASE_INITIALIZING,
        .rootPath    = storeInfo.primaryPath.c_str(),
        .currentPath = storeInfo.primaryPath.c_str(),
    };
    NotifyServerEvent(session.options, SEARCH_SERVICE_SERVER_EVENT_REQUEST_RECEIVED, S_OK, session.handledRequests, &requestDetails);

    if (session.maintenanceState != nullptr)
    {
        SetQueuedMaintenanceState(*session.maintenanceState, false);
        session.maintenanceState->running = true;
    }

    const ServerEventDetails runningDetails{
        .requestType = SEARCH_SERVICE_SERVER_REQUEST_COMPACT,
        .phase       = FILESYSTEM_SEARCH_PHASE_INITIALIZING,
        .rootPath    = storeInfo.primaryPath.c_str(),
        .currentPath = storeInfo.primaryPath.c_str(),
    };
    NotifyServerEvent(session.options, SEARCH_SERVICE_SERVER_EVENT_MAINTENANCE_RUNNING, S_OK, session.handledRequests, &runningDetails);

    SqliteIndexStore::ManualMaintenanceResult maintenance{};
    const HRESULT hr = SqliteIndexStore::RunManualMaintenance(storeInfo.primaryPath, &maintenance);

    if (session.maintenanceState != nullptr)
    {
        session.maintenanceState->running = false;
        RefreshAutomaticMaintenanceQueue(session.options, *session.maintenanceState, session.handledRequests);
    }

    const std::wstring_view completedPath =
        maintenance.after.databasePath.empty() ? std::wstring_view(storeInfo.primaryPath) : std::wstring_view(maintenance.after.databasePath);
    const ServerEventDetails completedDetails{
        .requestType = SEARCH_SERVICE_SERVER_REQUEST_COMPACT,
        .phase       = FILESYSTEM_SEARCH_PHASE_COMPLETED,
        .rootPath    = completedPath.data(),
        .currentPath = completedPath.data(),
    };
    NotifyServerEvent(session.options, SEARCH_SERVICE_SERVER_EVENT_MAINTENANCE_COMPLETED, hr, session.handledRequests, &completedDetails);

    if (SUCCEEDED(hr))
    {
        Debug::Info(L"SearchServiceBroker: compact request completed path='{}' db {} -> {} wal {} -> {} freePages {} -> {}",
                    completedPath,
                    maintenance.before.databaseBytes,
                    maintenance.after.databaseBytes,
                    maintenance.before.writeAheadLogBytes,
                    maintenance.after.writeAheadLogBytes,
                    maintenance.before.freelistPageCount,
                    maintenance.after.freelistPageCount);
    }
    else
    {
        Debug::Warning(L"SearchServiceBroker: compact request failed path='{}' hr=0x{:08X}", storeInfo.primaryPath, static_cast<unsigned long>(hr));
    }

    return SendAck(session.pipe.get(), session.options.protocolVersion, hr, &session.ioContext);
}

HRESULT HandleShutdownRequest(SessionContext& session, std::span<const std::byte> payloadBytes) noexcept
{
    if (! payloadBytes.empty())
    {
        return SendProtocolError(session.pipe.get(), session.options.protocolVersion, kProtocolErrorHr, L"Invalid shutdown request.", &session.ioContext);
    }

    if (! session.options.allowShutdownRequests)
    {
        return SendAck(session.pipe.get(), session.options.protocolVersion, E_ACCESSDENIED, &session.ioContext);
    }

    const ServerEventDetails details{
        .requestType = SEARCH_SERVICE_SERVER_REQUEST_SHUTDOWN,
        .phase       = FILESYSTEM_SEARCH_PHASE_INITIALIZING,
    };
    NotifyServerEvent(session.options, SEARCH_SERVICE_SERVER_EVENT_REQUEST_RECEIVED, S_OK, session.handledRequests, &details);
    session.shutdownRequested = true;
    return SendAck(session.pipe.get(), session.options.protocolVersion, S_OK, &session.ioContext);
}

HRESULT HandleQueryRequest(SessionContext& session, std::span<const std::byte> payloadBytes) noexcept
{
    QueryRequestPayload payload{};
    if (! ReadPod(payloadBytes, payload))
    {
        return SendProtocolError(session.pipe.get(), session.options.protocolVersion, kProtocolErrorHr, L"Invalid query request.", &session.ioContext);
    }

    std::wstring rootPath;
    std::wstring namePattern;
    if (! ReadUtf16(payloadBytes, payload.rootPathBytes, rootPath) || ! ReadUtf16(payloadBytes, payload.namePatternBytes, namePattern))
    {
        return SendProtocolError(session.pipe.get(), session.options.protocolVersion, kProtocolErrorHr, L"Invalid query strings.", &session.ioContext);
    }

    std::wstring authorizedRoot;
    HRESULT rootHr = ValidateAndAuthorizeClientRoot(session, rootPath, authorizedRoot);
    if (FAILED(rootHr))
    {
        Debug::Warning(L"SearchServiceBroker: query request rejected root='{}'. hr=0x{:08X}", rootPath, static_cast<unsigned long>(rootHr));
        return SendProtocolError(session.pipe.get(), session.options.protocolVersion, rootHr, L"Rejected query root.", &session.ioContext);
    }

    LocalSearchIndexCore::QueryPlan plan{};
    plan.rootPath           = std::move(authorizedRoot);
    plan.namePattern        = std::move(namePattern);
    plan.nameMode           = static_cast<FileSystemSearchNameMode>(payload.nameMode);
    plan.matchCaseName      = (payload.flags & FILESYSTEM_SEARCH_MATCH_CASE_NAME) != 0;
    plan.recursive          = (payload.flags & FILESYSTEM_SEARCH_RECURSIVE) != 0;
    plan.includeFiles       = (payload.flags & FILESYSTEM_SEARCH_INCLUDE_FILES) != 0;
    plan.includeDirectories = (payload.flags & FILESYSTEM_SEARCH_INCLUDE_DIRECTORIES) != 0;
    plan.maxResults         = payload.maxResults;

    const ServerEventDetails receivedDetails{
        .requestType            = SEARCH_SERVICE_SERVER_REQUEST_QUERY,
        .phase                  = FILESYSTEM_SEARCH_PHASE_INITIALIZING,
        .requestFlags           = payload.flags,
        .requestNameMode        = payload.nameMode,
        .batchesSent            = 0u,
        .scannedDirectories     = 0u,
        .scannedFiles           = 0u,
        .candidateFiles         = 0u,
        .matchedEntries         = 0u,
        .snapshotFileBytes      = 0u,
        .estimatedMemoryBytes   = 0u,
        .ensureReadyDurationMs  = 0u,
        .executeQueryDurationMs = 0u,
        .rootPath               = plan.rootPath.c_str(),
        .namePattern            = plan.namePattern.c_str(),
        .currentPath            = plan.rootPath.c_str(),
    };
    NotifyServerEvent(session.options, SEARCH_SERVICE_SERVER_EVENT_REQUEST_RECEIVED, S_OK, session.handledRequests, &receivedDetails);

    LocalSearchIndexCore::QueryStats stats{};
    ServerQueryContext queryContext{
        .pipe            = session.pipe.get(),
        .protocolVersion = session.options.protocolVersion,
        .ioContext       = &session.ioContext,
        .handledRequests = session.handledRequests,
        .options         = &session.options,
        .stats           = &stats,
        .requestFlags    = payload.flags,
        .requestNameMode = payload.nameMode,
        .rootPath        = &plan.rootPath,
        .namePattern     = &plan.namePattern,
    };
    ServerCandidateBatchState batchState{
        .pipe                   = session.pipe.get(),
        .protocolVersion        = session.options.protocolVersion,
        .ioContext              = &session.ioContext,
        .disconnectAfterBatches = session.options.disconnectAfterBatches,
        .handledRequests        = session.handledRequests,
        .options                = &session.options,
        .stats                  = &stats,
        .requestFlags           = payload.flags,
        .requestNameMode        = payload.nameMode,
        .rootPath               = plan.rootPath,
        .namePattern            = plan.namePattern,
#if defined(RS_SEARCH_TEST_HOOKS)
        .testHook = session.options.testHook,
#endif
    };
    const auto buildCompletionDetails = [&]() noexcept
    {
        return ServerEventDetails{
            .requestType            = SEARCH_SERVICE_SERVER_REQUEST_QUERY,
            .phase                  = FILESYSTEM_SEARCH_PHASE_COMPLETED,
            .requestFlags           = payload.flags,
            .requestNameMode        = payload.nameMode,
            .batchesSent            = batchState.batchesSent,
            .scannedDirectories     = stats.directoryCount,
            .scannedFiles           = stats.fileCount,
            .candidateFiles         = stats.candidateCount,
            .matchedEntries         = 0u,
            .snapshotFileBytes      = stats.snapshotFileBytes,
            .estimatedMemoryBytes   = stats.estimatedMemoryBytes,
            .ensureReadyDurationMs  = stats.ensureReadyDurationMs,
            .executeQueryDurationMs = stats.executeQueryDurationMs,
            .warningFlags           = QueryStatsWarningFlags(stats) | batchState.warningFlags,
            .storeState =
                static_cast<uint32_t>(stats.usedLiveScanFallback ? ResolveFallbackStoreState(stats.fallbackReason) : LocalSearchIndexCore::StoreState::Ready),
            .syncPhase = static_cast<uint32_t>(stats.usedLiveScanFallback ? LocalSearchIndexCore::SyncPhase::Idle : LocalSearchIndexCore::SyncPhase::Watching),
            .queryExecutionMode = static_cast<uint32_t>(stats.queryExecutionMode),
            .fallbackReason     = static_cast<uint32_t>(stats.fallbackReason),
            .rootPath           = plan.rootPath.c_str(),
            .namePattern        = plan.namePattern.c_str(),
            .currentPath        = plan.rootPath.c_str(),
            .activeRoot         = plan.rootPath.c_str(),
        };
    };
    HRESULT hr = session.options.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite
                     ? session.repository->EnumerateNoWait(
                           plan, &ServerCancelCheck, &queryContext, &StreamServerCandidate, &batchState, &stats, &StreamServerQueryProgress, &queryContext)
                     : session.repository->Enumerate(
                           plan, &ServerCancelCheck, &queryContext, &StreamServerCandidate, &batchState, &stats, &StreamServerQueryProgress, &queryContext);
    stats.warningFlags |= batchState.warningFlags;
    if (FAILED(hr) && hr == HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE))
    {
        return hr;
    }

    if (SUCCEEDED(hr))
    {
        hr = FlushServerCandidates(batchState);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    if (FAILED(hr))
    {
        Debug::Warning(L"SearchServiceBroker: query failed root='{}' pattern='{}' mode={} hr=0x{:08X} candidates={} snapshotBytes={} memoryBytes={} "
                       L"readyMs={} queryMs={}",
                       plan.rootPath,
                       plan.namePattern,
                       static_cast<uint32_t>(plan.nameMode),
                       static_cast<unsigned long>(hr),
                       stats.candidateCount,
                       stats.snapshotFileBytes,
                       stats.estimatedMemoryBytes,
                       stats.ensureReadyDurationMs,
                       stats.executeQueryDurationMs);
        const ServerEventDetails completeDetails = buildCompletionDetails();
        NotifyServerEvent(session.options, SEARCH_SERVICE_SERVER_EVENT_QUERY_COMPLETED, hr, session.handledRequests, &completeDetails);
        return SendQueryComplete(session.pipe.get(), session.options.protocolVersion, hr, stats, &session.ioContext);
    }

    Debug::Info(L"SearchServiceBroker: query root='{}' pattern='{}' mode={} recursive={} includeFiles={} includeDirs={} candidates={} "
                L"snapshotBytes={} memoryBytes={} readyMs={} queryMs={} batches={}",
                plan.rootPath,
                plan.namePattern,
                static_cast<uint32_t>(plan.nameMode),
                plan.recursive,
                plan.includeFiles,
                plan.includeDirectories,
                stats.candidateCount,
                stats.snapshotFileBytes,
                stats.estimatedMemoryBytes,
                stats.ensureReadyDurationMs,
                stats.executeQueryDurationMs,
                batchState.batchesSent);
    const ServerEventDetails completeDetails = buildCompletionDetails();
    NotifyServerEvent(session.options, SEARCH_SERVICE_SERVER_EVENT_QUERY_COMPLETED, S_OK, session.handledRequests, &completeDetails);
    return SendQueryComplete(session.pipe.get(), session.options.protocolVersion, S_OK, stats, &session.ioContext);
}

HRESULT HandleClient(SessionContext& session) noexcept
{
    FrameHeader header{};
    std::vector<std::byte> payload;
    HRESULT hr = ReceiveFrame(session.pipe.get(), header, payload, &session.ioContext);
    if (FAILED(hr))
    {
        return hr;
    }

    if (header.protocolVersion != session.options.protocolVersion)
    {
        return SendProtocolError(session.pipe.get(), session.options.protocolVersion, kProtocolErrorHr, L"Protocol version mismatch.", &session.ioContext);
    }

    std::span<const std::byte> payloadBytes(payload.data(), payload.size());
    switch (static_cast<MessageType>(header.messageType))
    {
        case MessageType::StatusRequest: return HandleStatusRequest(session);
        case MessageType::QueryRequest: return HandleQueryRequest(session, payloadBytes);
        case MessageType::RebuildRequest: return HandleRebuildRequest(session, payloadBytes);
        case MessageType::CompactRequest: return HandleCompactRequest(session);
        case MessageType::ShutdownRequest: return HandleShutdownRequest(session, payloadBytes);
        case MessageType::StatusResponse:
        case MessageType::QueryProgress:
        case MessageType::QueryBatch:
        case MessageType::QueryComplete:
        case MessageType::Ack:
        case MessageType::Error:
        default: return SendProtocolError(session.pipe.get(), session.options.protocolVersion, kProtocolErrorHr, L"Unsupported request.", &session.ioContext);
    }
}

HRESULT CreatePipeSecurity(SECURITY_ATTRIBUTES& outAttributes, wil::unique_hlocal& outDescriptor) noexcept
{
    outAttributes = {};
    outDescriptor.reset();

    static constexpr wchar_t kPipeSddl[] = L"D:P(A;;GA;;;SY)(A;;GA;;;BA)(A;;GRGW;;;IU)";

    PSECURITY_DESCRIPTOR descriptor = nullptr;
    if (::ConvertStringSecurityDescriptorToSecurityDescriptorW(kPipeSddl, SDDL_REVISION_1, &descriptor, nullptr) == 0)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    outDescriptor.reset(descriptor);
    outAttributes.nLength              = sizeof(outAttributes);
    outAttributes.lpSecurityDescriptor = outDescriptor.get();
    outAttributes.bInheritHandle       = FALSE;
    return S_OK;
}

HRESULT CreateServerPipe(const std::wstring& pipeName, SECURITY_ATTRIBUTES& securityAttributes, wil::unique_handle& outPipe) noexcept
{
    outPipe.reset(::CreateNamedPipeW(pipeName.c_str(),
                                     PIPE_ACCESS_DUPLEX | FILE_FLAG_OVERLAPPED,
                                     PIPE_TYPE_BYTE | PIPE_READMODE_BYTE | PIPE_WAIT | PIPE_REJECT_REMOTE_CLIENTS,
                                     PIPE_UNLIMITED_INSTANCES,
                                     256u * 1024u,
                                     256u * 1024u,
                                     0u,
                                     &securityAttributes));
    if (! outPipe)
    {
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    return S_OK;
}

HRESULT WaitForClientConnection(HANDLE pipe, HANDLE stopEvent, DWORD timeoutMs) noexcept
{
    wil::unique_event connectEvent;
    connectEvent.create();
    if (! connectEvent)
    {
        return E_OUTOFMEMORY;
    }

    OVERLAPPED overlapped{};
    overlapped.hEvent = connectEvent.get();

    if (::ConnectNamedPipe(pipe, &overlapped) != 0)
    {
        return S_OK;
    }

    const DWORD error = ::GetLastError();
    if (error == ERROR_PIPE_CONNECTED)
    {
        return S_OK;
    }

    if (error != ERROR_IO_PENDING)
    {
        return HRESULT_FROM_WIN32(error);
    }

    HANDLE waits[2]              = {stopEvent, connectEvent.get()};
    const DWORD waitCount        = stopEvent ? 2u : 1u;
    const DWORD effectiveTimeout = timeoutMs;
    const DWORD waitResult       = ::WaitForMultipleObjects(waitCount, stopEvent ? waits : waits + 1, FALSE, effectiveTimeout);
    if (stopEvent && waitResult == WAIT_OBJECT_0)
    {
        static_cast<void>(::CancelIoEx(pipe, &overlapped));
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    if (waitResult == WAIT_TIMEOUT)
    {
        static_cast<void>(::CancelIoEx(pipe, &overlapped));
        DWORD ignoredBytes = 0u;
        static_cast<void>(::GetOverlappedResult(pipe, &overlapped, &ignoredBytes, FALSE));
        return S_FALSE;
    }
    if (waitResult == WAIT_FAILED)
    {
        static_cast<void>(::CancelIoEx(pipe, &overlapped));
        return HRESULT_FROM_WIN32(::GetLastError());
    }

    DWORD bytesTransferred = 0u;
    if (::GetOverlappedResult(pipe, &overlapped, &bytesTransferred, FALSE) == 0)
    {
        const DWORD resultError = ::GetLastError();
        if (resultError == ERROR_PIPE_CONNECTED)
        {
            return S_OK;
        }

        return HRESULT_FROM_WIN32(resultError);
    }

    return S_OK;
}

[[nodiscard]] bool IsRecoverableClientFailure(HRESULT hr) noexcept
{
    return hr == kProtocolErrorHr || hr == HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE) || hr == HRESULT_FROM_WIN32(ERROR_NO_DATA) ||
           hr == HRESULT_FROM_WIN32(ERROR_PIPE_NOT_CONNECTED) || hr == HRESULT_FROM_WIN32(ERROR_SEM_TIMEOUT) ||
           hr == HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW) || hr == HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME) || hr == E_INVALIDARG || hr == E_ACCESSDENIED ||
           hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
}
} // namespace

#ifdef ENABLE_TESTS
HRESULT DecodeQueryBatchForTests(std::span<const std::byte> payloadBytes, std::vector<LocalSearchIndexCore::Candidate>& outCandidates) noexcept
{
    return DecodeQueryBatchPayload(payloadBytes, outCandidates);
}
#endif

std::wstring GetDefaultPipeName() noexcept
{
    return BuildPipeName({});
}

std::wstring GetConfiguredPipeName() noexcept
{
    const std::wstring envPipe = GetEnvironmentValue(kPipeNameEnvVar);
    return BuildPipeName(envPipe);
}

std::wstring GetProgramDataSearchIndexRoot() noexcept
{
    const std::wstring root = GetEnvironmentValue(L"ProgramData");
    if (root.empty())
    {
        return {};
    }

#ifdef _DEBUG
    return (std::filesystem::path(root) / L"RedSalamander" / L"SearchIndex.Debug").wstring();
#else
    return (std::filesystem::path(root) / L"RedSalamander" / L"SearchIndex").wstring();
#endif
}

HRESULT GetStatus(ServiceStatus& outStatus) noexcept
{
    try
    {
        outStatus               = {};
        std::wstring_view stage = L"connect";
        const auto logFailure   = [&](const HRESULT hr) noexcept
        {
            Debug::Warning(L"SearchServiceBroker: status probe failed at stage='{}' hr=0x{:08X}", stage, static_cast<unsigned long>(hr));
            return hr;
        };

        ClientIoContext ioContext{};
        ioContext.operationDeadline = MakeDeadline(kClientControlOperationTimeoutMs);

        wil::unique_handle pipe;
        HRESULT hr = ConnectClientPipe(pipe, &ioContext);
        if (FAILED(hr))
        {
            return logFailure(hr);
        }

        std::vector<std::byte> emptyPayload;
        stage = L"send status request";
        hr    = ClientSendFrame(pipe.get(), MessageType::StatusRequest, kProtocolVersion, emptyPayload, &ioContext);
        if (FAILED(hr))
        {
            return logFailure(hr);
        }

        FrameHeader header{};
        std::vector<std::byte> payload;
        stage = L"receive status response";
        hr    = ClientReceiveFrame(pipe.get(), header, payload, &ioContext);
        if (FAILED(hr))
        {
            return logFailure(hr);
        }

        if (header.protocolVersion != kProtocolVersion)
        {
            return logFailure(kProtocolErrorHr);
        }

        std::span<const std::byte> remaining(payload.data(), payload.size());
        if (static_cast<MessageType>(header.messageType) == MessageType::Error)
        {
            ErrorPayload error{};
            return ReadPod(remaining, error) ? logFailure(static_cast<HRESULT>(error.result)) : logFailure(kProtocolErrorHr);
        }

        if (static_cast<MessageType>(header.messageType) != MessageType::StatusResponse)
        {
            return logFailure(kProtocolErrorHr);
        }

        StatusResponsePayload statusPayload{};
        if (! ReadPod(remaining, statusPayload) || ! ReadUtf16(remaining, statusPayload.pipeNameBytes, outStatus.pipeName) ||
            ! ReadUtf16(remaining, statusPayload.storageRootBytes, outStatus.storageRootDirectory))
        {
            outStatus = {};
            return logFailure(kProtocolErrorHr);
        }

        outStatus.protocolVersion = header.protocolVersion;
        outStatus.processId       = statusPayload.processId;
        if (remaining.size_bytes() >= sizeof(StatusResponseExtendedPayload))
        {
            StatusResponseExtendedPayload extended{};
            if (! ReadPod(remaining, extended) || ! ReadUtf16(remaining, extended.persistentStorePathBytes, outStatus.persistentStorePath) ||
                ! ReadUtf16(remaining, extended.writeAheadLogPathBytes, outStatus.writeAheadLogPath))
            {
                outStatus = {};
                return logFailure(kProtocolErrorHr);
            }

            outStatus.persistentStoreKind                = static_cast<LocalSearchIndexCore::PersistentStoreKind>(extended.persistentStoreKind);
            outStatus.persistentStoreBytes               = extended.persistentStoreBytes;
            outStatus.writeAheadLogBytes                 = extended.writeAheadLogBytes;
            outStatus.persistentStoreInspectionSucceeded = (extended.flags & STATUS_RESPONSE_FLAG_STORE_INSPECTED) != 0u;
            outStatus.readyForQueryCutover               = (extended.flags & STATUS_RESPONSE_FLAG_READY_FOR_QUERY_CUTOVER) != 0u;
            outStatus.indexedVolumeCount                 = extended.indexedVolumeCount;
            outStatus.indexedEntryCount                  = extended.indexedEntryCount;
            outStatus.legacyImportVolumeCount            = extended.legacyImportVolumeCount;
            outStatus.autoCheckpointEnabled              = (extended.flags & STATUS_RESPONSE_FLAG_AUTO_CHECKPOINT) != 0u;
            outStatus.autoCheckpointTargetBytes          = extended.autoCheckpointTargetBytes;
            outStatus.autoCompactionEnabled              = (extended.flags & STATUS_RESPONSE_FLAG_AUTO_COMPACTION) != 0u;
            outStatus.autoCompactionFragmentationPercent = extended.autoCompactionFragmentationPercent;
            outStatus.autoCompactionMinBytes             = extended.autoCompactionMinBytes;

            if (remaining.size_bytes() >= sizeof(StatusResponseMaintenancePayload))
            {
                StatusResponseMaintenancePayload maintenance{};
                if (! ReadPod(remaining, maintenance) || ! ReadUtf16(remaining, maintenance.lastCheckpointUtcBytes, outStatus.lastCheckpointUtc) ||
                    ! ReadUtf16(remaining, maintenance.lastCompactionUtcBytes, outStatus.lastCompactionUtc))
                {
                    outStatus = {};
                    return logFailure(kProtocolErrorHr);
                }

                outStatus.persistentStorePageCount         = maintenance.pageCount;
                outStatus.persistentStoreFreelistPageCount = maintenance.freelistPageCount;
                outStatus.maintenanceQueued                = (maintenance.flags & STATUS_RESPONSE_MAINTENANCE_FLAG_QUEUED) != 0u;
                outStatus.maintenanceRunning               = (maintenance.flags & STATUS_RESPONSE_MAINTENANCE_FLAG_RUNNING) != 0u;
            }

            if (remaining.size_bytes() >= sizeof(StatusResponseDiscoveryPayload))
            {
                StatusResponseDiscoveryPayload discovery{};
                std::wstring discoveredRootsText;
                if (! ReadPod(remaining, discovery) || ! ReadUtf16(remaining, discovery.discoveredRootsBytes, discoveredRootsText))
                {
                    outStatus = {};
                    return logFailure(kProtocolErrorHr);
                }

                outStatus.discoveredRootCount = discovery.discoveredRootCount;
                outStatus.discoveredRoots.clear();

                size_t offset = 0u;
                while (offset < discoveredRootsText.size())
                {
                    const size_t separator = discoveredRootsText.find(L'\n', offset);
                    const size_t tokenEnd  = separator == std::wstring::npos ? discoveredRootsText.size() : separator;
                    if (tokenEnd > offset)
                    {
                        outStatus.discoveredRoots.emplace_back(discoveredRootsText.substr(offset, tokenEnd - offset));
                    }
                    if (separator == std::wstring::npos)
                    {
                        break;
                    }

                    offset = separator + 1u;
                }
            }

            if (remaining.size_bytes() >= sizeof(StatusResponseStartupWarmupPayload))
            {
                StatusResponseStartupWarmupPayload warmup{};
                if (! ReadPod(remaining, warmup) || ! ReadUtf16(remaining, warmup.currentRootBytes, outStatus.startupWarmupCurrentRoot))
                {
                    outStatus = {};
                    return logFailure(kProtocolErrorHr);
                }

                outStatus.startupWarmupEnabled        = (warmup.flags & STATUS_RESPONSE_STARTUP_WARMUP_FLAG_ENABLED) != 0u;
                outStatus.startupWarmupRunning        = (warmup.flags & STATUS_RESPONSE_STARTUP_WARMUP_FLAG_RUNNING) != 0u;
                outStatus.startupWarmupTotalRoots     = warmup.totalRoots;
                outStatus.startupWarmupCompletedRoots = warmup.completedRoots;
                outStatus.startupWarmupFailedRoots    = warmup.failedRoots;
                outStatus.startupWarmupHasFailure     = (warmup.flags & STATUS_RESPONSE_STARTUP_WARMUP_FLAG_HAS_FAILURE) != 0u;
            }

            if (remaining.size_bytes() >= sizeof(StatusResponseStartupWarmupFailurePayload))
            {
                StatusResponseStartupWarmupFailurePayload warmupFailure{};
                if (! ReadPod(remaining, warmupFailure) || ! ReadUtf16(remaining, warmupFailure.lastFailureRootBytes, outStatus.startupWarmupLastFailureRoot))
                {
                    outStatus = {};
                    return logFailure(kProtocolErrorHr);
                }

                outStatus.startupWarmupHasFailure =
                    outStatus.startupWarmupHasFailure || ((warmupFailure.flags & STATUS_RESPONSE_STARTUP_WARMUP_FLAG_HAS_FAILURE) != 0u);
                outStatus.startupWarmupLastFailureHr = warmupFailure.lastFailureHr;
            }

            if (remaining.size_bytes() >= sizeof(StatusResponseRuntimePayload))
            {
                StatusResponseRuntimePayload runtime{};
                if (! ReadPod(remaining, runtime) || ! ReadUtf16(remaining, runtime.activeRootBytes, outStatus.activeRoot))
                {
                    outStatus = {};
                    return logFailure(kProtocolErrorHr);
                }

                outStatus.storeState         = static_cast<LocalSearchIndexCore::StoreState>(runtime.storeState);
                outStatus.syncPhase          = static_cast<LocalSearchIndexCore::SyncPhase>(runtime.syncPhase);
                outStatus.queryExecutionMode = static_cast<LocalSearchIndexCore::QueryExecutionMode>(runtime.queryExecutionMode);
                outStatus.fallbackReason     = static_cast<LocalSearchIndexCore::FallbackReason>(runtime.fallbackReason);
                outStatus.completedRoots     = runtime.completedRoots;
                outStatus.totalRoots         = runtime.totalRoots;
            }
        }
#if defined(_DEBUG) || defined(__SANITIZE_ADDRESS__)
        Debug::Info(L"SearchServiceBroker: status probe succeeded pid={} pipe='{}' storage='{}' backend='{}' store='{}' inspected={} ready={} volumes={} "
                    L"entries={} legacyImports={} pages={} freePages={} checkpoint='{}' compact='{}' queued={} running={} discoveredRoots={} startupWarmup={} "
                    L"{}/{} failed={} root='{}' lastFailure={} lastFailureRoot='{}'",
                    outStatus.processId,
                    outStatus.pipeName,
                    outStatus.storageRootDirectory,
                    LocalSearchIndexCore::GetPersistentStoreKindText(outStatus.persistentStoreKind),
                    outStatus.persistentStorePath,
                    outStatus.persistentStoreInspectionSucceeded,
                    outStatus.readyForQueryCutover,
                    outStatus.indexedVolumeCount,
                    outStatus.indexedEntryCount,
                    outStatus.legacyImportVolumeCount,
                    outStatus.persistentStorePageCount,
                    outStatus.persistentStoreFreelistPageCount,
                    outStatus.lastCheckpointUtc,
                    outStatus.lastCompactionUtc,
                    outStatus.maintenanceQueued,
                    outStatus.maintenanceRunning,
                    outStatus.discoveredRootCount,
                    outStatus.startupWarmupRunning,
                    outStatus.startupWarmupCompletedRoots,
                    outStatus.startupWarmupTotalRoots,
                    outStatus.startupWarmupFailedRoots,
                    outStatus.startupWarmupCurrentRoot,
                    static_cast<unsigned long>(outStatus.startupWarmupLastFailureHr),
                    outStatus.startupWarmupLastFailureRoot);
#endif
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"SearchServiceBroker: GetStatus failed with an unexpected std::exception.");
        outStatus = {};
        return E_FAIL;
    }
}

HRESULT Query(const QueryRequest& request,
              ProgressCallbackFn progressCallback,
              void* progressCookie,
              LocalSearchIndexCore::CancelCheckFn cancelCheck,
              void* cancelCookie,
              std::vector<LocalSearchIndexCore::Candidate>& outCandidates,
              LocalSearchIndexCore::QueryStats* outStats,
              CandidateBatchCallbackFn candidateBatchCallback,
              void* candidateBatchCookie) noexcept
{
    try
    {
        outCandidates.clear();
        if (outStats != nullptr)
        {
            *outStats = {};
        }

        if (request.rootPath.empty())
        {
            return E_INVALIDARG;
        }

        std::wstring_view stage = L"connect";
        const auto logFailure   = [&](const HRESULT hr) noexcept
        {
            Debug::Warning(L"SearchServiceBroker: query failed stage='{}' root='{}' pattern='{}' mode={} flags=0x{:08X} maxResults={} hr=0x{:08X}",
                           stage,
                           request.rootPath,
                           request.namePattern,
                           static_cast<uint32_t>(request.nameMode),
                           static_cast<unsigned long>(request.flags),
                           request.maxResults,
                           static_cast<unsigned long>(hr));
            return hr;
        };

        ClientIoContext ioContext{};
        ioContext.cancelCheck       = cancelCheck;
        ioContext.cancelCookie      = cancelCookie;
        ioContext.operationDeadline = MakeDeadline(kClientQueryOperationTimeoutMs);

        wil::unique_handle pipe;
        HRESULT hr = ConnectClientPipe(pipe, &ioContext);
        if (FAILED(hr))
        {
            return logFailure(hr);
        }
#if defined(_DEBUG) || defined(__SANITIZE_ADDRESS__)
        Debug::Info(L"SearchServiceBroker: query start root='{}' pattern='{}' mode={} flags=0x{:08X} maxResults={}",
                    request.rootPath,
                    request.namePattern,
                    static_cast<uint32_t>(request.nameMode),
                    static_cast<unsigned long>(request.flags),
                    request.maxResults);
#endif

        QueryRequestPayload payload{};
        payload.nameMode         = static_cast<uint32_t>(request.nameMode);
        payload.flags            = static_cast<uint32_t>(request.flags);
        payload.rootPathBytes    = static_cast<uint32_t>(request.rootPath.size() * sizeof(wchar_t));
        payload.namePatternBytes = static_cast<uint32_t>(request.namePattern.size() * sizeof(wchar_t));
        payload.maxResults       = request.maxResults;

        std::vector<std::byte> requestBuffer;
        requestBuffer.reserve(sizeof(payload) + payload.rootPathBytes + payload.namePatternBytes);
        AppendBytes(requestBuffer, &payload, sizeof(payload));
        AppendUtf16(requestBuffer, request.rootPath);
        AppendUtf16(requestBuffer, request.namePattern);

        stage = L"send query request";
        hr    = ClientSendFrame(pipe.get(), MessageType::QueryRequest, kProtocolVersion, requestBuffer, &ioContext);
        if (FAILED(hr))
        {
            return logFailure(hr);
        }
#if defined(_DEBUG) || defined(__SANITIZE_ADDRESS__)
        Debug::Info(L"SearchServiceBroker: query request sent root='{}' payloadBytes={}", request.rootPath, static_cast<unsigned long>(requestBuffer.size()));
#endif

        QueryProgress lastProgress{};
        bool sawProgress = false;
#if defined(_DEBUG) || defined(__SANITIZE_ADDRESS__)
        bool loggedFirstResponse = false;
#endif
        uint64_t consumedCandidates         = 0u;
        uint64_t bufferedCandidateBytes     = 0u;
        const size_t bufferedCandidateLimit = ResolveClientBufferedCandidateLimit(request.maxResults);
        for (;;)
        {
            if (cancelCheck != nullptr)
            {
                stage = L"cancel check";
                hr    = cancelCheck(cancelCookie);
                if (FAILED(hr))
                {
                    return logFailure(hr);
                }
            }

            FrameHeader header{};
            std::vector<std::byte> payloadBytes;
            stage = L"receive query response";
            hr    = ClientReceiveFrame(pipe.get(), header, payloadBytes, &ioContext);
            if (FAILED(hr))
            {
                return logFailure(hr);
            }

            if (header.protocolVersion != kProtocolVersion)
            {
                return logFailure(kProtocolErrorHr);
            }

            std::span<const std::byte> remaining(payloadBytes.data(), payloadBytes.size());
            switch (static_cast<MessageType>(header.messageType))
            {
                case MessageType::QueryProgress:
                {
                    ProgressPayload progressPayload{};
                    QueryProgress progress{};
                    if (! ReadPod(remaining, progressPayload) || ! ReadUtf16(remaining, progressPayload.currentPathBytes, progress.currentPath))
                    {
                        return logFailure(kProtocolErrorHr);
                    }

                    progress.phase              = static_cast<FileSystemSearchPhase>(progressPayload.phase);
                    progress.warningFlags       = progressPayload.warningFlags;
                    progress.statusHint         = progressPayload.statusHint;
                    progress.scannedDirectories = progressPayload.scannedDirectories;
                    progress.scannedFiles       = progressPayload.scannedFiles;
                    progress.candidateFiles     = progressPayload.candidateFiles;
                    progress.matchedEntries     = progressPayload.matchedEntries;
                    if (remaining.size_bytes() >= sizeof(ProgressRuntimePayload))
                    {
                        ProgressRuntimePayload runtime{};
                        if (! ReadPod(remaining, runtime) || ! ReadUtf16(remaining, runtime.activeRootBytes, progress.activeRoot))
                        {
                            return logFailure(kProtocolErrorHr);
                        }

                        progress.storeState         = static_cast<LocalSearchIndexCore::StoreState>(runtime.storeState);
                        progress.syncPhase          = static_cast<LocalSearchIndexCore::SyncPhase>(runtime.syncPhase);
                        progress.queryExecutionMode = static_cast<LocalSearchIndexCore::QueryExecutionMode>(runtime.queryExecutionMode);
                        progress.fallbackReason     = static_cast<LocalSearchIndexCore::FallbackReason>(runtime.fallbackReason);
                        progress.completedRoots     = runtime.completedRoots;
                        progress.totalRoots         = runtime.totalRoots;
                    }
                    lastProgress = progress;
                    sawProgress  = true;

#if defined(_DEBUG) || defined(__SANITIZE_ADDRESS__)
                    if (! loggedFirstResponse)
                    {
                        Debug::Info(L"SearchServiceBroker: first query progress root='{}' phase={} current='{}' dirs={} files={} candidates={} matches={}",
                                    request.rootPath,
                                    static_cast<uint32_t>(progress.phase),
                                    progress.currentPath,
                                    progress.scannedDirectories,
                                    progress.scannedFiles,
                                    progress.candidateFiles,
                                    progress.matchedEntries);
                        loggedFirstResponse = true;
                    }
#endif

                    if (progressCallback != nullptr)
                    {
                        stage = L"progress callback";
                        hr    = progressCallback(&progress, progressCookie);
                        if (FAILED(hr))
                        {
                            return logFailure(hr);
                        }
                    }
                    break;
                }

                case MessageType::QueryBatch:
                {
                    std::vector<LocalSearchIndexCore::Candidate> batchCandidates;
                    hr = DecodeQueryBatchPayload(remaining, batchCandidates);
                    if (FAILED(hr))
                    {
                        return logFailure(hr);
                    }

                    if (candidateBatchCallback != nullptr && ! batchCandidates.empty())
                    {
                        size_t batchConsumed = 0u;
                        stage                = L"candidate batch callback";
                        hr                   = candidateBatchCallback(batchCandidates.data(), batchCandidates.size(), &batchConsumed, candidateBatchCookie);
                        if (batchConsumed > batchCandidates.size())
                        {
                            return logFailure(E_FAIL);
                        }

                        if (hr == S_OK && batchConsumed != batchCandidates.size())
                        {
                            return logFailure(E_FAIL);
                        }

                        consumedCandidates += static_cast<uint64_t>(batchConsumed);
                        if (hr == S_FALSE)
                        {
                            if (outStats != nullptr)
                            {
                                outStats->candidateCount = consumedCandidates;
                                if (sawProgress)
                                {
                                    outStats->directoryCount     = lastProgress.scannedDirectories;
                                    outStats->fileCount          = lastProgress.scannedFiles;
                                    outStats->queryExecutionMode = lastProgress.queryExecutionMode;
                                    outStats->fallbackReason     = lastProgress.fallbackReason;
                                    outStats->usedLiveScanFallback =
                                        lastProgress.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback;
                                }
                            }
#if defined(_DEBUG) || defined(__SANITIZE_ADDRESS__)
                            Debug::Info(
                                L"SearchServiceBroker: query stopped early after {} consumed candidate(s) for root='{}'", consumedCandidates, request.rootPath);
#endif
                            return S_OK;
                        }
                        if (FAILED(hr))
                        {
                            return logFailure(hr);
                        }
                    }
                    else
                    {
                        uint64_t batchBufferedBytes = 0u;
                        if (! CanAppendClientBufferedCandidates(
                                outCandidates.size(),
                                bufferedCandidateBytes,
                                std::span<const LocalSearchIndexCore::Candidate>(batchCandidates.data(), batchCandidates.size()),
                                bufferedCandidateLimit,
                                batchBufferedBytes))
                        {
                            outCandidates.clear();
                            stage = L"candidate accumulation limit";
                            return logFailure(HRESULT_FROM_WIN32(ERROR_BUFFER_OVERFLOW));
                        }

                        consumedCandidates += static_cast<uint64_t>(batchCandidates.size());
                        outCandidates.insert(
                            outCandidates.end(), std::make_move_iterator(batchCandidates.begin()), std::make_move_iterator(batchCandidates.end()));
                        bufferedCandidateBytes += batchBufferedBytes;
                    }
                    break;
                }

                case MessageType::QueryComplete:
                {
                    QueryCompletePayload complete{};
                    if (! ReadPod(remaining, complete))
                    {
                        return logFailure(kProtocolErrorHr);
                    }

                    if (outStats != nullptr)
                    {
                        outStats->fileSystemKind         = static_cast<LocalSearchIndexCore::FileSystemKind>(complete.fileSystemKind);
                        outStats->entryCount             = complete.entryCount;
                        outStats->fileCount              = complete.fileCount;
                        outStats->directoryCount         = complete.directoryCount;
                        outStats->candidateCount         = complete.candidateCount;
                        outStats->nextUsn                = complete.nextUsn;
                        outStats->journalId              = complete.journalId;
                        outStats->snapshotFileBytes      = complete.snapshotFileBytes;
                        outStats->estimatedMemoryBytes   = complete.estimatedMemoryBytes;
                        outStats->ensureReadyDurationMs  = complete.ensureReadyDurationMs;
                        outStats->executeQueryDurationMs = complete.executeQueryDurationMs;
                        UnpackQueryStatsFlags(complete.flags, *outStats);
                        outStats->queryExecutionMode   = lastProgress.queryExecutionMode;
                        outStats->fallbackReason       = lastProgress.fallbackReason;
                        outStats->usedLiveScanFallback = lastProgress.queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback;
                        if (remaining.size_bytes() >= sizeof(QueryCompleteRuntimePayload))
                        {
                            QueryCompleteRuntimePayload runtime{};
                            if (! ReadPod(remaining, runtime))
                            {
                                return logFailure(kProtocolErrorHr);
                            }

                            outStats->queryExecutionMode   = static_cast<LocalSearchIndexCore::QueryExecutionMode>(runtime.queryExecutionMode);
                            outStats->fallbackReason       = static_cast<LocalSearchIndexCore::FallbackReason>(runtime.fallbackReason);
                            outStats->usedLiveScanFallback = outStats->queryExecutionMode == LocalSearchIndexCore::QueryExecutionMode::LiveScanFallback;
                            outStats->warningFlags         = runtime.warningFlags;
                        }
                    }

                    const HRESULT result = static_cast<HRESULT>(complete.result);
#if defined(_DEBUG) || defined(__SANITIZE_ADDRESS__)
                    Debug::Info(L"SearchServiceBroker: query complete root='{}' hr=0x{:08X} candidates={} dirs={} files={} readyMs={} queryMs={} "
                                L"sqliteStore={} prefilter={} cutoverBlocked={}",
                                request.rootPath,
                                static_cast<unsigned long>(result),
                                complete.candidateCount,
                                complete.directoryCount,
                                complete.fileCount,
                                complete.ensureReadyDurationMs,
                                complete.executeQueryDurationMs,
                                (complete.flags & QUERY_COMPLETE_FLAG_USED_SQLITE_STORE) != 0u,
                                (complete.flags & QUERY_COMPLETE_FLAG_USED_NAME_PREFILTER) != 0u,
                                (complete.flags & QUERY_COMPLETE_FLAG_SQLITE_CUTOVER_BLOCKED) != 0u);
#endif
                    return FAILED(result) ? logFailure(result) : result;
                }

                case MessageType::Error:
                {
                    ErrorPayload error{};
                    return ReadPod(remaining, error) ? logFailure(static_cast<HRESULT>(error.result)) : logFailure(kProtocolErrorHr);
                }

                case MessageType::StatusRequest:
                case MessageType::StatusResponse:
                case MessageType::QueryRequest:
                case MessageType::RebuildRequest:
                case MessageType::CompactRequest:
                case MessageType::ShutdownRequest:
                case MessageType::Ack:
                default: return logFailure(kProtocolErrorHr);
            }
        }
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"SearchServiceBroker: Query failed with an unexpected std::exception.");
        outCandidates.clear();
        if (outStats != nullptr)
        {
            *outStats = {};
        }
        return E_FAIL;
    }
}

HRESULT RequestRebuild(std::wstring_view rootPath) noexcept
{
    try
    {
        wil::unique_handle pipe;
        HRESULT hr = ConnectClientPipe(pipe);
        if (FAILED(hr))
        {
            return hr;
        }

        ClientIoContext ioContext{};
        ioContext.operationDeadline = MakeDeadline(kClientControlOperationTimeoutMs);

        RebuildRequestPayload payload{};
        payload.rootPathBytes = static_cast<uint32_t>(rootPath.size() * sizeof(wchar_t));

        std::vector<std::byte> buffer;
        buffer.reserve(sizeof(payload) + payload.rootPathBytes);
        AppendBytes(buffer, &payload, sizeof(payload));
        AppendUtf16(buffer, rootPath);

        hr = ClientSendFrame(pipe.get(), MessageType::RebuildRequest, kProtocolVersion, buffer, &ioContext);
        if (FAILED(hr))
        {
            return hr;
        }

        FrameHeader header{};
        std::vector<std::byte> payloadBytes;
        hr = ClientReceiveFrame(pipe.get(), header, payloadBytes, &ioContext);
        if (FAILED(hr))
        {
            return hr;
        }

        if (header.protocolVersion != kProtocolVersion)
        {
            return kProtocolErrorHr;
        }

        std::span<const std::byte> remaining(payloadBytes.data(), payloadBytes.size());
        if (static_cast<MessageType>(header.messageType) == MessageType::Ack)
        {
            AckPayload ack{};
            return ReadPod(remaining, ack) ? static_cast<HRESULT>(ack.result) : kProtocolErrorHr;
        }

        if (static_cast<MessageType>(header.messageType) == MessageType::Error)
        {
            ErrorPayload error{};
            return ReadPod(remaining, error) ? static_cast<HRESULT>(error.result) : kProtocolErrorHr;
        }

        return kProtocolErrorHr;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"SearchServiceBroker: RequestRebuild failed with an unexpected std::exception.");
        return E_FAIL;
    }
}

HRESULT RequestCompact() noexcept
{
    try
    {
        wil::unique_handle pipe;
        HRESULT hr = ConnectClientPipe(pipe);
        if (FAILED(hr))
        {
            return hr;
        }

        ClientIoContext ioContext{};
        ioContext.operationDeadline = MakeDeadline(kClientControlOperationTimeoutMs);

        hr = ClientSendFrame(pipe.get(), MessageType::CompactRequest, kProtocolVersion, {}, &ioContext);
        if (FAILED(hr))
        {
            return hr;
        }

        FrameHeader header{};
        std::vector<std::byte> payloadBytes;
        hr = ClientReceiveFrame(pipe.get(), header, payloadBytes, &ioContext);
        if (FAILED(hr))
        {
            return hr;
        }

        if (header.protocolVersion != kProtocolVersion)
        {
            return kProtocolErrorHr;
        }

        std::span<const std::byte> remaining(payloadBytes.data(), payloadBytes.size());
        if (static_cast<MessageType>(header.messageType) == MessageType::Ack)
        {
            AckPayload ack{};
            return ReadPod(remaining, ack) ? static_cast<HRESULT>(ack.result) : kProtocolErrorHr;
        }

        if (static_cast<MessageType>(header.messageType) == MessageType::Error)
        {
            ErrorPayload error{};
            return ReadPod(remaining, error) ? static_cast<HRESULT>(error.result) : kProtocolErrorHr;
        }

        return kProtocolErrorHr;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"SearchServiceBroker: RequestCompact failed with an unexpected std::exception.");
        return E_FAIL;
    }
}

HRESULT RequestShutdown(std::wstring_view pipeName, uint32_t timeoutMs) noexcept
{
    try
    {
        if (pipeName.empty() || timeoutMs == 0u)
        {
            return E_INVALIDARG;
        }

        wil::unique_handle pipe;
        const DWORD connectRetryMs = (std::min)(static_cast<DWORD>(timeoutMs), kDefaultMissingPipeRetryWindowMs);
        HRESULT hr                 = ConnectClientPipeForName(pipeName, connectRetryMs, pipe);
        if (FAILED(hr))
        {
            return hr;
        }

        ClientIoContext ioContext{};
        ioContext.operationDeadline = MakeDeadline(timeoutMs);

        hr = ClientSendFrame(pipe.get(), MessageType::ShutdownRequest, kProtocolVersion, {}, &ioContext);
        if (FAILED(hr))
        {
            return hr;
        }

        FrameHeader header{};
        std::vector<std::byte> payloadBytes;
        hr = ClientReceiveFrame(pipe.get(), header, payloadBytes, &ioContext);
        if (FAILED(hr))
        {
            return hr;
        }

        if (header.protocolVersion != kProtocolVersion)
        {
            return kProtocolErrorHr;
        }

        std::span<const std::byte> remaining(payloadBytes.data(), payloadBytes.size());
        if (static_cast<MessageType>(header.messageType) == MessageType::Ack)
        {
            AckPayload ack{};
            return ReadPod(remaining, ack) ? static_cast<HRESULT>(ack.result) : kProtocolErrorHr;
        }

        if (static_cast<MessageType>(header.messageType) == MessageType::Error)
        {
            ErrorPayload error{};
            return ReadPod(remaining, error) ? static_cast<HRESULT>(error.result) : kProtocolErrorHr;
        }

        return kProtocolErrorHr;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"SearchServiceBroker: RequestShutdown failed with an unexpected std::exception.");
        return E_FAIL;
    }
}

HRESULT RunServer(const ServerOptions& options, HANDLE stopEvent, ServerRunResult* outResult) noexcept
{
    try
    {
        ServerRunResult result{};
        const bool storeLocationOverridden = ! options.storageRootDirectory.empty() || ! options.sqliteDatabasePath.empty();

        ServerOptions effectiveOptions = options;
        effectiveOptions.pipeName      = BuildPipeName(options.pipeName);
        if (effectiveOptions.storageRootDirectory.empty())
        {
            effectiveOptions.storageRootDirectory = GetProgramDataSearchIndexRoot();
        }
        const LocalSearchIndexCore::RepositoryOptions repositoryOptions     = BuildRepositoryOptions(effectiveOptions);
        const LocalSearchIndexCore::PersistentStoreInfo configuredStoreInfo = LocalSearchIndexCore::GetPersistentStoreInfo(repositoryOptions);
        if (effectiveOptions.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite)
        {
            SqliteIndexStore::StoreInfo sqliteStoreInfo{};
            HRESULT hr = SqliteIndexStore::EnsureBootstrap(configuredStoreInfo.primaryPath, &sqliteStoreInfo);
            if (FAILED(hr))
            {
                Debug::Warning(L"SearchServiceBroker: SQLite bootstrap unavailable for '{}'. hr=0x{:08X}. continuing in degraded mode.",
                               configuredStoreInfo.primaryPath,
                               static_cast<unsigned long>(hr));
            }
            else
            {
                Debug::Info(L"SearchServiceBroker: SQLite store ready path='{}' schemaVersion={} wal={} autoVacuum={}",
                            sqliteStoreInfo.databasePath,
                            sqliteStoreInfo.schemaVersion,
                            sqliteStoreInfo.walEnabled,
                            sqliteStoreInfo.incrementalAutoVacuumEnabled);
            }
        }
        const LocalSearchIndexCore::PersistentStoreInfo storeInfo = LocalSearchIndexCore::GetPersistentStoreInfo(repositoryOptions);

        Debug::Info(
            L"SearchServiceBroker: server start pipe='{}' storage='{}' protocol={} allowRebuild={} allowShutdown={} maxRequests={} disconnectAfterBatches={}",
            effectiveOptions.pipeName,
            effectiveOptions.storageRootDirectory,
            effectiveOptions.protocolVersion,
            effectiveOptions.allowRebuildRequests,
            effectiveOptions.allowShutdownRequests,
            effectiveOptions.maxRequestsBeforeExit,
            effectiveOptions.disconnectAfterBatches);
        Debug::Info(L"SearchServiceBroker: server store backend='{}' root='{}' primary='{}' checkpoint={} compaction={}",
                    LocalSearchIndexCore::GetPersistentStoreKindText(storeInfo.kind),
                    storeInfo.rootDirectory,
                    storeInfo.primaryPath,
                    storeInfo.autoCheckpointEnabled,
                    storeInfo.autoCompactionEnabled);
        NotifyServerEvent(effectiveOptions, SEARCH_SERVICE_SERVER_EVENT_STARTING, S_OK, result.handledRequests);

        LocalSearchIndexCore::Repository repository(repositoryOptions);
        const StartupDiscoveryRoots startupDiscovery = ResolveStartupDiscoveryRoots();
        uint32_t primedRoots                         = 0u;
        for (const std::wstring& root : startupDiscovery.roots)
        {
            const HRESULT primeHr = repository.PrimeRoot(root);
            if (SUCCEEDED(primeHr))
            {
                ++primedRoots;
                continue;
            }

            Debug::Warning(L"SearchServiceBroker: failed to prime discovered root='{}'. hr=0x{:08X}", root, static_cast<unsigned long>(primeHr));
        }
        std::vector<std::wstring> startupRoots;
        repository.CollectCachedRoots(startupRoots);

        Debug::Info(L"SearchServiceBroker: discovered {} startup root(s), primed {} cache entry/caches for startup. override={}",
                    startupDiscovery.roots.size(),
                    primedRoots,
                    startupDiscovery.usedOverride);
        Debug::Info(L"SearchServiceBroker: startup discovery mode={} primedRoots={}",
                    startupDiscovery.usedOverride ? L"override" : L"fixed-local-volumes",
                    startupRoots.size());

        const bool enableStartupWarmup = effectiveOptions.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite &&
                                         (! storeLocationOverridden || startupDiscovery.usedOverride);
        ServerStartupWarmupState startupWarmupState{};
        UpdateStartupWarmupState(startupWarmupState, enableStartupWarmup, false, static_cast<uint32_t>(startupRoots.size()), 0u, 0u, {});
        std::jthread startupWarmupThread;
        if (enableStartupWarmup && ! startupRoots.empty())
        {
            const uint32_t startupWarmupDelayMs = GetStartupWarmupDelayMs();
            startupWarmupThread =
                std::jthread([&repository, &startupWarmupState, &effectiveOptions, roots = std::move(startupRoots), stopEvent, startupWarmupDelayMs](
                                 std::stop_token stopToken) noexcept
            {
                StartupWarmupCancelContext cancelContext{
                    .stopEvent = stopEvent,
                    .stopToken = std::move(stopToken),
                };

                uint32_t warmedRoots = 0u;
                uint32_t failedRoots = 0u;
                for (const std::wstring& root : roots)
                {
                    UpdateStartupWarmupState(startupWarmupState, true, true, static_cast<uint32_t>(roots.size()), warmedRoots, failedRoots, root);
                    const ServerEventDetails progressDetails{
                        .startupWarmupTotalRoots     = static_cast<uint32_t>(roots.size()),
                        .startupWarmupCompletedRoots = warmedRoots,
                        .startupWarmupFailedRoots    = failedRoots,
                        .currentPath                 = root.c_str(),
                    };
                    NotifyServerEvent(effectiveOptions, SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_PROGRESS, S_OK, 0u, &progressDetails);

                    if (startupWarmupDelayMs != 0u)
                    {
                        ::Sleep(startupWarmupDelayMs);
                    }

                    LocalSearchIndexCore::QueryStats stats{};
                    const HRESULT warmHr = repository.EnsureReadyForRoot(root, &StartupWarmupCancelCheck, &cancelContext, &stats);
                    if (warmHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
                    {
                        UpdateStartupWarmupState(startupWarmupState, true, false, static_cast<uint32_t>(roots.size()), warmedRoots, failedRoots, {});
                        Debug::Info(L"SearchServiceBroker: startup warmup cancelled after {} root(s).", warmedRoots);
                        return;
                    }

                    if (FAILED(warmHr))
                    {
                        ++failedRoots;
                        UpdateStartupWarmupState(startupWarmupState, true, true, static_cast<uint32_t>(roots.size()), warmedRoots, failedRoots, root);
                        SetStartupWarmupFailure(startupWarmupState, warmHr, root);
                        const ServerStartupWarmupSnapshot failedSnapshot = CaptureStartupWarmupSnapshot(&startupWarmupState);
                        const ServerEventDetails failedDetails{
                            .startupWarmupTotalRoots      = failedSnapshot.totalRoots,
                            .startupWarmupCompletedRoots  = failedSnapshot.completedRoots,
                            .startupWarmupFailedRoots     = failedSnapshot.failedRoots,
                            .startupWarmupLastFailureHr   = failedSnapshot.lastFailureHr,
                            .currentPath                  = root.c_str(),
                            .startupWarmupLastFailureRoot = failedSnapshot.lastFailureRoot.c_str(),
                        };
                        NotifyServerEvent(effectiveOptions, SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_FAILED, warmHr, 0u, &failedDetails);
                        Debug::Warning(L"SearchServiceBroker: startup warmup failed for root='{}'. hr=0x{:08X}", root, static_cast<unsigned long>(warmHr));
                        continue;
                    }

                    ++warmedRoots;
                    UpdateStartupWarmupState(startupWarmupState, true, true, static_cast<uint32_t>(roots.size()), warmedRoots, failedRoots, root);
                    Debug::Info(L"SearchServiceBroker: startup warmup ready root='{}' entries={} files={} dirs={} journalAvailable={} sqliteStore={}",
                                root,
                                stats.entryCount,
                                stats.fileCount,
                                stats.directoryCount,
                                stats.journalAvailable,
                                stats.usedSqliteStore);
                }

                UpdateStartupWarmupState(startupWarmupState, true, false, static_cast<uint32_t>(roots.size()), warmedRoots, failedRoots, {});
                const ServerStartupWarmupSnapshot completionSnapshot = CaptureStartupWarmupSnapshot(&startupWarmupState);
                const ServerEventDetails completionDetails{
                    .startupWarmupTotalRoots      = completionSnapshot.totalRoots,
                    .startupWarmupCompletedRoots  = completionSnapshot.completedRoots,
                    .startupWarmupFailedRoots     = completionSnapshot.failedRoots,
                    .startupWarmupLastFailureHr   = completionSnapshot.lastFailureHr,
                    .startupWarmupLastFailureRoot = completionSnapshot.lastFailureRoot.c_str(),
                };
                NotifyServerEvent(effectiveOptions, SEARCH_SERVICE_SERVER_EVENT_STARTUP_WARMUP_COMPLETED, S_OK, 0u, &completionDetails);
                Debug::Info(L"SearchServiceBroker: startup warmup completed {} root(s).", warmedRoots);
            });
        }
        else if (effectiveOptions.persistentStoreKind == LocalSearchIndexCore::PersistentStoreKind::Sqlite)
        {
            Debug::Info(L"SearchServiceBroker: startup warmup disabled. storeLocationOverridden={} overrideRoots={}",
                        storeLocationOverridden,
                        startupDiscovery.usedOverride);
        }
        ServerMaintenanceState maintenanceState{};
        RefreshAutomaticMaintenanceQueue(effectiveOptions, maintenanceState, result.handledRequests);

        SECURITY_ATTRIBUTES securityAttributes{};
        wil::unique_hlocal securityDescriptor;
        HRESULT hr = CreatePipeSecurity(securityAttributes, securityDescriptor);
        if (FAILED(hr))
        {
            NotifyServerEvent(effectiveOptions, SEARCH_SERVICE_SERVER_EVENT_ERROR, hr, result.handledRequests);
            return hr;
        }

        for (;;)
        {
            if (stopEvent != nullptr && ::WaitForSingleObject(stopEvent, 0u) == WAIT_OBJECT_0)
            {
                NotifyServerEvent(effectiveOptions, SEARCH_SERVICE_SERVER_EVENT_STOPPING, S_OK, result.handledRequests);
                break;
            }

            if (ShouldRunAutomaticMaintenanceNow(maintenanceState))
            {
                const HRESULT maintenanceHr = RunQueuedAutomaticMaintenance(effectiveOptions, maintenanceState, result.handledRequests);
                if (FAILED(maintenanceHr))
                {
                    RefreshAutomaticMaintenanceQueue(effectiveOptions, maintenanceState, result.handledRequests);
                }
                continue;
            }

            wil::unique_handle pipe;
            hr = CreateServerPipe(effectiveOptions.pipeName, securityAttributes, pipe);
            if (FAILED(hr))
            {
                NotifyServerEvent(effectiveOptions, SEARCH_SERVICE_SERVER_EVENT_ERROR, hr, result.handledRequests);
                return hr;
            }

            NotifyServerEvent(effectiveOptions, SEARCH_SERVICE_SERVER_EVENT_WAITING_FOR_CLIENT, S_OK, result.handledRequests);
            hr = WaitForClientConnection(pipe.get(), stopEvent, GetMaintenanceWaitTimeoutMs(maintenanceState));
            if (hr == S_FALSE)
            {
                continue;
            }
            if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
            {
                NotifyServerEvent(effectiveOptions, SEARCH_SERVICE_SERVER_EVENT_STOPPING, S_OK, result.handledRequests);
                break;
            }
            if (FAILED(hr))
            {
                NotifyServerEvent(effectiveOptions, SEARCH_SERVICE_SERVER_EVENT_ERROR, hr, result.handledRequests);
                return hr;
            }

            NotifyServerEvent(effectiveOptions, SEARCH_SERVICE_SERVER_EVENT_CLIENT_CONNECTED, S_OK, result.handledRequests);

            SessionContext session{};
            session.pipe                     = std::move(pipe);
            session.repository               = &repository;
            session.maintenanceState         = &maintenanceState;
            session.startupWarmupState       = &startupWarmupState;
            session.options                  = effectiveOptions;
            session.ioContext.stopEvent      = stopEvent;
            session.ioContext.frameTimeoutMs = GetServerFrameTimeoutMs();
            session.handledRequests          = result.handledRequests;

            hr = HandleClient(session);
            if (FAILED(hr) && ! IsRecoverableClientFailure(hr))
            {
                NotifyServerEvent(effectiveOptions, SEARCH_SERVICE_SERVER_EVENT_ERROR, hr, result.handledRequests);
                return hr;
            }

            ++result.handledRequests;
            NotifyServerEvent(effectiveOptions, SEARCH_SERVICE_SERVER_EVENT_REQUEST_HANDLED, hr, result.handledRequests);
            static_cast<void>(::FlushFileBuffers(session.pipe.get()));
            static_cast<void>(::DisconnectNamedPipe(session.pipe.get()));

            if (session.shutdownRequested)
            {
                NotifyServerEvent(effectiveOptions, SEARCH_SERVICE_SERVER_EVENT_STOPPING, S_OK, result.handledRequests);
                break;
            }

            RefreshAutomaticMaintenanceQueue(effectiveOptions, maintenanceState, result.handledRequests);

            if (effectiveOptions.maxRequestsBeforeExit != 0u && result.handledRequests >= effectiveOptions.maxRequestsBeforeExit)
            {
                NotifyServerEvent(effectiveOptions, SEARCH_SERVICE_SERVER_EVENT_STOPPING, S_OK, result.handledRequests);
                break;
            }
        }

        if (outResult != nullptr)
        {
            *outResult = result;
        }

        Debug::Info(L"SearchServiceBroker: server stop pipe='{}' handledRequests={}", effectiveOptions.pipeName, result.handledRequests);
        NotifyServerEvent(effectiveOptions, SEARCH_SERVICE_SERVER_EVENT_STOPPED, S_OK, result.handledRequests);
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"SearchServiceBroker: RunServer failed with an unexpected std::exception.");
        if (outResult != nullptr)
        {
            *outResult = {};
        }
        NotifyServerEvent(options, SEARCH_SERVICE_SERVER_EVENT_ERROR, E_FAIL, 0u);
        return E_FAIL;
    }
}
} // namespace SearchServiceBroker
