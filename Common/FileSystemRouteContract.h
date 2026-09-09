#pragma once

#include "FileSystemPathIdentity.h"
#include "PlugInterfaces/FileSystem.h"

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>

namespace FileSystemRouteContract
{
inline constexpr unsigned long kNormalArenaBytes = 4u * 1024u;
inline constexpr unsigned long kMaximumArenaBytes = 64u * 1024u;

enum class QueryState : uint8_t
{
    Available,
    Unsupported,
    ContractViolation,
};

struct Snapshot final
{
    std::wstring providerId;
    std::wstring pathProfileId;
    std::wstring rootId;
    FileSystemRouteAvailability availability = FILESYSTEM_ROUTE_UNSUPPORTED;
    FileSystemCancellationRoute cancellationRoute = FILESYSTEM_CANCELLATION_UNCONTAINED;
    FileSystemNamespaceKind namespaceKind = FILESYSTEM_NAMESPACE_PROVIDER_VIRTUAL_FOLDER;
    FileSystemRouteNormalization normalization = FILESYSTEM_ROUTE_NORMALIZATION_NONE;
    uint32_t proofFlags = FILESYSTEM_ROUTE_PROOF_NONE;
    uint32_t providerWatchdogTimeoutMs = 0u;
    uint32_t copyMoveMaxConcurrency = 0u;
    uint32_t deleteMaxConcurrency = 0u;
    uint32_t deleteRecycleBinMaxConcurrency = 0u;
    uint64_t maxComponentUtf16 = 0u;

    bool copyOperation = false;
    bool moveOperation = false;
    bool nativeMoveOperation = false;
    bool deleteOperation = false;
    bool renameOperation = false;
    bool createDirectoryOperation = false;
    bool properties = false;
    bool read = false;
    bool write = false;
    bool recycleOperation = false;
    bool boundDelete = false;
    bool conditionalDelete = false;
    bool exclusiveStage = false;
    bool conditionalPublish = false;
    bool committedSize = false;
    bool preserveFileLink = false;
    bool preserveDirectoryLink = false;
    bool retargetInTree = false;
    bool exactLinkRemoval = false;
    bool cancellationAbort = false;
    bool cancellationDeadline = false;
    std::optional<FileSystemPathIdentity> pathIdentity;
};

struct QueryResult final
{
    QueryState state = QueryState::ContractViolation;
    HRESULT status = E_UNEXPECTED;
    Snapshot snapshot;
    bool usedArenaFallback = false;
};

// Queries and immediately copies every provider-owned fact. Missing typed support and
// malformed output fail closed; capability JSON is never consulted.
[[nodiscard]] QueryResult Query(IFileSystem* fileSystem,
                                std::wstring_view path,
                                FileSystemOperation operation,
                                std::wstring_view expectedProviderId) noexcept;
[[nodiscard]] QueryResult Query(IFileSystemRouteCapabilities* routeCapabilities,
                                std::wstring_view path,
                                FileSystemOperation operation,
                                std::wstring_view expectedProviderId) noexcept;
#ifdef ENABLE_TESTS
// Deterministic fallback-allocation fault used by the contract suite. The
// provider must first request more than kNormalArenaBytes for this path to run.
[[nodiscard]] QueryResult QueryWithAllocationFailureForSelfTest(IFileSystemRouteCapabilities* routeCapabilities,
                                                                std::wstring_view path,
                                                                FileSystemOperation operation,
                                                                std::wstring_view expectedProviderId) noexcept;
#endif

struct BooleanResult final
{
    QueryState state = QueryState::ContractViolation;
    HRESULT status = E_UNEXPECTED;
    bool value = false;
};

[[nodiscard]] BooleanResult QueryTransferPeerAllowed(IFileSystem* fileSystem,
                                                     std::wstring_view path,
                                                     FileSystemOperation operation,
                                                     FileSystemTransferPeerRole role,
                                                     std::wstring_view peerPluginId) noexcept;
[[nodiscard]] BooleanResult QueryTransferPeerAllowed(IFileSystemRouteCapabilities* routeCapabilities,
                                                     std::wstring_view path,
                                                     FileSystemOperation operation,
                                                     FileSystemTransferPeerRole role,
                                                     std::wstring_view peerPluginId) noexcept;

struct ChildNameResult final
{
    QueryState state = QueryState::ContractViolation;
    HRESULT status = E_UNEXPECTED;
    FileSystemChildNameStatus nameStatus = FILESYSTEM_CHILD_NAME_UNSUPPORTED;
    HRESULT failureStatus = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
};

[[nodiscard]] ChildNameResult ValidateChildName(IFileSystem* fileSystem,
                                                std::wstring_view parentPath,
                                                std::wstring_view childName,
                                                FileSystemOperation operation) noexcept;
[[nodiscard]] ChildNameResult ValidateChildName(IFileSystemRouteCapabilities* routeCapabilities,
                                                std::wstring_view parentPath,
                                                std::wstring_view childName,
                                                FileSystemOperation operation) noexcept;

struct ChildNameContractResult final
{
    QueryState state = QueryState::ContractViolation;
    HRESULT status = E_UNEXPECTED;
    FileSystemChildNameStatus nameStatus = FILESYSTEM_CHILD_NAME_UNSUPPORTED;
    HRESULT failureStatus = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    std::wstring joinedPath;
    std::wstring collisionKey;
    uint32_t arenaFallbackCount = 0u;
};

// Queries one retained typed route generation and copies the provider's complete
// child-name decision. Invalid names preserve the provider HRESULT and return no
// joined path/key. Missing, unsupported, or malformed typed output fails closed.
[[nodiscard]] ChildNameContractResult QueryChildNameContract(IFileSystem* fileSystem,
                                                             std::wstring_view parentPath,
                                                             std::wstring_view childName,
                                                             FileSystemOperation operation,
                                                             std::wstring_view expectedProviderId) noexcept;
[[nodiscard]] ChildNameContractResult QueryChildNameContract(IFileSystemRouteCapabilities* routeCapabilities,
                                                             std::wstring_view parentPath,
                                                             std::wstring_view childName,
                                                             FileSystemOperation operation,
                                                             std::wstring_view expectedProviderId) noexcept;
#ifdef ENABLE_TESTS
[[nodiscard]] ChildNameContractResult QueryChildNameContractWithAllocationFailureForSelfTest(
    IFileSystemRouteCapabilities* routeCapabilities,
    std::wstring_view parentPath,
    std::wstring_view childName,
    FileSystemOperation operation,
    std::wstring_view expectedProviderId) noexcept;
#endif

struct StringResult final
{
    QueryState state = QueryState::ContractViolation;
    HRESULT status = E_UNEXPECTED;
    std::wstring value;
    bool usedArenaFallback = false;
};

[[nodiscard]] StringResult QueryChildNameCollisionKey(IFileSystem* fileSystem,
                                                      std::wstring_view parentPath,
                                                      std::wstring_view childName,
                                                      FileSystemOperation operation) noexcept;
[[nodiscard]] StringResult QueryChildNameCollisionKey(IFileSystemRouteCapabilities* routeCapabilities,
                                                      std::wstring_view parentPath,
                                                      std::wstring_view childName,
                                                      FileSystemOperation operation) noexcept;
[[nodiscard]] StringResult QueryJoinedPath(IFileSystem* fileSystem,
                                           std::wstring_view parentPath,
                                           std::wstring_view childName,
                                           FileSystemOperation operation) noexcept;
[[nodiscard]] StringResult QueryJoinedPath(IFileSystemRouteCapabilities* routeCapabilities,
                                           std::wstring_view parentPath,
                                           std::wstring_view childName,
                                           FileSystemOperation operation) noexcept;

enum class MutationClassification : uint8_t
{
    Unsupported,
    RetryableNoCommit,
    FailedKnown,
    Indeterminate,
    ContractViolation,
};

[[nodiscard]] inline bool IsValidItemMutationResultPrefix(const FileSystemItemMutationResult& result) noexcept
{
    return FileSystemItemMutationResultHasSupportedSize(result) && (result.outcomeKnown == FALSE || result.outcomeKnown == TRUE) &&
           (result.mutationCommitted == FALSE || result.mutationCommitted == TRUE) &&
           (result.originalStillPresent == FALSE || result.originalStillPresent == TRUE) &&
           (! FileSystemItemMutationResultHasOwnedStageDisposition(result) || FileSystemOwnedStageDispositionIsValid(result.ownedStageDisposition));
}

// Copy only advertised, understood fields while the callback owns the provider buffer.
// A V1 provider may supply exactly 16 bytes; struct assignment would read its absent
// stage field. Future suffixes are ignored and must not be advertised by our snapshot.
[[nodiscard]] inline std::optional<FileSystemItemMutationResult> SnapshotItemMutationResult(const FileSystemItemMutationResult* result) noexcept
{
    if (result == nullptr || ! IsValidItemMutationResultPrefix(*result))
    {
        return std::nullopt;
    }
    const bool hasStage = FileSystemItemMutationResultHasOwnedStageDisposition(*result);
    return FileSystemItemMutationResult{
        hasStage ? static_cast<uint32_t>(sizeof(FileSystemItemMutationResult)) : FILESYSTEM_ITEM_MUTATION_RESULT_V1_SIZE,
        result->outcomeKnown,
        result->mutationCommitted,
        result->originalStillPresent,
        hasStage ? result->ownedStageDisposition : FileSystemOwnedStageDisposition::NotApplicable,
    };
}

// Classifies a failed destructive attempt. A successful committed receipt stays on
// the normal completion path and need not call this helper. Header-owned so provider
// DLLs share this policy without depending on the host's route-query implementation.
[[nodiscard]] inline MutationClassification ClassifyFailedMutation(bool routeSupported, HRESULT status, const FileSystemItemMutationResult* receipt) noexcept
{
    if (! routeSupported)
    {
        return MutationClassification::Unsupported;
    }
    if (receipt == nullptr)
    {
        return MutationClassification::Indeterminate;
    }
    if (! IsValidItemMutationResultPrefix(*receipt))
    {
        return MutationClassification::ContractViolation;
    }
    if (receipt->outcomeKnown != TRUE)
    {
        return MutationClassification::Indeterminate;
    }
    if (receipt->mutationCommitted == TRUE)
    {
        return MutationClassification::FailedKnown;
    }
    if (receipt->originalStillPresent != TRUE)
    {
        return MutationClassification::ContractViolation;
    }

    const FileSystemOwnedStageDisposition stage =
        FileSystemItemMutationResultHasOwnedStageDisposition(*receipt) ? receipt->ownedStageDisposition : FileSystemOwnedStageDisposition::NotApplicable;
    if (stage == FileSystemOwnedStageDisposition::Published)
    {
        return MutationClassification::ContractViolation;
    }
    if (stage == FileSystemOwnedStageDisposition::Unknown || stage == FileSystemOwnedStageDisposition::RetainedIncomplete)
    {
        return MutationClassification::Indeterminate;
    }
    if (stage == FileSystemOwnedStageDisposition::Retained || SUCCEEDED(status))
    {
        return MutationClassification::FailedKnown;
    }
    return MutationClassification::RetryableNoCommit;
}
} // namespace FileSystemRouteContract
