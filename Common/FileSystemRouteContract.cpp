#include "FileSystemRouteContract.h"


#include <algorithm>
#include <array>
#include <limits>
#include <memory>
#include <utility>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

#include <wil/com.h>

namespace FileSystemRouteContract
{
namespace
{
[[nodiscard]] bool IsStrictBool(BOOL value) noexcept
{
    return value == FALSE || value == TRUE;
}

[[nodiscard]] bool EqualsNoCase(std::wstring_view left, std::wstring_view right) noexcept
{
    if (left.size() != right.size() || left.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return false;
    }
    if (left.empty())
    {
        return true;
    }
    return CompareStringOrdinal(
               left.data(), static_cast<int>(left.size()), right.data(), static_cast<int>(right.size()), TRUE) == CSTR_EQUAL;
}

[[nodiscard]] bool IsValidOperation(FileSystemOperation operation) noexcept
{
    return operation <= FILESYSTEM_CREATE_DIRECTORY;
}

[[nodiscard]] bool IsValidAvailability(FileSystemRouteAvailability value) noexcept
{
    return value == FILESYSTEM_ROUTE_UNSUPPORTED || value == FILESYSTEM_ROUTE_AVAILABLE;
}

[[nodiscard]] bool IsValidCancellationRoute(FileSystemCancellationRoute value) noexcept
{
    return value <= FILESYSTEM_CANCELLATION_PROVIDER_WATCHDOG;
}

[[nodiscard]] bool IsValidNamespaceKind(FileSystemNamespaceKind value) noexcept
{
    return value >= FILESYSTEM_NAMESPACE_REAL_CONTAINER && value <= FILESYSTEM_NAMESPACE_FIXED_OBJECT_SET;
}

[[nodiscard]] bool IsValidComparison(FileSystemRouteComponentComparison value) noexcept
{
    return value == FILESYSTEM_ROUTE_COMPONENT_ORDINAL_IGNORE_CASE ||
           value == FILESYSTEM_ROUTE_COMPONENT_ORDINAL_CASE_SENSITIVE;
}

[[nodiscard]] bool IsValidNormalization(FileSystemRouteNormalization value) noexcept
{
    return value == FILESYSTEM_ROUTE_NORMALIZATION_NONE;
}

[[nodiscard]] bool IsValidCaseOnlyRename(FileSystemRouteCaseOnlyRename value) noexcept
{
    return value >= FILESYSTEM_ROUTE_CASE_ONLY_SUPPORTED && value <= FILESYSTEM_ROUTE_CASE_ONLY_NOT_APPLICABLE;
}

[[nodiscard]] bool IsValidProofFlags(uint32_t value) noexcept
{
    constexpr uint32_t known = FILESYSTEM_ROUTE_PROOF_HOST_READBACK | FILESYSTEM_ROUTE_PROOF_PROVIDER_BLAKE3 |
                               FILESYSTEM_ROUTE_PROOF_PROVIDER_REREAD | FILESYSTEM_ROUTE_PROOF_WRITER_DIGEST;
    return (value & ~known) == 0u;
}

[[nodiscard]] bool TryCopyArenaString(const FileSystemArena& arena, const wchar_t* value, std::wstring& output) noexcept
{
    output.clear();
    if (arena.buffer == nullptr || arena.usedBytes > arena.capacityBytes || value == nullptr)
    {
        return false;
    }

    const uintptr_t begin = reinterpret_cast<uintptr_t>(arena.buffer);
    const uintptr_t end = begin + arena.usedBytes;
    if (end < begin)
    {
        return false;
    }
    const uintptr_t address = reinterpret_cast<uintptr_t>(value);
    if (address < begin || address >= end || (address % alignof(wchar_t)) != 0u)
    {
        return false;
    }
    const size_t remainingBytes = static_cast<size_t>(end - address);
    if ((remainingBytes % sizeof(wchar_t)) != 0u)
    {
        return false;
    }
    const size_t remainingCharacters = remainingBytes / sizeof(wchar_t);
    const wchar_t* terminator = std::char_traits<wchar_t>::find(value, remainingCharacters, L'\0');
    if (terminator == nullptr)
    {
        return false;
    }
    output.assign(value, terminator);
    return true;
}

[[nodiscard]] bool TryMapPathIdentity(const FileSystemRouteFacts& facts,
                                      std::wstring acceptedSeparators,
                                      FileSystemPathIdentity& identity) noexcept
{
    if (acceptedSeparators.empty() || acceptedSeparators.find(facts.preferredSeparator) == std::wstring::npos)
    {
        return false;
    }

    identity.pathTextStableIdentity = facts.pathTextStableIdentity == TRUE;
    identity.componentComparison = facts.componentComparison == FILESYSTEM_ROUTE_COMPONENT_ORDINAL_IGNORE_CASE
        ? FileSystemPathComponentComparison::OrdinalIgnoreCase
        : FileSystemPathComponentComparison::OrdinalCaseSensitive;
    identity.preferredSeparator = facts.preferredSeparator;
    identity.acceptedSeparators = std::move(acceptedSeparators);
    identity.casePreserving = facts.casePreserving == TRUE;
    switch (facts.caseOnlyRename)
    {
        case FILESYSTEM_ROUTE_CASE_ONLY_SUPPORTED: identity.caseOnlyRename = FileSystemPathCaseOnlyRename::Supported; break;
        case FILESYSTEM_ROUTE_CASE_ONLY_NO_OP: identity.caseOnlyRename = FileSystemPathCaseOnlyRename::NoOp; break;
        case FILESYSTEM_ROUTE_CASE_ONLY_UNSUPPORTED: identity.caseOnlyRename = FileSystemPathCaseOnlyRename::Unsupported; break;
        case FILESYSTEM_ROUTE_CASE_ONLY_NOT_APPLICABLE: identity.caseOnlyRename = FileSystemPathCaseOnlyRename::NotApplicable; break;
        default: return false;
    }
    return true;
}

[[nodiscard]] bool ValidateFacts(const FileSystemRouteFacts& facts) noexcept
{
    const std::array strictBooleans{
        facts.copyOperation,
        facts.moveOperation,
        facts.nativeMoveOperation,
        facts.deleteOperation,
        facts.renameOperation,
        facts.createDirectoryOperation,
        facts.propertiesOperation,
        facts.readOperation,
        facts.writeOperation,
        facts.recycleOperation,
        facts.boundDelete,
        facts.conditionalDelete,
        facts.exclusiveStage,
        facts.conditionalPublish,
        facts.committedSize,
        facts.preserveFileLink,
        facts.preserveDirectoryLink,
        facts.retargetInTree,
        facts.exactLinkRemoval,
        facts.cancellationAbort,
        facts.cancellationDeadline,
        facts.pathTextStableIdentity,
        facts.casePreserving,
    };
    if (! std::ranges::all_of(strictBooleans, IsStrictBool) || facts.sizeBytes < sizeof(FileSystemRouteFacts) ||
        ! IsValidAvailability(facts.availability) || ! IsValidCancellationRoute(facts.cancellationRoute) ||
        ! IsValidNamespaceKind(facts.namespaceKind) || ! IsValidComparison(facts.componentComparison) ||
        ! IsValidNormalization(facts.normalization) || ! IsValidCaseOnlyRename(facts.caseOnlyRename) ||
        ! IsValidProofFlags(facts.proofFlags) || facts.copyMoveMaxConcurrency == 0u || facts.deleteMaxConcurrency == 0u ||
        facts.deleteRecycleBinMaxConcurrency == 0u || facts.maxComponentUtf16 == 0u)
    {
        return false;
    }

    if ((facts.cancellationRoute == FILESYSTEM_CANCELLATION_PROVIDER_WATCHDOG) !=
        (facts.providerWatchdogTimeoutMs != 0u))
    {
        return false;
    }
    if (facts.nativeMoveOperation == TRUE && facts.moveOperation != TRUE)
    {
        return false;
    }
    if (facts.retargetInTree == TRUE && facts.preserveFileLink != TRUE && facts.preserveDirectoryLink != TRUE)
    {
        return false;
    }
    return true;
}

[[nodiscard]] QueryResult MakeQueryFailure(QueryState state, HRESULT status) noexcept
{
    return QueryResult{.state = state, .status = status};
}

[[nodiscard]] HRESULT QueryFactsOnce(IFileSystemRouteCapabilities* route,
                                     std::wstring_view path,
                                     FileSystemOperation operation,
                                     FileSystemArena& arena,
                                     FileSystemRouteFacts& facts) noexcept
{
    const std::wstring pathCopy(path);
    arena.usedBytes = 0u;
    facts = {};
    facts.sizeBytes = sizeof(facts);
    return route->GetRouteFacts(pathCopy.c_str(), operation, &arena, &facts);
}

template<typename Invoke>
[[nodiscard]] StringResult QueryString(IFileSystemRouteCapabilities* route, Invoke&& invoke, const bool failFallbackAllocation = false) noexcept
{
    if (route == nullptr)
    {
        return StringResult{.state = QueryState::ContractViolation, .status = E_POINTER};
    }

    alignas(wchar_t) std::array<unsigned char, kNormalArenaBytes> stack{};
    FileSystemArena arena{stack.data(), static_cast<unsigned long>(stack.size()), 0u};
    const wchar_t* value = nullptr;
    unsigned long required = 0u;
    HRESULT hr = invoke(route, arena, &value, &required);
    bool fallback = false;
    std::unique_ptr<unsigned char[]> heap;
    if (hr == HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER))
    {
        if (required <= stack.size() || required > kMaximumArenaBytes)
        {
            return StringResult{.state = QueryState::ContractViolation, .status = HRESULT_FROM_WIN32(ERROR_INVALID_DATA)};
        }
        if (failFallbackAllocation)
        {
            return StringResult{.state = QueryState::ContractViolation, .status = E_OUTOFMEMORY};
        }
        heap.reset(new (std::nothrow) unsigned char[required]);
        if (! heap)
        {
            return StringResult{.state = QueryState::ContractViolation, .status = E_OUTOFMEMORY};
        }
        arena = FileSystemArena{heap.get(), required, 0u};
        value = nullptr;
        unsigned long retryRequired = 0u;
        hr = invoke(route, arena, &value, &retryRequired);
        fallback = true;
        if (retryRequired > required)
        {
            return StringResult{.state = QueryState::ContractViolation, .status = HRESULT_FROM_WIN32(ERROR_INVALID_DATA)};
        }
        required = retryRequired;
    }
    if (FAILED(hr))
    {
        const QueryState state = hr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) ? QueryState::Unsupported
                                                                               : QueryState::ContractViolation;
        return StringResult{.state = state, .status = hr, .usedArenaFallback = fallback};
    }
    std::wstring copied;
    if (required == 0u || required != arena.usedBytes || required > arena.capacityBytes || ! TryCopyArenaString(arena, value, copied))
    {
        return StringResult{
            .state = QueryState::ContractViolation, .status = HRESULT_FROM_WIN32(ERROR_INVALID_DATA), .usedArenaFallback = fallback};
    }
    return StringResult{.state = QueryState::Available, .status = S_OK, .value = std::move(copied), .usedArenaFallback = fallback};
}

[[nodiscard]] QueryResult QueryRouteCapabilities(IFileSystemRouteCapabilities* route,
                                                 std::wstring_view path,
                                                 FileSystemOperation operation,
                                                 std::wstring_view expectedProviderId,
                                                 bool failFallbackAllocation) noexcept
{
    if (route == nullptr || path.empty() || expectedProviderId.empty() || ! IsValidOperation(operation))
    {
        return MakeQueryFailure(QueryState::ContractViolation, E_INVALIDARG);
    }
    alignas(wchar_t) std::array<unsigned char, kNormalArenaBytes> stack{};
    FileSystemArena arena{stack.data(), static_cast<unsigned long>(stack.size()), 0u};
    FileSystemRouteFacts facts{};
    HRESULT hr = QueryFactsOnce(route, path, operation, arena, facts);
    bool fallback = false;
    std::unique_ptr<unsigned char[]> heap;
    if (hr == HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER))
    {
        if (facts.requiredArenaBytes <= stack.size() || facts.requiredArenaBytes > kMaximumArenaBytes)
        {
            return MakeQueryFailure(QueryState::ContractViolation, HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
        }
        if (failFallbackAllocation)
        {
            return MakeQueryFailure(QueryState::ContractViolation, E_OUTOFMEMORY);
        }
        heap.reset(new (std::nothrow) unsigned char[facts.requiredArenaBytes]);
        if (! heap)
        {
            return MakeQueryFailure(QueryState::ContractViolation, E_OUTOFMEMORY);
        }
        const unsigned long capacity = facts.requiredArenaBytes;
        arena = FileSystemArena{heap.get(), capacity, 0u};
        hr = QueryFactsOnce(route, path, operation, arena, facts);
        fallback = true;
        if (facts.requiredArenaBytes > capacity)
        {
            return MakeQueryFailure(QueryState::ContractViolation, HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
        }
    }
    if (FAILED(hr))
    {
        const QueryState state = hr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) ? QueryState::Unsupported
                                                                               : QueryState::ContractViolation;
        QueryResult failure = MakeQueryFailure(state, hr);
        failure.usedArenaFallback = fallback;
        return failure;
    }
    if (! ValidateFacts(facts) || facts.requiredArenaBytes == 0u || facts.requiredArenaBytes != arena.usedBytes ||
        arena.usedBytes > arena.capacityBytes)
    {
        QueryResult failure = MakeQueryFailure(QueryState::ContractViolation, HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
        failure.usedArenaFallback = fallback;
        return failure;
    }

    Snapshot snapshot{};
    std::wstring acceptedSeparators;
    if (! TryCopyArenaString(arena, facts.acceptedSeparators, acceptedSeparators) ||
        ! TryCopyArenaString(arena, facts.providerId, snapshot.providerId) ||
        ! TryCopyArenaString(arena, facts.pathProfileId, snapshot.pathProfileId) ||
        ! TryCopyArenaString(arena, facts.rootId, snapshot.rootId) || acceptedSeparators.empty() ||
        snapshot.providerId.empty() || snapshot.pathProfileId.empty() || snapshot.rootId.empty() ||
        ! EqualsNoCase(snapshot.providerId, expectedProviderId))
    {
        QueryResult failure = MakeQueryFailure(QueryState::ContractViolation, HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
        failure.usedArenaFallback = fallback;
        return failure;
    }

    FileSystemPathIdentity identity{};
    if (! TryMapPathIdentity(facts, std::move(acceptedSeparators), identity))
    {
        QueryResult failure = MakeQueryFailure(QueryState::ContractViolation, HRESULT_FROM_WIN32(ERROR_INVALID_DATA));
        failure.usedArenaFallback = fallback;
        return failure;
    }

    snapshot.availability = facts.availability;
    snapshot.cancellationRoute = facts.cancellationRoute;
    snapshot.namespaceKind = facts.namespaceKind;
    snapshot.normalization = facts.normalization;
    snapshot.proofFlags = facts.proofFlags;
    snapshot.providerWatchdogTimeoutMs = facts.providerWatchdogTimeoutMs;
    snapshot.copyMoveMaxConcurrency = facts.copyMoveMaxConcurrency;
    snapshot.deleteMaxConcurrency = facts.deleteMaxConcurrency;
    snapshot.deleteRecycleBinMaxConcurrency = facts.deleteRecycleBinMaxConcurrency;
    snapshot.maxComponentUtf16 = facts.maxComponentUtf16;
    snapshot.copyOperation = facts.copyOperation == TRUE;
    snapshot.moveOperation = facts.moveOperation == TRUE;
    snapshot.nativeMoveOperation = facts.nativeMoveOperation == TRUE;
    snapshot.deleteOperation = facts.deleteOperation == TRUE;
    snapshot.renameOperation = facts.renameOperation == TRUE;
    snapshot.createDirectoryOperation = facts.createDirectoryOperation == TRUE;
    snapshot.properties = facts.propertiesOperation == TRUE;
    snapshot.read = facts.readOperation == TRUE;
    snapshot.write = facts.writeOperation == TRUE;
    snapshot.recycleOperation = facts.recycleOperation == TRUE;
    snapshot.boundDelete = facts.boundDelete == TRUE;
    snapshot.conditionalDelete = facts.conditionalDelete == TRUE;
    snapshot.exclusiveStage = facts.exclusiveStage == TRUE;
    snapshot.conditionalPublish = facts.conditionalPublish == TRUE;
    snapshot.committedSize = facts.committedSize == TRUE;
    snapshot.preserveFileLink = facts.preserveFileLink == TRUE;
    snapshot.preserveDirectoryLink = facts.preserveDirectoryLink == TRUE;
    snapshot.retargetInTree = facts.retargetInTree == TRUE;
    snapshot.exactLinkRemoval = facts.exactLinkRemoval == TRUE;
    snapshot.cancellationAbort = facts.cancellationAbort == TRUE;
    snapshot.cancellationDeadline = facts.cancellationDeadline == TRUE;
    snapshot.pathIdentity = std::move(identity);

    const QueryState state = facts.availability == FILESYSTEM_ROUTE_AVAILABLE && snapshot.pathIdentity->pathTextStableIdentity
        ? QueryState::Available
        : QueryState::Unsupported;
    return QueryResult{.state = state, .status = state == QueryState::Available ? S_OK : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                       .snapshot = std::move(snapshot), .usedArenaFallback = fallback};
}

[[nodiscard]] ChildNameContractResult QueryChildNameContractInternal(IFileSystemRouteCapabilities* route,
                                                                     const std::wstring_view parentPath,
                                                                     const std::wstring_view childName,
                                                                     const FileSystemOperation operation,
                                                                     const std::wstring_view expectedProviderId,
                                                                     const bool failFallbackAllocation) noexcept
{
    if (route == nullptr || parentPath.empty() || childName.empty() || expectedProviderId.empty() || ! IsValidOperation(operation))
    {
        return ChildNameContractResult{.state = QueryState::ContractViolation, .status = E_INVALIDARG};
    }

    const QueryResult routeResult = Query(route, parentPath, operation, expectedProviderId);
    ChildNameContractResult result{
        .state = routeResult.state,
        .status = routeResult.status,
        .arenaFallbackCount = routeResult.usedArenaFallback ? 1u : 0u,
    };
    if (routeResult.state != QueryState::Available || ! routeResult.snapshot.pathIdentity.has_value())
    {
        return result;
    }

    const ChildNameResult validation = ValidateChildName(route, parentPath, childName, operation);
    result.state = validation.state;
    result.status = validation.status;
    result.nameStatus = validation.nameStatus;
    result.failureStatus = validation.failureStatus;
    if (validation.state != QueryState::Available || validation.nameStatus != FILESYSTEM_CHILD_NAME_VALID)
    {
        return result;
    }

    const std::wstring parentCopy(parentPath);
    const std::wstring childCopy(childName);
    const StringResult joined = QueryString(
        route,
        [&](IFileSystemRouteCapabilities* capabilities,
            FileSystemArena& arena,
            const wchar_t** value,
            unsigned long* required) noexcept
        { return capabilities->JoinPath(parentCopy.c_str(), childCopy.c_str(), operation, &arena, value, required); },
        failFallbackAllocation);
    result.arenaFallbackCount += joined.usedArenaFallback ? 1u : 0u;
    if (joined.state != QueryState::Available)
    {
        result.state = joined.state;
        result.status = joined.status;
        return result;
    }

    const StringResult collision = QueryString(
        route,
        [&](IFileSystemRouteCapabilities* capabilities,
            FileSystemArena& arena,
            const wchar_t** value,
            unsigned long* required) noexcept
        { return capabilities->GetChildNameCollisionKey(parentCopy.c_str(), childCopy.c_str(), operation, &arena, value, required); },
        failFallbackAllocation);
    result.arenaFallbackCount += collision.usedArenaFallback ? 1u : 0u;
    if (collision.state != QueryState::Available)
    {
        result.state = collision.state;
        result.status = collision.status;
        return result;
    }

    const FileSystemPathIdentity& identity = routeResult.snapshot.pathIdentity.value();
    std::wstring joinedParent;
    std::wstring joinedLeaf;
    if (joined.value.empty() || collision.value.empty() ||
        ! TryGetFileSystemParentPath(identity, joined.value, joinedParent) ||
        ! TryGetFileSystemLeafName(identity, joined.value, joinedLeaf) ||
        ! EquivalentPath(identity, joinedParent, parentPath) || ! EquivalentComponent(identity, joinedLeaf, childName))
    {
        result.state = QueryState::ContractViolation;
        result.status = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        return result;
    }

    result.state = QueryState::Available;
    result.status = S_OK;
    result.joinedPath = joined.value;
    result.collisionKey = collision.value;
    return result;
}
}

QueryResult Query(IFileSystem* fileSystem,
                  std::wstring_view path,
                  FileSystemOperation operation,
                  std::wstring_view expectedProviderId) noexcept
{
    if (fileSystem == nullptr || path.empty() || expectedProviderId.empty() || ! IsValidOperation(operation))
    {
        return MakeQueryFailure(QueryState::ContractViolation, E_INVALIDARG);
    }

    wil::com_ptr<IFileSystemRouteCapabilities> route;
    const HRESULT queryHr = fileSystem->QueryInterface(__uuidof(IFileSystemRouteCapabilities), route.put_void());
    if (FAILED(queryHr) || ! route)
    {
        return MakeQueryFailure(QueryState::ContractViolation, FAILED(queryHr) ? queryHr : E_NOINTERFACE);
    }
    return Query(route.get(), path, operation, expectedProviderId);
}

QueryResult Query(IFileSystemRouteCapabilities* route,
                  std::wstring_view path,
                  FileSystemOperation operation,
                  std::wstring_view expectedProviderId) noexcept
{
    return QueryRouteCapabilities(route, path, operation, expectedProviderId, false);
}

#ifdef ENABLE_TESTS
QueryResult QueryWithAllocationFailureForSelfTest(IFileSystemRouteCapabilities* route,
                                                  std::wstring_view path,
                                                  FileSystemOperation operation,
                                                  std::wstring_view expectedProviderId) noexcept
{
    return QueryRouteCapabilities(route, path, operation, expectedProviderId, true);
}
#endif

BooleanResult QueryTransferPeerAllowed(IFileSystem* fileSystem,
                                       std::wstring_view path,
                                       FileSystemOperation operation,
                                       FileSystemTransferPeerRole role,
                                       std::wstring_view peerPluginId) noexcept
{
    if (fileSystem == nullptr || path.empty() || peerPluginId.empty() ||
        (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE) ||
        (role != FILESYSTEM_TRANSFER_PEER_EXPORT && role != FILESYSTEM_TRANSFER_PEER_IMPORT))
    {
        return BooleanResult{.state = QueryState::ContractViolation, .status = E_INVALIDARG};
    }
    wil::com_ptr<IFileSystemRouteCapabilities> route;
    const HRESULT hr = fileSystem->QueryInterface(__uuidof(IFileSystemRouteCapabilities), route.put_void());
    if (FAILED(hr) || ! route)
    {
        return BooleanResult{.state = QueryState::ContractViolation, .status = FAILED(hr) ? hr : E_NOINTERFACE};
    }
    return QueryTransferPeerAllowed(route.get(), path, operation, role, peerPluginId);
}

BooleanResult QueryTransferPeerAllowed(IFileSystemRouteCapabilities* route,
                                       std::wstring_view path,
                                       FileSystemOperation operation,
                                       FileSystemTransferPeerRole role,
                                       std::wstring_view peerPluginId) noexcept
{
    if (route == nullptr || path.empty() || peerPluginId.empty() ||
        (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE) ||
        (role != FILESYSTEM_TRANSFER_PEER_EXPORT && role != FILESYSTEM_TRANSFER_PEER_IMPORT))
    {
        return BooleanResult{.state = QueryState::ContractViolation, .status = E_INVALIDARG};
    }
    const std::wstring pathCopy(path);
    const std::wstring peerCopy(peerPluginId);
    BOOL allowed = FALSE;
    const HRESULT hr = route->IsTransferPeerAllowed(pathCopy.c_str(), operation, role, peerCopy.c_str(), &allowed);
    if (FAILED(hr))
    {
        return BooleanResult{.state = hr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) ? QueryState::Unsupported
                                                                                    : QueryState::ContractViolation,
                             .status = hr};
    }
    if (! IsStrictBool(allowed))
    {
        return BooleanResult{.state = QueryState::ContractViolation, .status = HRESULT_FROM_WIN32(ERROR_INVALID_DATA)};
    }
    return BooleanResult{.state = QueryState::Available, .status = S_OK, .value = allowed == TRUE};
}

ChildNameResult ValidateChildName(IFileSystem* fileSystem,
                                  std::wstring_view parentPath,
                                  std::wstring_view childName,
                                  FileSystemOperation operation) noexcept
{
    if (fileSystem == nullptr || parentPath.empty() || childName.empty() || ! IsValidOperation(operation))
    {
        return ChildNameResult{.state = QueryState::ContractViolation, .status = E_INVALIDARG};
    }
    wil::com_ptr<IFileSystemRouteCapabilities> route;
    const HRESULT hr = fileSystem->QueryInterface(__uuidof(IFileSystemRouteCapabilities), route.put_void());
    if (FAILED(hr) || ! route)
    {
        return ChildNameResult{.state = QueryState::ContractViolation, .status = FAILED(hr) ? hr : E_NOINTERFACE};
    }
    return ValidateChildName(route.get(), parentPath, childName, operation);
}

ChildNameResult ValidateChildName(IFileSystemRouteCapabilities* route,
                                  std::wstring_view parentPath,
                                  std::wstring_view childName,
                                  FileSystemOperation operation) noexcept
{
    if (route == nullptr || parentPath.empty() || childName.empty() || ! IsValidOperation(operation))
    {
        return ChildNameResult{.state = QueryState::ContractViolation, .status = E_INVALIDARG};
    }
    const std::wstring parentCopy(parentPath);
    const std::wstring childCopy(childName);
    FileSystemChildNameValidation validation{.sizeBytes = sizeof(validation)};
    const HRESULT hr = route->ValidateChildName(parentCopy.c_str(), childCopy.c_str(), operation, &validation);
    if (FAILED(hr))
    {
        return ChildNameResult{.state = hr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) ? QueryState::Unsupported
                                                                                      : QueryState::ContractViolation,
                               .status = hr};
    }
    const bool statusValid = validation.sizeBytes >= sizeof(validation) && validation.status <= FILESYSTEM_CHILD_NAME_INVALID;
    const bool resultValid = (validation.status == FILESYSTEM_CHILD_NAME_VALID && validation.failureStatus == S_OK) ||
        (validation.status == FILESYSTEM_CHILD_NAME_INVALID && FAILED(validation.failureStatus)) ||
        (validation.status == FILESYSTEM_CHILD_NAME_UNSUPPORTED && validation.failureStatus == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
    if (! statusValid || ! resultValid)
    {
        return ChildNameResult{.state = QueryState::ContractViolation, .status = HRESULT_FROM_WIN32(ERROR_INVALID_DATA)};
    }
    const QueryState state = validation.status == FILESYSTEM_CHILD_NAME_UNSUPPORTED ? QueryState::Unsupported : QueryState::Available;
    return ChildNameResult{.state = state,
                           .status = S_OK,
                           .nameStatus = validation.status,
                           .failureStatus = validation.failureStatus};
}

ChildNameContractResult QueryChildNameContract(IFileSystem* fileSystem,
                                               std::wstring_view parentPath,
                                               std::wstring_view childName,
                                               FileSystemOperation operation,
                                               std::wstring_view expectedProviderId) noexcept
{
    if (fileSystem == nullptr)
    {
        return ChildNameContractResult{.state = QueryState::ContractViolation, .status = E_POINTER};
    }
    wil::com_ptr<IFileSystemRouteCapabilities> route;
    const HRESULT hr = fileSystem->QueryInterface(__uuidof(IFileSystemRouteCapabilities), route.put_void());
    if (FAILED(hr) || ! route)
    {
        return ChildNameContractResult{.state = QueryState::ContractViolation, .status = FAILED(hr) ? hr : E_NOINTERFACE};
    }
    return QueryChildNameContract(route.get(), parentPath, childName, operation, expectedProviderId);
}

ChildNameContractResult QueryChildNameContract(IFileSystemRouteCapabilities* route,
                                               std::wstring_view parentPath,
                                               std::wstring_view childName,
                                               FileSystemOperation operation,
                                               std::wstring_view expectedProviderId) noexcept
{
    return QueryChildNameContractInternal(route, parentPath, childName, operation, expectedProviderId, false);
}

#ifdef ENABLE_TESTS
ChildNameContractResult QueryChildNameContractWithAllocationFailureForSelfTest(
    IFileSystemRouteCapabilities* route,
    std::wstring_view parentPath,
    std::wstring_view childName,
    FileSystemOperation operation,
    std::wstring_view expectedProviderId) noexcept
{
    return QueryChildNameContractInternal(route, parentPath, childName, operation, expectedProviderId, true);
}
#endif

StringResult QueryChildNameCollisionKey(IFileSystem* fileSystem,
                                        std::wstring_view parentPath,
                                        std::wstring_view childName,
                                        FileSystemOperation operation) noexcept
{
    if (fileSystem == nullptr)
    {
        return StringResult{.state = QueryState::ContractViolation, .status = E_POINTER};
    }
    wil::com_ptr<IFileSystemRouteCapabilities> route;
    const HRESULT queryHr = fileSystem->QueryInterface(__uuidof(IFileSystemRouteCapabilities), route.put_void());
    if (FAILED(queryHr) || ! route)
    {
        return StringResult{.state = QueryState::ContractViolation, .status = FAILED(queryHr) ? queryHr : E_NOINTERFACE};
    }
    return QueryChildNameCollisionKey(route.get(), parentPath, childName, operation);
}

StringResult QueryChildNameCollisionKey(IFileSystemRouteCapabilities* route,
                                        std::wstring_view parentPath,
                                        std::wstring_view childName,
                                        FileSystemOperation operation) noexcept
{
    const std::wstring parentCopy(parentPath);
    const std::wstring childCopy(childName);
    if (parentCopy.empty() || childCopy.empty() || ! IsValidOperation(operation))
    {
        return StringResult{.state = QueryState::ContractViolation, .status = E_INVALIDARG};
    }
    return QueryString(route, [&](IFileSystemRouteCapabilities* capabilities,
                                  FileSystemArena& arena,
                                       const wchar_t** value,
                                       unsigned long* required) noexcept
    { return capabilities->GetChildNameCollisionKey(parentCopy.c_str(), childCopy.c_str(), operation, &arena, value, required); });
}

StringResult QueryJoinedPath(IFileSystem* fileSystem,
                             std::wstring_view parentPath,
                             std::wstring_view childName,
                             FileSystemOperation operation) noexcept
{
    if (fileSystem == nullptr)
    {
        return StringResult{.state = QueryState::ContractViolation, .status = E_POINTER};
    }
    wil::com_ptr<IFileSystemRouteCapabilities> route;
    const HRESULT queryHr = fileSystem->QueryInterface(__uuidof(IFileSystemRouteCapabilities), route.put_void());
    if (FAILED(queryHr) || ! route)
    {
        return StringResult{.state = QueryState::ContractViolation, .status = FAILED(queryHr) ? queryHr : E_NOINTERFACE};
    }
    return QueryJoinedPath(route.get(), parentPath, childName, operation);
}

StringResult QueryJoinedPath(IFileSystemRouteCapabilities* route,
                             std::wstring_view parentPath,
                             std::wstring_view childName,
                             FileSystemOperation operation) noexcept
{
    const std::wstring parentCopy(parentPath);
    const std::wstring childCopy(childName);
    if (parentCopy.empty() || childCopy.empty() || ! IsValidOperation(operation))
    {
        return StringResult{.state = QueryState::ContractViolation, .status = E_INVALIDARG};
    }
    return QueryString(route, [&](IFileSystemRouteCapabilities* capabilities,
                                  FileSystemArena& arena,
                                       const wchar_t** value,
                                       unsigned long* required) noexcept
    { return capabilities->JoinPath(parentCopy.c_str(), childCopy.c_str(), operation, &arena, value, required); });
}

}
