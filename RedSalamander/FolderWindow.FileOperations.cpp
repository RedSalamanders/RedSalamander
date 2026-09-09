#include "FolderWindow.FileOperationsInternal.h"
#include "FolderWindow.FileOperations.State.Private.h"
#ifdef ENABLE_TESTS
#include "FolderWindow.FileOperations.SelfTest.h"
#endif
#include "FileSystemPathIdentity.h"
#include "FileSystemRouteContract.h"
#include "FileOperationArtifactRegistry.h"
#include "HostServices.h"
#include "NavigationLocation.h"
#include "PathUtils.h"

#include <limits>
#include <numeric>
#include <unordered_map>
#include <unordered_set>

namespace
{
using FileOperations::FileOperationPlan;
using FileOperations::PlanRejectionBucket;

constexpr ULONGLONG kVerificationCapabilityUiBudgetMs = 300ull;
#ifdef ENABLE_TESTS
std::atomic<bool> g_forceVerificationCapabilityBudgetExceeded{false};
#endif

[[nodiscard]] bool IsQualifiedEndpointValid(const FileOperations::QualifiedEndpoint& endpoint) noexcept
{
    return ! endpoint.pluginId.empty() && ! endpoint.instanceId.empty() && ! endpoint.profileId.empty() && ! endpoint.rootId.empty() &&
           endpoint.pathIdentity.has_value() && endpoint.pathIdentity->pathTextStableIdentity &&
           endpoint.cancellationRouteClass != FileOperations::CancellationRouteClass::Uncontained &&
           ((endpoint.cancellationRouteClass == FileOperations::CancellationRouteClass::ProviderWatchdog &&
             endpoint.providerWatchdogTimeoutMs != 0u) ||
            (endpoint.cancellationRouteClass == FileOperations::CancellationRouteClass::Bounded &&
             endpoint.providerWatchdogTimeoutMs == 0u));
}

[[nodiscard]] bool IsLocalFileSystemEndpointImpl(const FileOperations::QualifiedEndpoint& endpoint) noexcept
{
    return NavigationLocation::EqualsNoCase(endpoint.pluginId, L"builtin/file-system");
}

[[nodiscard]] bool QualifiedEndpointsMatchImpl(const FileOperations::QualifiedEndpoint& left,
                                                const FileOperations::QualifiedEndpoint& right) noexcept
{
    return NavigationLocation::EqualsNoCase(left.pluginId, right.pluginId) && left.instanceId == right.instanceId &&
           left.profileId == right.profileId && left.rootId == right.rootId;
}

[[nodiscard]] bool QualifiedEndpointsShareObjectIdentityDomainImpl(const FileOperations::QualifiedEndpoint& left,
                                                                    const FileOperations::QualifiedEndpoint& right) noexcept
{
    // rootId qualifies strategy topology, not the provider's object-ID namespace. Aliases such as
    // a drive path and its UNC administrative share can have different roots while BindObject
    // still returns comparable IDs from one provider instance/path profile.
    return NavigationLocation::EqualsNoCase(left.pluginId, right.pluginId) && left.instanceId == right.instanceId &&
           left.profileId == right.profileId;
}

// C1: whether a provider binds objects at all, from its interface alone (no provider I/O). A
// provider without the interface runs its contract-tested native delete; every other permanent
// delete root is pinned by Preparing on the task thread.
enum class ObjectBindingSupport : uint8_t
{
    Supported,
    Unsupported,
    Failed,
};

[[nodiscard]] ObjectBindingSupport QueryObjectBindingSupport(IFileSystem* fileSystem, HRESULT& status) noexcept
{
    status = E_POINTER;
    if (fileSystem == nullptr)
    {
        return ObjectBindingSupport::Failed;
    }
    wil::com_ptr<IFileSystemObjectBinding> binding;
    status = fileSystem->QueryInterface(__uuidof(IFileSystemObjectBinding), binding.put_void());
    if (status == E_NOINTERFACE && ! binding)
    {
        return ObjectBindingSupport::Unsupported;
    }
    if (FAILED(status) || ! binding)
    {
        return ObjectBindingSupport::Failed;
    }
    return ObjectBindingSupport::Supported;
}

[[nodiscard]] bool TryGetFileSystemPluginIdentity(IFileSystem* fileSystem,
                                                   std::wstring& pluginId,
                                                   std::wstring& pluginShortId) noexcept
{
    pluginId.clear();
    pluginShortId.clear();
    if (fileSystem == nullptr)
    {
        return false;
    }

    wil::com_ptr<IInformations> informations;
    HRESULT hr = fileSystem->QueryInterface(IID_PPV_ARGS(informations.addressof()));
    if (FAILED(hr) || ! informations)
    {
        return false;
    }

    const PluginMetaData* metadata = nullptr;
    hr = informations->GetMetaData(&metadata);
    if (FAILED(hr) || metadata == nullptr || metadata->id == nullptr || metadata->id[0] == L'\0')
    {
        return false;
    }

    pluginId.assign(metadata->id);
    if (metadata->shortId != nullptr)
    {
        pluginShortId.assign(metadata->shortId);
    }
    return true;
}

[[nodiscard]] HRESULT RejectPlan(PlanRejectionBucket bucket, PlanRejectionBucket* rejectionBucket) noexcept
{
    if (rejectionBucket)
    {
        *rejectionBucket = bucket;
    }
    return E_INVALIDARG;
}

[[nodiscard]] HRESULT ValidatePlanImpl(const FileOperationPlan& plan,
                                       PlanRejectionBucket* rejectionBucket,
                                       bool allowPendingPermanentConsent) noexcept
{
    if (rejectionBucket)
    {
        *rejectionBucket = PlanRejectionBucket::None;
    }

    return std::visit(
        [&](const auto& typedPlan) noexcept -> HRESULT
        {
            using Plan = std::remove_cvref_t<decltype(typedPlan)>;
            if constexpr (std::is_same_v<Plan, FileOperations::TransferPlan>)
            {
                if (! IsQualifiedEndpointValid(typedPlan.sourceEndpoint) || ! IsQualifiedEndpointValid(typedPlan.destinationEndpoint))
                {
                    return RejectPlan(PlanRejectionBucket::MalformedEndpoint, rejectionBucket);
                }
                if (typedPlan.selectedItems.empty())
                {
                    return RejectPlan(PlanRejectionBucket::EmptySelection, rejectionBucket);
                }
                if ((typedPlan.intent == FileOperations::TransferIntent::Copy &&
                     typedPlan.strategy != FileOperations::OperationStrategy::Copy) ||
                    (typedPlan.strategy == FileOperations::OperationStrategy::Native &&
                     ! QualifiedEndpointsMatchImpl(typedPlan.sourceEndpoint, typedPlan.destinationEndpoint)))
                {
                    return RejectPlan(PlanRejectionBucket::InvalidStrategy, rejectionBucket);
                }
                if (typedPlan.intent == FileOperations::TransferIntent::Move && typedPlan.strategy == FileOperations::OperationStrategy::Copy)
                {
                    return RejectPlan(PlanRejectionBucket::InvalidStrategy, rejectionBucket);
                }
                if (typedPlan.destination.providerFolderPath.empty())
                {
                    return RejectPlan(PlanRejectionBucket::MalformedDestination, rejectionBucket);
                }
                const bool localSource      = IsLocalFileSystemEndpointImpl(typedPlan.sourceEndpoint);
                const bool localDestination = IsLocalFileSystemEndpointImpl(typedPlan.destinationEndpoint);
                if ((localDestination && ! Common::Paths::IsSupportedLocalFileOperationPath(typedPlan.destination.providerFolderPath)) ||
                    (localSource && std::ranges::any_of(typedPlan.selectedItems, [](const FileOperations::QualifiedSourceItem& item) noexcept
                                                       { return ! Common::Paths::IsSupportedLocalFileOperationPath(item.providerPath); })))
                {
                    return RejectPlan(PlanRejectionBucket::UnsupportedDeviceNamespace, rejectionBucket);
                }
                if (std::ranges::any_of(typedPlan.selectedItems, [](const FileOperations::QualifiedSourceItem& item) noexcept
                                        { return item.providerPath.empty(); }))
                {
                    return RejectPlan(PlanRejectionBucket::MalformedSource, rejectionBucket);
                }
                if (typedPlan.moveClipboardSequence.has_value() &&
                    (typedPlan.intent != FileOperations::TransferIntent::Move || typedPlan.moveClipboardSequence->windowsSequenceNumber == 0u))
                {
                    return RejectPlan(PlanRejectionBucket::InvalidClipboardSequence, rejectionBucket);
                }
                if (! typedPlan.explicitMappings.empty())
                {
                    if (typedPlan.explicitMappings.size() != typedPlan.selectedItems.size())
                    {
                        return RejectPlan(PlanRejectionBucket::InvalidExplicitMappings, rejectionBucket);
                    }
                    std::vector<bool> mapped(typedPlan.selectedItems.size(), false);
                    for (const FileOperations::TransferDestinationMapping& mapping : typedPlan.explicitMappings)
                    {
                        if (mapping.sourceIndex >= mapped.size() || mapped[mapping.sourceIndex] || mapping.destinationProviderPath.empty())
                        {
                            return RejectPlan(PlanRejectionBucket::InvalidExplicitMappings, rejectionBucket);
                        }
                        const FileSystemPathIdentity& destinationIdentity = typedPlan.destinationEndpoint.pathIdentity.value();
                        const bool mappingInDestination = EquivalentPath(destinationIdentity,
                                                                         typedPlan.destination.providerFolderPath,
                                                                         mapping.destinationProviderPath) ||
                            IsStrictDescendantPath(destinationIdentity,
                                                   typedPlan.destination.providerFolderPath,
                                                   mapping.destinationProviderPath);
                        if ((localDestination && ! Common::Paths::IsSupportedLocalFileOperationPath(mapping.destinationProviderPath)) || ! mappingInDestination)
                        {
                            return RejectPlan(localDestination && ! Common::Paths::IsSupportedLocalFileOperationPath(mapping.destinationProviderPath)
                                                  ? PlanRejectionBucket::UnsupportedDeviceNamespace
                                                  : PlanRejectionBucket::EscapingDestinationMapping,
                                              rejectionBucket);
                        }
                        mapped[mapping.sourceIndex] = true;
                    }
                }
                if (QualifiedEndpointsMatchImpl(typedPlan.sourceEndpoint, typedPlan.destinationEndpoint))
                {
                    const FileSystemPathIdentity& pathIdentity = typedPlan.sourceEndpoint.pathIdentity.value();
                    for (size_t sourceIndex = 0; sourceIndex < typedPlan.selectedItems.size(); ++sourceIndex)
                    {
                        const std::wstring& sourcePath = typedPlan.selectedItems[sourceIndex].providerPath;
                        if (typedPlan.explicitMappings.empty())
                        {
                            // The ordinary destination is destinationFolder/sourceLeaf. If the
                            // folder itself is the source or lies below it, publication must recurse
                            // into the selected item. Reject this structural envelope before binding.
                            if (EquivalentPath(pathIdentity, sourcePath, typedPlan.destination.providerFolderPath) ||
                                IsStrictDescendantPath(pathIdentity, sourcePath, typedPlan.destination.providerFolderPath))
                            {
                                return RejectPlan(PlanRejectionBucket::DestinationInsideSource, rejectionBucket);
                            }
                            continue;
                        }

                        const auto mapping = std::ranges::find_if(typedPlan.explicitMappings, [sourceIndex](const auto& candidate) noexcept
                        { return candidate.sourceIndex == sourceIndex; });
                        if (mapping != typedPlan.explicitMappings.end() &&
                            IsStrictDescendantPath(pathIdentity, sourcePath, mapping->destinationProviderPath))
                        {
                            return RejectPlan(PlanRejectionBucket::DestinationInsideSource, rejectionBucket);
                        }
                    }
                }
                if (typedPlan.intent == FileOperations::TransferIntent::Move &&
                    QualifiedEndpointsMatchImpl(typedPlan.sourceEndpoint, typedPlan.destinationEndpoint))
                {
                    for (size_t sourceIndex = 0; sourceIndex < typedPlan.selectedItems.size(); ++sourceIndex)
                    {
                        const std::wstring& sourcePath = typedPlan.selectedItems[sourceIndex].providerPath;
                        std::wstring destinationParent = typedPlan.destination.providerFolderPath;
                        if (! typedPlan.explicitMappings.empty())
                        {
                            const auto mapping = std::ranges::find_if(typedPlan.explicitMappings, [sourceIndex](const auto& candidate) noexcept
                            { return candidate.sourceIndex == sourceIndex; });
                            if (mapping != typedPlan.explicitMappings.end())
                            {
                                if (! TryGetFileSystemParentPath(typedPlan.destinationEndpoint.pathIdentity.value(),
                                                                  mapping->destinationProviderPath,
                                                                  destinationParent))
                                {
                                    return RejectPlan(PlanRejectionBucket::InvalidExplicitMappings, rejectionBucket);
                                }
                            }
                        }
                        std::wstring sourceParent;
                        if (! TryGetFileSystemParentPath(typedPlan.sourceEndpoint.pathIdentity.value(), sourcePath, sourceParent))
                        {
                            return RejectPlan(PlanRejectionBucket::MalformedSource, rejectionBucket);
                        }
                        if (EquivalentPath(typedPlan.sourceEndpoint.pathIdentity.value(), sourceParent, destinationParent))
                        {
                            return RejectPlan(PlanRejectionBucket::SameFolderMove, rejectionBucket);
                        }
                    }
                }
                return S_OK;
            }
            else if constexpr (std::is_same_v<Plan, FileOperations::DeletePlan>)
            {
                if (! IsQualifiedEndpointValid(typedPlan.endpoint))
                {
                    return RejectPlan(PlanRejectionBucket::MalformedEndpoint, rejectionBucket);
                }
                if (typedPlan.selectedItems.empty())
                {
                    return RejectPlan(PlanRejectionBucket::EmptySelection, rejectionBucket);
                }
                if (std::ranges::any_of(typedPlan.selectedItems, [](const FileOperations::QualifiedSourceItem& item) noexcept
                                        { return item.providerPath.empty(); }))
                {
                    return RejectPlan(PlanRejectionBucket::MalformedSource, rejectionBucket);
                }
                if (IsLocalFileSystemEndpointImpl(typedPlan.endpoint) &&
                    std::ranges::any_of(typedPlan.selectedItems, [](const FileOperations::QualifiedSourceItem& item) noexcept
                    { return ! Common::Paths::IsSupportedLocalFileOperationPath(item.providerPath); }))
                {
                    return RejectPlan(PlanRejectionBucket::UnsupportedDeviceNamespace, rejectionBucket);
                }
                // C1: a Permanent Delete root is pinned by Preparing, not snapshotted at admission.
                if (typedPlan.nativeAuthority &&
                    std::ranges::any_of(typedPlan.selectedItems, [](const FileOperations::QualifiedSourceItem& item) noexcept
                    { return item.ingressSnapshot.has_value(); }))
                {
                    return RejectPlan(PlanRejectionBucket::MalformedSource, rejectionBucket);
                }
                if (typedPlan.mode == FileOperations::DeleteMode::Permanent && ! typedPlan.initialConsent.has_value() &&
                    ! allowPendingPermanentConsent)
                {
                    return RejectPlan(PlanRejectionBucket::MissingDestructiveConsent, rejectionBucket);
                }
                if (typedPlan.initialConsent.has_value() && typedPlan.initialConsent->taskNonce == 0u)
                {
                    return RejectPlan(PlanRejectionBucket::MissingDestructiveConsent, rejectionBucket);
                }
                return S_OK;
            }
            else
            {
                const bool scheduledRename = typedPlan.origin == FileOperations::RenameOrigin::BatchRename ||
                                             typedPlan.origin == FileOperations::RenameOrigin::ChangeCase;
                const bool originValid = typedPlan.origin == FileOperations::RenameOrigin::InlineRename || scheduledRename;
                if (! originValid || ! IsQualifiedEndpointValid(typedPlan.endpoint) || typedPlan.finalMappings.empty())
                {
                    return RejectPlan(! originValid || typedPlan.finalMappings.empty() ? PlanRejectionBucket::InvalidRename
                                                                                       : PlanRejectionBucket::MalformedEndpoint,
                                      rejectionBucket);
                }
                for (const FileOperations::RenameStep& step : typedPlan.finalMappings)
                {
                    // An inline rename publishes with its join and collision key pending; Preparing
                    // fills both from the provider's contract before any interlock or mutation.
                    const bool joinPending = typedPlan.origin == FileOperations::RenameOrigin::InlineRename &&
                        step.providerJoinedPath.empty() && step.providerCollisionKey.empty();
                    if (step.source.providerPath.empty() || step.finalLeafName.empty() ||
                        (! joinPending && (step.providerJoinedPath.empty() || step.providerCollisionKey.empty())))
                    {
                        return RejectPlan(PlanRejectionBucket::InvalidRename, rejectionBucket);
                    }
                    // A leaf never spells a path. This check is lexical and needs no provider I/O, so it
                    // stays at admission even when the provider join is filled in Preparing.
                    const std::wstring_view leafSeparators =
                        typedPlan.endpoint.pathIdentity.has_value() ? std::wstring_view(typedPlan.endpoint.pathIdentity->acceptedSeparators) : L"\\/";
                    if (step.finalLeafName.find_first_of(leafSeparators) != std::wstring::npos)
                    {
                        return RejectPlan(PlanRejectionBucket::InvalidRename, rejectionBucket);
                    }
                    if (IsLocalFileSystemEndpointImpl(typedPlan.endpoint) && ! Common::Paths::IsSupportedLocalFileOperationPath(step.source.providerPath))
                    {
                        return RejectPlan(PlanRejectionBucket::UnsupportedDeviceNamespace, rejectionBucket);
                    }
                }
                if (scheduledRename &&
                    std::ranges::any_of(typedPlan.finalMappings, [&](const FileOperations::RenameStep& step) noexcept
                    {
                        return ! step.source.ingressSnapshot.has_value() ||
                               step.source.ingressSnapshot->objectId.empty() ||
                               step.source.ingressSnapshot->pathProfileId != typedPlan.endpoint.profileId ||
                               step.providerParentKey.empty() || step.providerSourceCollisionKey.empty();
                    }))
                {
                    return RejectPlan(PlanRejectionBucket::MalformedSource, rejectionBucket);
                }
                if (scheduledRename)
                {
                    std::unordered_set<std::string> sourceIdentities;
                    sourceIdentities.reserve(typedPlan.finalMappings.size());
                    for (const FileOperations::RenameStep& step : typedPlan.finalMappings)
                    {
                        const std::vector<std::byte>& objectId = step.source.ingressSnapshot->objectId;
                        const std::string identityKey(reinterpret_cast<const char*>(objectId.data()), objectId.size());
                        if (! sourceIdentities.emplace(identityKey).second)
                        {
                            // A physical object has one state transition. Treat hard-link/path
                            // aliases as a duplicate source rather than inventing row ordering.
                            return RejectPlan(PlanRejectionBucket::InvalidRename, rejectionBucket);
                        }
                    }
                }
                if (typedPlan.origin == FileOperations::RenameOrigin::InlineRename)
                {
                    if (! typedPlan.schedule.layers.empty() || ! typedPlan.schedule.cycleOperationIndices.empty())
                    {
                        return RejectPlan(PlanRejectionBucket::InvalidRename, rejectionBucket);
                    }
                    return S_OK;
                }
                if (! typedPlan.schedule.cycleOperationIndices.empty())
                {
                    return RejectPlan(PlanRejectionBucket::InvalidRename, rejectionBucket);
                }
                std::vector<bool> scheduled(typedPlan.finalMappings.size(), false);
                size_t scheduledCount = 0u;
                for (const BatchRenameExecutionLayer& layer : typedPlan.schedule.layers)
                {
                    if (layer.operationIndices.empty())
                    {
                        return RejectPlan(PlanRejectionBucket::InvalidRename, rejectionBucket);
                    }
                    for (const size_t index : layer.operationIndices)
                    {
                        if (index >= scheduled.size() || scheduled[index])
                        {
                            return RejectPlan(PlanRejectionBucket::InvalidRename, rejectionBucket);
                        }
                        scheduled[index] = true;
                        ++scheduledCount;
                    }
                }
                return scheduledCount == typedPlan.finalMappings.size()
                           ? S_OK
                           : RejectPlan(PlanRejectionBucket::InvalidRename, rejectionBucket);
            }
        },
        plan);
}

[[nodiscard]] std::wstring_view PlanRejectionBucketToString(PlanRejectionBucket bucket) noexcept
{
    switch (bucket)
    {
        case PlanRejectionBucket::None: return L"none";
        case PlanRejectionBucket::UnsupportedOperation: return L"unsupported-operation";
        case PlanRejectionBucket::MissingFileSystem: return L"missing-filesystem";
        case PlanRejectionBucket::EmptySelection: return L"empty-selection";
        case PlanRejectionBucket::MalformedEndpoint: return L"malformed-endpoint";
        case PlanRejectionBucket::MalformedSource: return L"malformed-source";
        case PlanRejectionBucket::MalformedDestination: return L"malformed-destination";
        case PlanRejectionBucket::UnsupportedDeviceNamespace: return L"unsupported-device-namespace";
        case PlanRejectionBucket::EscapingDestinationMapping: return L"escaping-destination-mapping";
        case PlanRejectionBucket::DestinationInsideSource: return L"destination-inside-source";
        case PlanRejectionBucket::SameFolderMove: return L"same-folder-move";
        case PlanRejectionBucket::InvalidStrategy: return L"invalid-strategy";
        case PlanRejectionBucket::MixedSourceEndpoint: return L"mixed-source-endpoint";
        case PlanRejectionBucket::InvalidExplicitMappings: return L"invalid-explicit-mappings";
        case PlanRejectionBucket::InvalidClipboardSequence: return L"invalid-clipboard-sequence";
        case PlanRejectionBucket::MissingDestructiveConsent: return L"missing-destructive-consent";
        case PlanRejectionBucket::InvalidRename: return L"invalid-rename";
        case PlanRejectionBucket::InvalidDestinationName: return L"invalid-destination-name";
    }
    return L"unknown";
}

[[nodiscard]] uint64_t PlanPayloadBytes(const FileOperationPlan& plan) noexcept
{
    const auto stringBytes = [](std::wstring_view value) noexcept -> uint64_t
    { return static_cast<uint64_t>(value.size()) * sizeof(wchar_t); };
    const auto endpointBytes = [&](const FileOperations::QualifiedEndpoint& endpoint) noexcept -> uint64_t
    {
        uint64_t bytes = stringBytes(endpoint.pluginId) + stringBytes(endpoint.instanceId) + stringBytes(endpoint.profileId) +
            stringBytes(endpoint.rootId);
        if (endpoint.pathIdentity.has_value())
        {
            bytes += sizeof(endpoint.pathIdentity->pathTextStableIdentity) + sizeof(endpoint.pathIdentity->componentComparison) +
                sizeof(endpoint.pathIdentity->preferredSeparator) + stringBytes(endpoint.pathIdentity->acceptedSeparators) +
                sizeof(endpoint.pathIdentity->casePreserving) + sizeof(endpoint.pathIdentity->caseOnlyRename);
        }
        return bytes;
    };
    const auto itemBytes = [&](const FileOperations::QualifiedSourceItem& item) noexcept -> uint64_t
    {
        uint64_t bytes = stringBytes(item.providerPath);
        if (item.ingressSnapshot.has_value())
        {
            bytes += static_cast<uint64_t>(item.ingressSnapshot->objectId.size() + item.ingressSnapshot->revisionId.size());
            bytes += stringBytes(item.ingressSnapshot->pathProfileId);
        }
        return bytes;
    };

    return std::visit(
        [&](const auto& typedPlan) noexcept -> uint64_t
        {
            using Plan = std::remove_cvref_t<decltype(typedPlan)>;
            uint64_t bytes = 0;
            if constexpr (std::is_same_v<Plan, FileOperations::TransferPlan>)
            {
                bytes += endpointBytes(typedPlan.sourceEndpoint) + endpointBytes(typedPlan.destinationEndpoint);
                bytes += stringBytes(typedPlan.destination.providerFolderPath);
                for (const auto& item : typedPlan.selectedItems)
                    bytes += itemBytes(item);
                for (const auto& mapping : typedPlan.explicitMappings)
                    bytes += sizeof(mapping.sourceIndex) + stringBytes(mapping.destinationProviderPath);
            }
            else if constexpr (std::is_same_v<Plan, FileOperations::DeletePlan>)
            {
                bytes += endpointBytes(typedPlan.endpoint);
                for (const auto& item : typedPlan.selectedItems)
                    bytes += itemBytes(item);
            }
            else
            {
                bytes += endpointBytes(typedPlan.endpoint);
                for (const auto& step : typedPlan.finalMappings)
                    bytes += itemBytes(step.source) + stringBytes(step.finalLeafName) + stringBytes(step.providerJoinedPath) +
                             stringBytes(step.providerParentKey) + stringBytes(step.providerSourceCollisionKey) +
                             stringBytes(step.providerCollisionKey);
                for (const BatchRenameExecutionLayer& layer : typedPlan.schedule.layers)
                    bytes += static_cast<uint64_t>(layer.operationIndices.size()) * sizeof(size_t);
            }
            return bytes;
        },
        plan);
}

[[nodiscard]] uint64_t PlanSourceCount(const FileOperationPlan& plan) noexcept
{
    return std::visit(
        [](const auto& typedPlan) noexcept -> uint64_t
        {
            using Plan = std::remove_cvref_t<decltype(typedPlan)>;
            if constexpr (std::is_same_v<Plan, FileOperations::RenamePlan>)
                return static_cast<uint64_t>(typedPlan.finalMappings.size());
            else
                return static_cast<uint64_t>(typedPlan.selectedItems.size());
        },
        plan);
}

[[nodiscard]] std::wstring CanonicalInstanceId(std::wstring_view instanceId)
{
    if (instanceId.empty())
    {
        return L"host/default";
    }
    return std::format(L"host/context/{}", instanceId);
}

struct FileSystemCapabilitiesV2
{
    std::wstring providerId;
    std::wstring pathProfileId;
    std::wstring rootId;
    bool copyOperation = false;
    bool moveOperation = false;
    bool nativeMoveOperation = false;
    bool renameOperation = false;
    bool createDirectoryOperation = false;
    bool read = false;
    bool write = false;
    bool deleteOperation = false;
    bool properties = false;
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
    bool verificationHostReadback = false;
    bool verificationProviderBlake3Proof = false;
    bool verificationWriterDigestProof = false; // R3-2
    bool verificationCapabilityCheckDeferred = false;
    bool cancellationAbort = false;
    bool cancellationDeadline = false;
    FileOperations::CancellationRouteClass cancellationRouteClass = FileOperations::CancellationRouteClass::Uncontained;
    uint32_t providerWatchdogTimeoutMs = 0u;
    uint64_t maxComponentUtf16 = 0u;
    FileSystemNamespaceKind namespaceKind = FILESYSTEM_NAMESPACE_PROVIDER_VIRTUAL_FOLDER;
    std::optional<FileSystemPathIdentity> pathIdentity;

    // Retained only while the host is assembling the current plan. It lets pair
    // admission ask the provider's typed directional peer policy without copying
    // JSON lists into a second authority.
    wil::com_ptr<IFileSystem> routeFileSystem;
    std::wstring routePath;
};

[[nodiscard]] bool HasStablePathIdentity(const FileSystemCapabilitiesV2& capabilities) noexcept
{
    return capabilities.pathIdentity.has_value() && capabilities.pathIdentity->pathTextStableIdentity;
}

[[nodiscard]] bool HasContainedCancellationRoute(const FileSystemCapabilitiesV2& capabilities) noexcept
{
    return capabilities.cancellationRouteClass != FileOperations::CancellationRouteClass::Uncontained;
}

[[nodiscard]] bool SamePathIdentityContract(const FileSystemPathIdentity& left, const FileSystemPathIdentity& right) noexcept
{
    return left.pathTextStableIdentity == right.pathTextStableIdentity && left.componentComparison == right.componentComparison &&
           left.preferredSeparator == right.preferredSeparator && left.acceptedSeparators == right.acceptedSeparators &&
           left.casePreserving == right.casePreserving && left.caseOnlyRename == right.caseOnlyRename;
}

[[nodiscard]] bool CapabilitiesAllowSameFileSystemOperation(const FileSystemCapabilitiesV2& capabilities,
                                                            FileSystemOperation operation,
                                                            FileSystemFlags flags = FILESYSTEM_FLAG_NONE) noexcept
{
    if (! HasStablePathIdentity(capabilities) || ! HasContainedCancellationRoute(capabilities))
    {
        return false;
    }
    switch (operation)
    {
        case FILESYSTEM_COPY: return capabilities.copyOperation;
        case FILESYSTEM_MOVE: return capabilities.moveOperation;
        case FILESYSTEM_DELETE:
            return (flags & FILESYSTEM_FLAG_USE_RECYCLE_BIN) != 0 ? capabilities.recycleOperation : capabilities.deleteOperation;
        case FILESYSTEM_RENAME: return capabilities.renameOperation;
        case FILESYSTEM_CREATE_DIRECTORY: return capabilities.createDirectoryOperation;
        default: return true;
    }
}

[[nodiscard]] std::vector<FolderView::RemovalDisposition> BuildDenseRemovedSourceDispositions(
    const FolderWindow::FileOperationCompletedEvent& event)
{
    std::vector<FolderView::RemovalDisposition> dispositions(
        event.sourcePaths.size(), FolderView::RemovalDisposition::Retained);
    for (const FolderWindow::FileOperationItemOutcome& outcome : event.itemOutcomes)
    {
        if (outcome.sourceIndex < dispositions.size())
        {
            dispositions[outcome.sourceIndex] = outcome.sourceDisposition == FileOperations::SourceDisposition::Removed
                ? FolderView::RemovalDisposition::Removed
                : FolderView::RemovalDisposition::Retained;
        }
    }
    return dispositions;
}

// The typed route interface is mandatory for executable admission. Capability JSON is
// diagnostics-only and is never consulted here.
void ReportCapabilitiesContractViolationOnce(const IFileSystem* fileSystem, std::wstring_view pluginId, HRESULT hr) noexcept
{
    static std::mutex reportedMutex;
    static std::unordered_set<const IFileSystem*> reportedProviders;

    {
        // The only throwing operation here is the set insertion, and that only throws
        // std::bad_alloc; per the repo exception policy allocation failure terminates
        // (this function is noexcept), so no catch is needed.
        const std::lock_guard lock(reportedMutex);
        if (! reportedProviders.insert(fileSystem).second)
        {
            return;
        }
    }

    Debug::Error(L"FolderWindow filesystem provider '{}' violates the mandatory typed route contract (hr=0x{:08X}); "
                 L"capability-gated operations are disabled for this instance.",
                 pluginId.empty() ? std::wstring_view(L"<unknown>") : pluginId,
                 static_cast<unsigned long>(hr));
}

[[nodiscard]] std::optional<FileSystemCapabilitiesV2> TryGetCapabilities(const wil::com_ptr<IFileSystem>& fileSystem,
                                                                          std::wstring_view path,
                                                                          FileSystemOperation operation,
                                                                          std::wstring_view pluginId = {}) noexcept
{
    if (! fileSystem)
    {
        return std::nullopt;
    }
    if (path.empty())
    {
        Debug::Error(L"FolderWindow capability query rejected an empty provider path for plugin '{}'.", pluginId);
        return std::nullopt;
    }

    const ULONGLONG capabilityQueryStart = GetTickCount64();
    FileSystemRouteContract::QueryResult query =
        FileSystemRouteContract::Query(fileSystem.get(), path, operation, pluginId);
    const ULONGLONG capabilityQueryElapsed = GetTickCount64() - capabilityQueryStart;
    if (query.state != FileSystemRouteContract::QueryState::Available)
    {
        if (query.state == FileSystemRouteContract::QueryState::ContractViolation)
        {
            ReportCapabilitiesContractViolationOnce(fileSystem.get(), pluginId, query.status);
        }
        return std::nullopt;
    }

    const FileSystemRouteContract::Snapshot& facts = query.snapshot;
    FileSystemCapabilitiesV2 capabilities{
        .providerId = facts.providerId,
        .pathProfileId = facts.pathProfileId,
        .rootId = facts.rootId,
        .copyOperation = facts.copyOperation,
        .moveOperation = facts.moveOperation,
        .nativeMoveOperation = facts.nativeMoveOperation,
        .renameOperation = facts.renameOperation,
        .createDirectoryOperation = facts.createDirectoryOperation,
        .read = facts.read,
        .write = facts.write,
        .deleteOperation = facts.deleteOperation,
        .properties = facts.properties,
        .recycleOperation = facts.recycleOperation,
        .boundDelete = facts.boundDelete,
        .conditionalDelete = facts.conditionalDelete,
        .exclusiveStage = facts.exclusiveStage,
        .conditionalPublish = facts.conditionalPublish,
        .committedSize = facts.committedSize,
        .preserveFileLink = facts.preserveFileLink,
        .preserveDirectoryLink = facts.preserveDirectoryLink,
        .retargetInTree = facts.retargetInTree,
        .exactLinkRemoval = facts.exactLinkRemoval,
        .verificationHostReadback = (facts.proofFlags & FILESYSTEM_ROUTE_PROOF_HOST_READBACK) != 0u,
        .verificationProviderBlake3Proof = (facts.proofFlags & FILESYSTEM_ROUTE_PROOF_PROVIDER_BLAKE3) != 0u,
        .verificationWriterDigestProof = (facts.proofFlags & FILESYSTEM_ROUTE_PROOF_WRITER_DIGEST) != 0u,
        .verificationCapabilityCheckDeferred = capabilityQueryElapsed > kVerificationCapabilityUiBudgetMs
#ifdef ENABLE_TESTS
            || g_forceVerificationCapabilityBudgetExceeded.load(std::memory_order_acquire)
#endif
            ,
        .cancellationAbort = facts.cancellationAbort,
        .cancellationDeadline = facts.cancellationDeadline,
        .cancellationRouteClass = facts.cancellationRoute == FILESYSTEM_CANCELLATION_BOUNDED
            ? FileOperations::CancellationRouteClass::Bounded
            : (facts.cancellationRoute == FILESYSTEM_CANCELLATION_PROVIDER_WATCHDOG
                   ? FileOperations::CancellationRouteClass::ProviderWatchdog
                   : FileOperations::CancellationRouteClass::Uncontained),
        .providerWatchdogTimeoutMs = facts.providerWatchdogTimeoutMs,
        .maxComponentUtf16 = facts.maxComponentUtf16,
        .namespaceKind = facts.namespaceKind,
        .pathIdentity = facts.pathIdentity,
        .routeFileSystem = fileSystem,
        .routePath = std::wstring(path),
    };
    return capabilities;
}

[[nodiscard]] std::optional<FileSystemPathIdentity> TryGetStablePathIdentity(const wil::com_ptr<IFileSystem>& fileSystem,
                                                                             std::wstring_view providerPath,
                                                                             std::wstring_view pluginId) noexcept
{
    const std::optional<FileSystemCapabilitiesV2> capabilities = TryGetCapabilities(fileSystem, providerPath, FILESYSTEM_RENAME, pluginId);
    if (! capabilities.has_value() || ! HasStablePathIdentity(capabilities.value()))
    {
        return std::nullopt;
    }
    return capabilities->pathIdentity;
}

[[nodiscard]] bool CanSameFileSystemOperationFromCapabilities(const wil::com_ptr<IFileSystem>& fileSystem,
                                                              std::wstring_view providerPath,
                                                              FileSystemOperation operation,
                                                              std::wstring_view pluginId = {},
                                                              FileSystemFlags flags = FILESYSTEM_FLAG_NONE) noexcept
{
    const std::optional<FileSystemCapabilitiesV2> capabilities = TryGetCapabilities(fileSystem, providerPath, operation, pluginId);
    if (! capabilities.has_value() || ! HasContainedCancellationRoute(capabilities.value()))
    {
        return false;
    }

    if (operation != FILESYSTEM_MOVE || (capabilities->moveOperation && capabilities->nativeMoveOperation))
    {
        return CapabilitiesAllowSameFileSystemOperation(capabilities.value(), operation, flags);
    }

    const std::optional<FileSystemCapabilitiesV2> copyCapabilities = TryGetCapabilities(fileSystem, providerPath, FILESYSTEM_COPY, pluginId);
    return copyCapabilities.has_value() && HasContainedCancellationRoute(copyCapabilities.value()) &&
        CapabilitiesAllowSameFileSystemOperation(copyCapabilities.value(), FILESYSTEM_COPY);
}

[[nodiscard]] bool TypedPeerAllows(const FileSystemCapabilitiesV2& capabilities,
                                   FileSystemOperation operation,
                                   FileSystemTransferPeerRole role,
                                   std::wstring_view otherPluginId) noexcept
{
    if (! capabilities.routeFileSystem || capabilities.routePath.empty() || otherPluginId.empty())
    {
        return false;
    }
    const FileSystemRouteContract::BooleanResult result = FileSystemRouteContract::QueryTransferPeerAllowed(
        capabilities.routeFileSystem.get(), capabilities.routePath, operation, role, otherPluginId);
    return result.state == FileSystemRouteContract::QueryState::Available && result.value;
}

[[nodiscard]] bool CanCrossFileSystemCopyMove(const wil::com_ptr<IFileSystem>& sourceFileSystem,
                                              std::wstring_view sourceProviderPath,
                                              std::wstring_view sourcePluginId,
                                              const wil::com_ptr<IFileSystem>& destinationFileSystem,
                                              std::wstring_view destinationProviderPath,
                                              std::wstring_view destinationPluginId,
                                              FileSystemOperation operation) noexcept
{
    if (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE)
    {
        return false;
    }

    const FileSystemOperation pairOperation = operation == FILESYSTEM_MOVE ? FILESYSTEM_COPY : operation;
    const std::optional<FileSystemCapabilitiesV2> sourceCaps =
        TryGetCapabilities(sourceFileSystem, sourceProviderPath, pairOperation, sourcePluginId);
    const std::optional<FileSystemCapabilitiesV2> destCaps =
        TryGetCapabilities(destinationFileSystem, destinationProviderPath, pairOperation, destinationPluginId);
    if (! sourceCaps.has_value() || ! destCaps.has_value() || ! HasStablePathIdentity(sourceCaps.value()) ||
        ! HasStablePathIdentity(destCaps.value()) || ! HasContainedCancellationRoute(sourceCaps.value()) ||
        ! HasContainedCancellationRoute(destCaps.value()))
    {
        return false;
    }

    if (! sourceCaps->read || ! destCaps->write)
    {
        return false;
    }

    if (operation == FILESYSTEM_MOVE)
    {
        const std::optional<FileSystemCapabilitiesV2> sourceMoveCaps =
            TryGetCapabilities(sourceFileSystem, sourceProviderPath, FILESYSTEM_MOVE, sourcePluginId);
        const std::optional<FileSystemCapabilitiesV2> destinationMoveCaps =
            TryGetCapabilities(destinationFileSystem, destinationProviderPath, FILESYSTEM_MOVE, destinationPluginId);
        if (! sourceMoveCaps.has_value() || ! destinationMoveCaps.has_value() ||
            ! HasContainedCancellationRoute(sourceMoveCaps.value()) || ! HasContainedCancellationRoute(destinationMoveCaps.value()))
        {
            return false;
        }
    }

    return TypedPeerAllows(sourceCaps.value(), FILESYSTEM_COPY, FILESYSTEM_TRANSFER_PEER_EXPORT, destinationPluginId) &&
        TypedPeerAllows(destCaps.value(), FILESYSTEM_COPY, FILESYSTEM_TRANSFER_PEER_IMPORT, sourcePluginId);
}

[[nodiscard]] bool CapabilityPairAllowsCrossFileSystemCopyMove(const FileSystemCapabilitiesV2& sourceCaps,
                                                               std::wstring_view sourcePluginId,
                                                               const FileSystemCapabilitiesV2& destinationCaps,
                                                               std::wstring_view destinationPluginId,
                                                               FileSystemOperation operation) noexcept
{
    if ((operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE) || ! HasStablePathIdentity(sourceCaps) ||
        ! HasStablePathIdentity(destinationCaps) || ! HasContainedCancellationRoute(sourceCaps) ||
        ! HasContainedCancellationRoute(destinationCaps) || ! sourceCaps.read || ! destinationCaps.write)
    {
        return false;
    }

    return TypedPeerAllows(sourceCaps, operation, FILESYSTEM_TRANSFER_PEER_EXPORT, destinationPluginId) &&
        TypedPeerAllows(destinationCaps, operation, FILESYSTEM_TRANSFER_PEER_IMPORT, sourcePluginId);
}

[[nodiscard]] FileOperations::ObjectBindingResult CaptureBoundObjectAuthorityImpl(
    wil::com_ptr<IFileSystemBoundObject> bound,
    std::wstring_view pathProfileId) noexcept
{
    FileOperations::ObjectBindingResult result{};
    if (! bound)
    {
        result.state  = FileOperations::ObjectBindingState::ProviderContractViolation;
        result.status = E_UNEXPECTED;
        return result;
    }

    FileSystemBoundObjectSnapshot snapshot{};
    snapshot.sizeBytes = sizeof(snapshot);
    const HRESULT snapshotHr = bound->GetSnapshot(&snapshot);
    constexpr uint32_t kMaxIdentityTokenBytes = 64u * 1024u;
    const bool kindValid = snapshot.kind >= FILESYSTEM_BOUND_REGULAR_FILE && snapshot.kind <= FILESYSTEM_BOUND_OTHER;
    const bool objectValid = snapshot.objectId != nullptr && snapshot.objectIdBytes > 0u && snapshot.objectIdBytes <= kMaxIdentityTokenBytes;
    const bool revisionValid = (snapshot.revisionIdBytes == 0u && snapshot.revisionId == nullptr) ||
                               (snapshot.revisionIdBytes > 0u && snapshot.revisionIdBytes <= kMaxIdentityTokenBytes && snapshot.revisionId != nullptr);
    if (FAILED(snapshotHr))
    {
        result.state  = FileOperations::ObjectBindingState::Indeterminate;
        result.status = snapshotHr;
        return result;
    }
    if (snapshot.sizeBytes != sizeof(snapshot) || ! kindValid || ! objectValid || ! revisionValid)
    {
        Debug::Error(L"File Operations object-binding provider returned a malformed successful snapshot.");
        result.state  = FileOperations::ObjectBindingState::ProviderContractViolation;
        result.status = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        return result;
    }

    result.authority.boundObject = std::move(bound);
    result.authority.identity.pathProfileId.assign(pathProfileId);
    const auto* objectBegin = static_cast<const std::byte*>(snapshot.objectId);
    result.authority.identity.objectId.assign(objectBegin, objectBegin + snapshot.objectIdBytes);
    if (snapshot.revisionIdBytes > 0u)
    {
        const auto* revisionBegin = static_cast<const std::byte*>(snapshot.revisionId);
        result.authority.identity.revisionId.assign(revisionBegin, revisionBegin + snapshot.revisionIdBytes);
    }
    result.authority.kind               = static_cast<FileSystemBoundObjectKind>(snapshot.kind);
    result.authority.committedSizeBytes = snapshot.committedSizeBytes;
    result.state                        = FileOperations::ObjectBindingState::Bound;
    result.status                       = S_OK;
    return result;
}

[[nodiscard]] HRESULT CrossCheckBoundObjectIdentityImpl(const FileOperations::BoundObjectAuthority& expected,
                                                         const FileOperations::BoundObjectAuthority& current,
                                                         bool& sameObject,
                                                         bool& sameRevision) noexcept
{
    sameObject   = false;
    sameRevision = false;
    if (! expected.boundObject || ! current.boundObject || expected.identity.objectId.empty() || current.identity.objectId.empty() ||
        expected.identity.pathProfileId.empty() || expected.identity.pathProfileId != current.identity.pathProfileId)
    {
        return E_INVALIDARG;
    }

    BOOL providerSame = FALSE;
    const HRESULT sameHr = expected.boundObject->IsSameObject(current.boundObject.get(), &providerSame);
    if (FAILED(sameHr))
    {
        return sameHr;
    }

    const bool snapshotSame = expected.identity.objectId == current.identity.objectId;
    if ((providerSame != FALSE) != snapshotSame)
    {
        Debug::Error(L"File Operations provider IsSameObject result disagrees with its immutable object identity snapshot.");
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    sameObject   = snapshotSame;
    sameRevision = expected.identity.revisionId == current.identity.revisionId;
    return S_OK;
}
} // namespace

#ifdef ENABLE_TESTS
void SetFileOpsVerificationCapabilityBudgetExceededForSelfTest(bool enabled) noexcept
{
    g_forceVerificationCapabilityBudgetExceeded.store(enabled, std::memory_order_release);
}
#endif

HRESULT FileOperations::CrossCheckBoundObjectIdentity(const BoundObjectAuthority& expected,
                                                       const BoundObjectAuthority& current,
                                                       bool& sameObject,
                                                       bool& sameRevision) noexcept
{
    return CrossCheckBoundObjectIdentityImpl(expected, current, sameObject, sameRevision);
}

HRESULT FileOperations::ValidatePlan(const FileOperationPlan& plan, PlanRejectionBucket* rejectionBucket) noexcept
{
    return ValidatePlanImpl(plan, rejectionBucket, false);
}

HRESULT FileOperations::ValidatePlan(const FileOperationPlan& plan, PlanRejectionBucket* rejectionBucket, const bool allowPendingPermanentConsent) noexcept
{
    return ValidatePlanImpl(plan, rejectionBucket, allowPendingPermanentConsent);
}

std::optional<FileOperations::OperationStrategy> FileOperations::SelectMoveStrategy(
    const MoveStrategyQualificationFacts& facts) noexcept
{
    if (facts.nativeMoveQualified)
    {
        return OperationStrategy::Native;
    }

    if (! facts.copyPairQualified)
    {
        return std::nullopt;
    }

    // R3-2: a destination without object binding qualifies when its atomic-final writer proves the
    // published content; the bridge then removes each source only after that proof matches.
    const bool destinationQualified =
        (facts.destinationExclusiveStage && facts.destinationConditionalPublish && facts.destinationBindingAvailable) ||
        facts.destinationWriterDigestProof;
    if (facts.movePairQualified && facts.sourceBoundDelete && facts.sourceConditionalDelete && facts.sourceBindingAvailable &&
        destinationQualified)
    {
        return OperationStrategy::Managed;
    }

    return OperationStrategy::CopyOnly;
}

bool FileOperations::IsProviderKeepBothDestinationEligible(bool providerHandlesNestedKeepBoth,
                                                            std::wstring_view conflictDestinationPath,
                                                            std::wstring_view selectedDestinationPath) noexcept
{
    return providerHandlesNestedKeepBoth ||
           (! selectedDestinationPath.empty() && NavigationLocation::EqualsNoCase(conflictDestinationPath, selectedDestinationPath));
}

FileOperations::ManagedCleanupAttemptDisposition FileOperations::ClassifyManagedCleanupMutation(
    const ManagedCleanupMutationFacts& facts) noexcept
{
    if (! facts.outcomeKnown)
    {
        return ManagedCleanupAttemptDisposition::Indeterminate;
    }

    if (facts.mutationCommitted)
    {
        return facts.originalStillPresent ? ManagedCleanupAttemptDisposition::ProviderContractViolation
                                          : ManagedCleanupAttemptDisposition::Removed;
    }

    if (! facts.originalStillPresent)
    {
        // Another actor may have removed the already-revalidated exact object before the
        // conditional call committed. Source disposition is still truthfully Removed.
        return ManagedCleanupAttemptDisposition::Removed;
    }

    if (SUCCEEDED(facts.status))
    {
        return ManagedCleanupAttemptDisposition::ProviderContractViolation;
    }

    return ManagedCleanupAttemptDisposition::Retained;
}

bool FileOperations::TryResolveTransferDestinationProviderPath(const TransferPlan& plan,
                                                               const size_t sourceIndex,
                                                               std::wstring& destinationOut) noexcept
{
    destinationOut.clear();
    if (sourceIndex >= plan.selectedItems.size() || ! plan.sourceEndpoint.pathIdentity.has_value() ||
        ! plan.destinationEndpoint.pathIdentity.has_value() || plan.destination.providerFolderPath.empty())
    {
        return false;
    }

    if (! plan.explicitMappings.empty())
    {
        const auto mapping = std::ranges::find_if(plan.explicitMappings, [sourceIndex](const TransferDestinationMapping& candidate) noexcept
        { return candidate.sourceIndex == sourceIndex; });
        if (mapping == plan.explicitMappings.end() || mapping->destinationProviderPath.empty())
        {
            return false;
        }
        destinationOut = mapping->destinationProviderPath;
        return true;
    }

    std::wstring sourceLeaf;
    if (! TryGetFileSystemLeafName(plan.sourceEndpoint.pathIdentity.value(),
                                    plan.selectedItems[sourceIndex].providerPath,
                                    sourceLeaf))
    {
        return false;
    }

    destinationOut = JoinFileSystemPath(plan.destinationEndpoint.pathIdentity.value(),
                                        plan.destination.providerFolderPath,
                                        sourceLeaf);
    return ! destinationOut.empty();
}

bool FileOperations::IsLocalFileSystemEndpoint(const QualifiedEndpoint& endpoint) noexcept
{
    return IsLocalFileSystemEndpointImpl(endpoint);
}

bool FileOperations::QualifiedEndpointsReferToSameRoot(const QualifiedEndpoint& left,
                                                       const QualifiedEndpoint& right) noexcept
{
    return QualifiedEndpointsMatchImpl(left, right);
}

bool FileOperations::QualifiedEndpointsShareObjectIdentityDomain(const QualifiedEndpoint& left,
                                                                 const QualifiedEndpoint& right) noexcept
{
    return QualifiedEndpointsShareObjectIdentityDomainImpl(left, right);
}

FileOperations::ObjectBindingResult FileOperations::BindObjectAuthority(IFileSystem* fileSystem,
                                                                         std::wstring_view providerPath,
                                                                         std::wstring_view pathProfileId,
                                                                         FileSystemBindFlags flags) noexcept
{
    Debug::Perf::Scope perf(L"fileops.identity.bind_us");
    perf.SetValue0(static_cast<uint64_t>(providerPath.size()) * sizeof(wchar_t));

    ObjectBindingResult result{};
    const auto finish = [&](ObjectBindingState state, HRESULT status) noexcept -> ObjectBindingResult
    {
        result.state  = state;
        result.status = status;
        perf.SetValue1(static_cast<uint64_t>(state));
        perf.SetHr(status);
        return std::move(result);
    };
    if (fileSystem == nullptr || providerPath.empty() || pathProfileId.empty())
    {
        return finish(ObjectBindingState::ProviderContractViolation, E_INVALIDARG);
    }

    wil::com_ptr<IFileSystemObjectBinding> binding;
    const HRESULT queryHr = fileSystem->QueryInterface(__uuidof(IFileSystemObjectBinding), binding.put_void());
    if (queryHr == E_NOINTERFACE)
    {
        return finish(! binding ? ObjectBindingState::Unsupported : ObjectBindingState::ProviderContractViolation,
                      ! binding ? HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) : E_UNEXPECTED);
    }
    if (FAILED(queryHr))
    {
        return finish(ObjectBindingState::Indeterminate, queryHr);
    }
    if (! binding)
    {
        Debug::Error(L"File Operations object-binding provider returned QueryInterface success with a null interface.");
        return finish(ObjectBindingState::ProviderContractViolation, E_UNEXPECTED);
    }

    wil::com_ptr<IFileSystemBoundObject> bound;
    const std::wstring path(providerPath);
    const HRESULT bindHr = binding->BindObject(path.c_str(), flags, bound.put());
    if (bindHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) || bindHr == E_NOTIMPL)
    {
        return finish(ObjectBindingState::Unsupported, HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
    }
    if (bindHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || bindHr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND) ||
        bindHr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND))
    {
        return finish(ObjectBindingState::Missing, bindHr);
    }
    if (FAILED(bindHr))
    {
        return finish(ObjectBindingState::Indeterminate, bindHr);
    }
    if (! bound)
    {
        Debug::Error(L"File Operations object-binding provider returned success with a null bound object.");
        return finish(ObjectBindingState::ProviderContractViolation, E_UNEXPECTED);
    }

    ObjectBindingResult captured = CaptureBoundObjectAuthorityImpl(std::move(bound), pathProfileId);
    result.authority = std::move(captured.authority);
    return finish(captured.state, captured.status);
}

HRESULT FileOperationArtifacts::ProjectProviderObject(IFileSystem* const fileSystem,
                                                       const std::wstring_view providerPath,
                                                       const std::wstring_view pluginId,
                                                       const std::wstring_view instanceContext,
                                                       Projection& out) noexcept
{
    out = {};
    Candidate candidate{};
    const HRESULT captureHr = CaptureProviderObjectCandidate(fileSystem, providerPath, pluginId, instanceContext, candidate);
    if (captureHr != S_OK)
    {
        return captureHr;
    }
    out = ProjectCandidate(candidate);
    return out.classification == Classification::Ordinary ? S_FALSE : S_OK;
}

HRESULT FileOperationArtifacts::CaptureProviderObjectCandidate(IFileSystem* const fileSystem,
                                                                const std::wstring_view providerPath,
                                                                const std::wstring_view pluginId,
                                                                const std::wstring_view instanceContext,
                                                                Candidate& out) noexcept
{
    out = {};
    if (providerPath.empty())
    {
        return E_INVALIDARG;
    }

    const std::wstring canonicalInstance = CanonicalInstanceId(instanceContext);
    const bool possibleName = HasPossibleArtifactName(
        std::filesystem::path(std::wstring(providerPath)).filename().native());
    const auto possibleFallback = [&]() noexcept -> HRESULT
    {
        if (! possibleName)
        {
            return S_FALSE;
        }
        out.endpoint.pluginId = std::wstring(pluginId);
        out.endpoint.instanceId = canonicalInstance;
        out.pathIdentity = FileSystemPathIdentity{
            .pathTextStableIdentity = false,
            .componentComparison = FileSystemPathComponentComparison::OrdinalCaseSensitive,
        };
        out.path = std::wstring(providerPath);
        out.probeState = ProbeState::Indeterminate;
        return S_OK;
    };

    if (fileSystem == nullptr || pluginId.empty())
    {
        return possibleFallback();
    }
    wil::com_ptr<IFileSystem> retainedFileSystem = fileSystem;
    const std::optional<FileSystemCapabilitiesV2> capabilities =
        TryGetCapabilities(retainedFileSystem, providerPath, FILESYSTEM_RENAME, pluginId);
    if (! capabilities.has_value() || ! HasStablePathIdentity(capabilities.value()))
    {
        return possibleFallback();
    }

    out = Candidate{
        .endpoint = Endpoint{
            .pluginId = std::wstring(pluginId),
            .instanceId = canonicalInstance,
            .profileId = capabilities->pathProfileId,
            .rootId = capabilities->rootId,
        },
        .pathIdentity = capabilities->pathIdentity.value(),
        .path = std::wstring(providerPath),
        .probeState = ProbeState::Indeterminate,
    };
    if (ClassifyCandidate(out).classification == Classification::Ordinary)
    {
        out = {};
        return S_FALSE;
    }

    FileOperations::ObjectBindingResult bound = FileOperations::BindObjectAuthority(
        fileSystem,
        providerPath,
        capabilities->pathProfileId,
        static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA));
    Debug::Perf::EmitValue(L"fileops.artifact.capture.bind_count", 1u, bound.status);
    if (bound.state == FileOperations::ObjectBindingState::Missing)
    {
        out = {};
        return S_FALSE;
    }
    if (bound.state == FileOperations::ObjectBindingState::Bound)
    {
        out.probeState = ProbeState::Present;
        out.currentIdentity = Identity{
            .objectId = std::move(bound.authority.identity.objectId),
            .revisionId = std::move(bound.authority.identity.revisionId),
            .pathProfileId = std::move(bound.authority.identity.pathProfileId),
            .kind = bound.authority.kind,
        };
    }
    return S_OK;
}

HRESULT FileOperationArtifacts::RevalidateProviderTouchGuard(
    IFileSystem* const fileSystem,
    const std::wstring_view pluginId,
    const std::wstring_view instanceContext,
    const TouchGuardReceipt& receipt,
    const std::span<const std::filesystem::path> providerPaths) noexcept
{
    Debug::Perf::Scope perf(L"fileops.artifact.touch.direct_revalidate.us");
    perf.SetValue1(static_cast<uint64_t>(providerPaths.size()));
    const auto finish = [&](const HRESULT status, const size_t guardedCount = 0u) noexcept -> HRESULT
    {
        perf.SetValue0(static_cast<uint64_t>(guardedCount));
        perf.SetHr(status);
        return status;
    };
    if (fileSystem == nullptr || pluginId.empty())
    {
        return finish(E_INVALIDARG);
    }
    if (receipt.items.empty() || providerPaths.empty())
    {
        return finish(S_OK);
    }

    TouchGuardReceipt subset;
    subset.items.reserve(std::min(receipt.items.size(), providerPaths.size()));
    for (const TouchGuardItem& accepted : receipt.items)
    {
        const bool selected = std::ranges::any_of(providerPaths, [&](const std::filesystem::path& path) noexcept
        {
            return EquivalentPath(accepted.pathIdentity, accepted.path.native(), path.native());
        });
        if (selected)
        {
            subset.items.push_back(accepted);
        }
    }
    if (subset.items.empty())
    {
        return finish(S_OK);
    }

    std::vector<Candidate> current;
    current.reserve(subset.items.size());
    for (const TouchGuardItem& accepted : subset.items)
    {
        Candidate candidate{};
        const HRESULT captureHr =
            CaptureProviderObjectCandidate(fileSystem, accepted.path.native(), pluginId, instanceContext, candidate);
        if (captureHr != S_OK)
        {
            return finish(FAILED(captureHr) ? captureHr : HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH), subset.items.size());
        }
        current.push_back(std::move(candidate));
    }
    return finish(RevalidateTouchGuard(subset, current), subset.items.size());
}

HRESULT FileOperationArtifacts::ProjectProviderChildObject(IFileSystem* const fileSystem,
                                                            const std::wstring_view providerFolderPath,
                                                            const std::wstring_view childLeaf,
                                                            const std::wstring_view pluginId,
                                                            const std::wstring_view instanceContext,
                                                            Projection& out) noexcept
{
    out = {};
    if (childLeaf.empty())
    {
        return E_INVALIDARG;
    }

    std::wstring childPath;
    if (fileSystem != nullptr && ! providerFolderPath.empty() && ! pluginId.empty())
    {
        wil::com_ptr<IFileSystem> retainedFileSystem = fileSystem;
        const std::optional<FileSystemCapabilitiesV2> folderCapabilities =
            TryGetCapabilities(retainedFileSystem, providerFolderPath, FILESYSTEM_RENAME, pluginId);
        if (folderCapabilities.has_value() && HasStablePathIdentity(folderCapabilities.value()))
        {
            childPath = JoinFileSystemPath(folderCapabilities->pathIdentity.value(), providerFolderPath, childLeaf);
        }
    }
    if (childPath.empty())
    {
        childPath = (std::filesystem::path(std::wstring(providerFolderPath)) / std::wstring(childLeaf)).native();
    }
    return ProjectProviderObject(fileSystem, childPath, pluginId, instanceContext, out);
}


FileOperations::ObjectBindingResult FileOperations::CaptureReturnedObjectAuthority(
    wil::com_ptr<IFileSystemBoundObject> bound,
    std::wstring_view pathProfileId) noexcept
{
    return CaptureBoundObjectAuthorityImpl(std::move(bound), pathProfileId);
}

FileOperations::ObjectRevalidationResult FileOperations::RevalidateObjectAuthority(IFileSystem* fileSystem,
                                                                                    std::wstring_view providerPath,
                                                                                    std::wstring_view pathProfileId,
                                                                                    FileSystemBindFlags flags,
                                                                                    const BoundObjectAuthority& expected) noexcept
{
    Debug::Perf::Scope perf(L"fileops.identity.revalidate_us");
    perf.SetValue0(static_cast<uint64_t>(providerPath.size()) * sizeof(wchar_t));

    ObjectRevalidationResult result{};
    const auto finish = [&](ObjectRevalidationState state, HRESULT status) noexcept -> ObjectRevalidationResult
    {
        result.state  = state;
        result.status = status;
        perf.SetValue1(static_cast<uint64_t>(state));
        perf.SetHr(status);
        return std::move(result);
    };
    if (! expected.boundObject || expected.identity.objectId.empty())
    {
        return finish(ObjectRevalidationState::ProviderContractViolation, E_INVALIDARG);
    }

    ObjectBindingResult rebound = BindObjectAuthority(fileSystem, providerPath, pathProfileId, flags);
    switch (rebound.state)
    {
        case ObjectBindingState::Unsupported: return finish(ObjectRevalidationState::Unsupported, rebound.status);
        case ObjectBindingState::Missing: return finish(ObjectRevalidationState::Missing, rebound.status);
        case ObjectBindingState::Indeterminate: return finish(ObjectRevalidationState::Indeterminate, rebound.status);
        case ObjectBindingState::ProviderContractViolation:
            return finish(ObjectRevalidationState::ProviderContractViolation, rebound.status);
        case ObjectBindingState::Bound: break;
    }

    result.current = std::move(rebound.authority);
    bool sameObject   = false;
    bool sameRevision = false;
    const HRESULT sameHr = CrossCheckBoundObjectIdentity(expected, result.current, sameObject, sameRevision);
    if (FAILED(sameHr))
    {
        return finish(sameHr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) || sameHr == E_INVALIDARG
                          ? ObjectRevalidationState::ProviderContractViolation
                          : ObjectRevalidationState::Indeterminate,
                      sameHr);
    }
    return finish(sameObject && sameRevision && expected.kind == result.current.kind
                      ? ObjectRevalidationState::Same
                      : ObjectRevalidationState::Changed,
                  S_OK);
}

FileOperations::TransferMutationGuard FileOperations::PrepareTransferMutationGuard(
    IFileSystem* sourceFileSystem,
    IFileSystem* destinationFileSystem,
    const QualifiedEndpoint& sourceEndpoint,
    const QualifiedEndpoint& destinationEndpoint,
    TransferIntent intent,
    std::wstring_view sourcePath,
    std::wstring_view destinationPath) noexcept
{
    Debug::Perf::Scope perf(L"fileops.identity.guard_us");
    perf.SetValue0(static_cast<uint64_t>(sourcePath.size() + destinationPath.size()) * sizeof(wchar_t));

    TransferMutationGuard guard{};
    const auto finish = [&](TransferSafetyState state, HRESULT status) noexcept -> TransferMutationGuard
    {
        guard.state  = state;
        guard.status = status;
        perf.SetValue1(static_cast<uint64_t>(state));
        perf.SetHr(status);
        return std::move(guard);
    };
    if (sourceFileSystem == nullptr || sourcePath.empty() || destinationPath.empty() ||
        ! IsQualifiedEndpointValid(sourceEndpoint) || ! IsQualifiedEndpointValid(destinationEndpoint))
    {
        return finish(TransferSafetyState::ProviderContractViolation, E_INVALIDARG);
    }
    if (destinationFileSystem == nullptr)
    {
        destinationFileSystem = sourceFileSystem;
    }

    const bool sameRoot = QualifiedEndpointsMatchImpl(sourceEndpoint, destinationEndpoint);
    const bool sameObjectIdentityDomain = QualifiedEndpointsShareObjectIdentityDomainImpl(sourceEndpoint, destinationEndpoint);
    if (sameRoot)
    {
        guard.samePathText = EquivalentPath(sourceEndpoint.pathIdentity.value(), sourcePath, destinationPath);
        if (IsStrictDescendantPath(sourceEndpoint.pathIdentity.value(), sourcePath, destinationPath))
        {
            return finish(TransferSafetyState::DestinationInsideSource, HRESULT_FROM_WIN32(ERROR_INVALID_NAME));
        }
    }

    constexpr FileSystemBindFlags bindFlags =
        static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA);
    const auto bindingFailure = [](ObjectBindingState state, TransferSafetyState missingState) noexcept -> TransferSafetyState
    {
        switch (state)
        {
            case ObjectBindingState::Missing: return missingState;
            case ObjectBindingState::Unsupported: return TransferSafetyState::Unsupported;
            case ObjectBindingState::ProviderContractViolation: return TransferSafetyState::ProviderContractViolation;
            case ObjectBindingState::Indeterminate: return TransferSafetyState::Indeterminate;
            case ObjectBindingState::Bound: break;
        }
        return TransferSafetyState::Indeterminate;
    };
    const auto sameObject = [](const BoundObjectAuthority& left, const BoundObjectAuthority& right, BOOL& same) noexcept -> HRESULT
    {
        same = FALSE;
        bool snapshotSame = false;
        bool sameRevision = false;
        const HRESULT hr = CrossCheckBoundObjectIdentity(left, right, snapshotSame, sameRevision);
        same = SUCCEEDED(hr) && snapshotSame ? TRUE : FALSE;
        return hr;
    };

    ObjectBindingResult source = BindObjectAuthority(sourceFileSystem, sourcePath, sourceEndpoint.profileId, bindFlags);
    if (source.state != ObjectBindingState::Bound)
    {
        return finish(bindingFailure(source.state, TransferSafetyState::SourceMissing), source.status);
    }
    guard.source = std::move(source.authority);

    ObjectBindingResult destination =
        BindObjectAuthority(destinationFileSystem, destinationPath, destinationEndpoint.profileId, bindFlags);
    if (destination.state == ObjectBindingState::Bound)
    {
        // The final destination is a collision object, not part of the traversal ancestry.
        // Retain a no-follow authority for links so the typed conflict surface can offer only
        // link-safe actions. Links encountered while walking destination ancestors remain a
        // hard AncestryLink stop below.
        if (sameObjectIdentityDomain)
        {
            BOOL same = FALSE;
            const HRESULT sameHr = sameObject(guard.source, destination.authority, same);
            if (FAILED(sameHr))
            {
                return finish(sameHr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) || sameHr == E_INVALIDARG
                                  ? TransferSafetyState::ProviderContractViolation
                                  : TransferSafetyState::Indeterminate,
                              sameHr);
            }
            if (same != FALSE)
            {
                return finish(TransferSafetyState::SameObject, HRESULT_FROM_WIN32(ERROR_INVALID_NAME));
            }
        }
        guard.destination = RetainedPathAuthority{
            .providerPath = std::wstring(destinationPath),
            .authority    = std::move(destination.authority),
        };
    }
    else if (destination.state == ObjectBindingState::Missing)
    {
        guard.destinationWasMissing = true;
    }
    else
    {
        return finish(bindingFailure(destination.state, TransferSafetyState::DestinationChanged), destination.status);
    }

    if (sameObjectIdentityDomain && intent == TransferIntent::Move)
    {
        std::wstring sourceParent;
        std::wstring destinationParent;
        if (! TryGetFileSystemParentPath(sourceEndpoint.pathIdentity.value(), sourcePath, sourceParent) ||
            ! TryGetFileSystemParentPath(destinationEndpoint.pathIdentity.value(), destinationPath, destinationParent))
        {
            return finish(TransferSafetyState::Unsupported, HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
        }

        ObjectBindingResult boundSourceParent =
            BindObjectAuthority(sourceFileSystem, sourceParent, sourceEndpoint.profileId, bindFlags);
        ObjectBindingResult boundDestinationParent =
            BindObjectAuthority(destinationFileSystem, destinationParent, destinationEndpoint.profileId, bindFlags);
        if (boundSourceParent.state != ObjectBindingState::Bound)
        {
            return finish(bindingFailure(boundSourceParent.state, TransferSafetyState::SourceChanged), boundSourceParent.status);
        }
        if (boundSourceParent.authority.kind == FILESYSTEM_BOUND_LINK)
        {
            return finish(TransferSafetyState::AncestryLink, HRESULT_FROM_WIN32(ERROR_REPARSE_TAG_INVALID));
        }

        guard.sourceParent = RetainedPathAuthority{
            .providerPath = sourceParent,
            .authority    = std::move(boundSourceParent.authority),
        };

        if (boundDestinationParent.state == ObjectBindingState::Bound)
        {
            if (boundDestinationParent.authority.kind == FILESYSTEM_BOUND_LINK)
            {
                return finish(TransferSafetyState::AncestryLink, HRESULT_FROM_WIN32(ERROR_REPARSE_TAG_INVALID));
            }

            BOOL same = FALSE;
            const HRESULT sameHr = sameObject(guard.sourceParent->authority, boundDestinationParent.authority, same);
            if (FAILED(sameHr))
            {
                return finish(sameHr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) || sameHr == E_INVALIDARG
                                  ? TransferSafetyState::ProviderContractViolation
                                  : TransferSafetyState::Indeterminate,
                              sameHr);
            }
            if (same != FALSE)
            {
                return finish(TransferSafetyState::SameFolderMove, HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS));
            }
        }
        else if (boundDestinationParent.state != ObjectBindingState::Missing)
        {
            return finish(bindingFailure(boundDestinationParent.state, TransferSafetyState::DestinationChanged), boundDestinationParent.status);
        }
    }

    if (sameObjectIdentityDomain)
    {
        const bool sourceIsDirectory = guard.source.kind == FILESYSTEM_BOUND_DIRECTORY;
        std::wstring current;
        if (! TryGetFileSystemParentPath(destinationEndpoint.pathIdentity.value(), destinationPath, current))
        {
            return finish(TransferSafetyState::Unsupported, HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
        }
        for (size_t depth = 0u; ! current.empty() && depth < 1024u; ++depth)
        {
            ObjectBindingResult ancestor =
                BindObjectAuthority(destinationFileSystem, current, destinationEndpoint.profileId, bindFlags);
            if (ancestor.state == ObjectBindingState::Bound)
            {
                if (ancestor.authority.kind == FILESYSTEM_BOUND_LINK)
                {
                    return finish(TransferSafetyState::AncestryLink, HRESULT_FROM_WIN32(ERROR_REPARSE_TAG_INVALID));
                }
                if (sourceIsDirectory)
                {
                    BOOL same = FALSE;
                    const HRESULT sameHr = sameObject(guard.source, ancestor.authority, same);
                    if (FAILED(sameHr))
                    {
                        return finish(sameHr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) || sameHr == E_INVALIDARG
                                          ? TransferSafetyState::ProviderContractViolation
                                          : TransferSafetyState::Indeterminate,
                                      sameHr);
                    }
                    if (same != FALSE)
                    {
                        return finish(TransferSafetyState::DestinationInsideSource, HRESULT_FROM_WIN32(ERROR_INVALID_NAME));
                    }
                }
                guard.destinationAncestors.emplace_back(RetainedPathAuthority{
                    .providerPath = current,
                    .authority    = std::move(ancestor.authority),
                });
            }
            else if (ancestor.state != ObjectBindingState::Missing)
            {
                return finish(bindingFailure(ancestor.state, TransferSafetyState::DestinationChanged), ancestor.status);
            }

            std::wstring parent;
            if (! TryGetFileSystemParentPath(destinationEndpoint.pathIdentity.value(), current, parent) ||
                EquivalentPath(destinationEndpoint.pathIdentity.value(), parent, current))
            {
                current.clear();
                break;
            }
            current = std::move(parent);
        }
        if (! current.empty())
        {
            return finish(TransferSafetyState::Unsupported, HRESULT_FROM_WIN32(ERROR_FILENAME_EXCED_RANGE));
        }
    }

    return finish(TransferSafetyState::Ready, S_OK);
}

FileOperations::TransferSafetyState FileOperations::RevalidateTransferMutationGuard(
    IFileSystem* sourceFileSystem,
    IFileSystem* destinationFileSystem,
    const QualifiedEndpoint& sourceEndpoint,
    const QualifiedEndpoint& destinationEndpoint,
    TransferIntent intent,
    std::wstring_view sourcePath,
    std::wstring_view destinationPath,
    const TransferMutationGuard& guard,
    HRESULT& status,
    bool allowMissingDestinationDirectoryMerge) noexcept
{
    Debug::Perf::Scope perf(L"fileops.identity.guard_revalidate_us");
    status = E_UNEXPECTED;
    const auto finish = [&](TransferSafetyState state, HRESULT hr) noexcept -> TransferSafetyState
    {
        status = hr;
        perf.SetValue0(static_cast<uint64_t>(state));
        perf.SetHr(hr);
        return state;
    };
    if (guard.state != TransferSafetyState::Ready || ! guard.source.boundObject || sourceFileSystem == nullptr)
    {
        return finish(TransferSafetyState::ProviderContractViolation, E_INVALIDARG);
    }
    if (destinationFileSystem == nullptr)
    {
        destinationFileSystem = sourceFileSystem;
    }
    constexpr FileSystemBindFlags bindFlags =
        static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA);
    const auto revalidationFailure = [](ObjectRevalidationState state, bool source) noexcept -> TransferSafetyState
    {
        switch (state)
        {
            case ObjectRevalidationState::Changed:
            case ObjectRevalidationState::Missing: return source ? TransferSafetyState::SourceChanged : TransferSafetyState::DestinationChanged;
            case ObjectRevalidationState::Unsupported: return TransferSafetyState::Unsupported;
            case ObjectRevalidationState::ProviderContractViolation: return TransferSafetyState::ProviderContractViolation;
            case ObjectRevalidationState::Indeterminate: return TransferSafetyState::Indeterminate;
            case ObjectRevalidationState::Same: break;
        }
        return TransferSafetyState::Indeterminate;
    };
    const auto revalidationFailureStatus = [](const ObjectRevalidationResult& result) noexcept -> HRESULT
    {
        return FAILED(result.status) ? result.status : HRESULT_FROM_WIN32(ERROR_RETRY);
    };

    ObjectRevalidationResult source = RevalidateObjectAuthority(
        sourceFileSystem, sourcePath, sourceEndpoint.profileId, bindFlags, guard.source);
    if (source.state != ObjectRevalidationState::Same)
    {
        return finish(revalidationFailure(source.state, true), revalidationFailureStatus(source));
    }
    if (guard.sourceParent.has_value())
    {
        ObjectRevalidationResult parent = RevalidateObjectAuthority(sourceFileSystem,
                                                                    guard.sourceParent->providerPath,
                                                                    sourceEndpoint.profileId,
                                                                    bindFlags,
                                                                    guard.sourceParent->authority);
        if (parent.state != ObjectRevalidationState::Same)
        {
            return finish(revalidationFailure(parent.state, true), revalidationFailureStatus(parent));
        }
    }
    if (guard.destination.has_value())
    {
        ObjectRevalidationResult destination = RevalidateObjectAuthority(destinationFileSystem,
                                                                         destinationPath,
                                                                         destinationEndpoint.profileId,
                                                                         bindFlags,
                                                                         guard.destination->authority);
        if (destination.state != ObjectRevalidationState::Same)
        {
            return finish(revalidationFailure(destination.state, false), revalidationFailureStatus(destination));
        }
    }
    for (const RetainedPathAuthority& ancestor : guard.destinationAncestors)
    {
        ObjectRevalidationResult current = RevalidateObjectAuthority(destinationFileSystem,
                                                                     ancestor.providerPath,
                                                                     destinationEndpoint.profileId,
                                                                     bindFlags,
                                                                     ancestor.authority);
        if (current.state != ObjectRevalidationState::Same)
        {
            return finish(revalidationFailure(current.state, false), revalidationFailureStatus(current));
        }
    }

    TransferMutationGuard current = PrepareTransferMutationGuard(sourceFileSystem,
                                                                 destinationFileSystem,
                                                                 sourceEndpoint,
                                                                 destinationEndpoint,
                                                                 intent,
                                                                 sourcePath,
                                                                 destinationPath);
    if (current.state != TransferSafetyState::Ready)
    {
        return finish(current.state, current.status);
    }
    if (guard.destinationWasMissing && ! current.destinationWasMissing)
    {
        // A resolved DirectoryShell is an idempotent folder-merge instruction. Another item in
        // the same plan may have created that ordinary directory while ensuring its parent.
        // This exception never admits a link or a non-directory and never applies to a file
        // publication, so external destination replacement remains fail-closed everywhere else.
        if (allowMissingDestinationDirectoryMerge && current.destination.has_value() &&
            current.destination->authority.kind == FILESYSTEM_BOUND_DIRECTORY)
        {
            return finish(TransferSafetyState::Ready, S_OK);
        }
        return finish(TransferSafetyState::DestinationChanged, HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS));
    }
    return finish(TransferSafetyState::Ready, S_OK);
}

#ifdef ENABLE_TESTS
FileOperations::ObjectBindingResult FileOperations::ValidateSuccessfulBoundObjectForSelfTest(
    IFileSystemBoundObject* bound,
    std::wstring_view pathProfileId) noexcept
{
    wil::com_ptr<IFileSystemBoundObject> retained;
    retained = bound;
    return CaptureReturnedObjectAuthority(std::move(retained), pathProfileId);
}

bool FileOperations::TryQualifyEndpointForSelfTest(const wil::com_ptr<IFileSystem>& fileSystem,
                                                    std::wstring_view providerPath,
                                                    FileSystemOperation operation,
                                                    std::wstring_view pluginId,
                                                    std::wstring_view instanceId,
                                                    QualifiedEndpoint& endpoint) noexcept
{
    endpoint = {};
    if (pluginId.empty() || instanceId.empty())
    {
        return false;
    }
    const std::optional<FileSystemCapabilitiesV2> capabilities =
        TryGetCapabilities(fileSystem, providerPath, operation, pluginId);
    if (! capabilities.has_value() || ! HasStablePathIdentity(capabilities.value()))
    {
        return false;
    }
    endpoint = QualifiedEndpoint{
        .pluginId   = std::wstring(pluginId),
        .instanceId = std::wstring(instanceId),
        .profileId  = capabilities->pathProfileId,
        .rootId     = capabilities->rootId,
        .pathIdentity = capabilities->pathIdentity,
        .verificationHostReadback = capabilities->verificationHostReadback,
        .verificationProviderBlake3Proof = capabilities->verificationProviderBlake3Proof,
        .verificationWriterDigestProof = capabilities->verificationWriterDigestProof,
        .verificationCapabilityCheckDeferred = capabilities->verificationCapabilityCheckDeferred,
        .cancellationAbort = capabilities->cancellationAbort,
        .cancellationDeadline = capabilities->cancellationDeadline,
        .cancellationRouteClass = capabilities->cancellationRouteClass,
        .providerWatchdogTimeoutMs = capabilities->providerWatchdogTimeoutMs,
    };
    return IsQualifiedEndpointValid(endpoint);
}

bool CanCrossFileSystemCopyMoveForSelfTest(const wil::com_ptr<IFileSystem>& sourceFileSystem,
                                           std::wstring_view sourceProviderPath,
                                           std::wstring_view sourcePluginId,
                                           const wil::com_ptr<IFileSystem>& destinationFileSystem,
                                           std::wstring_view destinationProviderPath,
                                           std::wstring_view destinationPluginId,
                                           FileSystemOperation operation) noexcept
{
    return CanCrossFileSystemCopyMove(sourceFileSystem,
                                      sourceProviderPath,
                                      sourcePluginId,
                                      destinationFileSystem,
                                      destinationProviderPath,
                                      destinationPluginId,
                                      operation);
}
#endif

bool CanSameFileSystemOperation(const wil::com_ptr<IFileSystem>& fileSystem,
                                std::wstring_view providerPath,
                                FileSystemOperation operation,
                                std::wstring_view pluginId,
                                FileSystemFlags flags) noexcept
{
    return CanSameFileSystemOperationFromCapabilities(fileSystem, providerPath, operation, pluginId, flags);
}

void FolderWindow::FileOperationStateDeleter::operator()(FileOperationState* state) const noexcept
{
    std::default_delete<FileOperationState>{}(state);
}

FolderWindow::FileOperationPromptDispatchScope::FileOperationPromptDispatchScope(FolderWindow& owner) noexcept
    : _owner(owner)
{
    ++_owner._fileOperationPromptDispatchDepth;
}

FolderWindow::FileOperationPromptDispatchScope::~FileOperationPromptDispatchScope() noexcept
{
    --_owner._fileOperationPromptDispatchDepth;
    if (_owner._fileOperationPromptDispatchDepth == 0u && _owner._fileOperationResetDeferred)
    {
        _owner._fileOperationResetDeferred = false;
        _owner._fileOperations.reset();
    }
}

HRESULT FolderWindow::ConfirmExternalArtifactTouch(
    const Pane pane,
    const std::span<const std::filesystem::path> providerPaths,
    FileOperationArtifacts::TouchGuardReceipt* const receiptOut) noexcept
{
    if (providerPaths.empty())
    {
        return E_INVALIDARG;
    }
    EnsureFileOperations();
    if (! _fileOperations)
    {
        return E_UNEXPECTED;
    }
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! state.fileSystem || state.pluginId.empty())
    {
        return E_UNEXPECTED;
    }
    return ConfirmExternalArtifactTouchForProvider(
        state.fileSystem.get(), state.pluginId, state.instanceContext, providerPaths, receiptOut);
}

HRESULT FolderWindow::ConfirmExternalArtifactTouchForProvider(
    IFileSystem* fileSystem,
    const std::wstring_view pluginId,
    const std::wstring_view instanceContext,
    const std::span<const std::filesystem::path> providerPaths,
    FileOperationArtifacts::TouchGuardReceipt* const receiptOut) noexcept
{
    if (! _fileOperations || ! fileSystem || pluginId.empty() || providerPaths.empty())
    {
        return E_INVALIDARG;
    }
    const FileOperationPromptDispatchScope promptDispatch(*this);
    return _fileOperations->ConfirmExternalArtifactTouch(fileSystem, pluginId, instanceContext, providerPaths, receiptOut);
}


void FolderWindow::EnsureFileOperations()
{
    if (_fileOperations)
    {
        return;
    }

    auto state = std::make_unique<FileOperationState>(*this);
    _fileOperations.reset(state.release());
}

uint64_t FolderWindow::CreateOrUpdateInformationalTask(const InformationalTaskUpdate& update) noexcept
{
    EnsureFileOperations();
    if (! _fileOperations)
    {
        return 0;
    }

    return _fileOperations->CreateOrUpdateInformationalTask(update);
}

void FolderWindow::DismissInformationalTask(uint64_t taskId) noexcept
{
    if (taskId == 0 || ! _fileOperations)
    {
        return;
    }

    _fileOperations->DismissInformationalTask(taskId);
}

HRESULT FolderWindow::FileOperationState::AdmitOperation(FileSystemOperation operation,
                                                         FolderWindow::Pane sourcePane,
                                                         std::optional<FolderWindow::Pane> destinationPane,
                                                         const wil::com_ptr<IFileSystem>& fileSystem,
                                                         std::vector<std::filesystem::path> sourcePaths,
                                                         std::filesystem::path destinationFolder,
                                                         FileSystemFlags flags,
                                                         bool waitForOthers,
                                                         uint64_t initialSpeedLimitBytesPerSecond,
                                                         ExecutionMode executionMode,
                                                         bool requireConfirmation,
                                                         wil::com_ptr<IFileSystem> destinationFileSystem,
                                                         uint64_t* taskIdOut,
                                                         std::vector<FolderWindow::ResolvedFileOperationItem> resolvedItems,
                                                         std::wstring confirmationMessage,
                                                         std::wstring sourcePluginIdOverride,
                                                         std::wstring sourcePluginShortIdOverride,
                                                         std::optional<std::wstring> sourceInstanceIdOverride,
                                                          FileOperations::DeleteOrigin deleteOrigin,
                                                          std::optional<FileOperations::ConsentKind> capturedConsentKind,
                                                          std::optional<uint32_t> moveClipboardSequence,
                                                          std::function<HRESULT()> preWorkerReleaseBarrier,
                                                          std::function<HRESULT()> preConsumptionDecisionGate,
                                                          std::function<void(std::shared_ptr<const FileOperations::PreparationSnapshot>)> preparationObserver)
{
    const auto startedAt = std::chrono::steady_clock::now();
    if (taskIdOut)
    {
        *taskIdOut = 0u;
    }

    const auto reject = [&](HRESULT hr, FileOperations::PlanRejectionBucket bucket) noexcept -> HRESULT
    {
        Debug::Perf::Emit(L"fileops.plan.construct_us",
                          PlanRejectionBucketToString(bucket),
                          Debug::Perf::ElapsedUs(startedAt),
                          static_cast<uint64_t>(sourcePaths.size()),
                          0u,
                          hr);
        Debug::Perf::EmitValue(L"fileops.plan.rejection_bucket", static_cast<uint64_t>(bucket), hr);
        return hr;
    };

    if (! fileSystem)
    {
        return reject(E_POINTER, FileOperations::PlanRejectionBucket::MissingFileSystem);
    }
    if (sourcePaths.empty())
    {
        return reject(S_FALSE, FileOperations::PlanRejectionBucket::EmptySelection);
    }
    if (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE && operation != FILESYSTEM_DELETE)
    {
        return reject(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), FileOperations::PlanRejectionBucket::UnsupportedOperation);
    }
    if (capturedConsentKind.has_value() &&
        (operation != FILESYSTEM_DELETE || capturedConsentKind.value() != FileOperations::ConsentKind::ArchiveDeleteAfter ||
         (deleteOrigin != FileOperations::DeleteOrigin::PackCleanup && deleteOrigin != FileOperations::DeleteOrigin::UnpackCleanup)))
    {
        return reject(E_INVALIDARG, FileOperations::PlanRejectionBucket::MissingDestructiveConsent);
    }
    const FolderWindow::PaneState& sourceState =
        sourcePane == FolderWindow::Pane::Left ? _owner._leftPane : _owner._rightPane;
    const bool sourceUsesPaneFileSystem = fileSystem.get() == sourceState.fileSystem.get();
    std::wstring sourcePluginId = std::move(sourcePluginIdOverride);
    std::wstring sourcePluginShortId = std::move(sourcePluginShortIdOverride);
    if (sourcePluginId.empty())
    {
        if (sourceUsesPaneFileSystem)
        {
            sourcePluginId      = sourceState.pluginId;
            sourcePluginShortId = sourceState.pluginShortId;
        }
        else if (! TryGetFileSystemPluginIdentity(fileSystem.get(), sourcePluginId, sourcePluginShortId))
        {
            return reject(E_INVALIDARG, FileOperations::PlanRejectionBucket::MalformedEndpoint);
        }
    }
    else if (sourcePluginShortId.empty() && sourceUsesPaneFileSystem)
    {
        sourcePluginShortId = sourceState.pluginShortId;
    }
    const std::wstring sourceInstanceContext = sourceInstanceIdOverride.has_value()
        ? sourceInstanceIdOverride.value()
        : (sourceUsesPaneFileSystem ? sourceState.instanceContext : std::wstring{});
    const std::wstring sourceInstanceId = CanonicalInstanceId(sourceInstanceContext);
    if (sourcePluginId.empty())
    {
        return reject(E_INVALIDARG, FileOperations::PlanRejectionBucket::MalformedEndpoint);
    }

    struct SourcePartition final
    {
        FileOperations::QualifiedEndpoint endpoint;
        std::vector<size_t> originalIndices;
        std::vector<FileSystemCapabilitiesV2> capabilities;
    };
    std::vector<SourcePartition> sourcePartitions;
    sourcePartitions.reserve(sourcePaths.size());
    for (size_t index = 0; index < sourcePaths.size(); ++index)
    {
        const std::optional<FileSystemCapabilitiesV2> capabilities =
            TryGetCapabilities(fileSystem, sourcePaths[index].native(), operation, sourcePluginId);
        if (! capabilities.has_value())
        {
            return reject(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), FileOperations::PlanRejectionBucket::MalformedEndpoint);
        }
        if (! capabilities->pathIdentity.has_value() || ! HasStablePathIdentity(capabilities.value()) ||
            ! HasContainedCancellationRoute(capabilities.value()))
        {
            return reject(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), FileOperations::PlanRejectionBucket::MalformedEndpoint);
        }

        const auto matching = std::ranges::find_if(sourcePartitions, [&](const SourcePartition& partition) noexcept
        {
            return partition.endpoint.profileId == capabilities->pathProfileId && partition.endpoint.rootId == capabilities->rootId;
        });
        if (matching != sourcePartitions.end())
        {
            if (! matching->endpoint.pathIdentity.has_value() ||
                ! SamePathIdentityContract(matching->endpoint.pathIdentity.value(), capabilities->pathIdentity.value()))
            {
                return reject(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), FileOperations::PlanRejectionBucket::MalformedEndpoint);
            }
            matching->originalIndices.push_back(index);
            matching->capabilities.push_back(capabilities.value());
            continue;
        }

        SourcePartition partition{};
        partition.endpoint = FileOperations::QualifiedEndpoint{
            .pluginId   = sourcePluginId,
            .instanceId = sourceInstanceId,
            .profileId  = capabilities->pathProfileId,
            .rootId     = capabilities->rootId,
            .pathIdentity = capabilities->pathIdentity,
            .verificationHostReadback = capabilities->verificationHostReadback,
            .verificationProviderBlake3Proof = capabilities->verificationProviderBlake3Proof,
            .verificationWriterDigestProof = capabilities->verificationWriterDigestProof,
            .verificationCapabilityCheckDeferred = capabilities->verificationCapabilityCheckDeferred,
            .cancellationAbort = capabilities->cancellationAbort,
            .cancellationDeadline = capabilities->cancellationDeadline,
            .cancellationRouteClass = capabilities->cancellationRouteClass,
            .providerWatchdogTimeoutMs = capabilities->providerWatchdogTimeoutMs,
        };
        partition.originalIndices.push_back(index);
        partition.capabilities.push_back(capabilities.value());
        sourcePartitions.emplace_back(std::move(partition));
    }

    FileOperations::OperationOptions options{};
    if (_owner._settings)
    {
        if (operation == FILESYSTEM_COPY)
        {
            // Preserve/Skip is a Copy option. A Move relocates every link object as stored.
            const auto policy = FolderWindowFileOperationsStateInternal::GetReparsePointPolicyFromSettings(*_owner._settings, sourcePluginId);
            options.linkPolicy = policy == FolderWindowFileOperationsStateInternal::ReparsePointPolicy::Preserve
                ? FileOperations::LinkPolicy::Preserve
                : FileOperations::LinkPolicy::Skip;
        }
        if (_owner._settings->fileOperations.has_value())
        {
            options.verifyAfterCopy = _owner._settings->fileOperations->verifyAfterCopy;
        }
    }
    options.executionMode = waitForOthers ? FileOperations::ExecutionMode::Queue : FileOperations::ExecutionMode::Parallel;
    const uint64_t effectiveBandwidthLimit = initialSpeedLimitBytesPerSecond == 0u
        ? FolderWindowFileOperationsStateInternal::GetDefaultBandwidthLimitBytesPerSecondFromSettings(_owner._settings)
        : initialSpeedLimitBytesPerSecond;
    if (effectiveBandwidthLimit != 0u)
    {
        options.bandwidthLimitBytesPerSecond = effectiveBandwidthLimit;
    }

    if (! resolvedItems.empty() && resolvedItems.size() != sourcePaths.size())
    {
        return reject(E_INVALIDARG, FileOperations::PlanRejectionBucket::InvalidExplicitMappings);
    }
    if (moveClipboardSequence.has_value())
    {
        if (operation != FILESYSTEM_MOVE || moveClipboardSequence.value() == 0u || ! preWorkerReleaseBarrier)
        {
            return reject(E_INVALIDARG, FileOperations::PlanRejectionBucket::InvalidClipboardSequence);
        }
    }
    else if (preWorkerReleaseBarrier)
    {
        return reject(E_INVALIDARG, FileOperations::PlanRejectionBucket::InvalidClipboardSequence);
    }

    FileOperations::FileOperationPlanGroup plans;
    plans.reserve(sourcePaths.size());
    std::vector<std::filesystem::path> groupedSourcePaths;
    groupedSourcePaths.reserve(sourcePaths.size());
    std::vector<FolderWindow::ResolvedFileOperationItem> groupedResolvedItems;
    groupedResolvedItems.reserve(resolvedItems.size());
    std::wstring destinationPluginShortId = sourcePluginShortId;
    std::wstring destinationInstanceContext = sourceInstanceContext;
    if (operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE)
    {
        if (destinationFolder.empty())
        {
            return reject(E_INVALIDARG, FileOperations::PlanRejectionBucket::MalformedDestination);
        }

        const wil::com_ptr<IFileSystem>& effectiveDestinationFileSystem = destinationFileSystem ? destinationFileSystem : fileSystem;
        std::wstring destinationPluginId = sourcePluginId;
        std::wstring destinationInstanceId = sourceInstanceId;
        const FolderWindow::PaneState* destinationState = nullptr;
        if (destinationPane.has_value())
        {
            destinationState =
                destinationPane.value() == FolderWindow::Pane::Left ? &_owner._leftPane : &_owner._rightPane;
            if (effectiveDestinationFileSystem.get() == destinationState->fileSystem.get())
            {
                destinationPluginId          = destinationState->pluginId;
                destinationPluginShortId     = destinationState->pluginShortId;
                destinationInstanceContext   = destinationState->instanceContext;
                destinationInstanceId        = CanonicalInstanceId(destinationState->instanceContext);
            }
        }
        const bool destinationUsesPaneFileSystem =
            destinationState != nullptr && effectiveDestinationFileSystem.get() == destinationState->fileSystem.get();
        if (destinationFileSystem && ! destinationUsesPaneFileSystem)
        {
            if (! TryGetFileSystemPluginIdentity(effectiveDestinationFileSystem.get(), destinationPluginId, destinationPluginShortId))
            {
                return reject(E_INVALIDARG, FileOperations::PlanRejectionBucket::MalformedEndpoint);
            }
            destinationInstanceContext.clear();
            destinationInstanceId = CanonicalInstanceId(std::wstring_view{});
        }
        if (destinationPluginId.empty())
        {
            return reject(E_INVALIDARG, FileOperations::PlanRejectionBucket::MalformedEndpoint);
        }

        const std::optional<FileSystemCapabilitiesV2> destinationCapabilities =
            TryGetCapabilities(effectiveDestinationFileSystem, destinationFolder.native(), operation, destinationPluginId);
        if (! destinationCapabilities.has_value())
        {
            return reject(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), FileOperations::PlanRejectionBucket::MalformedEndpoint);
        }
        if (! destinationCapabilities->pathIdentity.has_value() || ! HasStablePathIdentity(destinationCapabilities.value()) ||
            ! HasContainedCancellationRoute(destinationCapabilities.value()))
        {
            return reject(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), FileOperations::PlanRejectionBucket::MalformedEndpoint);
        }

        const FileOperations::QualifiedEndpoint destinationEndpoint{
            .pluginId   = destinationPluginId,
            .instanceId = std::move(destinationInstanceId),
            .profileId  = destinationCapabilities->pathProfileId,
            .rootId     = destinationCapabilities->rootId,
            .pathIdentity = destinationCapabilities->pathIdentity,
            .verificationHostReadback = destinationCapabilities->verificationHostReadback,
            .verificationProviderBlake3Proof = destinationCapabilities->verificationProviderBlake3Proof,
            .verificationWriterDigestProof = destinationCapabilities->verificationWriterDigestProof,
            .verificationCapabilityCheckDeferred = destinationCapabilities->verificationCapabilityCheckDeferred,
            .cancellationAbort = destinationCapabilities->cancellationAbort,
            .cancellationDeadline = destinationCapabilities->cancellationDeadline,
            .cancellationRouteClass = destinationCapabilities->cancellationRouteClass,
            .providerWatchdogTimeoutMs = destinationCapabilities->providerWatchdogTimeoutMs,
        };
        wil::com_ptr<IFileSystemObjectBinding> sourceBinding;
        wil::com_ptr<IFileSystemObjectBinding> destinationBinding;
        if (operation == FILESYSTEM_MOVE)
        {
            static_cast<void>(fileSystem->QueryInterface(__uuidof(IFileSystemObjectBinding), sourceBinding.put_void()));
            static_cast<void>(effectiveDestinationFileSystem->QueryInterface(
                __uuidof(IFileSystemObjectBinding), destinationBinding.put_void()));
        }


        for (const SourcePartition& partition : sourcePartitions)
        {
            const bool sameEndpoint = QualifiedEndpointsMatchImpl(partition.endpoint, destinationEndpoint);
            const bool providerNativeMoveAdvertised = operation == FILESYSTEM_MOVE && sameEndpoint &&
                destinationCapabilities->moveOperation && destinationCapabilities->nativeMoveOperation &&
                std::ranges::all_of(partition.capabilities, [](const FileSystemCapabilitiesV2& sourceCapabilities) noexcept
                {
                    return sourceCapabilities.moveOperation && sourceCapabilities.nativeMoveOperation;
                });
            const bool nativeMoveQualified = providerNativeMoveAdvertised;
            std::optional<FileSystemCapabilitiesV2> destinationCopyCapabilities;
            std::vector<FileSystemCapabilitiesV2> sourceCopyCapabilities;
            if (operation == FILESYSTEM_MOVE && ! nativeMoveQualified)
            {
                destinationCopyCapabilities =
                    TryGetCapabilities(effectiveDestinationFileSystem, destinationFolder.native(), FILESYSTEM_COPY, destinationPluginId);
                if (! destinationCopyCapabilities.has_value() ||
                    destinationCopyCapabilities->pathProfileId != destinationCapabilities->pathProfileId ||
                    destinationCopyCapabilities->rootId != destinationCapabilities->rootId ||
                    ! destinationCopyCapabilities->pathIdentity.has_value() ||
                    ! SamePathIdentityContract(destinationCopyCapabilities->pathIdentity.value(),
                                               destinationCapabilities->pathIdentity.value()))
                {
                    return reject(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), FileOperations::PlanRejectionBucket::MalformedEndpoint);
                }

                sourceCopyCapabilities.reserve(partition.capabilities.size());
                for (size_t capabilityIndex = 0u; capabilityIndex < partition.capabilities.size(); ++capabilityIndex)
                {
                    const size_t originalIndex = partition.originalIndices[capabilityIndex];
                    std::optional<FileSystemCapabilitiesV2> sourceCopy =
                        TryGetCapabilities(fileSystem, sourcePaths[originalIndex].native(), FILESYSTEM_COPY, sourcePluginId);
                    if (! sourceCopy.has_value() ||
                        sourceCopy->pathProfileId != partition.capabilities[capabilityIndex].pathProfileId ||
                        sourceCopy->rootId != partition.capabilities[capabilityIndex].rootId ||
                        ! sourceCopy->pathIdentity.has_value() ||
                        ! SamePathIdentityContract(sourceCopy->pathIdentity.value(),
                                                   partition.capabilities[capabilityIndex].pathIdentity.value()))
                    {
                        return reject(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), FileOperations::PlanRejectionBucket::MalformedEndpoint);
                    }
                    sourceCopyCapabilities.emplace_back(std::move(sourceCopy.value()));
                }
            }

            bool copyPairQualified = operation == FILESYSTEM_COPY;
            if (destinationCopyCapabilities.has_value() || operation == FILESYSTEM_COPY)
            {
                const FileSystemCapabilitiesV2& effectiveDestinationCapabilities = operation == FILESYSTEM_MOVE
                    ? destinationCopyCapabilities.value()
                    : destinationCapabilities.value();
                copyPairQualified = true;
                for (size_t capabilityIndex = 0u; capabilityIndex < partition.capabilities.size(); ++capabilityIndex)
                {
                    const FileSystemCapabilitiesV2& sourceCapabilities = operation == FILESYSTEM_MOVE
                        ? sourceCopyCapabilities[capabilityIndex]
                        : partition.capabilities[capabilityIndex];
                    const bool itemCopyQualified = sameEndpoint
                        ? CapabilitiesAllowSameFileSystemOperation(sourceCapabilities, FILESYSTEM_COPY) &&
                              CapabilitiesAllowSameFileSystemOperation(effectiveDestinationCapabilities, FILESYSTEM_COPY)
                        : CapabilityPairAllowsCrossFileSystemCopyMove(sourceCapabilities,
                                                                      sourcePluginId,
                                                                      effectiveDestinationCapabilities,
                                                                      destinationPluginId,
                                                                      FILESYSTEM_COPY);
                    if (! itemCopyQualified)
                    {
                        copyPairQualified = false;
                        break;
                    }
                }
            }

            FileOperations::OperationStrategy strategy = FileOperations::OperationStrategy::Copy;
            if (operation == FILESYSTEM_MOVE)
            {
                const bool movePairQualified = sameEndpoint
                    ? destinationCapabilities->moveOperation &&
                          std::ranges::all_of(partition.capabilities, [](const FileSystemCapabilitiesV2& capabilities) noexcept
                          { return capabilities.moveOperation; })
                    : std::ranges::all_of(partition.capabilities, [&](const FileSystemCapabilitiesV2& capabilities) noexcept
                      {
                          return CapabilityPairAllowsCrossFileSystemCopyMove(capabilities,
                                                                             sourcePluginId,
                                                                             destinationCapabilities.value(),
                                                                             destinationPluginId,
                                                                             FILESYSTEM_MOVE);
                      });
                const bool sourceBoundDelete = ! sourceCopyCapabilities.empty() &&
                    std::ranges::all_of(partition.capabilities, [](const FileSystemCapabilitiesV2& capabilities) noexcept
                    { return capabilities.boundDelete; }) &&
                    std::ranges::all_of(sourceCopyCapabilities, [](const FileSystemCapabilitiesV2& capabilities) noexcept
                    { return capabilities.boundDelete; });
                const bool sourceConditionalDelete = ! sourceCopyCapabilities.empty() &&
                    std::ranges::all_of(partition.capabilities, [](const FileSystemCapabilitiesV2& capabilities) noexcept
                    { return capabilities.conditionalDelete; }) &&
                    std::ranges::all_of(sourceCopyCapabilities, [](const FileSystemCapabilitiesV2& capabilities) noexcept
                    { return capabilities.conditionalDelete; });
                const bool destinationExclusiveStage = destinationCopyCapabilities.has_value() &&
                    destinationCapabilities->exclusiveStage && destinationCopyCapabilities->exclusiveStage;
                const bool destinationConditionalPublish = destinationCopyCapabilities.has_value() &&
                    destinationCapabilities->conditionalPublish && destinationCopyCapabilities->conditionalPublish;
                const FileOperations::MoveStrategyQualificationFacts facts{
                    .nativeMoveQualified = nativeMoveQualified,
                    .copyPairQualified = copyPairQualified,
                    .movePairQualified = movePairQualified,
                    .sourceBoundDelete = sourceBoundDelete,
                    .sourceConditionalDelete = sourceConditionalDelete,
                    .destinationExclusiveStage = destinationExclusiveStage,
                    .destinationConditionalPublish = destinationConditionalPublish,
                    .sourceBindingAvailable = static_cast<bool>(sourceBinding),
                    .destinationBindingAvailable = static_cast<bool>(destinationBinding),
                    .destinationWriterDigestProof = destinationCapabilities->verificationWriterDigestProof,
                };
                const std::optional<FileOperations::OperationStrategy> selected = FileOperations::SelectMoveStrategy(facts);
                if (! selected.has_value())
                {
                    return reject(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                                  FileOperations::PlanRejectionBucket::UnsupportedOperation);
                }
                strategy = selected.value();
            }
            else if (! copyPairQualified)
            {
                return reject(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                              FileOperations::PlanRejectionBucket::UnsupportedOperation);
            }

            FileOperations::TransferPlan transfer{};
            transfer.intent = operation == FILESYSTEM_MOVE ? FileOperations::TransferIntent::Move : FileOperations::TransferIntent::Copy;
            transfer.strategy            = strategy;
            transfer.sourceEndpoint      = partition.endpoint;
            transfer.destinationEndpoint = destinationEndpoint;
            transfer.destination.providerFolderPath = destinationFolder.native();
            transfer.options = options;
            transfer.selectedItems.reserve(partition.originalIndices.size());
            if (! resolvedItems.empty())
            {
                transfer.explicitMappings.reserve(partition.originalIndices.size());
            }
            for (size_t childIndex = 0; childIndex < partition.originalIndices.size(); ++childIndex)
            {
                const size_t originalIndex = partition.originalIndices[childIndex];
                const std::filesystem::path& sourcePath = sourcePaths[originalIndex];
                transfer.selectedItems.emplace_back(FileOperations::QualifiedSourceItem{.providerPath = sourcePath.native()});
                groupedSourcePaths.push_back(sourcePath);
                std::wstring destinationLeaf;
                if (! resolvedItems.empty())
                {
                    if (resolvedItems[originalIndex].sourcePath != sourcePath || resolvedItems[originalIndex].destinationPath.empty())
                    {
                        return reject(E_INVALIDARG, FileOperations::PlanRejectionBucket::InvalidExplicitMappings);
                    }
                    transfer.explicitMappings.emplace_back(FileOperations::TransferDestinationMapping{
                        .sourceIndex = childIndex,
                        .destinationProviderPath = resolvedItems[originalIndex].destinationPath.native(),
                    });
                    groupedResolvedItems.push_back(resolvedItems[originalIndex]);
                    if (! TryGetFileSystemLeafName(destinationEndpoint.pathIdentity.value(),
                                                   resolvedItems[originalIndex].destinationPath.native(),
                                                   destinationLeaf))
                    {
                        return reject(E_INVALIDARG, FileOperations::PlanRejectionBucket::InvalidExplicitMappings);
                    }
                }
                else if (! TryGetFileSystemLeafName(partition.endpoint.pathIdentity.value(), sourcePath.native(), destinationLeaf))
                {
                    return reject(E_INVALIDARG, FileOperations::PlanRejectionBucket::MalformedSource);
                }
                // R0-RC3 (9) / R1d-OR1: the destination provider's child-name contract admits every
                // leaf before any path is joined, but that is a provider call, so it runs in
                // Preparing on the task thread (`Task::PrepareTransferDestinationNames`) under the
                // cancel watch, not here on the UI thread. Admission only proves the leaf has a shape.
            }
            if (moveClipboardSequence.has_value())
            {
                transfer.moveClipboardSequence = FileOperations::ClipboardSequence{.windowsSequenceNumber = moveClipboardSequence.value()};
            }
            plans.emplace_back(std::move(transfer));
        }
    }
    else
    {
        for (const SourcePartition& partition : sourcePartitions)
        {
            if (std::ranges::any_of(partition.capabilities, [&](const FileSystemCapabilitiesV2& capabilities) noexcept
                {
                    return ! CapabilitiesAllowSameFileSystemOperation(capabilities, operation, flags);
                }))
            {
                return reject(HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                              FileOperations::PlanRejectionBucket::UnsupportedOperation);
            }

            FileOperations::DeletePlan deletion{};
            deletion.endpoint = partition.endpoint;
            deletion.mode = (flags & FILESYSTEM_FLAG_USE_RECYCLE_BIN) != 0 ? FileOperations::DeleteMode::Recycle
                                                                           : FileOperations::DeleteMode::Permanent;
            if (deletion.mode == FileOperations::DeleteMode::Permanent)
            {
                // C1: no provider call here. A provider without bound objects runs its
                // contract-tested native delete, which consumes the selected path and returns its
                // own receipt (R0f). Every other root is pinned by Preparing on the task thread, and
                // the confirmation follows on the card once it is.
                HRESULT bindingStatus = S_OK;
                switch (QueryObjectBindingSupport(fileSystem.get(), bindingStatus))
                {
                    case ObjectBindingSupport::Unsupported: deletion.nativeAuthority = true; break;
                    case ObjectBindingSupport::Failed:
                        return reject(FAILED(bindingStatus) ? bindingStatus : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED),
                                      FileOperations::PlanRejectionBucket::MalformedSource);
                    case ObjectBindingSupport::Supported: break;
                }
            }
            deletion.selectedItems.reserve(partition.originalIndices.size());
            for (const size_t originalIndex : partition.originalIndices)
            {
                const std::filesystem::path& sourcePath = sourcePaths[originalIndex];
                deletion.selectedItems.emplace_back(FileOperations::QualifiedSourceItem{.providerPath = sourcePath.native()});
                groupedSourcePaths.push_back(sourcePath);
            }
            deletion.origin    = deleteOrigin;
            deletion.recursive = (flags & FILESYSTEM_FLAG_RECURSIVE) != 0;
            deletion.options   = options;
            plans.emplace_back(std::move(deletion));
        }
    }

    sourcePaths = std::move(groupedSourcePaths);
    if (! resolvedItems.empty())
    {
        resolvedItems = std::move(groupedResolvedItems);
    }

    uint64_t sourceCount = 0u;
    uint64_t payloadBytes = 0u;
    for (const FileOperations::FileOperationPlan& plan : plans)
    {
        FileOperations::PlanRejectionBucket rejectionBucket = FileOperations::PlanRejectionBucket::None;
        const HRESULT validationHr = ValidatePlanImpl(plan, &rejectionBucket, true);
        if (FAILED(validationHr))
        {
            return reject(validationHr, rejectionBucket);
        }
        sourceCount += PlanSourceCount(plan);
        payloadBytes += PlanPayloadBytes(plan);
    }

    const uint64_t constructionUs = Debug::Perf::ElapsedUs(startedAt);
    Debug::Perf::Emit(L"fileops.plan.construct_us",
                      plans.size() == 1u ? L"accepted-envelope" : L"accepted-mixed-root-group",
                      constructionUs,
                      sourceCount,
                      payloadBytes,
                      S_OK);
    Debug::Perf::EmitValue(L"fileops.plan.source_count", sourceCount);
    Debug::Perf::EmitValue(L"fileops.plan.payload_bytes", payloadBytes);

    OperationAdmission admission{};
    admission.plans                      = std::move(plans);
    admission.sourcePane                 = sourcePane;
    admission.destinationPane            = destinationPane;
    admission.fileSystem                 = fileSystem;
    admission.destinationFileSystem      = std::move(destinationFileSystem);
    admission.flags                      = flags;
    admission.executionMode              = executionMode;
    admission.requireConfirmation        = requireConfirmation;
    admission.resolvedItems              = std::move(resolvedItems);
    admission.confirmationMessage        = std::move(confirmationMessage);
    admission.sourcePluginShortId        = sourcePluginShortId;
    admission.sourceInstanceContext      = sourceInstanceContext;
    admission.destinationPluginShortId   = std::move(destinationPluginShortId);
    admission.destinationInstanceContext = std::move(destinationInstanceContext);
    admission.capturedConsentKind         = capturedConsentKind;
    admission.preWorkerReleaseBarrier    = std::move(preWorkerReleaseBarrier);
    admission.preConsumptionDecisionGate = std::move(preConsumptionDecisionGate);
    admission.preparationObserver        = std::move(preparationObserver);

    const auto admissionStartedAt = std::chrono::steady_clock::now();
    const HRESULT admissionHr = StartOperation(std::move(admission), taskIdOut);
    Debug::Perf::Emit(L"fileops.plan.admit_us",
                      admissionHr == S_OK ? L"accepted" : L"not-published",
                      Debug::Perf::ElapsedUs(admissionStartedAt),
                      sourceCount,
                      payloadBytes,
                      admissionHr);
    return admissionHr;
}

HRESULT FolderWindow::FileOperationState::AdmitInlineRename(FolderWindow::Pane sourcePane,
                                                             const wil::com_ptr<IFileSystem>& fileSystem,
                                                             std::filesystem::path sourcePath,
                                                             std::wstring finalLeafName,
                                                             uint64_t* taskIdOut,
                                                             std::filesystem::path* providerParentPathOut)
{
#ifdef ENABLE_TESTS
    FolderWindowFileOperationsStateInternal::g_fileOpsInlineRenameAdmissionThreadId.store(GetCurrentThreadId(), std::memory_order_release);
#endif
    if (taskIdOut)
    {
        *taskIdOut = 0u;
    }
    if (providerParentPathOut)
    {
        providerParentPathOut->clear();
    }
    if (! fileSystem)
    {
        return E_POINTER;
    }
    if (sourcePath.empty() || finalLeafName.empty())
    {
        return E_INVALIDARG;
    }

    const FolderWindow::PaneState& sourceState =
        sourcePane == FolderWindow::Pane::Left ? _owner._leftPane : _owner._rightPane;
    if (sourceState.fileSystem.get() != fileSystem.get() || sourceState.pluginId.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const std::optional<FileSystemCapabilitiesV2> capabilities =
        TryGetCapabilities(fileSystem, sourcePath.native(), FILESYSTEM_RENAME, sourceState.pluginId);
    if (! capabilities.has_value() || ! capabilities->renameOperation || ! HasStablePathIdentity(capabilities.value()) ||
        ! HasContainedCancellationRoute(capabilities.value()))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    std::wstring providerParentPath;
    if (! TryGetFileSystemParentPath(capabilities->pathIdentity.value(), sourcePath.native(), providerParentPath))
    {
        return E_INVALIDARG;
    }
    // C1 (b0): the destination provider's child-name contract is a provider call, so it runs in
    // Preparing on the task thread (`Task::PrepareTransferDestinationNames`), which fills the
    // join and collision key below or fails the task on its card with the provider's reason.
    // Admission only proves the leaf is not empty and the source has a parent.

    FileOperations::RenamePlan rename{};
    rename.origin = FileOperations::RenameOrigin::InlineRename;
    rename.endpoint = FileOperations::QualifiedEndpoint{
        .pluginId   = sourceState.pluginId,
        .instanceId = CanonicalInstanceId(sourceState.instanceContext),
        .profileId  = capabilities->pathProfileId,
        .rootId     = capabilities->rootId,
        .pathIdentity = capabilities->pathIdentity,
        .verificationHostReadback = capabilities->verificationHostReadback,
        .verificationProviderBlake3Proof = capabilities->verificationProviderBlake3Proof,
        .verificationWriterDigestProof = capabilities->verificationWriterDigestProof,
        .verificationCapabilityCheckDeferred = capabilities->verificationCapabilityCheckDeferred,
        .cancellationAbort = capabilities->cancellationAbort,
        .cancellationDeadline = capabilities->cancellationDeadline,
        .cancellationRouteClass = capabilities->cancellationRouteClass,
        .providerWatchdogTimeoutMs = capabilities->providerWatchdogTimeoutMs,
    };
    rename.finalMappings.emplace_back(FileOperations::RenameStep{
        .source        = FileOperations::QualifiedSourceItem{.providerPath = sourcePath.native()},
        .finalLeafName = std::move(finalLeafName),
    });
    // Inline F2 is the explicit global-Queue exception. Overlapping mutation roots remain gated
    // by EnterOperation's identity-aware interlock.
    rename.options.executionMode = FileOperations::ExecutionMode::Parallel;

    FileOperations::PlanRejectionBucket rejection = FileOperations::PlanRejectionBucket::None;
    const HRESULT validationHr = FileOperations::ValidatePlan(rename, &rejection);
    if (FAILED(validationHr))
    {
        Debug::Perf::EmitValue(L"fileops.plan.rejection_bucket", static_cast<uint64_t>(rejection), validationHr);
        return validationHr;
    }

    OperationAdmission admission{};
    admission.plans.emplace_back(std::move(rename));
    admission.sourcePane         = sourcePane;
    admission.fileSystem         = fileSystem;
    admission.executionMode      = ExecutionMode::PerItem;
    admission.sourcePluginShortId = sourceState.pluginShortId;
    admission.sourceInstanceContext = sourceState.instanceContext;
    const HRESULT startHr = StartOperation(std::move(admission), taskIdOut);
    if (SUCCEEDED(startHr) && providerParentPathOut)
    {
        *providerParentPathOut = std::filesystem::path(std::move(providerParentPath));
    }
    return startHr;
}

HRESULT FolderWindow::FileOperationState::Task::PrepareBatchRenameAdmission() noexcept
{
    if (! _pendingBatchRenameAdmission.has_value())
    {
        return S_OK;
    }

    BatchRenameAdmissionInput input = std::move(_pendingBatchRenameAdmission.value());
    _pendingBatchRenameAdmission.reset();
    const auto workerStartedAt = std::chrono::steady_clock::now();
    const uint64_t rowCount = static_cast<uint64_t>(input.operations.size());
    uint64_t capabilityCallCount = 0u;
    uint64_t bindCallCount = 0u;
    uint64_t providerNameQueryCount = 0u;
    uint64_t providerNameArenaFallbackCount = 0u;
    uint64_t cycleCount = 0u;
    HRESULT result = E_UNEXPECTED;
    std::wstring_view detail = L"not-published";
    const auto emitAdmission = wil::scope_exit([&]() noexcept
    {
        const uint64_t wallUs = Debug::Perf::ElapsedUs(input.startedAt);
        const uint64_t workerWallUs = Debug::Perf::ElapsedUs(workerStartedAt);
        Debug::Perf::Emit(L"batchrename.admit.us", detail, wallUs, rowCount, cycleCount, result);
        Debug::Perf::Emit(L"batchrename.admission.wall_us", detail, wallUs, rowCount, 0u, result);
        Debug::Perf::Emit(L"batchrename.admission.worker_wall_us", detail, workerWallUs, rowCount, 0u, result);
        Debug::Perf::EmitValue(L"batchrename.admission.ui_capture_us", input.uiCaptureUs, result);
        Debug::Perf::EmitValue(L"batchrename.admission.capability_calls", capabilityCallCount, result);
        Debug::Perf::EmitValue(L"batchrename.admission.bind_calls", bindCallCount, result);
        Debug::Perf::EmitValue(L"batchrename.admission.provider_name_queries", providerNameQueryCount, result);
        Debug::Perf::EmitValue(L"batchrename.admission.provider_name_arena_fallbacks", providerNameArenaFallbackCount, result);
        Debug::Perf::EmitValue(L"batchrename.admission.retained_bindings", 0u, result);
        if (input.origin == FileOperations::RenameOrigin::ChangeCase)
        {
            Debug::Perf::Emit(L"changecase.admission.wall_us", detail, wallUs, rowCount, cycleCount, result);
            Debug::Perf::Emit(L"changecase.admission.worker_wall_us", detail, workerWallUs, rowCount, 0u, result);
            Debug::Perf::EmitValue(L"changecase.admission.ui_capture_us", input.uiCaptureUs, result);
            Debug::Perf::EmitValue(L"changecase.admission.capability_calls", capabilityCallCount, result);
            Debug::Perf::EmitValue(L"changecase.admission.bind_calls", bindCallCount, result);
            Debug::Perf::EmitValue(L"changecase.admission.provider_name_queries", providerNameQueryCount, result);
            Debug::Perf::EmitValue(L"changecase.admission.retained_bindings", 0u, result);
        }
        if (FolderWindowFileOperationsStateInternal::IsCancellationStatus(result))
        {
            const ULONGLONG cancelRequestedTick = _cancelRequestedTick.load(std::memory_order_acquire);
            const ULONGLONG nowTick             = GetTickCount64();
            const uint64_t cancelLatencyUs      = cancelRequestedTick != 0u && nowTick >= cancelRequestedTick
                                                      ? (nowTick - cancelRequestedTick) * 1'000u
                                                      : wallUs;
            Debug::Perf::EmitValue(L"batchrename.admission.cancel_latency_us", cancelLatencyUs, result);
            if (input.origin == FileOperations::RenameOrigin::ChangeCase)
            {
                Debug::Perf::EmitValue(L"changecase.admission.cancel_latency_us", cancelLatencyUs, result);
            }
        }
    });

    if (! _fileSystem)
    {
        result = E_POINTER;
        return result;
    }
    if (input.operations.empty() || _sourcePluginId.empty())
    {
        result = E_INVALIDARG;
        return result;
    }
    if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
    {
        result = HRESULT_FROM_WIN32(ERROR_CANCELLED);
        return result;
    }

    FileOperations::RenamePlan rename{};
    rename.origin = input.origin;
    rename.options.executionMode = input.executionMode;
    rename.finalMappings.reserve(input.operations.size());

    constexpr FileSystemBindFlags ingressBindFlags = static_cast<FileSystemBindFlags>(
        FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_RENAME);
    wil::com_ptr<IFileSystemRouteCapabilities> route;
    const HRESULT routeHr = _fileSystem->QueryInterface(__uuidof(IFileSystemRouteCapabilities), route.put_void());
    if (FAILED(routeHr) || ! route)
    {
        result = FAILED(routeHr) ? routeHr : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        return result;
    }
    std::optional<FileSystemCapabilitiesV2> firstCapabilities;
    for (BatchRenameExecutionOp& operation : input.operations)
    {
        if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
        {
            result = HRESULT_FROM_WIN32(ERROR_CANCELLED);
            return result;
        }
        if (operation.originalSource.empty() || operation.finalLeaf.empty() || operation.providerFinalPath.empty() ||
            operation.providerParentKey.empty() || operation.providerSourceCollisionKey.empty() || operation.providerFinalCollisionKey.empty())
        {
            result = E_INVALIDARG;
            return result;
        }
        operation.depth = PathDepthKey(operation.originalSource);

        ++capabilityCallCount;
        std::optional<FileSystemCapabilitiesV2> capabilities =
            TryGetCapabilities(_fileSystem, operation.originalSource.native(), FILESYSTEM_RENAME, _sourcePluginId);
        if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
        {
            result = HRESULT_FROM_WIN32(ERROR_CANCELLED);
            return result;
        }
        if (! capabilities.has_value() || ! capabilities->renameOperation || ! HasStablePathIdentity(capabilities.value()) ||
            ! HasContainedCancellationRoute(capabilities.value()))
        {
            result = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            return result;
        }
        if (! firstCapabilities.has_value())
        {
            firstCapabilities = capabilities;
        }
        else if (capabilities->pathProfileId != firstCapabilities->pathProfileId ||
                 capabilities->rootId != firstCapabilities->rootId ||
                 ! SamePathIdentityContract(capabilities->pathIdentity.value(), firstCapabilities->pathIdentity.value()))
        {
            // One RenamePlan owns one executable identity domain. A mixed-root request must be
            // split by its caller before the single admission barrier.
            result = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            return result;
        }

        std::wstring parentPath;
        if (! TryGetFileSystemParentPath(capabilities->pathIdentity.value(), operation.originalSource.native(), parentPath))
        {
            result = E_INVALIDARG;
            return result;
        }
        const std::optional<std::wstring> parentKey = TryMakePathKey(capabilities->pathIdentity.value(), parentPath);
        std::wstring sourceLeaf;
        if (! TryGetFileSystemLeafName(capabilities->pathIdentity.value(), operation.originalSource.native(), sourceLeaf))
        {
            result = E_INVALIDARG;
            return result;
        }
        ++providerNameQueryCount;
        const FileSystemRouteContract::ChildNameContractResult sourceName = FileSystemRouteContract::QueryChildNameContract(
            route.get(), parentPath, sourceLeaf, FILESYSTEM_RENAME, _sourcePluginId);
        providerNameArenaFallbackCount += sourceName.arenaFallbackCount;
        ++providerNameQueryCount;
        const FileSystemRouteContract::ChildNameContractResult finalName = FileSystemRouteContract::QueryChildNameContract(
            route.get(), parentPath, operation.finalLeaf, FILESYSTEM_RENAME, _sourcePluginId);
        providerNameArenaFallbackCount += finalName.arenaFallbackCount;
        const auto nameFailure = [](const FileSystemRouteContract::ChildNameContractResult& contract) noexcept
        {
            if (contract.state == FileSystemRouteContract::QueryState::Available && contract.nameStatus == FILESYSTEM_CHILD_NAME_INVALID)
            {
                return FAILED(contract.failureStatus) ? contract.failureStatus : HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
            }
            return FAILED(contract.status) ? contract.status : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        };
        if (! parentKey.has_value() || sourceName.state != FileSystemRouteContract::QueryState::Available ||
            sourceName.nameStatus != FILESYSTEM_CHILD_NAME_VALID || sourceName.collisionKey.empty())
        {
            result = ! parentKey.has_value() ? E_INVALIDARG : nameFailure(sourceName);
            return result;
        }
        if (finalName.state != FileSystemRouteContract::QueryState::Available || finalName.nameStatus != FILESYSTEM_CHILD_NAME_VALID ||
            finalName.joinedPath.empty() || finalName.collisionKey.empty())
        {
            result = nameFailure(finalName);
            return result;
        }
        if (operation.providerFinalPath.native() != finalName.joinedPath || operation.providerParentKey != parentKey.value() ||
            operation.providerSourceCollisionKey != sourceName.collisionKey || operation.providerFinalCollisionKey != finalName.collisionKey)
        {
            result = HRESULT_FROM_WIN32(ERROR_RETRY);
            return result;
        }

        ++bindCallCount;
        FileOperations::ObjectBindingResult ingress = FileOperations::BindObjectAuthority(
            _fileSystem.get(), operation.originalSource.native(), capabilities->pathProfileId, ingressBindFlags);
        if (_cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested())
        {
            result = HRESULT_FROM_WIN32(ERROR_CANCELLED);
            return result;
        }
        if (ingress.state != FileOperations::ObjectBindingState::Bound)
        {
            result = FAILED(ingress.status) ? ingress.status : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            return result;
        }
        operation.isDirectory = ingress.authority.kind == FILESYSTEM_BOUND_DIRECTORY;
        rename.finalMappings.emplace_back(FileOperations::RenameStep{
            .source = FileOperations::QualifiedSourceItem{
                .providerPath = operation.originalSource.native(),
                .ingressSnapshot = std::move(ingress.authority.identity),
            },
            .finalLeafName = operation.finalLeaf,
            .providerJoinedPath = finalName.joinedPath,
            .providerParentKey = parentKey.value(),
            .providerSourceCollisionKey = sourceName.collisionKey,
            .providerCollisionKey = finalName.collisionKey,
        });
    }

    const FileSystemCapabilitiesV2& capabilities = firstCapabilities.value();
    rename.endpoint = FileOperations::QualifiedEndpoint{
        .pluginId = _sourcePluginId,
        .instanceId = CanonicalInstanceId(_sourceInstanceContext),
        .profileId = capabilities.pathProfileId,
        .rootId = capabilities.rootId,
        .pathIdentity = capabilities.pathIdentity,
        .verificationHostReadback = capabilities.verificationHostReadback,
        .verificationProviderBlake3Proof = capabilities.verificationProviderBlake3Proof,
        .verificationWriterDigestProof = capabilities.verificationWriterDigestProof,
        .verificationCapabilityCheckDeferred = capabilities.verificationCapabilityCheckDeferred,
        .cancellationAbort = capabilities.cancellationAbort,
        .cancellationDeadline = capabilities.cancellationDeadline,
        .cancellationRouteClass = capabilities.cancellationRouteClass,
        .providerWatchdogTimeoutMs = capabilities.providerWatchdogTimeoutMs,
    };

    // Build the one immutable acyclic schedule after capability and identity qualification. The
    // same builder powers preview and direct execution; a cycle is rejected here before the plan
    // can be published or any mutation interlock can be entered.
    const FileSystemPathIdentity& pathIdentity = capabilities.pathIdentity.value();
    const HRESULT scheduleHr = BuildBatchRenameExecutionSchedule(pathIdentity, input.operations, rename.schedule);
    cycleCount = static_cast<uint64_t>(rename.schedule.cycleOperationIndices.size());
    if (FAILED(scheduleHr))
    {
        detail = scheduleHr == HRESULT_FROM_WIN32(ERROR_CIRCULAR_DEPENDENCY) ? L"dependency-cycle" : L"invalid-schedule";
        result = scheduleHr;
        return result;
    }

    FileOperations::PlanRejectionBucket rejection = FileOperations::PlanRejectionBucket::None;
    const HRESULT validationHr = FileOperations::ValidatePlan(rename, &rejection);
    if (FAILED(validationHr))
    {
        Debug::Perf::EmitValue(L"fileops.plan.rejection_bucket", static_cast<uint64_t>(rejection), validationHr);
        result = validationHr;
        return result;
    }

    FileOperations::FileOperationPlanGroup plans;
    plans.emplace_back(std::move(rename));
    StorePlans(std::make_shared<const FileOperations::FileOperationPlanGroup>(std::move(plans)));
    detail = L"accepted";
    result = S_OK;
    return result;
}

HRESULT FolderWindow::FileOperationState::AdmitScheduledRename(
    FolderWindow::Pane sourcePane,
    const wil::com_ptr<IFileSystem>& fileSystem,
    const FileOperations::RenameOrigin origin,
    std::vector<BatchRenameExecutionOp> operations,
    std::function<void(uint64_t completedItems, uint64_t totalItems)> progressCallback,
    std::function<void(uint64_t taskId)> publishedCallback,
    uint64_t* taskIdOut,
    std::function<HRESULT()> preConsumptionDecisionGate,
    std::function<void(std::shared_ptr<const FileOperations::PreparationSnapshot>)> preparationObserver)
{
    const auto startedAt = std::chrono::steady_clock::now();
    if (taskIdOut)
    {
        *taskIdOut = 0u;
    }
    if (_completionShutdown.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS);
    }
    if (! fileSystem)
    {
        return E_POINTER;
    }
    if (operations.empty())
    {
        return S_FALSE;
    }
    if (origin != FileOperations::RenameOrigin::BatchRename && origin != FileOperations::RenameOrigin::ChangeCase)
    {
        return E_INVALIDARG;
    }

    const FolderWindow::PaneState& sourceState =
        sourcePane == FolderWindow::Pane::Left ? _owner._leftPane : _owner._rightPane;
    if (sourceState.fileSystem.get() != fileSystem.get() || sourceState.pluginId.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    std::vector<std::filesystem::path> sourcePaths;
    sourcePaths.reserve(operations.size());
    for (const BatchRenameExecutionOp& operation : operations)
    {
        if (operation.originalSource.empty() || operation.finalLeaf.empty())
        {
            return E_INVALIDARG;
        }
        sourcePaths.push_back(operation.originalSource);
    }

    const FileOperations::ExecutionMode planExecutionMode =
        ShouldQueueNewTask() ? FileOperations::ExecutionMode::Queue : FileOperations::ExecutionMode::Parallel;
    const uint64_t uiCaptureUs = Debug::Perf::ElapsedUs(startedAt);
    uint64_t admittedTaskId = 0u;
    {
        std::scoped_lock lock(_mutex);
        admittedTaskId = _nextTaskId++;
    }

    const ULONGLONG presentationAdmissionTick = TaskPresentationNowTick();
    auto task = std::make_unique<Task>(*this);
    task->_taskId = admittedTaskId;
    task->_queueOrderKey.store(admittedTaskId, std::memory_order_release);
    task->_operation = FILESYSTEM_RENAME;
    task->_executionMode = ExecutionMode::PerItem;
    task->_sourcePane = sourcePane;
    task->_sourcePluginId = sourceState.pluginId;
    task->_sourcePluginShortId = sourceState.pluginShortId;
    task->_sourceInstanceContext = sourceState.instanceContext;
    task->_fileSystem = fileSystem;
    task->_sourcePaths = std::move(sourcePaths);
    task->_waitForOthers.store(planExecutionMode == FileOperations::ExecutionMode::Queue, std::memory_order_release);
    task->_externalProgressCallback = std::move(progressCallback);
    task->_preConsumptionDecisionGate = std::move(preConsumptionDecisionGate);
    task->_preparationObserver = std::move(preparationObserver);
    task->_pendingBatchRenameAdmission.emplace(BatchRenameAdmissionInput{
        .origin = origin,
        .operations = std::move(operations),
        .executionMode = planExecutionMode,
        .startedAt = startedAt,
        .uiCaptureUs = uiCaptureUs,
    });
    task->_presentationState.store(Task::TaskPresentationState::Hidden, std::memory_order_release);
    task->_presentationDeadlineTick =
        presentationAdmissionTick > std::numeric_limits<ULONGLONG>::max() - FileOperations::kTaskCardRevealDelayMs
            ? std::numeric_limits<ULONGLONG>::max()
            : presentationAdmissionTick + FileOperations::kTaskCardRevealDelayMs;
    task->InitializeSourceItemResultBuilders();
    task->SetWaitingInQueue(false);

    {
        const size_t itemCount = task->_sourcePaths.size();
        std::scoped_lock lock(task->_topLevelCompletionMutex);
        task->_topLevelItemKinds.assign(itemCount, Task::TopLevelItemKind::Unknown);
        task->_topLevelItemCompleted.assign(itemCount, 0u);
    }
    {
        std::scoped_lock lock(task->_progressPathMutex);
        task->_progressSourcePath = task->_sourcePaths.front().native();
    }

    Task* const rawTask = task.get();
    {
        std::scoped_lock lock(_mutex);
        _tasks.emplace_back(std::move(task));
    }

    try
    {
        rawTask->_thread = std::jthread([rawTask](std::stop_token stopToken) noexcept { rawTask->ThreadMain(stopToken); });
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::system_error& error)
    {
        Debug::Error(L"File Operations could not create the Batch Rename admission worker (taskId={}, code={}).",
                     admittedTaskId,
                     error.code().value());
        std::scoped_lock lock(_mutex);
        const auto taskIt = std::ranges::find_if(_tasks, [rawTask](const auto& candidate) noexcept
        { return candidate.get() == rawTask; });
        if (taskIt != _tasks.end())
        {
            _tasks.erase(taskIt);
        }
        return HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY);
    }

    RequestTaskPresentationRefresh();
    if (publishedCallback)
    {
        try
        {
            publishedCallback(admittedTaskId);
        }
        catch (const std::bad_alloc&)
        {
            // Completion registration is part of task publication. OOM is fatal by process policy;
            // never unwind with the worker permanently parked behind its release gate.
            std::terminate();
        }
    }
    const uint64_t uiRowCount = static_cast<uint64_t>(rawTask->_sourcePaths.size());
    if (taskIdOut)
    {
        *taskIdOut = admittedTaskId;
    }
    rawTask->_workerReleased.store(true, std::memory_order_release);
    rawTask->_workerReleased.notify_all();
    Debug::Perf::Emit(L"batchrename.admission.ui_capture_us",
                      L"published",
                      uiCaptureUs,
                      uiRowCount,
                      0u,
                      S_OK);
    Debug::Perf::Emit(L"batchrename.admission.ui_thread_us",
                      L"published",
                      Debug::Perf::ElapsedUs(startedAt),
                      uiRowCount,
                      0u,
                      S_OK);
    if (origin == FileOperations::RenameOrigin::ChangeCase)
    {
        Debug::Perf::Emit(L"changecase.admission.ui_thread_us",
                          L"published",
                          Debug::Perf::ElapsedUs(startedAt),
                          uiRowCount,
                          0u,
                          S_OK);
    }
    return S_OK;
}

HRESULT FolderWindow::FileOperationState::AdmitBatchRename(
    FolderWindow::Pane sourcePane,
    const wil::com_ptr<IFileSystem>& fileSystem,
    std::vector<BatchRenameExecutionOp> operations,
    std::function<void(uint64_t completedItems, uint64_t totalItems)> progressCallback,
    std::function<void(uint64_t taskId)> publishedCallback,
    uint64_t* taskIdOut,
    std::function<HRESULT()> preConsumptionDecisionGate,
    std::function<void(std::shared_ptr<const FileOperations::PreparationSnapshot>)> preparationObserver)
{
    return AdmitScheduledRename(sourcePane,
                                fileSystem,
                                FileOperations::RenameOrigin::BatchRename,
                                std::move(operations),
                                std::move(progressCallback),
                                std::move(publishedCallback),
                                taskIdOut,
                                std::move(preConsumptionDecisionGate),
                                std::move(preparationObserver));
}

HRESULT FolderWindow::FileOperationState::QualifyCreateDirectory(
    const wil::com_ptr<IFileSystem>& fileSystem,
    std::wstring_view pluginId,
    std::wstring_view instanceContext,
    const std::filesystem::path& providerParentPath,
    std::wstring_view leafName,
    FileOperations::CreateDirectoryAdmission& admissionOut) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    HRESULT result       = HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    std::wstring_view detail = L"rejected";
    const auto emitAdmission = wil::scope_exit([&]() noexcept
    {
        Debug::Perf::Emit(L"fileops.create_directory.admit_us",
                          detail,
                          Debug::Perf::ElapsedUs(startedAt),
                          1u,
                          static_cast<uint64_t>(leafName.size()),
                          result);
    });

    admissionOut = {};
    if (! fileSystem || pluginId.empty() || providerParentPath.empty() || leafName.empty() || leafName == L"." || leafName == L"..")
    {
        result = E_INVALIDARG;
        return result;
    }

    const std::optional<FileSystemCapabilitiesV2> parentCapabilities =
        TryGetCapabilities(fileSystem, providerParentPath.native(), FILESYSTEM_CREATE_DIRECTORY, pluginId);
    if (! parentCapabilities.has_value() || ! parentCapabilities->createDirectoryOperation ||
        ! HasStablePathIdentity(parentCapabilities.value()) || ! HasContainedCancellationRoute(parentCapabilities.value()) ||
        parentCapabilities->rootId.empty())
    {
        return result;
    }

    const FileSystemPathIdentity& parentIdentity = parentCapabilities->pathIdentity.value();
    const FileSystemRouteContract::ChildNameContractResult nameContract = FileSystemRouteContract::QueryChildNameContract(
        fileSystem.get(), providerParentPath.native(), leafName, FILESYSTEM_CREATE_DIRECTORY, pluginId);
    if (nameContract.state != FileSystemRouteContract::QueryState::Available ||
        nameContract.nameStatus != FILESYSTEM_CHILD_NAME_VALID)
    {
        result = nameContract.state == FileSystemRouteContract::QueryState::Available && FAILED(nameContract.failureStatus)
            ? nameContract.failureStatus
            : (FAILED(nameContract.status) ? nameContract.status : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED));
        return result;
    }
    const std::wstring& candidatePath = nameContract.joinedPath;

    const std::optional<FileSystemCapabilitiesV2> candidateCapabilities =
        TryGetCapabilities(fileSystem, candidatePath, FILESYSTEM_CREATE_DIRECTORY, pluginId);
    if (! candidateCapabilities.has_value() || ! candidateCapabilities->createDirectoryOperation ||
        ! HasStablePathIdentity(candidateCapabilities.value()) || ! HasContainedCancellationRoute(candidateCapabilities.value()) ||
        candidateCapabilities->rootId.empty() ||
        candidateCapabilities->pathProfileId != parentCapabilities->pathProfileId ||
        candidateCapabilities->rootId != parentCapabilities->rootId ||
        ! SamePathIdentityContract(candidateCapabilities->pathIdentity.value(), parentIdentity))
    {
        return result;
    }

    std::wstring derivedParent;
    std::wstring derivedLeaf;
    if (! TryGetFileSystemParentPath(candidateCapabilities->pathIdentity.value(), candidatePath, derivedParent) ||
        ! EquivalentPath(candidateCapabilities->pathIdentity.value(), derivedParent, providerParentPath.native()) ||
        ! TryGetFileSystemLeafName(candidateCapabilities->pathIdentity.value(), candidatePath, derivedLeaf) ||
        ! EquivalentComponent(candidateCapabilities->pathIdentity.value(), derivedLeaf, leafName))
    {
        result = HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        return result;
    }

    const bool localEndpoint = NavigationLocation::EqualsNoCase(pluginId, L"builtin/file-system");

    admissionOut.endpoint = FileOperations::QualifiedEndpoint{
        .pluginId   = std::wstring(pluginId),
        .instanceId = CanonicalInstanceId(instanceContext),
        .profileId  = candidateCapabilities->pathProfileId,
        .rootId     = candidateCapabilities->rootId,
        .pathIdentity = candidateCapabilities->pathIdentity,
        .verificationHostReadback = candidateCapabilities->verificationHostReadback,
        .verificationProviderBlake3Proof = candidateCapabilities->verificationProviderBlake3Proof,
        .verificationWriterDigestProof = candidateCapabilities->verificationWriterDigestProof,
        .verificationCapabilityCheckDeferred = candidateCapabilities->verificationCapabilityCheckDeferred,
        .cancellationAbort = candidateCapabilities->cancellationAbort,
        .cancellationDeadline = candidateCapabilities->cancellationDeadline,
        .cancellationRouteClass = candidateCapabilities->cancellationRouteClass,
        .providerWatchdogTimeoutMs = candidateCapabilities->providerWatchdogTimeoutMs,
    };
    admissionOut.candidateProviderPath = std::filesystem::path(candidatePath);
    admissionOut.providerCollisionKey = nameContract.collisionKey;
    admissionOut.allowLocalNativeFallback = localEndpoint;
    result = S_OK;
    detail = localEndpoint ? L"accepted-local" : L"accepted-provider";
    return result;
}

HRESULT FolderWindow::StartFileOperationFromFolderView(Pane pane, FolderView::FileOperationRequest request) noexcept
{
    PaneState& destinationState = pane == Pane::Left ? _leftPane : _rightPane;
    if (! destinationState.fileSystem)
    {
        return E_POINTER;
    }

    EnsureFileOperations();
    if (! _fileOperations)
    {
        return E_FAIL;
    }

    if (request.operation == FILESYSTEM_RENAME)
    {
        if (request.origin != FolderView::FileOperationRequest::Origin::InlineF2 || request.sourcePaths.size() != 1u ||
            ! request.renameLeafName.has_value() || request.renameLeafName->empty() || request.destinationFolder.has_value() ||
            request.moveClipboardSequence.has_value() || request.preWorkerReleaseBarrier)
        {
            return E_INVALIDARG;
        }

        const std::filesystem::path sourcePath = std::move(request.sourcePaths.front());
        const std::wstring finalLeafName = std::move(request.renameLeafName.value());
        std::filesystem::path providerParentPath;
        uint64_t taskId = 0u;
        const FileOperationPromptDispatchScope promptDispatch(*this);
        const HRESULT startHr = _fileOperations->AdmitInlineRename(
            pane, destinationState.fileSystem, sourcePath, finalLeafName, &taskId, &providerParentPath);
        if (SUCCEEDED(startHr) && taskId != 0u)
        {
            std::function<void(HRESULT)> requestCompletion = std::move(request.completionCallback);
            _fileOperationRequestCompletionCallbacks.insert_or_assign(
                taskId,
                [this,
                 pane,
                 providerParentPath,
                 finalLeafName,
                 completion = std::move(requestCompletion)](const FileOperationCompletedEvent& event) mutable
            {
                const auto renamed = std::ranges::find_if(event.itemOutcomes, [](const FileOperationItemOutcome& outcome) noexcept
                {
                    return outcome.sourceIndex == 0u &&
                           outcome.publication == FileOperations::PublicationState::Published &&
                           outcome.sourceDisposition == FileOperations::SourceDisposition::Removed &&
                           outcome.completion == FileOperations::ItemCompletion::Completed;
                });
                if (SUCCEEDED(event.hr) && renamed != event.itemOutcomes.end())
                {
                    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
                    state.folderView.RememberFocusedItemForFolder(providerParentPath, finalLeafName);
                }
                if (completion)
                {
                    completion(event.hr);
                }
            });
        }
        return startHr;
    }

    const bool isCopyMove = request.operation == FILESYSTEM_COPY || request.operation == FILESYSTEM_MOVE;
    if (request.sourcePaths.empty())
    {
        return S_FALSE;
    }

    Pane sourcePane                      = pane;
    std::optional<Pane> destinationPane  = std::nullopt;
    wil::com_ptr<IFileSystem> fileSystem = destinationState.fileSystem;
    wil::com_ptr<IFileSystem> destinationFileSystem;
    std::wstring sourcePluginIdOverride;
    std::wstring sourcePluginShortIdOverride;
    std::optional<std::wstring> sourceInstanceIdOverride;

    if (isCopyMove && ! request.sourceContextSpecified &&
        CompareStringOrdinal(destinationState.pluginId.c_str(), -1, L"builtin/file-system", -1, TRUE) != CSTR_EQUAL)
    {
        const FileSystemPluginManager::PluginEntry* localEntry = FileSystemPluginManager::GetInstance().FindPluginById(L"builtin/file-system");
        if (! localEntry || ! localEntry->fileSystem || localEntry->disabled || ! localEntry->loadable || localEntry->shortId.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        const std::filesystem::path destinationCapabilityPath =
            request.destinationFolder.value_or(destinationState.folderView.GetFolderPath().value_or(std::filesystem::path{}));
        if (destinationCapabilityPath.empty() ||
            ! CanCrossFileSystemCopyMove(localEntry->fileSystem,
                                         request.sourcePaths.front().native(),
                                         localEntry->id,
                                         destinationState.fileSystem,
                                         destinationCapabilityPath.native(),
                                         destinationState.pluginId,
                                         request.operation))
        {
            destinationState.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation,
                                                         FolderView::OverlaySeverity::Error,
                                                         LoadStringResource(nullptr, IDS_CAPTION_ERROR),
                                                         LoadStringResource(nullptr, IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS));
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        fileSystem                 = localEntry->fileSystem;
        destinationFileSystem      = destinationState.fileSystem;
        destinationPane            = pane;
        sourcePluginIdOverride     = localEntry->id;
        sourcePluginShortIdOverride = localEntry->shortId;
        sourceInstanceIdOverride    = std::wstring{};
    }

    if (isCopyMove && request.sourceContextSpecified)
    {
        const auto contextMatches = [&](const PaneState& paneState) noexcept -> bool
        {
            return CompareStringOrdinal(paneState.pluginId.c_str(), -1, request.sourcePluginId.c_str(), -1, TRUE) == CSTR_EQUAL &&
                   NavigationLocation::EqualsNoCase(paneState.instanceContext, request.sourceInstanceContext);
        };

        const bool leftMatches  = contextMatches(_leftPane);
        const bool rightMatches = contextMatches(_rightPane);

        if (leftMatches ^ rightMatches)
        {
            sourcePane = leftMatches ? Pane::Left : Pane::Right;
        }
        else if (leftMatches && rightMatches)
        {
            const auto isUnderFolder = [](std::wstring_view folder, std::wstring_view path) noexcept -> bool
            {
                while (! folder.empty() && (folder.back() == L'\\' || folder.back() == L'/'))
                {
                    folder.remove_suffix(1);
                }
                if (folder.empty() || path.size() <= folder.size())
                {
                    return false;
                }

                if (! OrdinalString::StartsWithNoCase(path, folder))
                {
                    return false;
                }

                const wchar_t next = path[folder.size()];
                return next == L'\\' || next == L'/';
            };

            bool inferredSourcePane = false;
            if (! request.sourcePaths.empty())
            {
                const std::wstring_view firstPath = request.sourcePaths.front().native();

                const auto leftFolder  = _leftPane.folderView.GetFolderPath();
                const auto rightFolder = _rightPane.folderView.GetFolderPath();

                const bool underLeft  = leftFolder.has_value() && isUnderFolder(leftFolder->native(), firstPath);
                const bool underRight = rightFolder.has_value() && isUnderFolder(rightFolder->native(), firstPath);

                if (underLeft ^ underRight)
                {
                    sourcePane         = underLeft ? Pane::Left : Pane::Right;
                    inferredSourcePane = true;
                }
            }

            if (! inferredSourcePane)
            {
                return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            }
        }
        else
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        if (sourcePane != pane)
        {
            PaneState& sourceState = sourcePane == Pane::Left ? _leftPane : _rightPane;
            if (! sourceState.fileSystem)
            {
                return E_POINTER;
            }

            if (! SanityCheckBothPanes(sourceState, destinationState, request.operation))
            {
                return E_FAIL;
            }

            fileSystem      = sourceState.fileSystem;
            destinationPane = pane;

            const bool contextSame = CompareStringOrdinal(sourceState.pluginId.c_str(), -1, destinationState.pluginId.c_str(), -1, TRUE) == CSTR_EQUAL &&
                                     NavigationLocation::EqualsNoCase(sourceState.instanceContext, destinationState.instanceContext);
            destinationFileSystem  = contextSame ? nullptr : destinationState.fileSystem;
        }
    }

    const std::wstring& sourcePluginIdForGate =
        sourcePluginIdOverride.empty() ? (sourcePane == Pane::Left ? _leftPane.pluginId : _rightPane.pluginId) : sourcePluginIdOverride;
    if (! destinationFileSystem &&
        ! CanSameFileSystemOperation(fileSystem, request.sourcePaths.front().native(), request.operation, sourcePluginIdForGate, request.flags))
    {
        Debug::Error(L"FolderWindow::StartFileOperationFromFolderView provider rejected same-filesystem operation plugin:{} op:{}.",
                     sourcePluginIdForGate,
                     static_cast<unsigned int>(request.operation));
        destinationState.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation,
                                                     FolderView::OverlaySeverity::Error,
                                                     LoadStringResource(nullptr, IDS_CAPTION_ERROR),
                                                     LoadStringResource(nullptr, IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS));
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const bool waitForOthers                = _fileOperations->ShouldQueueNewTask();
    std::filesystem::path destinationFolder = request.destinationFolder.value_or(std::filesystem::path{});
    PaneState& sourceStateForFocus           = sourcePane == Pane::Left ? _leftPane : _rightPane;
    const bool tracksRemovalFocus = request.operation == FILESYSTEM_DELETE || request.operation == FILESYSTEM_MOVE;
    const std::optional<FileSystemPathIdentity> removalPathIdentity =
        tracksRemovalFocus ? TryGetStablePathIdentity(fileSystem, request.sourcePaths.front().native(), sourcePluginIdForGate) : std::nullopt;
    const uint64_t removalFocusToken = removalPathIdentity.has_value()
        ? sourceStateForFocus.folderView.BeginRemovalFocusTracking(request.sourcePaths, removalPathIdentity.value())
        : 0u;
    std::function<void(HRESULT)> requestCompletion = std::move(request.completionCallback);
    const FileOperations::DeleteOrigin deleteOrigin = [&]() noexcept
    {
        switch (request.origin)
        {
            case FolderView::FileOperationRequest::Origin::PackCleanup: return FileOperations::DeleteOrigin::PackCleanup;
            case FolderView::FileOperationRequest::Origin::UnpackCleanup: return FileOperations::DeleteOrigin::UnpackCleanup;
            case FolderView::FileOperationRequest::Origin::PaneCommand:
            case FolderView::FileOperationRequest::Origin::ClipboardPaste:
            case FolderView::FileOperationRequest::Origin::InternalDrop:
            case FolderView::FileOperationRequest::Origin::ExternalDrop: return FileOperations::DeleteOrigin::PaneCommand;
            case FolderView::FileOperationRequest::Origin::InlineF2: return FileOperations::DeleteOrigin::PaneCommand;
        }
        return FileOperations::DeleteOrigin::PaneCommand;
    }();
    const std::optional<FileOperations::ConsentKind> capturedConsentKind = request.archiveDeleteAfterConfirmed
        ? std::optional<FileOperations::ConsentKind>{FileOperations::ConsentKind::ArchiveDeleteAfter}
        : std::nullopt;
    uint64_t taskId = 0;
    const FileOperationPromptDispatchScope promptDispatch(*this);
    HRESULT startHr = _fileOperations->AdmitOperation(request.operation,
                                                      sourcePane,
                                                      destinationPane,
                                                      fileSystem,
                                                      std::move(request.sourcePaths),
                                                      std::move(destinationFolder),
                                                      request.flags,
                                                      waitForOthers,
                                                      0,
                                                      FileOperationState::ExecutionMode::PerItem,
                                                      false,
                                                      std::move(destinationFileSystem),
                                                      &taskId,
                                                      {},
                                                      {},
                                                      std::move(sourcePluginIdOverride),
                                                      std::move(sourcePluginShortIdOverride),
                                                      std::move(sourceInstanceIdOverride),
                                                      deleteOrigin,
                                                      capturedConsentKind,
                                                      request.moveClipboardSequence,
                                                      std::move(request.preWorkerReleaseBarrier));
    if (SUCCEEDED(startHr) && taskId != 0u && (removalFocusToken != 0u || requestCompletion))
    {
        _fileOperationRequestCompletionCallbacks.insert_or_assign(
            taskId,
            [this, sourcePane, removalFocusToken, completion = std::move(requestCompletion)](const FileOperationCompletedEvent& event) mutable
        {
            PaneState& sourceState = sourcePane == Pane::Left ? _leftPane : _rightPane;
            const std::vector<FolderView::RemovalDisposition> dispositions = BuildDenseRemovedSourceDispositions(event);
            sourceState.folderView.CompleteRemovalFocusTracking(removalFocusToken, dispositions);
            if (completion)
            {
                completion(event.hr);
            }
        });
    }
    else
    {
        sourceStateForFocus.folderView.CompleteRemovalFocusTracking(removalFocusToken, std::span<const FolderView::RemovalDisposition>{});
    }
    return startHr;
}

std::optional<FolderWindow::Pane> FolderWindow::ResolveSourcePaneForResolvedPaths(std::wstring_view sourcePluginId,
                                                                                  std::wstring_view sourceInstanceContext,
                                                                                  const std::vector<std::filesystem::path>& sourcePaths) const noexcept
{
    if (sourcePluginId.empty() || sourcePaths.empty())
    {
        return std::nullopt;
    }

    const auto contextMatches = [&](const PaneState& state) noexcept -> bool
    {
        return CompareStringOrdinal(state.pluginId.c_str(), -1, sourcePluginId.data(), static_cast<int>(sourcePluginId.size()), TRUE) == CSTR_EQUAL &&
               NavigationLocation::EqualsNoCase(state.instanceContext, sourceInstanceContext);
    };

    const bool leftMatches  = contextMatches(_leftPane);
    const bool rightMatches = contextMatches(_rightPane);
    if (! leftMatches && ! rightMatches)
    {
        return std::nullopt;
    }

    const auto isSameOrUnderFolder = [](std::wstring_view folder, std::wstring_view path) noexcept -> bool
    {
        while (! folder.empty() && (folder.back() == L'\\' || folder.back() == L'/'))
        {
            folder.remove_suffix(1);
        }
        if (folder.empty() || path.size() < folder.size())
        {
            return false;
        }

        if (! OrdinalString::StartsWithNoCase(path, folder))
        {
            return false;
        }

        return path.size() == folder.size() || path[folder.size()] == L'\\' || path[folder.size()] == L'/';
    };

    const auto pathFitsPane = [&](Pane pane) noexcept -> bool
    {
        const PaneState& state                            = pane == Pane::Left ? _leftPane : _rightPane;
        const std::optional<std::filesystem::path> folder = state.folderView.GetFolderPath();
        return folder.has_value() && isSameOrUnderFolder(folder->native(), sourcePaths.front().native());
    };

    Pane sourcePane = Pane::Left;
    if (leftMatches && rightMatches)
    {
        const Pane focusedPane = GetFocusedPane();
        if (pathFitsPane(focusedPane))
        {
            sourcePane = focusedPane;
        }
        else if (pathFitsPane(_activePane))
        {
            sourcePane = _activePane;
        }
        else if (pathFitsPane(Pane::Left))
        {
            sourcePane = Pane::Left;
        }
        else if (pathFitsPane(Pane::Right))
        {
            sourcePane = Pane::Right;
        }
        else
        {
            sourcePane = focusedPane;
        }
    }
    else
    {
        sourcePane = leftMatches ? Pane::Left : Pane::Right;
    }

    return sourcePane;
}

std::optional<std::filesystem::path> FolderWindow::GetOtherPaneDestinationForResolvedPaths(std::wstring_view sourcePluginId,
                                                                                           std::wstring_view sourceInstanceContext,
                                                                                           const std::vector<std::filesystem::path>& sourcePaths) const noexcept
{
    const std::optional<Pane> sourcePane = ResolveSourcePaneForResolvedPaths(sourcePluginId, sourceInstanceContext, sourcePaths);
    if (! sourcePane.has_value())
    {
        return std::nullopt;
    }

    const Pane destinationPane        = OppositePane(sourcePane.value());
    const PaneState& destinationState = destinationPane == Pane::Left ? _leftPane : _rightPane;
    return destinationState.folderView.GetFolderPath();
}

HRESULT FolderWindow::StartFileOperationForResolvedPaths(std::wstring_view sourcePluginId,
                                                         std::wstring_view sourceInstanceContext,
                                                         FileSystemOperation operation,
                                                         std::vector<std::filesystem::path> sourcePaths,
                                                         FileSystemFlags flags,
                                                         bool requireConfirmation,
                                                         uint64_t* taskIdOut) noexcept
{
    if (taskIdOut)
    {
        *taskIdOut = 0;
    }

    if (sourcePluginId.empty())
    {
        return E_INVALIDARG;
    }

    if (sourcePaths.empty())
    {
        return S_FALSE;
    }

    EnsureFileOperations();
    if (! _fileOperations)
    {
        return E_FAIL;
    }

    const std::optional<Pane> resolvedSourcePane = ResolveSourcePaneForResolvedPaths(sourcePluginId, sourceInstanceContext, sourcePaths);
    if (! resolvedSourcePane.has_value())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const Pane sourcePane  = resolvedSourcePane.value();
    PaneState& sourceState = sourcePane == Pane::Left ? _leftPane : _rightPane;
    if (! sourceState.fileSystem)
    {
        return E_POINTER;
    }

    if (! CanSameFileSystemOperation(sourceState.fileSystem, sourcePaths.front().native(), operation, sourceState.pluginId, flags))
    {
        Debug::Error(L"FolderWindow::StartFileOperationForResolvedPaths provider rejected operation plugin:{} op:{}.",
                     sourceState.pluginId,
                     static_cast<unsigned int>(operation));
        sourceState.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation,
                                                FolderView::OverlaySeverity::Error,
                                                LoadStringResource(nullptr, IDS_CAPTION_ERROR),
                                                LoadStringResource(nullptr, IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS));
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const bool waitForOthers = _fileOperations->ShouldQueueNewTask();
    const FileOperationPromptDispatchScope promptDispatch(*this);
    return _fileOperations->AdmitOperation(operation,
                                           sourcePane,
                                           std::nullopt,
                                           sourceState.fileSystem,
                                           std::move(sourcePaths),
                                           {},
                                           flags,
                                           waitForOthers,
                                           0,
                                           FileOperationState::ExecutionMode::PerItem,
                                           requireConfirmation,
                                           nullptr,
                                           taskIdOut,
                                           {},
                                           {},
                                           {},
                                           {},
                                           std::nullopt,
                                           FileOperations::DeleteOrigin::FindResults);
}

HRESULT FolderWindow::StartFileOperationForResolvedItemsToOtherPane(std::wstring_view sourcePluginId,
                                                                    std::wstring_view sourceInstanceContext,
                                                                    FileSystemOperation operation,
                                                                    std::vector<ResolvedFileOperationItem> items,
                                                                    std::optional<std::filesystem::path>* outDestinationFolder,
                                                                    uint64_t* taskIdOut,
                                                                    std::wstring confirmationMessage) noexcept
{
    if (taskIdOut)
    {
        *taskIdOut = 0;
    }

    if (outDestinationFolder)
    {
        outDestinationFolder->reset();
    }

    if (sourcePluginId.empty() || (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE))
    {
        return E_INVALIDARG;
    }

    if (items.empty())
    {
        return S_FALSE;
    }

    std::vector<std::filesystem::path> sourcePaths;
    sourcePaths.reserve(items.size());
    for (const ResolvedFileOperationItem& item : items)
    {
        if (item.sourcePath.empty() || item.destinationPath.empty())
        {
            return E_INVALIDARG;
        }
        sourcePaths.push_back(item.sourcePath);
    }

    EnsureFileOperations();
    if (! _fileOperations)
    {
        return E_FAIL;
    }

    const std::optional<Pane> resolvedSourcePane = ResolveSourcePaneForResolvedPaths(sourcePluginId, sourceInstanceContext, sourcePaths);
    if (! resolvedSourcePane.has_value())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const Pane sourcePane = resolvedSourcePane.value();
    const Pane destPane   = OppositePane(sourcePane);
    PaneState& src        = sourcePane == Pane::Left ? _leftPane : _rightPane;
    PaneState& dest       = destPane == Pane::Left ? _leftPane : _rightPane;
    if (! src.fileSystem || ! dest.fileSystem)
    {
        return E_POINTER;
    }

    if (! SanityCheckBothPanes(src, dest, operation))
    {
        return E_FAIL;
    }

    const std::optional<std::filesystem::path> destinationFolder = dest.folderView.GetFolderPath();
    if (! destinationFolder.has_value())
    {
        return E_FAIL;
    }
    if (outDestinationFolder)
    {
        *outDestinationFolder = destinationFolder.value();
    }

    const bool waitForOthers                        = _fileOperations->ShouldQueueNewTask();
    const bool contextSame                          = CompareStringOrdinal(src.pluginId.c_str(), -1, dest.pluginId.c_str(), -1, TRUE) == CSTR_EQUAL &&
                                                      NavigationLocation::EqualsNoCase(src.instanceContext, dest.instanceContext);
    wil::com_ptr<IFileSystem> destinationFileSystem = contextSame ? nullptr : dest.fileSystem;
    if (! destinationFileSystem && ! CanSameFileSystemOperation(src.fileSystem, sourcePaths.front().native(), operation, src.pluginId))
    {
        Debug::Error(L"FolderWindow::StartFileOperationForResolvedItemsToOtherPane provider rejected same-filesystem operation plugin:{} op:{}.",
                     src.pluginId,
                     static_cast<unsigned int>(operation));
        src.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation,
                                        FolderView::OverlaySeverity::Error,
                                        LoadStringResource(nullptr, IDS_CAPTION_ERROR),
                                        LoadStringResource(nullptr, IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS));
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const bool requireConfirmation = operation == FILESYSTEM_MOVE;
    const FileOperationPromptDispatchScope promptDispatch(*this);
    return _fileOperations->AdmitOperation(operation,
                                           sourcePane,
                                           destPane,
                                           src.fileSystem,
                                           std::move(sourcePaths),
                                           destinationFolder.value(),
                                           FILESYSTEM_FLAG_NONE,
                                           waitForOthers,
                                           0,
                                           FileOperationState::ExecutionMode::PerItem,
                                           requireConfirmation,
                                           std::move(destinationFileSystem),
                                           taskIdOut,
                                           std::move(items),
                                           std::move(confirmationMessage));
}

HRESULT FolderWindow::StartFileOperationForResolvedPathsToOtherPane(std::wstring_view sourcePluginId,
                                                                    std::wstring_view sourceInstanceContext,
                                                                    FileSystemOperation operation,
                                                                    std::vector<std::filesystem::path> sourcePaths,
                                                                    FileSystemFlags flags,
                                                                    std::optional<std::filesystem::path>* outDestinationFolder,
                                                                    uint64_t* taskIdOut) noexcept
{
    if (taskIdOut)
    {
        *taskIdOut = 0;
    }

    if (outDestinationFolder)
    {
        outDestinationFolder->reset();
    }

    if (sourcePluginId.empty() || (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE))
    {
        return E_INVALIDARG;
    }

    if (sourcePaths.empty())
    {
        return S_FALSE;
    }

    EnsureFileOperations();
    if (! _fileOperations)
    {
        return E_FAIL;
    }

    const std::optional<Pane> resolvedSourcePane = ResolveSourcePaneForResolvedPaths(sourcePluginId, sourceInstanceContext, sourcePaths);
    if (! resolvedSourcePane.has_value())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const Pane sourcePane = resolvedSourcePane.value();
    const Pane destPane   = OppositePane(sourcePane);
    PaneState& src        = sourcePane == Pane::Left ? _leftPane : _rightPane;
    PaneState& dest       = destPane == Pane::Left ? _leftPane : _rightPane;
    if (! src.fileSystem || ! dest.fileSystem)
    {
        return E_POINTER;
    }

    if (! SanityCheckBothPanes(src, dest, operation))
    {
        return E_FAIL;
    }

    const std::optional<std::filesystem::path> destinationFolder = dest.folderView.GetFolderPath();
    if (! destinationFolder.has_value())
    {
        return E_FAIL;
    }
    if (outDestinationFolder)
    {
        *outDestinationFolder = destinationFolder.value();
    }

    const bool waitForOthers                        = _fileOperations->ShouldQueueNewTask();
    const bool contextSame                          = CompareStringOrdinal(src.pluginId.c_str(), -1, dest.pluginId.c_str(), -1, TRUE) == CSTR_EQUAL &&
                                                      NavigationLocation::EqualsNoCase(src.instanceContext, dest.instanceContext);
    wil::com_ptr<IFileSystem> destinationFileSystem = contextSame ? nullptr : dest.fileSystem;
    if (! destinationFileSystem && ! CanSameFileSystemOperation(src.fileSystem, sourcePaths.front().native(), operation, src.pluginId, flags))
    {
        Debug::Error(L"FolderWindow::StartFileOperationForResolvedPathsToOtherPane provider rejected same-filesystem operation plugin:{} op:{}.",
                     src.pluginId,
                     static_cast<unsigned int>(operation));
        src.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation,
                                        FolderView::OverlaySeverity::Error,
                                        LoadStringResource(nullptr, IDS_CAPTION_ERROR),
                                        LoadStringResource(nullptr, IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS));
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    const FileOperationPromptDispatchScope promptDispatch(*this);
    return _fileOperations->AdmitOperation(operation,
                                           sourcePane,
                                           destPane,
                                           src.fileSystem,
                                           std::move(sourcePaths),
                                           destinationFolder.value(),
                                           flags,
                                           waitForOthers,
                                           0,
                                           FileOperationState::ExecutionMode::PerItem,
                                           false,
                                           std::move(destinationFileSystem),
                                           taskIdOut);
}

HRESULT FolderWindow::StartFileOperationForResolvedPathsToDestination(std::wstring_view sourcePluginId,
                                                                      std::wstring_view sourceInstanceContext,
                                                                      FileSystemOperation operation,
                                                                      std::vector<std::filesystem::path> sourcePaths,
                                                                      std::filesystem::path destinationFolder,
                                                                      FileSystemFlags flags,
                                                                      uint64_t* taskIdOut) noexcept
{
    if (taskIdOut)
    {
        *taskIdOut = 0;
    }

    if (sourcePluginId.empty() || (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE))
    {
        return E_INVALIDARG;
    }

    if (sourcePaths.empty())
    {
        return S_FALSE;
    }

    if (destinationFolder.empty())
    {
        return E_INVALIDARG;
    }

    EnsureFileOperations();
    if (! _fileOperations)
    {
        return E_FAIL;
    }

    const std::optional<Pane> resolvedSourcePane = ResolveSourcePaneForResolvedPaths(sourcePluginId, sourceInstanceContext, sourcePaths);
    if (! resolvedSourcePane.has_value())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const Pane sourcePane = resolvedSourcePane.value();
    PaneState& src        = sourcePane == Pane::Left ? _leftPane : _rightPane;
    if (! src.fileSystem)
    {
        return E_POINTER;
    }

    if (! CanSameFileSystemOperation(src.fileSystem, sourcePaths.front().native(), operation, src.pluginId, flags))
    {
        Debug::Error(L"FolderWindow::StartFileOperationForResolvedPathsToDestination provider rejected same-filesystem operation plugin:{} op:{}.",
                     src.pluginId,
                     static_cast<unsigned int>(operation));
        src.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation,
                                        FolderView::OverlaySeverity::Error,
                                        LoadStringResource(nullptr, IDS_CAPTION_ERROR),
                                        LoadStringResource(nullptr, IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS));
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const bool waitForOthers = _fileOperations->ShouldQueueNewTask();
    const FileOperationPromptDispatchScope promptDispatch(*this);
    return _fileOperations->AdmitOperation(operation,
                                           sourcePane,
                                           std::nullopt,
                                           src.fileSystem,
                                           std::move(sourcePaths),
                                           std::move(destinationFolder),
                                           flags,
                                           waitForOthers,
                                           0,
                                           FileOperationState::ExecutionMode::PerItem,
                                           false,
                                           nullptr,
                                           taskIdOut);
}

void FolderWindow::ShutdownFileOperations() noexcept
{
    _fileOperationsDeferredCloseTarget = nullptr;
    if (_fileOperations)
    {
        _fileOperations->Shutdown();
    }

    // HostShowPrompt pumps FolderWindow messages. If shutdown is dispatched from that
    // nested loop, keep only the already-shut-down state object alive until the prompt
    // handler unwinds; its task storage has already been canceled, joined, and cleared.
    if (_fileOperationPromptDispatchDepth != 0u)
    {
        _fileOperationResetDeferred = true;
        return;
    }

    _fileOperationResetDeferred = false;
    _fileOperations.reset();
}

void FolderWindow::ApplyFileOperationsTheme() noexcept
{
    if (_fileOperations)
    {
        _fileOperations->ApplyTheme(_theme);
    }
}

void FolderWindow::CommandToggleFileOperationsIssuesPane()
{
    EnsureFileOperations();
    if (! _fileOperations)
    {
        return;
    }

    _fileOperations->ToggleIssuesPane();
}

void FolderWindow::CommandShowFileOperations()
{
    EnsureFileOperations();
    if (_fileOperations)
    {
        _fileOperations->ShowPopup();
    }
}

bool FolderWindow::IsFileOperationsIssuesPaneVisible() noexcept
{
    if (! _fileOperations)
    {
        return false;
    }

    return _fileOperations->IsIssuesPaneVisible();
}

#ifdef ENABLE_TESTS
FolderWindow::FileOperationState* FolderWindow::DebugGetFileOperationState() noexcept
{
    EnsureFileOperations();
    return _fileOperations.get();
}

bool FolderWindow::DebugWasFileOperationStateRetainedDuringNestedPromptShutdown() const noexcept
{
    return _debugFileOperationStateRetainedDuringNestedPromptShutdown;
}

void FolderWindow::DebugSetFileOperationRequestCallbackEnabled(Pane pane, bool enabled) noexcept
{
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! enabled)
    {
        state.folderView.SetFileOperationRequestCallback({});
        return;
    }

    state.folderView.SetFileOperationRequestCallback([this, pane](FolderView::FileOperationRequest request) noexcept -> HRESULT
    { return StartFileOperationFromFolderView(pane, std::move(request)); });
}
#endif

bool FolderWindow::ConfirmCancelAllFileOperations(HWND ownerWindow) noexcept
{
    if (! _fileOperations || ! _fileOperations->HasActiveOperations())
    {
        return true;
    }

    // The caller's own window is the one to close again once the drain is quiet; the prompt owner
    // below may fall back to this folder window.
    const HWND closeTarget = ownerWindow;
    if (! ownerWindow)
    {
        ownerWindow = _hWnd.get();
    }

    const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_FILEOPS_EXIT);
    const std::wstring message = LoadStringResource(nullptr, IDS_MSG_FILEOPS_CANCEL_ALL_EXIT);

    HostPromptRequest prompt{};
    prompt.sizeBytes     = sizeof(prompt);
    prompt.scope         = HOST_ALERT_SCOPE_WINDOW;
    prompt.severity      = HOST_ALERT_INFO;
    prompt.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
    prompt.targetWindow  = ownerWindow;
    prompt.title         = title.c_str();
    prompt.message       = message.c_str();
    prompt.defaultResult = HOST_PROMPT_RESULT_CANCEL;

    const FileOperationPromptDispatchScope promptDispatch(*this);
    HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
    const HRESULT hrPrompt        = HostShowPrompt(prompt, nullptr, &promptResult);
    if (FAILED(hrPrompt) || promptResult != HOST_PROMPT_RESULT_OK)
    {
        return false;
    }

    return CancelAllFileOperationsThenClose(closeTarget);
}

bool FolderWindow::CancelAllFileOperationsThenClose(HWND closeTarget) noexcept
{
    if (! _fileOperations || ! _fileOperations->HasActiveOperations())
    {
        _fileOperationsDeferredCloseTarget = nullptr;
        return true;
    }
    // C1: cancellation drains off the UI thread. The window stays alive, its popup shows every
    // task Stopping, and the close resumes from the completion path once the last task has been
    // reaped there; the WM_DESTROY join then finds nothing to wait for.
    _fileOperations->CancelAll();
    _fileOperationsDeferredCloseTarget = closeTarget != nullptr ? closeTarget : GetAncestor(_hWnd.get(), GA_ROOT);
    return false;
}

void FolderWindow::ResumeDeferredCloseIfFileOperationsQuiet() noexcept
{
    if (_fileOperationsDeferredCloseTarget == nullptr || (_fileOperations && _fileOperations->HasActiveOperations()))
    {
        return;
    }
    const HWND closeTarget             = _fileOperationsDeferredCloseTarget;
    _fileOperationsDeferredCloseTarget = nullptr;
    if (IsWindow(closeTarget) == FALSE || PostMessageW(closeTarget, WM_CLOSE, 0, 0) == FALSE)
    {
        Debug::Warning(L"FolderWindow: the deferred application close could not resume after file operations went quiet.");
    }
}

void FolderWindow::CommandDelete(Pane pane)
{
    SetActivePane(pane);
    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! _fileOperations)
    {
        state.folderView.CommandDelete();
        return;
    }

    if (! state.fileSystem)
    {
        return;
    }

    std::vector<std::filesystem::path> paths = state.folderView.GetSelectedOrFocusedPaths();
    if (paths.empty())
    {
        return;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_USE_RECYCLE_BIN);

    if (! CanSameFileSystemOperation(state.fileSystem, paths.front().native(), FILESYSTEM_DELETE, state.pluginId, flags))
    {
        Debug::Error(L"FolderWindow::CommandDelete provider rejected delete plugin:{}.", state.pluginId);
        state.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation,
                                          FolderView::OverlaySeverity::Error,
                                          LoadStringResource(nullptr, IDS_CAPTION_ERROR),
                                          LoadStringResource(nullptr, IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS));
        return;
    }

    // Recycle deletes go through ONE bulk DeleteItems call: the plugin batches same-parent
    // items into a single IFileOperation (up to 1000 per batch), which per-item routing would
    // degrade to one shell operation + one STA thread per item.
    const bool waitForOthers = _fileOperations->ShouldQueueNewTask();
    const std::optional<FileSystemPathIdentity> removalPathIdentity =
        TryGetStablePathIdentity(state.fileSystem, paths.front().native(), state.pluginId);
    const uint64_t removalFocusToken = removalPathIdentity.has_value()
        ? state.folderView.BeginRemovalFocusTracking(paths, removalPathIdentity.value())
        : 0u;
    uint64_t taskId                  = 0u;
    const FileOperationPromptDispatchScope promptDispatch(*this);
    const HRESULT startHr            = _fileOperations->AdmitOperation(FILESYSTEM_DELETE,
                                                            pane,
                                                            std::nullopt,
                                                            state.fileSystem,
                                                            std::move(paths),
                                                            {},
                                                            flags,
                                                            waitForOthers,
                                                            0,
                                                            FileOperationState::ExecutionMode::BulkItems,
                                                            false,
                                                            nullptr,
                                                            &taskId);
    if (SUCCEEDED(startHr) && taskId != 0u && removalFocusToken != 0u)
    {
        _fileOperationRequestCompletionCallbacks.insert_or_assign(taskId, [this, pane, removalFocusToken](const FileOperationCompletedEvent& event)
        {
            PaneState& sourceState = pane == Pane::Left ? _leftPane : _rightPane;
            const std::vector<FolderView::RemovalDisposition> dispositions = BuildDenseRemovedSourceDispositions(event);
            sourceState.folderView.CompleteRemovalFocusTracking(removalFocusToken, dispositions);
        });
    }
    else
    {
        state.folderView.CompleteRemovalFocusTracking(removalFocusToken, std::span<const FolderView::RemovalDisposition>{});
    }
}

void FolderWindow::CommandPermanentDelete(Pane pane)
{
    SetActivePane(pane);
    EnsureFileOperations();

    PaneState& state = pane == Pane::Left ? _leftPane : _rightPane;
    if (! _fileOperations || ! state.fileSystem)
    {
        return;
    }

    std::vector<std::filesystem::path> paths = state.folderView.GetSelectedOrFocusedPaths();
    if (paths.empty())
    {
        return;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);

    if (! CanSameFileSystemOperation(state.fileSystem, paths.front().native(), FILESYSTEM_DELETE, state.pluginId, flags))
    {
        Debug::Error(L"FolderWindow::CommandPermanentDelete provider rejected delete plugin:{}.", state.pluginId);
        state.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation,
                                          FolderView::OverlaySeverity::Error,
                                          LoadStringResource(nullptr, IDS_CAPTION_ERROR),
                                          LoadStringResource(nullptr, IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS));
        return;
    }

    const bool waitForOthers = _fileOperations->ShouldQueueNewTask();
    const std::optional<FileSystemPathIdentity> removalPathIdentity =
        TryGetStablePathIdentity(state.fileSystem, paths.front().native(), state.pluginId);
    const uint64_t removalFocusToken = removalPathIdentity.has_value()
        ? state.folderView.BeginRemovalFocusTracking(paths, removalPathIdentity.value())
        : 0u;
    uint64_t taskId                  = 0u;
    const FileOperationPromptDispatchScope promptDispatch(*this);
    const HRESULT startHr            = _fileOperations->AdmitOperation(FILESYSTEM_DELETE,
                                                            pane,
                                                            std::nullopt,
                                                            state.fileSystem,
                                                            std::move(paths),
                                                            {},
                                                            flags,
                                                            waitForOthers,
                                                            0,
                                                            FileOperationState::ExecutionMode::PerItem,
                                                            true,
                                                            nullptr,
                                                            &taskId);
    if (SUCCEEDED(startHr) && taskId != 0u && removalFocusToken != 0u)
    {
        _fileOperationRequestCompletionCallbacks.insert_or_assign(taskId, [this, pane, removalFocusToken](const FileOperationCompletedEvent& event)
        {
            PaneState& sourceState = pane == Pane::Left ? _leftPane : _rightPane;
            const std::vector<FolderView::RemovalDisposition> dispositions = BuildDenseRemovedSourceDispositions(event);
            sourceState.folderView.CompleteRemovalFocusTracking(removalFocusToken, dispositions);
        });
    }
    else
    {
        state.folderView.CompleteRemovalFocusTracking(removalFocusToken, std::span<const FolderView::RemovalDisposition>{});
    }
}

bool FolderWindow::SanityCheckBothPanes(FolderWindow::PaneState& src, FolderWindow::PaneState& dest, FileSystemOperation operation)
{
    bool ok             = true;
    bool sameFolder     = false;
    bool contextsDiffer = false;
    bool destinationUnsettled = false;
    const std::optional<std::filesystem::path> sourceFolder = src.folderView.GetFolderPath();
    const std::optional<std::filesystem::path> destinationFolder = dest.folderView.GetFolderPath();
    if (! _fileOperations)
    {
        Debug::Error(L"FolderWindow::SanityCheckBothPanes No active file operations.");
        ok = false;
    }

    if (ok && (! src.fileSystem || ! dest.fileSystem))
    {
        Debug::Error(L"FolderWindow::SanityCheckBothPanes Source or destination pane has no file system.");
        ok = false;
    }

    if (ok && (src.pluginId.empty() || dest.pluginId.empty()))
    {
        Debug::Error(L"FolderWindow::SanityCheckBothPanes Source or destination pane has no file system metadata.");
        ok = false;
    }

    if (ok && (! sourceFolder.has_value() || ! destinationFolder.has_value()))
    {
        Debug::Error(L"FolderWindow::SanityCheckBothPanes No source or destination path.");
        ok = false;
    }

    if (ok && ! dest.folderView.IsCurrentFolderEnumerated())
    {
        Debug::Warning(L"FolderWindow::SanityCheckBothPanes rejected an operation while the destination folder is still loading.");
        destinationUnsettled = true;
        ok                   = false;
    }

    if (ok)
    {
        const bool contextSame = CompareStringOrdinal(src.pluginId.c_str(), -1, dest.pluginId.c_str(), -1, TRUE) == CSTR_EQUAL &&
                                 NavigationLocation::EqualsNoCase(src.instanceContext, dest.instanceContext);
        contextsDiffer         = ! contextSame;

        if (operation == FILESYSTEM_MOVE && contextSame && sourceFolder.has_value() && destinationFolder.has_value() &&
            NavigationLocation::EqualsNoCase(sourceFolder->native(), destinationFolder->native()))
        {
            Debug::Error(L"FolderWindow::SanityCheckBothPanes Move source and destination folder are the same: {}.", sourceFolder->native());
            sameFolder = true;
            ok         = false;
        }
    }

    if (ok && contextsDiffer && (operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE))
    {
        if (! CanCrossFileSystemCopyMove(src.fileSystem,
                                         sourceFolder->native(),
                                         src.pluginId,
                                         dest.fileSystem,
                                         destinationFolder->native(),
                                         dest.pluginId,
                                         operation))
        {
            Debug::Error(L"FolderWindow::SanityCheckBothPanes Cross-filesystem operation not allowed src:{} dest:{} op:{}.",
                         src.pluginId,
                         dest.pluginId,
                         static_cast<unsigned int>(operation));
            ok = false;
        }
    }
    else if (ok && ! contextsDiffer &&
             ! CanSameFileSystemOperation(src.fileSystem, sourceFolder->native(), operation, src.pluginId))
    {
        Debug::Error(L"FolderWindow::SanityCheckBothPanes provider rejected same-filesystem operation plugin:{} op:{}.",
                     src.pluginId,
                     static_cast<unsigned int>(operation));
        contextsDiffer = true;
        ok             = false;
    }

    if (! ok && _hWnd)
    {
        const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
        int messageId              = destinationUnsettled ? IDS_MSG_PANE_OP_DESTINATION_LOADING
                                     : sameFolder        ? IDS_MSG_PANE_OP_REQUIRES_DIFFERENT_FOLDER
                                     : contextsDiffer    ? IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS
                                                         : IDS_MSG_PANE_OP_REQUIRES_SAME_FS;
        const std::wstring message = LoadStringResource(nullptr, static_cast<UINT>(messageId));
        src.folderView.ShowAlertOverlay(FolderView::ErrorOverlayKind::Operation, FolderView::OverlaySeverity::Error, title, message);
        return false;
    }

    return ok;
}

void FolderWindow::CommandCopyToOtherPane(Pane sourcePane, bool withOptions)
{
    SetActivePane(sourcePane);
    const Pane destPane = sourcePane == Pane::Left ? Pane::Right : Pane::Left;

    PaneState& src  = sourcePane == Pane::Left ? _leftPane : _rightPane;
    PaneState& dest = destPane == Pane::Left ? _leftPane : _rightPane;

    if (! SanityCheckBothPanes(src, dest, FILESYSTEM_COPY))
    {
        return;
    }

    std::vector<std::filesystem::path> paths = src.folderView.GetSelectedOrFocusedPaths();
    if (paths.empty())
    {
        auto srcPath = src.currentPath.has_value() ? src.currentPath.value().c_str() : L"(unknown)";
        Debug::Error(L"FolderWindow::CommandCopyToOtherPane No selected paths: {}", srcPath);
        return;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);

    const bool waitForOthers                        = _fileOperations->ShouldQueueNewTask();
    const bool contextSame                          = CompareStringOrdinal(src.pluginId.c_str(), -1, dest.pluginId.c_str(), -1, TRUE) == CSTR_EQUAL &&
                                                      NavigationLocation::EqualsNoCase(src.instanceContext, dest.instanceContext);
    wil::com_ptr<IFileSystem> destinationFileSystem = contextSame ? nullptr : dest.fileSystem;
    const FileOperationPromptDispatchScope promptDispatch(*this);
    static_cast<void>(_fileOperations->AdmitOperation(FILESYSTEM_COPY,
                                                      sourcePane,
                                                      destPane,
                                                      src.fileSystem,
                                                      std::move(paths),
                                                      dest.folderView.GetFolderPath().value(),
                                                      flags,
                                                      waitForOthers,
                                                      0,
                                                      FileOperationState::ExecutionMode::PerItem,
                                                      withOptions,
                                                      std::move(destinationFileSystem)));
}

void FolderWindow::CommandMoveToOtherPane(Pane sourcePane, bool withOptions)
{
    SetActivePane(sourcePane);
    const Pane destPane = sourcePane == Pane::Left ? Pane::Right : Pane::Left;

    PaneState& src  = sourcePane == Pane::Left ? _leftPane : _rightPane;
    PaneState& dest = destPane == Pane::Left ? _leftPane : _rightPane;

    if (! SanityCheckBothPanes(src, dest, FILESYSTEM_MOVE))
    {
        return;
    }

    std::vector<std::filesystem::path> paths = src.folderView.GetSelectedOrFocusedPaths();
    if (paths.empty())
    {
        auto srcPath = src.currentPath.has_value() ? src.currentPath.value().c_str() : L"(unknown)";
        Debug::Error(L"FolderWindow::CommandMoveToOtherPane No selected paths: {}", srcPath);
        return;
    }

    const FileSystemFlags flags = static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE);

    const bool waitForOthers                        = _fileOperations->ShouldQueueNewTask();
    const bool contextSame                          = CompareStringOrdinal(src.pluginId.c_str(), -1, dest.pluginId.c_str(), -1, TRUE) == CSTR_EQUAL &&
                                                      NavigationLocation::EqualsNoCase(src.instanceContext, dest.instanceContext);
    wil::com_ptr<IFileSystem> destinationFileSystem = contextSame ? nullptr : dest.fileSystem;
    const std::optional<FileSystemPathIdentity> removalPathIdentity =
        TryGetStablePathIdentity(src.fileSystem, paths.front().native(), src.pluginId);
    const uint64_t removalFocusToken = removalPathIdentity.has_value()
        ? src.folderView.BeginRemovalFocusTracking(paths, removalPathIdentity.value())
        : 0u;
    uint64_t taskId = 0u;
    const FileOperationPromptDispatchScope promptDispatch(*this);
    const HRESULT startHr = _fileOperations->AdmitOperation(FILESYSTEM_MOVE,
                                                            sourcePane,
                                                            destPane,
                                                            src.fileSystem,
                                                            std::move(paths),
                                                            dest.folderView.GetFolderPath().value(),
                                                            flags,
                                                            waitForOthers,
                                                            0,
                                                            FileOperationState::ExecutionMode::PerItem,
                                                            withOptions,
                                                            std::move(destinationFileSystem),
                                                            &taskId);
    if (SUCCEEDED(startHr) && taskId != 0u && removalFocusToken != 0u)
    {
        _fileOperationRequestCompletionCallbacks.insert_or_assign(taskId, [this, sourcePane, removalFocusToken](const FileOperationCompletedEvent& event)
        {
            PaneState& sourceState = sourcePane == Pane::Left ? _leftPane : _rightPane;
            const std::vector<FolderView::RemovalDisposition> dispositions = BuildDenseRemovedSourceDispositions(event);
            sourceState.folderView.CompleteRemovalFocusTracking(removalFocusToken, dispositions);
        });
    }
    else
    {
        src.folderView.CompleteRemovalFocusTracking(removalFocusToken, std::span<const FolderView::RemovalDisposition>{});
    }
}

LRESULT FolderWindow::OnFileOperationCompleted(WPARAM wp, LPARAM lp) noexcept
{
    auto posted = TakeMessagePayload<FileOperationState::TaskCompletedPayload>(lp);
    FileOperationState::TaskCompletedPayload payload{};
    if (posted)
    {
        payload = *posted;
    }
    else if (! _fileOperations || ! _fileOperations->TakeFallbackCompletedPayload(static_cast<uint64_t>(wp), payload))
    {
        return 0;
    }

    ApplyFileOperationCompletion(payload.taskId, payload.hr, payload.warningCount, payload.errorCount);
    ResumeDeferredCloseIfFileOperationsQuiet();
    return 0;
}

LRESULT FolderWindow::OnFileOperationBatchRenameArtifactPrompt(const WPARAM wp, const LPARAM lp) noexcept
{
    const uint64_t taskId = static_cast<uint64_t>(wp);
    auto payload = TakeMessagePayload<FileOperationState::BatchRenameArtifactPromptPayload>(lp);
    if (! _fileOperations)
    {
        return 0;
    }

    if (! payload)
    {
        _fileOperations->CompleteBatchRenameArtifactPromptByTaskId(
            taskId, HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS));
        return 0;
    }

    const FileOperationPromptDispatchScope promptDispatch(*this);
    _fileOperations->OnBatchRenameArtifactPrompt(std::move(payload));
    return 0;
}

LRESULT FolderWindow::OnFileOperationClipboardMoveReady(const WPARAM wp, const LPARAM lp) noexcept
{
    const uint64_t taskId = static_cast<uint64_t>(wp);
    auto payload = TakeMessagePayload<FileOperationState::ClipboardMoveReadyPayload>(lp);
    if (! _fileOperations)
    {
        return 0;
    }

    if (! payload)
    {
        _fileOperations->CompleteClipboardMoveAdmissionByTaskId(
            taskId, HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS));
        return 0;
    }

    // Clipboard/OLE can dispatch nested window messages. Retain FileOperationState during the
    // callback; the state method itself re-finds the task after the callback returns.
    const FileOperationPromptDispatchScope promptDispatch(*this);
    _fileOperations->OnClipboardMoveReady(std::move(payload));
    return 0;
}

void FolderWindow::OnFileOperationPresentationChanged() noexcept
{
    if (! _fileOperations)
    {
        KillTimer(_hWnd.get(), kFileOperationPresentationTimerId);
        return;
    }

    const std::optional<UINT> nextDelay = _fileOperations->RefreshDeferredTaskPresentation();
    if (! nextDelay.has_value())
    {
        KillTimer(_hWnd.get(), kFileOperationPresentationTimerId);
        return;
    }

    if (SetTimer(_hWnd.get(), kFileOperationPresentationTimerId, (std::max)(1u, nextDelay.value()), nullptr) == 0)
    {
        _fileOperations->RevealPendingTasksImmediately();
    }
}

void FolderWindow::ApplyFileOperationCompletion(uint64_t taskId, HRESULT hr, unsigned long warningCount, unsigned long errorCount) noexcept
{
    if (! _fileOperations)
    {
        return;
    }

#ifdef ENABLE_TESTS
    _fileOperations->DebugRecordCompletionApplyTidForSelfTest(GetCurrentThreadId());
    if (FileOperationsSelfTest::IsRunning())
    {
        FileOperationsSelfTest::NotifyTaskCompleted(taskId, hr);
    }
#endif

    FileOperationState::Task* task = _fileOperations->FindTask(taskId);
    if (! task)
    {
        if (auto completionIt = _fileOperationRequestCompletionCallbacks.find(taskId);
            completionIt != _fileOperationRequestCompletionCallbacks.end())
        {
            FileOperationCompletedCallback completion = std::move(completionIt->second);
            _fileOperationRequestCompletionCallbacks.erase(completionIt);
            if (completion)
            {
                FileOperationCompletedEvent missingTaskEvent{};
                missingTaskEvent.taskId = taskId;
                missingTaskEvent.hr = hr;
                completion(missingTaskEvent);
            }
        }
        return;
    }

    const Pane sourcePane                     = task->GetSourcePane();
    const std::optional<Pane> destinationPane = task->GetDestinationPane();
    const bool hasResolvedItems               = task->_resolvedItems.size() == task->_sourcePaths.size();
    const std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans = task->LoadPlans();

    const auto pathIdentityForIndex = [&](size_t index, bool destination) noexcept -> std::optional<FileSystemPathIdentity>
    {
        if (! plans)
        {
            return std::nullopt;
        }

        size_t currentIndex = 0u;
        for (const FileOperations::FileOperationPlan& plan : *plans)
        {
            if (const auto* transfer = std::get_if<FileOperations::TransferPlan>(&plan))
            {
                for ([[maybe_unused]] const FileOperations::QualifiedSourceItem& item : transfer->selectedItems)
                {
                    if (currentIndex == index)
                    {
                        return destination ? transfer->destinationEndpoint.pathIdentity : transfer->sourceEndpoint.pathIdentity;
                    }
                    ++currentIndex;
                }
                continue;
            }
            if (const auto* rename = std::get_if<FileOperations::RenamePlan>(&plan))
            {
                for ([[maybe_unused]] const FileOperations::RenameStep& step : rename->finalMappings)
                {
                    if (currentIndex == index)
                    {
                        return rename->endpoint.pathIdentity;
                    }
                    ++currentIndex;
                }
                continue;
            }
            if (const auto* deletePlan = std::get_if<FileOperations::DeletePlan>(&plan))
            {
                for ([[maybe_unused]] const FileOperations::QualifiedSourceItem& item : deletePlan->selectedItems)
                {
                    if (currentIndex == index)
                    {
                        return destination ? std::nullopt : deletePlan->endpoint.pathIdentity;
                    }
                    ++currentIndex;
                }
            }
        }
        return std::nullopt;
    };

    const auto resolvedProviderParentForIndex = [&](size_t index,
                                                    const std::filesystem::path& path,
                                                    bool destination) noexcept -> std::optional<std::filesystem::path>
    {
        const std::optional<FileSystemPathIdentity> identity = pathIdentityForIndex(index, destination);
        if (! identity.has_value())
        {
            return std::nullopt;
        }

        std::wstring providerParentPath;
        if (! TryGetFileSystemParentPath(identity.value(), path.native(), providerParentPath))
        {
            return std::nullopt;
        }
        return std::filesystem::path(std::move(providerParentPath));
    };

    const auto resolvedDestinationForIndex = [&](size_t index) noexcept -> std::optional<std::filesystem::path>
    {
        if ((task->GetOperation() == FILESYSTEM_COPY || task->GetOperation() == FILESYSTEM_MOVE) && plans)
        {
            size_t currentIndex = 0u;
            for (const FileOperations::FileOperationPlan& plan : *plans)
            {
                const auto* transfer = std::get_if<FileOperations::TransferPlan>(&plan);
                if (transfer == nullptr)
                {
                    return std::nullopt;
                }
                if (index < currentIndex + transfer->selectedItems.size())
                {
                    std::wstring destination;
                    if (! FileOperations::TryResolveTransferDestinationProviderPath(*transfer, index - currentIndex, destination))
                    {
                        return std::nullopt;
                    }
                    if (hasResolvedItems &&
                        (! transfer->destinationEndpoint.pathIdentity.has_value() ||
                         ! EquivalentPath(transfer->destinationEndpoint.pathIdentity.value(),
                                          destination,
                                          task->_resolvedItems[index].destinationPath.native())))
                    {
                        return std::nullopt;
                    }
                    return std::filesystem::path(std::move(destination));
                }
                currentIndex += transfer->selectedItems.size();
            }
            return std::nullopt;
        }

        if (task->GetOperation() == FILESYSTEM_RENAME && plans)
        {
            size_t currentIndex = 0u;
            for (const FileOperations::FileOperationPlan& plan : *plans)
            {
                const auto* rename = std::get_if<FileOperations::RenamePlan>(&plan);
                if (rename == nullptr)
                {
                    return std::nullopt;
                }
                for (const FileOperations::RenameStep& step : rename->finalMappings)
                {
                    if (currentIndex == index)
                    {
                        return step.providerJoinedPath.empty() ? std::nullopt
                                                               : std::optional<std::filesystem::path>(step.providerJoinedPath);
                    }
                    ++currentIndex;
                }
            }
        }

        return std::nullopt;
    };

    const auto resolvedRenameParentForIndex = [&](size_t index) noexcept -> std::optional<std::filesystem::path>
    {
        if (task->GetOperation() != FILESYSTEM_RENAME || index >= task->_sourcePaths.size())
        {
            return std::nullopt;
        }
        return resolvedProviderParentForIndex(index, task->_sourcePaths[index], false);
    };

    const auto isResolvedDirectoryShell = [&](size_t index) noexcept -> bool
    { return hasResolvedItems && task->_resolvedItems[index].kind == ResolvedFileOperationItemKind::DirectoryShell; };

    FileOperationCompletedEvent e{};
    e.taskId            = taskId;
    e.operation         = task->GetOperation();
    e.sourcePane        = sourcePane;
    e.destinationPane   = destinationPane;
    e.sourcePaths       = task->_sourcePaths;
    {
        std::scoped_lock lock(task->_sourceItemStatusMutex);
        e.itemOutcomes.reserve(task->_sourceItemResultBuilders.size());
        for (const FileOperationState::Task::SourceItemResultBuilder& builder : task->_sourceItemResultBuilders)
        {
            if (builder.terminal.has_value())
            {
                const FileOperations::FileOperationItemResult& result = builder.terminal.value();
                e.itemOutcomes.push_back(FileOperationItemOutcome{
                    .sourceIndex = result.sourceIndex,
                    .publication = result.publication,
                    .verification = result.verification,
                    .sourceDisposition = result.sourceDisposition,
                    .completion = result.completion,
                    .ownedStageDisposition = result.ownedStageDisposition,
                    .status = result.status,
                    .finalSourcePath = std::filesystem::path(result.finalSourcePath),
                    .finalDestinationPath = std::filesystem::path(result.finalDestinationPath),
                });
            }
        }
    }
    e.destinationFolder = task->GetDestinationFolder();
    if ((destinationPane.has_value() && (task->GetOperation() == FILESYSTEM_COPY || task->GetOperation() == FILESYSTEM_MOVE)) ||
        task->GetOperation() == FILESYSTEM_RENAME)
    {
        e.destinationPaths.reserve(task->_sourcePaths.size());
        for (size_t index = 0; index < task->_sourcePaths.size(); ++index)
        {
            if (const auto destination = resolvedDestinationForIndex(index); destination.has_value())
            {
                e.destinationPaths.push_back(destination.value());
            }
        }
    }
    e.hr = hr;

    if (auto completionIt = _fileOperationRequestCompletionCallbacks.find(taskId);
        completionIt != _fileOperationRequestCompletionCallbacks.end())
    {
        FileOperationCompletedCallback completion = std::move(completionIt->second);
        _fileOperationRequestCompletionCallbacks.erase(completionIt);
        if (completion)
        {
            completion(e);
        }
    }

    if (! _fileOperationCompletedCallbacks.empty())
    {
        // Iterate over a copy: a callback may unsubscribe (or subscribe) while handling the event.
        const auto subscriptions = _fileOperationCompletedCallbacks;
        for (const FileOperationCompletedSubscription& subscription : subscriptions)
        {
            if (subscription.hasLifetimeGuard && subscription.lifetimeGuard.expired())
            {
                continue;
            }
            if (subscription.callback)
            {
                subscription.callback(e);
            }
        }
    }

    PaneState& src            = sourcePane == Pane::Left ? _leftPane : _rightPane;
    DirectoryInfoCache& cache = DirectoryInfoCache::GetInstance();

    const auto trimTrailingSeparators = [](std::wstring_view path) noexcept -> std::wstring_view
    {
        while (path.size() > 1u && (path.back() == L'\\' || path.back() == L'/'))
        {
            path.remove_suffix(1);
        }
        return path;
    };

    const auto foldersEqual = [&](const std::filesystem::path& left, const std::filesystem::path& right) noexcept -> bool
    {
        const std::wstring_view leftText  = trimTrailingSeparators(left.native());
        const std::wstring_view rightText = trimTrailingSeparators(right.native());
        return NavigationLocation::EqualsNoCase(leftText, rightText);
    };

    const auto forceRefreshIfShowingFolder = [&](PaneState& paneState, const std::filesystem::path& folder) noexcept
    {
        if (folder.empty())
        {
            return;
        }

        const auto paneFolder = paneState.folderView.GetFolderPath();
        if (paneFolder.has_value() && foldersEqual(paneFolder.value(), folder))
        {
            paneState.folderView.ForceRefresh();
        }
    };

    const auto forceRefreshVisibleFolder = [&](const std::filesystem::path& folder) noexcept
    {
        forceRefreshIfShowingFolder(_leftPane, folder);
        forceRefreshIfShowingFolder(_rightPane, folder);
    };

    const auto forceRefreshVisibleSourceParents = [&]() noexcept
    {
        for (size_t index = 0; index < task->_sourcePaths.size(); ++index)
        {
            const auto parent = resolvedProviderParentForIndex(index, task->_sourcePaths[index], false);
            if (parent.has_value() && ! parent.value().empty())
            {
                forceRefreshVisibleFolder(parent.value());
            }
        }
    };

    const auto outcomeForIndex = [&](size_t index) noexcept -> const FileOperationItemOutcome*
    {
        const auto it = std::ranges::find_if(e.itemOutcomes, [index](const FileOperationItemOutcome& outcome) noexcept
        { return outcome.sourceIndex == index; });
        return it == e.itemOutcomes.end() ? nullptr : std::addressof(*it);
    };
    const auto sourceWasRemoved = [&](size_t index) noexcept
    {
        const FileOperationItemOutcome* const outcome = outcomeForIndex(index);
        return outcome != nullptr && outcome->sourceDisposition == FileOperations::SourceDisposition::Removed;
    };
    const auto destinationNeedsRefresh = [&](size_t index) noexcept
    {
        const FileOperationItemOutcome* const outcome = outcomeForIndex(index);
        return outcome == nullptr || outcome->publication == FileOperations::PublicationState::Published ||
               outcome->publication == FileOperations::PublicationState::Unknown;
    };

    const auto notifyDestinationParentsChanged = [&](PaneState& dst) noexcept
    {
        if (! dst.fileSystem)
        {
            return;
        }

        for (size_t index = 0; index < task->_sourcePaths.size(); ++index)
        {
            if (! destinationNeedsRefresh(index))
            {
                continue;
            }
            const auto destination = resolvedDestinationForIndex(index);
            if (! destination.has_value())
            {
                continue;
            }

            const auto parent = resolvedProviderParentForIndex(index, destination.value(), true);
            if (parent.has_value() && ! parent.value().empty())
            {
                cache.NotifyFolderContentsChanged(dst.fileSystem.get(), parent.value());
                forceRefreshVisibleFolder(parent.value());
            }
        }
    };

    const auto forceRefreshPane = [&](PaneState& paneState)
    {
        const auto folder = paneState.folderView.GetFolderPath();
        if (! paneState.fileSystem || ! folder.has_value() || ! cache.IsFolderWatched(paneState.fileSystem.get(), folder.value()))
        {
            paneState.folderView.ForceRefresh();
        }
    };

    {
        PaneState* dst         = destinationPane.has_value() ? &(destinationPane.value() == Pane::Left ? _leftPane : _rightPane) : nullptr;
        const bool sameContext = dst != nullptr && CompareStringOrdinal(task->_sourcePluginId.c_str(), -1, dst->pluginId.c_str(), -1, TRUE) == CSTR_EQUAL &&
                                 NavigationLocation::EqualsNoCase(src.instanceContext, dst->instanceContext);

        switch (task->GetOperation())
        {
            case FILESYSTEM_COPY:
                if (dst != nullptr && dst->fileSystem)
                {
                    cache.NotifyFolderContentsChanged(dst->fileSystem.get(), task->GetDestinationFolder());
                    notifyDestinationParentsChanged(*dst);
                }
                forceRefreshVisibleFolder(task->GetDestinationFolder());
                break;
            case FILESYSTEM_MOVE:
                if (sameContext && src.fileSystem)
                {
                    for (size_t index = 0; index < task->_sourcePaths.size(); ++index)
                    {
                        if (isResolvedDirectoryShell(index) || ! sourceWasRemoved(index) || ! destinationNeedsRefresh(index))
                        {
                            continue;
                        }
                        if (const auto destination = resolvedDestinationForIndex(index); destination.has_value())
                        {
                            cache.NotifyPathMoved(src.fileSystem.get(), task->_sourcePaths[index], destination.value());
                        }
                    }
                }
                else
                {
                    if (task->_fileSystem)
                    {
                        for (size_t index = 0; index < task->_sourcePaths.size(); ++index)
                        {
                            if (! isResolvedDirectoryShell(index) && sourceWasRemoved(index))
                            {
                                cache.NotifyPathDeleted(task->_fileSystem.get(), task->_sourcePaths[index]);
                            }
                        }
                    }
                    if (dst != nullptr && dst->fileSystem)
                    {
                        cache.NotifyFolderContentsChanged(dst->fileSystem.get(), task->GetDestinationFolder());
                        notifyDestinationParentsChanged(*dst);
                    }
                }
                forceRefreshVisibleSourceParents();
                forceRefreshVisibleFolder(task->GetDestinationFolder());
                break;
            case FILESYSTEM_DELETE:
                if (src.fileSystem)
                {
                    for (size_t index = 0; index < task->_sourcePaths.size(); ++index)
                    {
                        if (sourceWasRemoved(index))
                        {
                            cache.NotifyPathDeleted(src.fileSystem.get(), task->_sourcePaths[index]);
                        }
                    }
                }
                forceRefreshVisibleSourceParents();
                break;
            case FILESYSTEM_RENAME:
                if (src.fileSystem)
                {
                    for (size_t index = 0; index < task->_sourcePaths.size(); ++index)
                    {
                        if (sourceWasRemoved(index) && destinationNeedsRefresh(index))
                        {
                            if (const auto destination = resolvedDestinationForIndex(index); destination.has_value())
                            {
                                cache.NotifyPathMoved(src.fileSystem.get(), task->_sourcePaths[index], destination.value());
                                if (const auto parent = resolvedRenameParentForIndex(index); parent.has_value())
                                {
                                    forceRefreshVisibleFolder(parent.value());
                                }
                            }
                        }
                    }
                }
                forceRefreshPane(src);
                if (dst != nullptr)
                {
                    forceRefreshPane(*dst);
                }
                break;
            case FILESYSTEM_CREATE_DIRECTORY:
            default:
                forceRefreshPane(src);
                if (dst != nullptr)
                {
                    forceRefreshPane(*dst);
                }
                break;
        }
    }
    const bool autoDismissSuccess = _fileOperations->GetAutoDismissSuccess();
    _fileOperations->RemoveTask(taskId);
    if (autoDismissSuccess && IsAutoDismissableFileOperationCompletion(hr, warningCount, errorCount))
    {
        _fileOperations->DismissCompletedTask(taskId);
    }
}
