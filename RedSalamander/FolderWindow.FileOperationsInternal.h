#pragma once

// Internal implementation header for FolderWindow file operations.
// Keep this header private to the FolderWindow file-operation translation units.

#include "BatchRenameExecutionEngine.h"
#include "FileOperationArtifactRegistry.h"
#include "FileOperationMoveBreadcrumb.h"
#include "FolderWindowInternal.h"

#include <chrono>
#include <array>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <memory>
#include <optional>
#include <string_view>
#include <variant>
#include <vector>

namespace FileOperations
{
inline constexpr ULONGLONG kTaskCardRevealDelayMs         = 500u;
inline constexpr ULONGLONG kInlineRenameCardRevealDelayMs = kTaskCardRevealDelayMs;

enum class TransferIntent : uint8_t
{
    Copy,
    Move,
};

enum class LinkPolicy : uint8_t
{
    Preserve,
    Skip,
};

enum class ExecutionMode : uint8_t
{
    Queue,
    Parallel,
};

enum class OperationStrategy : uint8_t
{
    Copy,
    Native,
    Managed,
    CopyOnly,
};

struct MoveStrategyQualificationFacts
{
    bool nativeMoveQualified           = false;
    bool copyPairQualified             = false;
    bool movePairQualified             = false;
    bool sourceBoundDelete             = false;
    bool sourceConditionalDelete       = false;
    bool destinationExclusiveStage     = false;
    bool destinationConditionalPublish = false;
    bool sourceBindingAvailable        = false;
    bool destinationBindingAvailable   = false;
    bool destinationWriterDigestProof  = false; // R3-2: the destination writer proves the published content
};

enum class ManagedCleanupAttemptDisposition : uint8_t
{
    Removed,
    Retained,
    Indeterminate,
    ProviderContractViolation,
};

struct ManagedCleanupMutationFacts
{
    HRESULT status            = E_FAIL;
    bool outcomeKnown         = false;
    bool mutationCommitted    = false;
    bool originalStillPresent = true;
};

enum class DeleteMode : uint8_t
{
    Recycle,
    Permanent,
};

enum class DeleteOrigin : uint8_t
{
    PaneCommand,
    FindResults,
    CompareResults,
    PackCleanup,
    UnpackCleanup,
    RecycleEscalation,
};

enum class RenameOrigin : uint8_t
{
    Unspecified,
    InlineRename,
    BatchRename,
    ChangeCase,
};

enum class ConsentKind : uint8_t
{
    PermanentDelete,
    ArchiveDeleteAfter,
    RecycleEscalation,
};

enum class CancellationRouteClass : uint8_t
{
    Uncontained,
    Bounded,
    ProviderWatchdog,
};

struct QualifiedEndpoint
{
    std::wstring pluginId;
    std::wstring instanceId;
    std::wstring profileId;
    std::wstring rootId;
    std::optional<FileSystemPathIdentity> pathIdentity;
    bool verificationHostReadback            = false;
    bool verificationProviderBlake3Proof     = false;
    bool verificationWriterDigestProof       = false; // R3-2: the destination writer proves the published content
    bool verificationCapabilityCheckDeferred = false;
    bool cancellationAbort                   = false;
    bool cancellationDeadline                = false;
    CancellationRouteClass cancellationRouteClass = CancellationRouteClass::Uncontained;
    uint32_t providerWatchdogTimeoutMs            = 0u;
};

struct ProviderIdentitySnapshot
{
    std::vector<std::byte> objectId;
    std::vector<std::byte> revisionId;
    std::wstring pathProfileId;
};

enum class ObjectBindingState : uint8_t
{
    Bound,
    Unsupported,
    Missing,
    Indeterminate,
    ProviderContractViolation,
};

struct BoundObjectAuthority
{
    wil::com_ptr<IFileSystemBoundObject> boundObject;
    ProviderIdentitySnapshot identity;
    FileSystemBoundObjectKind kind = FILESYSTEM_BOUND_OTHER;
    uint64_t committedSizeBytes    = std::numeric_limits<uint64_t>::max();
};

struct ObjectBindingResult
{
    ObjectBindingState state = ObjectBindingState::Indeterminate;
    HRESULT status           = E_UNEXPECTED;
    BoundObjectAuthority authority;
};

enum class ObjectRevalidationState : uint8_t
{
    Same,
    Changed,
    Missing,
    Unsupported,
    Indeterminate,
    ProviderContractViolation,
};

struct ObjectRevalidationResult
{
    ObjectRevalidationState state = ObjectRevalidationState::Indeterminate;
    HRESULT status                = E_UNEXPECTED;
    BoundObjectAuthority current;
};

enum class TransferSafetyState : uint8_t
{
    Ready,
    SameObject,
    SameFolderMove,
    DestinationInsideSource,
    AncestryLink,
    SourceMissing,
    DestinationChanged,
    SourceChanged,
    Unsupported,
    Indeterminate,
    ProviderContractViolation,
};

struct RetainedPathAuthority
{
    std::wstring providerPath;
    BoundObjectAuthority authority;
};

struct TransferMutationGuard
{
    TransferSafetyState state = TransferSafetyState::Indeterminate;
    HRESULT status            = E_UNEXPECTED;
    BoundObjectAuthority source;
    std::optional<RetainedPathAuthority> sourceParent;
    std::optional<RetainedPathAuthority> destination;
    std::vector<RetainedPathAuthority> destinationAncestors;
    bool destinationWasMissing = false;
    bool samePathText          = false;
};

// A task holds every exact binding while it owns the interlock. Nodes form an interned parent chain,
// so a large sibling selection retains each common ancestor once instead of one handle chain per item.
// The scheduler walks from `root` toward the endpoint root to recognize alias-equivalent
// ancestor/descendant envelopes without serializing disjoint siblings that merely share a volume root.
struct MutationInterlockAuthorityNode
{
    QualifiedEndpoint endpoint;
    RetainedPathAuthority retained;
    std::shared_ptr<const MutationInterlockAuthorityNode> parent;
};

struct MutationInterlockAnchor
{
    std::shared_ptr<const MutationInterlockAuthorityNode> authority;
    // Provider-profile relative path from the exact authority to the requested mutation target.
    // Empty means the authority is the target itself.
    std::wstring relativePath;
    // Canonical provider-profile key computed before queue admission. A missing key keeps the
    // exact pairwise fallback; it never weakens identity or pathname comparison.
    std::optional<std::wstring> relativePathKey;
};

enum class MutationInterlockAccess : uint8_t
{
    ReadSource,
    WriteSource,
    PublishDestination,
};

[[nodiscard]] constexpr bool MutationInterlockAccessesConflict(MutationInterlockAccess left, MutationInterlockAccess right) noexcept
{
    return left != MutationInterlockAccess::ReadSource || right != MutationInterlockAccess::ReadSource;
}

struct MutationInterlockScope
{
    QualifiedEndpoint endpoint;
    std::wstring providerPath;
    std::shared_ptr<const MutationInterlockAuthorityNode> root;
    std::vector<MutationInterlockAnchor> anchors;
    MutationInterlockAccess access = MutationInterlockAccess::WriteSource;
    bool exactDeleteAuthority      = false;
    // Exact identity is explicitly unsupported. Such a scope conflicts conservatively with every
    // mutation scope in the same provider identity domain; it is never treated as path-only proof.
    bool conservativeIdentityDomain = false;
    // C10: the identity a provider without object binding resolved for this root while the task
    // prepared; the card's re-check and the delete compare the live identity with it.
    std::optional<FileSystemDeleteIdentity> pinnedDeleteIdentity;
};

enum class TaskLifecyclePhase : uint8_t
{
    Preparing,
    AwaitingAcceptance,
    Ready,
    Waiting,
    Running,
    Stopping,
    Terminal,
};

struct PreparationEndpointFact final
{
    std::wstring pluginId;
    std::wstring instanceId;
    std::wstring profileId;
    std::wstring rootId;
};

struct PreparationScopeFact final
{
    uint32_t endpointIndex = 0u;
    std::wstring providerPath;
    MutationInterlockAccess access = MutationInterlockAccess::WriteSource;
    bool exactDeleteAuthority      = false;
    bool conservativeIdentityDomain = false;
};

struct PreparationStrategyFact final
{
    OperationStrategy strategy = OperationStrategy::Copy;
    uint64_t selectedRootCount  = 0u;
};

// Immutable, authority-free observer facts published after worker-owned preparation. Bound
// objects, provider interfaces, descendant discoveries, and mutable execution state never enter
// this snapshot.
struct PreparationSnapshot final
{
    uint64_t taskId                    = 0u;
    FileSystemOperation operation      = FILESYSTEM_COPY;
    uint64_t selectedRootCount         = 0u;
    uint64_t copyOnlyCount             = 0u;
    uint64_t buildSnapshotUs           = 0u;
    uint64_t selectedRootReadinessUs   = 0u;
    uint64_t retainedBytes             = 0u;
    HRESULT status                     = E_PENDING;
    std::vector<PreparationEndpointFact> endpoints;
    std::vector<PreparationScopeFact> scopes;
    std::vector<PreparationStrategyFact> strategies;
};

#ifdef ENABLE_TESTS
[[nodiscard]] bool DebugMutationScopesOverlapForSelfTest(const MutationInterlockScope& left, const MutationInterlockScope& right) noexcept;
[[nodiscard]] bool DebugMutationScopeSetsOverlapForSelfTest(const std::vector<MutationInterlockScope>& left,
                                                            const std::vector<MutationInterlockScope>& right,
                                                            uint64_t& comparisons) noexcept;
#endif

struct QualifiedSourceItem
{
    std::wstring providerPath;
    // Comparison-only ingress evidence. Execution must bind and revalidate; this value is never mutation authority.
    std::optional<ProviderIdentitySnapshot> ingressSnapshot;
};

struct QualifiedDestination
{
    std::wstring providerFolderPath;
};

struct CreateDirectoryAdmission
{
    QualifiedEndpoint endpoint;
    std::filesystem::path candidateProviderPath;
    std::wstring providerCollisionKey;
    bool allowLocalNativeFallback = false;
};

struct TransferDestinationMapping
{
    size_t sourceIndex = 0;
    std::wstring destinationProviderPath;
};

struct ClipboardSequence
{
    uint32_t windowsSequenceNumber = 0;
};

struct RenameStep
{
    QualifiedSourceItem source;
    std::wstring finalLeafName;
    std::wstring providerJoinedPath;
    std::wstring providerParentKey;
    std::wstring providerSourceCollisionKey;
    std::wstring providerCollisionKey;
};

struct DestructiveConsentReceipt
{
    ConsentKind kind   = ConsentKind::PermanentDelete;
    uint64_t taskNonce = 0;
    std::optional<size_t> itemIndex;
};

struct OperationOptions
{
    LinkPolicy linkPolicy       = LinkPolicy::Preserve;
    bool verifyAfterCopy        = false;
    ExecutionMode executionMode = ExecutionMode::Queue;
    std::optional<uint64_t> bandwidthLimitBytesPerSecond;
};

struct TransferPlan
{
    TransferIntent intent      = TransferIntent::Copy;
    OperationStrategy strategy = OperationStrategy::Copy;
    QualifiedEndpoint sourceEndpoint;
    QualifiedEndpoint destinationEndpoint;
    std::vector<QualifiedSourceItem> selectedItems;
    QualifiedDestination destination;
    std::vector<TransferDestinationMapping> explicitMappings;
    OperationOptions options;
    std::optional<ClipboardSequence> moveClipboardSequence;
};

struct RenamePlan
{
    RenameOrigin origin = RenameOrigin::Unspecified;
    QualifiedEndpoint endpoint;
    std::vector<RenameStep> finalMappings;
    BatchRenameExecutionSchedule schedule;
    OperationOptions options;
};

struct DeletePlan
{
    QualifiedEndpoint endpoint;
    std::vector<QualifiedSourceItem> selectedItems;
    DeleteMode mode     = DeleteMode::Recycle;
    DeleteOrigin origin = DeleteOrigin::PaneCommand;
    bool recursive      = false;
    // The provider exposes no bound objects (FTP/SFTP/SCP, S3, MTP): a Permanent Delete runs the
    // provider's contract-tested native mutation by selected path and relies on the receipt that
    // mutation returns. Items carry no ingress snapshot and no exact-delete interlock authority.
    bool nativeAuthority = false;
    OperationOptions options;
    std::optional<DestructiveConsentReceipt> initialConsent;
};

using FileOperationPlan = std::variant<TransferPlan, RenamePlan, DeletePlan>;
// One user admission may contain several executable child plans when a selection spans provider
// roots. Every child remains qualified by exactly one source endpoint; the admission owns the
// single confirmation and clipboard-consumption barrier for the whole group.
using FileOperationPlanGroup = std::vector<FileOperationPlan>;

struct FileOperationItemResult
{
    size_t sourceIndex                          = 0;
    OperationStrategy strategy                  = OperationStrategy::Copy;
    PublicationState publication                = PublicationState::NotAttempted;
    VerificationState verification              = VerificationState::NotRequested;
    SourceDisposition sourceDisposition         = SourceDisposition::Retained;
    ItemCompletion completion                   = ItemCompletion::Failed;
    OwnedStageDisposition ownedStageDisposition = OwnedStageDisposition::NotApplicable;
    HRESULT status                              = E_PENDING;
    std::wstring finalSourcePath;
    std::wstring finalDestinationPath;
    // Present only when the executor retained exact no-follow source identity for an object that
    // remained. Completion actions must rebind and compare this snapshot before republishing a
    // new cut list; pathname existence alone is insufficient.
    std::optional<ProviderIdentitySnapshot> retainedSourceIdentity;
};

enum class ConflictClass : uint8_t
{
    RegularFileExists,
    ReadOnlyRegularFile,
    ReadOnlyRegularFileExists,
    TypeMismatch,
    DestinationLink,
    NameNotRepresentable,
    TargetConflict,
    SharingViolation,
    AccessDenied,
    DiskFull,
    PathTooLong,
    RecycleFailed,
    InsufficientSpace,
    SpaceUnknown,
    EfsPlaintext,
    SparseInflation,
    PlaceholderHydration,
    MetadataLoss,
    NetworkOffline,
    UnsupportedReparse,
    SameHostOverlap,
    SameHostLiveOutput,
    PermanentDeleteConfirmation, // C1: the initial permanent-delete confirmation, on the card after Preparing pinned the roots
    Unknown,
    Count,
};

enum class ConflictAction : uint8_t
{
    None,
    Overwrite,
    ReplaceReadOnly,
    ReplaceLink,
    PermanentDelete,
    Proceed,
    RetainSource,
    Retry,
    KeepBoth,
    Skip,
    SkipAll,
    Cancel,
    RunConcurrently,
    QueueUntilOtherTaskFinishes,
    InvalidateLiveOutput,
};

enum class DeferredConsentRisk : uint8_t
{
    InsufficientSpace,
    SpaceUnknown,
    EfsPlaintext,
    SparseInflation,
    PlaceholderHydration,
    MetadataLoss,
    RecycleEscalation,
    SameHostOverlap, // R4-A02-1: another live or queued task in this host overlaps this task's scopes
    SameHostLiveOutput, // R4-A02-2: this mutation would invalidate another concurrently-live task's output
    PermanentDelete, // C1: the initial permanent-delete confirmation, asked on the card after Preparing pinned the roots
};

// R4-A02-1: the concrete problem an overlap advisory names (roles of the first conflicting pair).
enum class SameHostOverlapProblem : uint8_t
{
    None,
    RemovesRead,       // this task writes/removes what the other task reads
    SameNames,         // both publish into the same destination names
    SameMembers,       // both remove or rename the same members
    RemovesPublished,  // this task removes/renames what the other task publishes
    ReplacesRead,      // this task publishes over what the other task reads
};

inline constexpr size_t kMaxSameHostOverlapRelations = 64u;

struct SameHostOverlapRelationSet final
{
    std::array<uint64_t, kMaxSameHostOverlapRelations> taskIds{};
    size_t taskIdCount = 0u;
};

struct DeferredConsentReceipt
{
    DeferredConsentRisk risk = DeferredConsentRisk::InsufficientSpace;
    uint64_t taskId          = 0;
    std::optional<size_t> itemIndex;
    std::wstring destinationRootId;
    std::optional<ProviderIdentitySnapshot> sourceIdentity;
    std::optional<ProviderIdentitySnapshot> destinationIdentity;
    // Allocated only for an overlap decision. Keeping the fixed relation set behind one shared
    // owner avoids multiplying it by every slot in Task's bounded receipt table.
    std::shared_ptr<const SameHostOverlapRelationSet> relatedTasks;
    ConflictAction decision = ConflictAction::Cancel;
};

[[nodiscard]] bool IsProviderKeepBothDestinationEligible(bool providerHandlesNestedKeepBoth,
                                                         std::wstring_view conflictDestinationPath,
                                                         std::wstring_view selectedDestinationPath) noexcept;

struct TypedConflictRecord
{
    ConflictClass conflictClass = ConflictClass::Unknown;
    std::vector<ConflictAction> allowedActions;
    ConflictAction safeDefault = ConflictAction::Cancel;
    bool applyToAllEligible    = false;
    std::wstring endpointProfileId;
    std::wstring sourcePath;
    std::wstring destinationPath;
    std::optional<ProviderIdentitySnapshot> sourceIdentity;
    std::optional<ProviderIdentitySnapshot> destinationIdentity;
};

enum class PlanRejectionBucket : uint8_t
{
    None,
    UnsupportedOperation,
    MissingFileSystem,
    EmptySelection,
    MalformedEndpoint,
    MalformedSource,
    MalformedDestination,
    UnsupportedDeviceNamespace,
    EscapingDestinationMapping,
    DestinationInsideSource,
    SameFolderMove,
    InvalidStrategy,
    MixedSourceEndpoint,
    InvalidExplicitMappings,
    InvalidClipboardSequence,
    MissingDestructiveConsent,
    InvalidRename,
    InvalidDestinationName, // R0-RC3 (9): a destination leaf the destination provider's child-name contract rejects
};

[[nodiscard]] HRESULT ValidatePlan(const FileOperationPlan& plan, PlanRejectionBucket* rejectionBucket = nullptr) noexcept;
// C1: a Permanent Delete plan may publish before its card confirmation records the consent receipt.
[[nodiscard]] HRESULT ValidatePlan(const FileOperationPlan& plan, PlanRejectionBucket* rejectionBucket, bool allowPendingPermanentConsent) noexcept;
[[nodiscard]] std::optional<OperationStrategy> SelectMoveStrategy(const MoveStrategyQualificationFacts& facts) noexcept;
[[nodiscard]] ManagedCleanupAttemptDisposition ClassifyManagedCleanupMutation(const ManagedCleanupMutationFacts& facts) noexcept;
[[nodiscard]] bool TryResolveTransferDestinationProviderPath(const TransferPlan& plan, size_t sourceIndex, std::wstring& destinationOut) noexcept;
[[nodiscard]] ObjectBindingResult BindObjectAuthority(IFileSystem* fileSystem,
                                                      std::wstring_view providerPath,
                                                      std::wstring_view pathProfileId,
                                                      FileSystemBindFlags flags) noexcept;
[[nodiscard]] ObjectBindingResult CaptureReturnedObjectAuthority(wil::com_ptr<IFileSystemBoundObject> bound, std::wstring_view pathProfileId) noexcept;
[[nodiscard]] ObjectRevalidationResult RevalidateObjectAuthority(IFileSystem* fileSystem,
                                                                 std::wstring_view providerPath,
                                                                 std::wstring_view pathProfileId,
                                                                 FileSystemBindFlags flags,
                                                                 const BoundObjectAuthority& expected) noexcept;
[[nodiscard]] HRESULT CrossCheckBoundObjectIdentity(const BoundObjectAuthority& expected,
                                                    const BoundObjectAuthority& current,
                                                    bool& sameObject,
                                                    bool& sameRevision) noexcept;
[[nodiscard]] bool IsLocalFileSystemEndpoint(const QualifiedEndpoint& endpoint) noexcept;
[[nodiscard]] bool QualifiedEndpointsReferToSameRoot(const QualifiedEndpoint& left, const QualifiedEndpoint& right) noexcept;
[[nodiscard]] bool QualifiedEndpointsShareObjectIdentityDomain(const QualifiedEndpoint& left, const QualifiedEndpoint& right) noexcept;
[[nodiscard]] TransferMutationGuard PrepareTransferMutationGuard(IFileSystem* sourceFileSystem,
                                                                 IFileSystem* destinationFileSystem,
                                                                 const QualifiedEndpoint& sourceEndpoint,
                                                                 const QualifiedEndpoint& destinationEndpoint,
                                                                 TransferIntent intent,
                                                                 std::wstring_view sourcePath,
                                                                 std::wstring_view destinationPath) noexcept;
[[nodiscard]] TransferSafetyState RevalidateTransferMutationGuard(IFileSystem* sourceFileSystem,
                                                                  IFileSystem* destinationFileSystem,
                                                                  const QualifiedEndpoint& sourceEndpoint,
                                                                  const QualifiedEndpoint& destinationEndpoint,
                                                                  TransferIntent intent,
                                                                  std::wstring_view sourcePath,
                                                                  std::wstring_view destinationPath,
                                                                  const TransferMutationGuard& guard,
                                                                  HRESULT& status,
                                                                  bool allowMissingDestinationDirectoryMerge = false) noexcept;
#if defined(ENABLE_TESTS)
[[nodiscard]] ObjectBindingResult ValidateSuccessfulBoundObjectForSelfTest(IFileSystemBoundObject* bound, std::wstring_view pathProfileId) noexcept;
[[nodiscard]] bool TryQualifyEndpointForSelfTest(const wil::com_ptr<IFileSystem>& fileSystem,
                                                 std::wstring_view providerPath,
                                                 FileSystemOperation operation,
                                                 std::wstring_view pluginId,
                                                 std::wstring_view instanceId,
                                                 QualifiedEndpoint& endpoint) noexcept;
#endif
} // namespace FileOperations

inline void RestoreActivePaneFolderViewFocusIfWindowHadFocusBeforeHide(FolderWindow& folderWindow, HWND containerWindow, HWND focusedBeforeHide) noexcept
{
    if (! containerWindow || ! focusedBeforeHide || (focusedBeforeHide != containerWindow && IsChild(containerWindow, focusedBeforeHide) == FALSE))
    {
        return;
    }

    const HWND folderView = folderWindow.GetFolderViewHwnd(folderWindow.GetActivePane());
    if (folderView && IsWindow(folderView) != FALSE)
    {
        folderWindow.RequestRestoreFolderViewFocus(folderView);
        return;
    }

    static_cast<void>(folderWindow.TryRestoreActivePaneFolderViewFocus());
}

struct FolderWindow::FileOperationState
{
    enum class DiagnosticSeverity : unsigned char
    {
        Debug,
        Info,
        Warning,
        Error,
    };

    enum class ExecutionMode : unsigned char
    {
        BulkItems,
        PerItem,
    };

    struct BatchRenameAdmissionInput final
    {
        FileOperations::RenameOrigin origin = FileOperations::RenameOrigin::BatchRename;
        std::vector<BatchRenameExecutionOp> operations;
        FileOperations::ExecutionMode executionMode = FileOperations::ExecutionMode::Queue;
        std::chrono::steady_clock::time_point startedAt{};
        uint64_t uiCaptureUs = 0u;
    };

    struct BatchRenameArtifactPromptPayload final
    {
        uint64_t taskId = 0u;
        FileOperationArtifacts::TouchGuardRequest request;
    };

    struct ClipboardMoveReadyPayload final
    {
        uint64_t taskId = 0u;
    };

    struct OperationAdmission
    {
        FileOperations::FileOperationPlanGroup plans;
        FolderWindow::Pane sourcePane = FolderWindow::Pane::Left;
        std::optional<FolderWindow::Pane> destinationPane;
        wil::com_ptr<IFileSystem> fileSystem;
        wil::com_ptr<IFileSystem> destinationFileSystem;
        FileSystemFlags flags = FILESYSTEM_FLAG_NONE;
        // Copy and Move are identity-guarded only by the per-item executor. Delete callers that
        // intentionally use provider batching must opt in to BulkItems explicitly.
        ExecutionMode executionMode = ExecutionMode::PerItem;
        bool requireConfirmation    = false;
        std::vector<FolderWindow::ResolvedFileOperationItem> resolvedItems;
        std::wstring confirmationMessage;
        // C1: a Permanent Delete confirms on its card after Preparing has pinned the roots. Admission
        // gathers the words here from the pane listing (no provider call); the test override that
        // used to answer the modal is captured with them (0 none, 1 confirm, 2 cancel).
        bool permanentDeleteConfirmationPending = false;
        std::wstring permanentDeleteConsentDetail;
        std::wstring permanentDeleteConsentFrom;
        int permanentDeleteTestOverride = 0;
        std::wstring sourcePluginShortId;
        std::wstring sourceInstanceContext;
        std::wstring destinationPluginShortId;
        std::wstring destinationInstanceContext;
        // The archive prompt owns this one destructive choice. StartOperation replaces the
        // captured kind with the admitted task nonce before final plan validation.
        std::optional<FileOperations::ConsentKind> capturedConsentKind;
        // Runs on the UI thread after worker-owned selected-root readiness and before mutation
        // release. Clipboard Move reaches this barrier through a tokenized asynchronous message;
        // a failure keeps the accepted task but permits only terminal no-mutation reduction.
        std::function<HRESULT()> preWorkerReleaseBarrier;
        // Worker-owned decision gate after immutable preparation facts are published and before
        // clipboard consumption or mutation release. It is legal for every operation family.
        std::function<HRESULT()> preConsumptionDecisionGate;
        // Optional non-owning observer. The callback must be noexcept in behavior and return
        // immediately; it receives authority-free immutable facts on the worker thread.
        std::function<void(std::shared_ptr<const FileOperations::PreparationSnapshot>)> preparationObserver;
        // Optional non-owning UI observer. The callback must be noexcept in behavior and return
        // immediately; Batch Rename uses it only to post a tokenized progress payload.
        std::function<void(uint64_t completedItems, uint64_t totalItems)> progressCallback;
    };

    struct TaskCompletedPayload
    {
        uint64_t taskId            = 0;
        HRESULT hr                 = S_OK;
        unsigned long warningCount = 0;
        unsigned long errorCount   = 0;
    };

    struct TaskDiagnosticEntry
    {
        SYSTEMTIME localTime{};
        uint64_t taskId                 = 0;
        FileSystemOperation operation   = FILESYSTEM_COPY;
        DiagnosticSeverity severity     = DiagnosticSeverity::Info;
        HRESULT status                  = S_OK;
        uint64_t processWorkingSetBytes = 0;
        uint64_t processPrivateBytes    = 0;
        std::wstring category;
        std::wstring message;
        std::wstring sourcePath;
        std::wstring destinationPath;
        std::wstring concurrencyMode;
        std::wstring storageType;
        std::wstring destinationStorageType;
        unsigned long autoTunedConcurrency       = 0;
        unsigned long effectiveConcurrencyBudget = 0;
    };

    struct CompletedTaskSummary
    {
        struct RetainedSourceActionItem final
        {
            std::filesystem::path providerPath;
            FileOperations::ProviderIdentitySnapshot identity;
        };

        uint64_t taskId               = 0;
        FileSystemOperation operation = FILESYSTEM_COPY;
        FolderWindow::Pane sourcePane = FolderWindow::Pane::Left;
        std::wstring sourcePluginId;
        std::wstring sourcePluginShortId;
        std::wstring sourceInstanceContext;
        std::optional<FolderWindow::Pane> destinationPane;
        std::wstring destinationPluginId;
        std::wstring destinationPluginShortId;
        std::wstring destinationInstanceContext;
        std::filesystem::path destinationFolder;
        FileSystemFlags flags = static_cast<FileSystemFlags>(0);
        std::filesystem::path diagnosticsLogPath;

        HRESULT resultHr             = S_OK;
        unsigned long totalItems     = 0;
        unsigned long completedItems = 0;
        uint64_t totalBytes          = 0;
        uint64_t completedBytes      = 0;

        // While the one traversal remains open, totals may be unknown; keep a best-effort
        // top-level type breakdown for UI even after discovery-ahead was skipped.
        bool discoverySkipped          = false;
        bool discoveryClosed           = false;
        unsigned long completedFiles   = 0;
        unsigned long completedFolders = 0;
        std::wstring sourcePath;
        std::wstring destinationPath;

        bool autoConcurrencyUsed                       = false;
        uint32_t autoConcurrencyStorageKind            = FILESYSTEM_STORAGE_UNKNOWN;
        uint32_t autoConcurrencyDestinationStorageKind = FILESYSTEM_STORAGE_UNKNOWN;
        unsigned int autoTunedConcurrency              = 0;
        unsigned int effectiveConcurrencyBudget        = 0;

        unsigned long warningCount = 0;
        unsigned long errorCount   = 0;
        std::wstring lastDiagnosticMessage;
        std::wstring resultSummary;
        unsigned long publishedItemCount               = 0;
        unsigned long unknownPublicationItemCount      = 0;
        unsigned long removedSourceCount               = 0;
        unsigned long retainedSourceCount              = 0;
        unsigned long retainedSourceNativeCount        = 0; // rename merges whose emptied source folder stayed
        unsigned long unknownSourceCount               = 0;
        unsigned long verifiedItemCount                = 0;
        unsigned long verificationFailedItemCount      = 0;
        unsigned long verificationUnavailableItemCount = 0;
        unsigned long verificationCanceledItemCount    = 0;
        unsigned long verificationProblemItemCount     = 0;
        unsigned long indeterminateItemCount           = 0;
        unsigned long retainedOwnedStageCount          = 0;
        unsigned long unknownOwnedStageCount           = 0;
        bool clipboardMoveAdmission                    = false;
        bool clipboardMoveConsumed                     = false;
        HRESULT clipboardMoveConsumptionStatus         = S_OK;
        std::vector<std::filesystem::path> retainedSourcePaths;
        std::vector<std::filesystem::path> unknownSourcePaths;
        std::vector<RetainedSourceActionItem> exactRetainedSourceItems;
        std::vector<TaskDiagnosticEntry> issueDiagnostics;

        // Restart-projected notices keep two explicitly different identities: taskId remains the
        // in-session card/action key, while interruptedOperationId is the durable breadcrumb's
        // historical operation ID. Dismissal acknowledges the exact breadcrumb path; neither ID
        // authorizes source or destination mutation.
        uint64_t interruptedOperationId = 0u;
        std::filesystem::path interruptedMoveBreadcrumbPath;
        bool interruptedMoveNotice = false;

        ULONGLONG lastProgressCallbackTick = 0;
        ULONGLONG completedTick            = 0;
    };

    struct Task final : public IFileSystemCallback, public IFileSystemOperationControl
    {
        // Maximum number of in-flight file lines the popup can display for a single task.
        // This should be >= the Copy/Move worker concurrency cap so parallel file copies can be represented.
        static constexpr size_t kMaxInFlightFiles = 16u;

        using ConflictBucket = FileOperations::ConflictClass;
        using ConflictAction = FileOperations::ConflictAction;
        using TaskLifecyclePhase = FileOperations::TaskLifecyclePhase;
        using PreparationSnapshot = FileOperations::PreparationSnapshot;

        enum class TaskPresentationState : uint8_t
        {
            NotApplicable,
            Hidden,
            RevealRequested,
            Presented,
            SuppressedCleanSuccess,
        };

        enum class ConflictItemKind : uint8_t
        {
            Unknown,
            RegularFile,
            Directory,
            Link,
        };

        enum class ConflictLinkKind : uint8_t
        {
            Unknown,
            SymbolicFile,
            SymbolicDirectory,
            Junction,
        };

        enum class SourceItemExecutionPhase : uint8_t
        {
            Preparing,
            MutationPossible,
        };

        enum class LiveOutputGuardDisposition : uint8_t
        {
            Proceed,
            Skip,
            RetryCurrentMutation,
            Cancel,
        };

        struct SourceItemResultBuilder
        {
            SourceItemExecutionPhase phase = SourceItemExecutionPhase::Preparing;
            std::optional<HRESULT> status;
            std::optional<FileSystemItemMutationResult> mutation;
            std::optional<FileOperations::FileOperationItemResult> terminal;
        };

        struct ConflictDecisionScope
        {
            ConflictBucket bucket                = ConflictBucket::Unknown;
            ConflictItemKind sourceKind          = ConflictItemKind::Unknown;
            ConflictItemKind destinationKind     = ConflictItemKind::Unknown;
            ConflictLinkKind destinationLinkKind = ConflictLinkKind::Unknown;
            std::wstring sourceProfileId;
            std::wstring destinationProfileId;
            std::wstring destinationRootId;
            bool applyToAllEligible = false;

            bool operator==(const ConflictDecisionScope&) const noexcept = default;
        };

        struct CachedConflictDecision
        {
            ConflictDecisionScope scope{};
            ConflictAction action = ConflictAction::None;
        };

        struct PerItemCallbackCookie
        {
            explicit PerItemCallbackCookie(size_t sourceIndex = 0u) noexcept : itemIndex(sourceIndex)
            {
            }
            PerItemCallbackCookie(const PerItemCallbackCookie&)            = delete;
            PerItemCallbackCookie& operator=(const PerItemCallbackCookie&) = delete;
            PerItemCallbackCookie(PerItemCallbackCookie&&)                 = delete;
            PerItemCallbackCookie& operator=(PerItemCallbackCookie&&)      = delete;

            size_t itemIndex = 0;
            std::wstring lastProgressSourcePath;
            std::wstring lastProgressDestinationPath;
            std::wstring operationDestinationPath;
            bool keepBothRequested = false;
            std::atomic<bool> explicitSkipObserved{false};
            bool sourceIsDirectory             = false;
            uint64_t lastDiscoveredBytes       = 0;
            uint64_t lastDiscoveredFiles       = 0;
            uint64_t lastDiscoveredDirectories = 0;
            std::array<unsigned int, static_cast<size_t>(ConflictBucket::Count)> issueRetryCounts{};
        };

        struct ConflictPromptState
        {
            static constexpr size_t kMaxActions = 8u;

            struct ItemMetadata
            {
                bool available            = false;
                bool objectKindKnown      = false;
                bool isDirectory          = false;
                bool isLink               = false;
                ConflictLinkKind linkKind = ConflictLinkKind::Unknown;
                bool sizeKnown            = false;
                uint64_t sizeBytes        = 0;
                __int64 lastWriteTime     = 0;
                unsigned long attributes  = 0;
            };

            bool active           = false;
            bool metadataLoading  = false;
            bool deferredConsent  = false;
            ConflictBucket bucket = ConflictBucket::Unknown;
            HRESULT status        = S_OK;
            std::wstring sourcePath;
            std::wstring destinationPath;
            ItemMetadata sourceMetadata;
            ItemMetadata destinationMetadata;
            ConflictDecisionScope decisionScope{};
            std::array<ConflictAction, kMaxActions> actions{};
            size_t actionCount = 0;
            std::array<ConflictAction, kMaxActions> primaryActions{};
            size_t primaryActionCount = 0;
            std::array<ConflictAction, kMaxActions> overflowActions{};
            size_t overflowActionCount   = 0;
            ConflictAction defaultAction = ConflictAction::None;
            ConflictAction escapeAction  = ConflictAction::None;
            bool applyToAllEligible      = false;
            bool applyToAllChecked       = false;
            bool skipAllEligible         = false;
            bool buttonsPublishable      = false;
            bool retryFailed             = false;
            unsigned int attemptCount    = 0u; // C2: attempts already made for this item and bucket; each explicit Retry is one
            bool factItemCountKnown      = false;
            uint64_t factItemCount       = 0;
            bool factBytesKnown          = false;
            uint64_t factBytes           = 0;
            uint8_t overlapProblem       = 0u; // FileOperations::SameHostOverlapProblem (R4-A02-1)
            uint64_t overlapTaskId       = 0u; // the first overlapping task named by the advisory
            std::wstring consentDetail;        // C1: what a permanent-delete confirmation deletes (the counts text)
            std::optional<FileOperations::ProviderIdentitySnapshot> sourceIdentity;
            std::optional<FileOperations::ProviderIdentitySnapshot> destinationIdentity;
        };

        struct ConflictArbiter
        {
            ConflictArbiter()                                  = default;
            ConflictArbiter(const ConflictArbiter&)            = delete;
            ConflictArbiter(ConflictArbiter&&)                 = delete;
            ConflictArbiter& operator=(const ConflictArbiter&) = delete;
            ConflictArbiter& operator=(ConflictArbiter&&)      = delete;
            ~ConflictArbiter()                                 = default;

            mutable std::mutex mutex;
            std::condition_variable cv;
            static constexpr size_t kMaxCachedConflictDecisions = 64u;
            std::array<std::optional<CachedConflictDecision>, kMaxCachedConflictDecisions> decisionCache{};
            static constexpr size_t kMaxDeferredConsentReceipts = 64u;
            std::array<std::optional<FileOperations::DeferredConsentReceipt>, kMaxDeferredConsentReceipts> consentReceipts{};
            ConflictPromptState prompt{};
            DWORD ownerThreadId = 0;
            std::optional<ConflictAction> decisionAction;
            bool decisionApplyToAll = false;
            wil::unique_event_nothrow decisionEvent;
        };

        struct InFlightFileProgress
        {
            const void* cookieKey     = nullptr;
            uint64_t progressStreamId = 0;
            std::wstring sourcePath;
            uint64_t totalBytes      = 0;
            uint64_t completedBytes  = 0;
            ULONGLONG lastUpdateTick = 0;
            bool completionCountedWhileDiscoveryOpen = false;
        };

        struct ProgressStreamPerf
        {
            const void* cookieKey           = nullptr;
            uint64_t progressStreamId       = 0;
            uint64_t callbackCount          = 0;
            uint64_t callbackUs             = 0;
            uint64_t lockWaitUs             = 0;
            uint64_t callbackGapCount       = 0;
            uint64_t callbackGapMs          = 0;
            uint64_t callbackGapBytes       = 0;
            uint64_t maxCallbackGapMs       = 0;
            uint64_t maxCallbackGapBytes    = 0;
            uint64_t maxCallbackDeltaBytes  = 0;
            uint64_t lastItemCompletedBytes = 0;
            ULONGLONG firstUpdateTick       = 0;
            ULONGLONG lastUpdateTick        = 0;
        };

        struct ConflictWorkerPerf
        {
            const void* cookieKey    = nullptr;
            uint64_t promptCount     = 0;
            uint64_t waitUs          = 0;
            ULONGLONG lastUpdateTick = 0;
        };

        struct DiagnosticPathSnapshot
        {
            std::wstring progressSourcePath;
            std::wstring progressDestinationPath;
            std::wstring lastProgressCallbackSourcePath;
            std::wstring lastProgressCallbackDestinationPath;
        };

        struct PerfStats
        {
            PerfStats()                            = default;
            PerfStats(const PerfStats&)            = delete;
            PerfStats(PerfStats&&)                 = delete;
            PerfStats& operator=(const PerfStats&) = delete;
            PerfStats& operator=(PerfStats&&)      = delete;

            uint64_t queueWaitUs        = 0;
            uint64_t interlockWaitUs    = 0;
            uint64_t interlockWaitCount = 0;
            std::atomic<uint64_t> schedulerWaitUs{0};
            std::atomic<uint64_t> schedulerWaitForWorkUs{0};
            std::atomic<uint64_t> schedulerProcessIndexUs{0};
            std::atomic<uint64_t> schedulerDequeueAttempts{0};
            std::atomic<uint64_t> schedulerDequeueSuccess{0};
            std::atomic<uint64_t> bridgeCopyUs{0};
            std::atomic<uint64_t> bridgeReaderWaitUs{0};
            std::atomic<uint64_t> bridgeWriterWaitUs{0};
            std::atomic<uint64_t> bridgeReadUs{0};
            std::atomic<uint64_t> bridgeWriteUs{0};
            std::atomic<uint64_t> bridgeStageCreateUs{0};
            std::atomic<uint64_t> bridgeStageCreateCount{0};
            std::atomic<uint64_t> bridgePublicationUs{0};
            std::atomic<uint64_t> bridgePublicationCount{0};
            std::atomic<uint64_t> bridgeStageRetainedCount{0};
            std::atomic<uint64_t> bridgeImmediateDestinationSizeProbeUs{0};
            std::atomic<uint64_t> bridgeImmediateDestinationSizeProbeCount{0};
            std::atomic<uint64_t> bridgeCommitSizeProofCount{0};
            std::atomic<uint64_t> bridgeCommitSizeProofFallbackCount{0};
            std::atomic<uint64_t> verificationUs{0};
            std::atomic<uint64_t> verificationReadBytes{0};
            std::atomic<uint64_t> verificationReadCalls{0};
            std::atomic<uint64_t> verificationProviderProofCount{0};
            std::atomic<uint64_t> verificationHostReadbackCount{0};
            uint64_t bridgeDirectoryEnsureCount          = 0;
            uint64_t bridgeFileAdmissionCount            = 0;
            uint64_t bridgeFileStartedBeforeProducerDone = 0;
            uint64_t bridgeAdmissionMaxQueueDepth        = 0;
            std::atomic<uint64_t> discoveryOpenUs{0};
            std::atomic<uint64_t> discoveryCallbackCount{0};
            std::atomic<uint64_t> discoveryCallbackUs{0};
            std::atomic<uint64_t> discoveryLockWaitUs{0};
            std::atomic<uint64_t> progressCallbackUs{0};
            uint64_t progressFirstCallbackDelayMs     = 0;
            uint64_t progressLockWaitUs               = 0;
            uint64_t progressLockHoldUs               = 0;
            uint64_t progressLockContentionCount      = 0;
            uint64_t progressPathUpdateBytes          = 0;
            uint64_t progressPathUpdateAppliedCount   = 0;
            uint64_t progressPathUpdateSkippedCount   = 0;
            uint64_t progressPathUpdateThrottledCount = 0;
            uint64_t progressInFlightEvictions        = 0;
            uint64_t perItemInFlightEvictions         = 0;
            std::atomic<uint64_t> pauseWaitUs{0};
            uint64_t conflictWaitUs = 0;
            std::atomic<uint64_t> conflictMetadataUs{0};
            uint64_t conflictConvergenceWaitUs           = 0;
            uint64_t conflictPromptCount                 = 0;
            uint64_t consentWaitUs                       = 0;
            uint64_t consentPromptCount                  = 0;
            uint64_t queueEnterCount                     = 0;
            uint64_t queueNotifyAllCount                 = 0;
            uint64_t queueCancelWhileWaiting             = 0;
            uint64_t queueDepthOnEnter                   = 0;
            uint64_t queueActiveOperations               = 0;
            uint64_t itemCompletedCallbackUs             = 0;
            uint64_t itemCompletedLockWaitUs             = 0;
            uint64_t itemCompletedLockHoldUs             = 0;
            uint64_t itemCompletedLockContentionCount    = 0;
            uint64_t itemCompletedPathUpdateBytes        = 0;
            uint64_t itemCompletedPathUpdateAppliedCount = 0;
            uint64_t itemCompletedPathUpdateSkippedCount = 0;
        };

        explicit Task(FileOperationState& state) noexcept;

        Task(const Task&)            = delete;
        Task(Task&&)                 = delete;
        Task& operator=(const Task&) = delete;
        Task& operator=(Task&&)      = delete;
        ~Task()                      = default;

        // IFileSystemCallback
        HRESULT STDMETHODCALLTYPE FileSystemProgress(FileSystemOperation operationType,
                                                     unsigned long totalItems,
                                                     unsigned long completedItems,
                                                     uint64_t totalBytes,
                                                     uint64_t completedBytes,
                                                     const wchar_t* currentSourcePath,
                                                     const wchar_t* currentDestinationPath,
                                                     uint64_t currentItemTotalBytes,
                                                     uint64_t currentItemCompletedBytes,
                                                     FileSystemOptions* options,
                                                     uint64_t progressStreamId,
                                                     void* cookie) noexcept override;

        HRESULT STDMETHODCALLTYPE FileSystemItemCompleted(FileSystemOperation operationType,
                                                          unsigned long itemIndex,
                                                          const wchar_t* sourcePath,
                                                          const wchar_t* destinationPath,
                                                          HRESULT status,
                                                          const FileSystemItemMutationResult* mutationResult,
                                                          FileSystemOptions* options,
                                                          void* cookie) noexcept override;

        HRESULT STDMETHODCALLTYPE FileSystemShouldCancel(BOOL* pCancel, void* cookie) noexcept override;

        HRESULT ReportVerificationProgress(const wchar_t* sourcePath,
                                           const wchar_t* destinationPath,
                                           uint64_t itemTotalBytes,
                                           uint64_t itemCompletedBytes,
                                           uint64_t completedBytes,
                                           bool active) noexcept;

        // IFileSystemOperationControl
        HRESULT STDMETHODCALLTYPE FileSystemShouldAbort(BOOL* abort, void* cookie) noexcept override;
        HRESULT STDMETHODCALLTYPE FileSystemGetDiscoveryMode(FileSystemDiscoveryMode* mode, void* cookie) noexcept override;
        HRESULT STDMETHODCALLTYPE FileSystemReportDiscoveryProgress(const FileSystemDiscoveryProgress* progress, void* cookie) noexcept override;

        HRESULT STDMETHODCALLTYPE FileSystemIssue(FileSystemOperation operationType,
                                                  const wchar_t* sourcePath,
                                                  const wchar_t* destinationPath,
                                                  HRESULT status,
                                                  FileSystemIssueAction* action,
                                                  IFileSystemBoundObject** expectedDestination,
                                                  FileSystemOptions* options,
                                                  void* cookie) noexcept override;

        void ThreadMain(std::stop_token stopToken) noexcept;
        void SkipDiscovery() noexcept;
        void RequestCancel() noexcept;
        // R0f cancel-watch predicate for this task's worker threads (see SynchronousIoCancelWatch.h).
        [[nodiscard]] static bool ShouldCancelSynchronousIo(void* context) noexcept;
        void SetPaused(bool paused) noexcept;
        void TogglePause() noexcept;
        void SetDesiredSpeedLimit(uint64_t bytesPerSecond) noexcept;
        void NoteLiveOutputPublished(std::wstring_view providerPath) noexcept;
        [[nodiscard]] LiveOutputGuardDisposition GuardLiveOutputBeforeInvalidation(
            std::wstring_view providerPath,
            FileOperations::MutationInterlockAccess access) noexcept;
        void SetWaitForOthers(bool wait) noexcept;
        void SetWaitingInQueue(bool waiting) noexcept;
        void SetQueuePaused(bool paused) noexcept;
        void ToggleConflictApplyToAllChecked() noexcept;
        void SubmitConflictDecision(ConflictAction action, bool applyToAllChecked) noexcept;
#ifdef ENABLE_TESTS
        HRESULT DebugRequestDeferredConsentForSelfTest(FileOperations::DeferredConsentRisk risk,
                                                       bool applyToAllEligible,
                                                       bool itemCountKnown,
                                                       uint64_t itemCount,
                                                       bool bytesKnown,
                                                       uint64_t bytes,
                                                       ConflictAction* selectedAction) noexcept;
        void DebugClearConcurrentOverlapRelationsForSelfTest() noexcept;
#endif

        bool HasStarted() const noexcept;
        bool HasEnteredOperation() const noexcept;
        [[nodiscard]] TaskLifecyclePhase GetLifecyclePhase() const noexcept;
        [[nodiscard]] std::shared_ptr<const PreparationSnapshot> LoadPreparationSnapshot() const noexcept;
        ULONGLONG GetEnteredOperationTick() const noexcept;
        bool IsPaused() const noexcept;
        bool IsWaitingForOthers() const noexcept;
        // R4-A02-1: the user chose Queue after these tasks for a same-host overlap; the admission wait is
        // that edge and the card offers no Start now.
        bool IsOverlapQueued() const noexcept;
        [[nodiscard]] bool AllowsConcurrentOverlapWith(uint64_t taskId) const noexcept;
        // The user or host asked for the cancel (durable, independent of which status a call returned).
        [[nodiscard]] bool CancelRequestedDurable() const noexcept;
        // A cancel status, or the aborted status of a call the cancel watch cut short under a requested cancel.
        [[nodiscard]] bool IsCancellationOutcome(HRESULT hr) const noexcept;
        bool IsWaitingInQueue() const noexcept;
        bool IsQueuePaused() const noexcept;
        uint64_t GetQueueOrderKey() const noexcept;
        [[nodiscard]] bool IsPresentationHidden() const noexcept;
        [[nodiscard]] bool IsPresentationVisible() const noexcept;
        void RequestPresentationReveal() noexcept;
        void RequestActionablePromptPresentation() noexcept;

        void SetDestinationFolder(const std::filesystem::path& folder);
        std::filesystem::path GetDestinationFolder() const;

        unsigned long GetPlannedItemCount() const noexcept;

        uint64_t GetId() const noexcept;
        HRESULT GetResult() const noexcept;

        FileSystemOperation GetOperation() const noexcept;
        FolderWindow::Pane GetSourcePane() const noexcept;
        std::optional<FolderWindow::Pane> GetDestinationPane() const noexcept;

        void WaitWhilePaused(const std::atomic<bool>* externalStop = nullptr) noexcept;
        void WakePauseWaiters() noexcept;
        void MarkRateSamplingStateChanged() noexcept;

        HRESULT ExecuteOperation() noexcept;
        [[nodiscard]] std::shared_ptr<const FileOperations::FileOperationPlanGroup> LoadPlans() const noexcept;
        void StorePlans(std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans) noexcept;
        HRESULT PrepareBatchRenameAdmission() noexcept;
        HRESULT PrepareBatchRenameArtifactGuard() noexcept;
        void CompleteBatchRenameArtifactPrompt(HRESULT status,
                                               std::optional<FileOperationArtifacts::TouchGuardReceipt> receipt) noexcept;
        HRESULT RevalidateArtifactTouchGuard() noexcept;
        HRESULT ExecuteInlineRename() noexcept;
        HRESULT ExecuteBatchRename() noexcept;
        HRESULT ExecuteBatchRenameMutation(size_t operationIndex,
                                           const std::filesystem::path& sourcePath,
                                           const std::filesystem::path& destinationPath,
                                           std::vector<FileOperations::BoundObjectAuthority>& authorities) noexcept;
        void InitializeSourceItemResultBuilders() noexcept;
        void MarkSourceItemsMutationPossible() noexcept;
        [[nodiscard]] static bool ClipboardMutationGateAllows(HRESULT readinessStatus,
                                                               HRESULT consumptionStatus,
                                                               bool consumed) noexcept;
        [[nodiscard]] bool StoreTypedItemResult(FileOperations::FileOperationItemResult result) noexcept;
        HRESULT FinalizeTypedItemResults(HRESULT operationStatus) noexcept;
        HRESULT PrepareForExecution() noexcept;
        HRESULT PrepareMutationInterlockScopes() noexcept;
        // R1d-OR1: the destination child-name contract for every transfer leaf, on the task thread.
        [[nodiscard]] HRESULT PrepareTransferDestinationNames(const FileOperations::FileOperationPlanGroup& plans) noexcept;
        HRESULT RunSameHostOverlapAdvisory() noexcept;
        HRESULT RunPermanentDeleteConfirmation() noexcept;
        // C10: pins the identity of a native-authority delete root through IFileSystemIdentityDelete
        // (S_OK without a pin when the provider has no such contract).
        HRESULT PinNativeDeleteIdentity(const FileOperations::QualifiedEndpoint& endpoint, std::wstring_view providerPath) noexcept;
        [[nodiscard]] bool PermanentDeleteRunsByNameOnly() const noexcept;
        [[nodiscard]] bool IsInlineRenameTask() const noexcept;
        HRESULT BuildPreparationSnapshot(uint64_t selectedRootReadinessUs) noexcept;
        HRESULT RunPreConsumptionDecisionGate() noexcept;
        HRESULT CreateMoveBreadcrumbAfterAcceptance() noexcept;
        void PublishLifecyclePhase(TaskLifecyclePhase phase) noexcept;
        void CompleteClipboardMoveAdmission(HRESULT status, bool consumed) noexcept;
        void BeginDiscovery() noexcept;
        void MarkDiscoveryItemClosed(size_t sourceIndex) noexcept;
        void CloseDiscovery() noexcept;
        void NoteDiscoveryCompletionWhileOpen(uint64_t observedAtPerfUs,
                                              uint64_t completedBytes,
                                              uint64_t completedMutations) noexcept;
        void InitializeFileSystemOptions(FileSystemOptions& options, void* operationControlCookie = nullptr) const noexcept;
        void LogDiagnostic(DiagnosticSeverity severity,
                           HRESULT status,
                           std::wstring_view category,
                           std::wstring_view message,
                           std::wstring_view sourcePath      = {},
                           std::wstring_view destinationPath = {}) noexcept;

        static HRESULT BuildPathArrayArena(const std::vector<std::filesystem::path>& paths,
                                           FileSystemArenaOwner& arenaOwner,
                                           const wchar_t*** outPaths,
                                           unsigned long* outCount) noexcept;

        FileOperationState* _state  = nullptr;
        FolderWindow* _folderWindow = nullptr;
        uint64_t _taskId            = 0;
        std::atomic<std::shared_ptr<const FileOperations::FileOperationPlanGroup>> _plans;
        std::atomic<TaskLifecyclePhase> _lifecyclePhase{TaskLifecyclePhase::Preparing};
        std::atomic<std::shared_ptr<const PreparationSnapshot>> _preparationSnapshot;
        std::optional<BatchRenameAdmissionInput> _pendingBatchRenameAdmission;
        FileSystemOperation _operation = FILESYSTEM_COPY;
        ExecutionMode _executionMode   = ExecutionMode::PerItem;
        FolderWindow::Pane _sourcePane = FolderWindow::Pane::Left;
        std::optional<FolderWindow::Pane> _destinationPane;
        std::wstring _sourcePluginId;
        std::wstring _sourcePluginShortId;
        std::wstring _sourceInstanceContext;
        std::wstring _destinationPluginId;
        std::wstring _destinationPluginShortId;
        std::wstring _destinationInstanceContext;
        wil::com_ptr<IFileSystem> _fileSystem;
        wil::com_ptr<IFileSystem> _destinationFileSystem;
        std::vector<std::filesystem::path> _sourcePaths;
        std::vector<FolderWindow::ResolvedFileOperationItem> _resolvedItems;
        // C1: the permanent-delete confirmation this task still owes on its card, in the modal's words.
        bool _permanentDeleteConfirmationPending = false;
        std::wstring _permanentDeleteConsentDetail;
        std::wstring _permanentDeleteConsentFrom;
        int _permanentDeleteTestOverride = 0;
        std::vector<DWORD> _sourcePathAttributesHint;
        std::vector<FileOperations::MutationInterlockScope> _mutationInterlockScopes;
        mutable std::mutex _operationMutex;
        std::filesystem::path _destinationFolder;
        FileSystemFlags _flags                  = FILESYSTEM_FLAG_NONE;
        unsigned long _crossFsBridgeBufferBytes = 4096u * 1024u;
        std::atomic<unsigned long> _resolvedCrossFsBridgeBufferBytes{0};
        std::atomic<bool> _autoConcurrencyUsed{false};
        std::atomic<uint32_t> _autoConcurrencyStorageKind{FILESYSTEM_STORAGE_UNKNOWN};
        std::atomic<uint32_t> _autoConcurrencyDestinationStorageKind{FILESYSTEM_STORAGE_UNKNOWN};
        std::atomic<unsigned int> _autoTunedConcurrency{0};
        std::atomic<unsigned int> _effectiveConcurrencyBudget{0};
        bool _clipboardMoveAdmission = false;
        std::atomic<bool> _clipboardMoveConsumed{false};
        std::atomic<HRESULT> _clipboardMoveConsumptionStatus{S_OK};
        std::atomic<bool> _clipboardMoveBarrierCompleted{false};
        // Readiness is worker-owned for every task. Selected roots, immutable plan validation,
        // interlocks, snapshot publication, and the common pre-consumption decision gate complete
        // before any clipboard consumption or mutation release.
        std::atomic<bool> _selectedRootReadinessComplete{false};
        std::atomic<HRESULT> _selectedRootReadinessStatus{E_PENDING};
        std::function<HRESULT()> _preWorkerReleaseBarrier;
        std::function<HRESULT()> _preConsumptionDecisionGate;
        std::function<void(std::shared_ptr<const PreparationSnapshot>)> _preparationObserver;
        std::atomic<bool> _workerReleased{false};
#ifdef ENABLE_TESTS
        std::atomic<bool> _debugWorkerStartGateWaiting{false};
#endif
        std::function<void(uint64_t completedItems, uint64_t totalItems)> _externalProgressCallback;
        std::optional<FileOperationArtifacts::TouchGuardReceipt> _artifactTouchReceipt;
        std::mutex _batchRenameArtifactPromptMutex;
        std::condition_variable _batchRenameArtifactPromptCv;
        HRESULT _batchRenameArtifactPromptStatus = E_PENDING;
        bool _batchRenameArtifactPromptCompleted = false;
        std::optional<FileOperationMoveBreadcrumb::Breadcrumb> _moveBreadcrumb;

        unsigned long _perItemTotalItems          = 0;
        unsigned int _perItemMaxConcurrencyBudget = 1;
        unsigned int _perItemMaxConcurrency       = 1;
        unsigned long _perItemCompletedItems      = 0;
        uint64_t _perItemCompletedEntryCount      = 0;
        uint64_t _perItemTotalEntryCount          = 0;
        uint64_t _perItemCompletedBytes           = 0;

        struct PerItemInFlightCall
        {
            const void* cookie           = nullptr;
            unsigned long completedItems = 0;
            uint64_t completedBytes      = 0;
            unsigned long totalItems     = 0;
            ULONGLONG lastUpdateTick     = 0;
        };

        std::mutex _perItemInFlightCallsMutex;
        std::array<PerItemInFlightCall, kMaxInFlightFiles> _perItemInFlightCalls{};
        size_t _perItemInFlightCallCount        = 0;
        uint64_t _perItemInFlightCompletedBytes = 0;
        uint64_t _perItemInFlightCompletedItems = 0;
        uint64_t _perItemInFlightTotalItems     = 0;

        enum class TopLevelItemKind : uint8_t
        {
            Unknown,
            File,
            Folder,
        };
        std::mutex _topLevelCompletionMutex;
        std::vector<TopLevelItemKind> _topLevelItemKinds;
        std::vector<uint8_t> _topLevelItemCompleted;
        unsigned long _plannedTopLevelFiles     = 0;
        unsigned long _plannedTopLevelFolders   = 0;
        unsigned long _completedTopLevelFiles   = 0;
        unsigned long _completedTopLevelFolders = 0;

        std::mutex _sourceItemStatusMutex;
        std::vector<SourceItemResultBuilder> _sourceItemResultBuilders;
        std::atomic<uint64_t> _duplicateTerminalStoreRejectCount{0u};
        std::atomic<uint64_t> _invalidTerminalStoreRejectCount{0u};

        std::atomic<bool> _waitForOthers{false};
        std::atomic<bool> _overlapQueued{false};
        // R4-A02: fixed, worker-owned relation sets. Queue names only older task IDs, which
        // prevents predecessor cycles; Run bypasses the root interlock only for the exact
        // relations disclosed by the warning.
        std::array<uint64_t, FileOperations::kMaxSameHostOverlapRelations> _overlapPredecessorTaskIds{};
        size_t _overlapPredecessorTaskIdCount = 0u;
        // Set (release) when this task's advisory has decided, on every exit; an older task that
        // publishes later may then name this task, which never adds an edge back to it.
        std::atomic<bool> _overlapAdvisoryDecided{false};
        std::array<uint64_t, FileOperations::kMaxSameHostOverlapRelations> _concurrentOverlapTaskIds{};
        size_t _concurrentOverlapTaskIdCount = 0u;
        std::atomic<bool> _hasPublishedLiveOutput{false};
        // Set while the scopes are built (before publication): rename plans publish under their
        // parent WriteSource scopes instead of PublishDestination scopes.
        bool _publishesUnderWriteSourceScopes = false;
        std::atomic<bool> _liveOutputIndexOverflowed{false};
        std::atomic<bool> _waitingInQueue{false};
        // Stable engine-owned ordering for tasks that have not entered the operation.
        // Popup reorder commands swap these keys and then reorder the wait deque.
        std::atomic<uint64_t> _queueOrderKey{0};
        std::atomic<bool> _enteredOperation{false};
        std::atomic<ULONGLONG> _enteredOperationTick{0};
        std::atomic<bool> _cancelled{false};
        std::atomic<ULONGLONG> _cancelRequestedTick{0};
        std::atomic<bool> _paused{false};
        std::atomic<bool> _queuePaused{false};
        std::atomic<ULONGLONG> _rateSamplingStateChangeTick{0};
        std::atomic<bool> _started{false};
        std::atomic<ULONGLONG> _operationStartTick{0};
        std::atomic<uint64_t> _desiredSpeedLimitBytesPerSecond{0};
        std::atomic<uint64_t> _appliedSpeedLimitBytesPerSecond{0};
        std::atomic<uint64_t> _effectiveSpeedLimitBytesPerSecond{0};
        std::atomic<HRESULT> _resultHr{S_OK};
        std::atomic<bool> _taskFinished{false};
        std::atomic<bool> _observedSkipAction{false};
        std::atomic<TaskPresentationState> _presentationState{TaskPresentationState::NotApplicable};
        ULONGLONG _presentationDeadlineTick = 0;
        bool _suppressCleanCompletionSummary = false;

        ConflictArbiter _conflictArbiter;

        // Single-pass discovery/execution state. Discovery totals are estimates until every selected
        // top-level traversal closes; execution is allowed to publish before that point.
        std::atomic<bool> _discoveryAheadActive{false};
        std::atomic<bool> _discoverySkipped{false};
        std::atomic<bool> _discoveryClosed{false};
        std::atomic<ULONGLONG> _discoveryStartTick{0};
        std::atomic<uint64_t> _discoveryStartPerfUs{0};
        std::atomic<ULONGLONG> _discoveryClosedTick{0};
        std::atomic<ULONGLONG> _discoverySkipRequestedTick{0};
        std::atomic<ULONGLONG> _discoveryReservationReleasedTick{0};
        std::atomic<uint64_t> _discoveryFirstMutationUs{0};
        std::atomic<uint64_t> _discoveryCompletedBytesWhileOpen{0};
        std::atomic<uint64_t> _discoveryCompletedMutationsWhileOpen{0};
        std::atomic<uint64_t> _discoveredTotalBytes{0};
        std::atomic<unsigned long> _discoveredFileCount{0};
        std::atomic<unsigned long> _discoveredDirectoryCount{0};
        std::atomic<uint64_t> _discoveryMaxQueueDepth{0};
        std::atomic<uint64_t> _discoveryStarvationCount{0};
        std::atomic<bool> _firstMutationBeforeDiscoveryClosed{false};
        std::mutex _discoveryMutex;
        std::vector<uint8_t> _discoveryItemClosed;
        size_t _closedDiscoveryItemCount = 0;

        std::stop_token _stopToken{};
        std::mutex _pauseMutex;
        std::condition_variable _pauseCv;

        std::mutex _progressMutex;
        unsigned long _progressTotalItems     = 0;
        unsigned long _progressCompletedItems = 0;
        uint64_t _progressTotalBytes          = 0;
        uint64_t _progressCompletedBytes      = 0;
        uint64_t _progressItemTotalBytes      = 0;
        uint64_t _progressItemCompletedBytes  = 0;
        std::atomic<unsigned long> _publishedProgressTotalItems{0};
        std::atomic<unsigned long> _publishedProgressCompletedItems{0};
        std::atomic<uint64_t> _publishedProgressTotalBytes{0};
        std::atomic<uint64_t> _publishedProgressCompletedBytes{0};
        std::atomic<uint64_t> _publishedProgressItemTotalBytes{0};
        std::atomic<uint64_t> _publishedProgressItemCompletedBytes{0};
        std::atomic<bool> _verificationRequested{false};
        std::atomic<bool> _verificationActive{false};
        std::atomic<uint64_t> _verificationTotalBytes{0};
        std::atomic<uint64_t> _verificationCompletedBytes{0};
        std::atomic<uint64_t> _verificationItemTotalBytes{0};
        std::atomic<uint64_t> _verificationItemCompletedBytes{0};
        std::atomic<uint64_t> _verificationReadCount{0};
        std::atomic<uint64_t> _verificationProviderProofCount{0};
        std::atomic<uint64_t> _verificationHostReadbackCount{0};
        std::atomic<unsigned long> _publishedCompletedTopLevelFiles{0};
        std::atomic<unsigned long> _publishedCompletedTopLevelFolders{0};
        std::mutex _progressPathMutex;
        std::wstring _progressSourcePath;
        std::wstring _progressDestinationPath;
        std::wstring _lastProgressCallbackSourcePath;
        std::wstring _lastProgressCallbackDestinationPath;
        std::atomic<std::shared_ptr<const DiagnosticPathSnapshot>> _publishedDiagnosticPathSnapshot{};
        ULONGLONG _lastVisibleProgressPathUpdateTick = 0;
        ULONGLONG _lastProgressCallbackTick          = 0;
        unsigned long _lastItemIndex                 = 0;
        HRESULT _lastItemHr                          = S_OK;
        std::atomic<uint64_t> _progressCallbackCount{0};
        std::atomic<uint64_t> _itemCompletedCallbackCount{0};

        std::mutex _inFlightFilesMutex;
        std::array<InFlightFileProgress, kMaxInFlightFiles> _inFlightFiles{};
        size_t _inFlightFileCount = 0;
        std::mutex _progressStreamPerfMutex;
        std::array<ProgressStreamPerf, kMaxInFlightFiles> _progressStreamPerf{};
        size_t _progressStreamPerfCount = 0;
        std::array<ConflictWorkerPerf, kMaxInFlightFiles> _conflictWorkerPerf{};
        size_t _conflictWorkerPerfCount = 0;

        PerfStats _perf{};
        std::atomic<uint64_t> _bridgeDirectoryEnsureCount{0};
        std::atomic<uint64_t> _bridgeSourceDirectoryEnumerationCount{0};
        std::atomic<uint64_t> _bridgeFileAdmissionCount{0};
        std::atomic<uint64_t> _bridgeFileStartedBeforeProducerDone{0};
        std::atomic<uint64_t> _bridgeAdmissionMaxQueueDepth{0};
        std::atomic<uint64_t> _bridgeTraversalMaxDepth{0};
        std::atomic<uint64_t> _bridgeTraversalMaxRetainedEntries{0};
        std::atomic<uint64_t> _bridgeTraversalMaxQueuedPathBytes{0};
        std::atomic<uint64_t> _bridgeTraversalMaxMetadataBytes{0};
        std::atomic<uint64_t> _bridgeTraversalLimitHitCount{0};
        std::atomic<uint64_t> _permanentDeleteIdentityMismatchCount{0};
        std::atomic<uint64_t> _permanentDeleteConditionalAttemptCount{0};
        std::atomic<uint64_t> _permanentDeleteRemovedCount{0};
        std::atomic<uint64_t> _permanentDeleteIndeterminateCount{0};
        std::atomic<uint64_t> _conflictExpectedDestinationBoundCount{0};
        std::atomic<uint64_t> _conflictExpectedDestinationUnavailableCount{0};
        std::atomic<uint64_t> _conflictExpectedDestinationReturnedCount{0};
        std::atomic<uint64_t> _conflictAtomicReplaceGrantCount{0}; // R3-1: replace decisions carried to an atomic-final writer (no bound authority)

#ifdef ENABLE_TESTS
        unsigned int _dbgConfiguredMaxConcurrency      = 1;
        ULONGLONG _dbgSingleInFlightStartTick          = 0;
        ULONGLONG _dbgLastSingleInFlightWarnTick       = 0;
        bool _dbgObservedMultipleInFlightFiles         = false;
        ULONGLONG _dbgLastPerItemInFlightEvictWarnTick = 0;
        std::atomic_uint32_t _dbgCallbackActiveScopeCount{0};
        std::atomic_uint32_t _dbgPauseWaiterCount{0};
        std::atomic<bool> _dbgConflictWaitBeforeSleepGate{false};
        std::atomic<bool> _dbgConflictWaitBeforeSleepReached{false};
#endif

        std::jthread _thread;
    };

    explicit FileOperationState(FolderWindow& owner);

    FileOperationState(const FileOperationState&)            = delete;
    FileOperationState(FileOperationState&&)                 = delete;
    FileOperationState& operator=(const FileOperationState&) = delete;
    FileOperationState& operator=(FileOperationState&&)      = delete;

    ~FileOperationState();

    HRESULT AdmitOperation(FileSystemOperation operation,
                           FolderWindow::Pane sourcePane,
                           std::optional<FolderWindow::Pane> destinationPane,
                           const wil::com_ptr<IFileSystem>& fileSystem,
                           std::vector<std::filesystem::path> sourcePaths,
                           std::filesystem::path destinationFolder,
                           FileSystemFlags flags,
                           bool waitForOthers,
                           uint64_t initialSpeedLimitBytesPerSecond                           = 0,
                           ExecutionMode executionMode                                        = ExecutionMode::PerItem,
                           bool requireConfirmation                                           = false,
                           wil::com_ptr<IFileSystem> destinationFileSystem                    = nullptr,
                           uint64_t* taskIdOut                                                = nullptr,
                           std::vector<FolderWindow::ResolvedFileOperationItem> resolvedItems = {},
                           std::wstring confirmationMessage                                   = {},
                           std::wstring sourcePluginIdOverride                                = {},
                           std::wstring sourcePluginShortIdOverride                           = {},
                           std::optional<std::wstring> sourceInstanceIdOverride               = std::nullopt,
                           FileOperations::DeleteOrigin deleteOrigin                          = FileOperations::DeleteOrigin::PaneCommand,
                           std::optional<FileOperations::ConsentKind> capturedConsentKind     = std::nullopt,
                           std::optional<uint32_t> moveClipboardSequence                      = std::nullopt,
                           std::function<HRESULT()> preWorkerReleaseBarrier                   = {},
                           std::function<HRESULT()> preConsumptionDecisionGate                = {},
                           std::function<void(std::shared_ptr<const FileOperations::PreparationSnapshot>)> preparationObserver = {});

    HRESULT AdmitInlineRename(FolderWindow::Pane sourcePane,
                              const wil::com_ptr<IFileSystem>& fileSystem,
                              std::filesystem::path sourcePath,
                              std::wstring finalLeafName,
                              uint64_t* taskIdOut                          = nullptr,
                              std::filesystem::path* providerParentPathOut = nullptr);

    HRESULT AdmitBatchRename(FolderWindow::Pane sourcePane,
                             const wil::com_ptr<IFileSystem>& fileSystem,
                             std::vector<BatchRenameExecutionOp> operations,
                             std::function<void(uint64_t completedItems, uint64_t totalItems)> progressCallback = {},
                             std::function<void(uint64_t taskId)> publishedCallback = {},
                             uint64_t* taskIdOut                                                                = nullptr,
                             std::function<HRESULT()> preConsumptionDecisionGate                               = {},
                             std::function<void(std::shared_ptr<const FileOperations::PreparationSnapshot>)> preparationObserver = {});

    HRESULT AdmitScheduledRename(FolderWindow::Pane sourcePane,
                                 const wil::com_ptr<IFileSystem>& fileSystem,
                                 FileOperations::RenameOrigin origin,
                                 std::vector<BatchRenameExecutionOp> operations,
                                 std::function<void(uint64_t completedItems, uint64_t totalItems)> progressCallback = {},
                                 std::function<void(uint64_t taskId)> publishedCallback = {},
                                 uint64_t* taskIdOut = nullptr,
                                 std::function<HRESULT()> preConsumptionDecisionGate = {},
                                 std::function<void(std::shared_ptr<const FileOperations::PreparationSnapshot>)> preparationObserver = {});

    void OnBatchRenameArtifactPrompt(std::unique_ptr<BatchRenameArtifactPromptPayload> payload) noexcept;
    void CompleteBatchRenameArtifactPromptByTaskId(
        uint64_t taskId,
        HRESULT status,
        std::optional<FileOperationArtifacts::TouchGuardReceipt> receipt = std::nullopt) noexcept;
    void OnClipboardMoveReady(std::unique_ptr<ClipboardMoveReadyPayload> payload) noexcept;
    void CompleteClipboardMoveAdmissionByTaskId(uint64_t taskId, HRESULT status) noexcept;

    void ShowPopup() noexcept;
    void OnPopupHiddenByUser() noexcept;

    HRESULT QualifyCreateDirectory(const wil::com_ptr<IFileSystem>& fileSystem,
                                   std::wstring_view pluginId,
                                   std::wstring_view instanceContext,
                                   const std::filesystem::path& providerParentPath,
                                   std::wstring_view leafName,
                                   FileOperations::CreateDirectoryAdmission& admissionOut) noexcept;

    HRESULT StartOperation(OperationAdmission admission, uint64_t* taskIdOut = nullptr);

    // External launch paths do not create a File Operations task, but they still consume the same
    // exact-set artifact warning receipt and immediately revalidate it before returning S_OK to the
    // process-launch caller. S_FALSE means the user canceled; any failure is fail-closed.
    HRESULT ConfirmExternalArtifactTouch(const wil::com_ptr<IFileSystem>& fileSystem,
                                         std::wstring_view pluginId,
                                         std::wstring_view instanceContext,
                                         std::span<const std::filesystem::path> providerPaths,
                                         FileOperationArtifacts::TouchGuardReceipt* receiptOut = nullptr) noexcept;

    void ApplyTheme(const AppTheme& theme);
    void Shutdown() noexcept;
    void NotifyQueueChanged();
    bool HasActiveOperations() noexcept;
    bool ShouldQueueNewTask() noexcept;
    void SetQueueNewTasks(bool queue) noexcept;
    bool GetQueueNewTasks() const noexcept;
    void ApplyQueueMode(bool queue) noexcept;
    bool RunQueuedTaskNow(uint64_t taskId) noexcept;
    bool MoveQueuedTask(uint64_t taskId, bool moveUp) noexcept;
    // A fresh queue-order key, larger than every task id and key handed out so far (it consumes an
    // id from the admission counter), for a task that must enter the queue behind a younger peer.
    [[nodiscard]] uint64_t AllocateLateQueueOrderKey() noexcept;
    void SetAllRunningTasksPaused(bool paused) noexcept;
    void CancelAll() noexcept;
    void CollectTasks(std::vector<Task*>& outTasks) noexcept;
    void CollectInformationalTasks(std::vector<FolderWindow::InformationalTaskUpdate>& outTasks) noexcept;
    void CollectCompletedTasks(std::vector<CompletedTaskSummary>& outTasks) noexcept;
    void CollectDiagnostics(std::vector<TaskDiagnosticEntry>& outEntries) noexcept;
    void CollectTaskDiagnosticSnapshot(uint64_t taskId, unsigned long& warningCount, unsigned long& errorCount, std::wstring& lastDiagnosticMessage) noexcept;
    [[nodiscard]] std::optional<UINT> RefreshDeferredTaskPresentation() noexcept;
    void RevealPendingTasksImmediately() noexcept;
    void DismissCompletedTask(uint64_t sessionTaskId) noexcept;
    uint64_t CreateOrUpdateInformationalTask(const FolderWindow::InformationalTaskUpdate& update) noexcept;
    void DismissInformationalTask(uint64_t taskId) noexcept;
    bool GetAutoDismissSuccess() const noexcept;
    void SetAutoDismissSuccess(bool enabled) noexcept;
    bool GetPopupFooterOnly() const noexcept;
    void SetPopupFooterOnly(bool footerOnly) noexcept;
    bool GetPopupCompactDensity() const noexcept;
    void SetPopupCompactDensity(bool compactDensity) noexcept;
    bool OpenDiagnosticsLogForTask(uint64_t taskId) noexcept;
    bool ExportTaskIssuesReport(uint64_t taskId, std::filesystem::path* reportPathOut = nullptr, bool openAfterExport = true) noexcept;
    bool OpenCompletedTaskSource(uint64_t taskId) noexcept;
    bool SelectCompletedTaskRetainedSources(uint64_t taskId) noexcept;
    bool CutCompletedTaskRetainedSourcesAgain(uint64_t taskId) noexcept;
    void ToggleIssuesPane() noexcept;
    bool IsIssuesPaneVisible() noexcept;
    bool TryGetIssuesPanePlacement(RECT& outRect, bool& outMaximized, UINT currentDpi) const noexcept;
    void SaveIssuesPanePlacement(HWND hwnd) noexcept;
    bool TryGetIssuesPaneViewState(std::wstring& outSortColumnId,
                                   bool& outSortDescending,
                                   std::vector<Common::Settings::GridColumnLayoutEntry>& outGridLayout) const noexcept;
    void SaveIssuesPaneViewState(std::wstring_view sortColumnId,
                                 bool sortDescending,
                                 const std::vector<Common::Settings::GridColumnLayoutEntry>& gridLayout) noexcept;
    bool TryGetPopupPlacement(RECT& outRect, bool& outMaximized, UINT currentDpi) const noexcept;
    bool TryGetPopupExpandedPlacement(RECT& outRect, UINT currentDpi) const noexcept;
    void SavePopupPlacement(HWND hwnd) noexcept;
    void SavePopupExpandedPlacement(HWND hwnd) noexcept;
    void OnPopupDestroyed(HWND hwnd) noexcept;
    void OnIssuesPaneDestroyed(HWND hwnd) noexcept;
    void UpdateLastPopupRect(const RECT& rect) noexcept;
    std::optional<RECT> GetLastPopupRect() noexcept;
#ifdef ENABLE_TESTS
    struct SettingsSaveDebugSnapshot
    {
        uint64_t queuedGeneration    = 0;
        uint64_t completedGeneration = 0;
        uint64_t coalescedCount      = 0;
        DWORD lastQueueThreadId      = 0;
        DWORD lastSaveThreadId       = 0;
        bool pending                 = false;
        bool saveInProgress          = false;
    };

    HWND GetPopupHwndForSelfTest() noexcept;
    void DebugEnsurePopupVisibleForSelfTest() noexcept;
    HWND GetIssuesPaneHwndForSelfTest() noexcept;
    void DebugResetIssuesPaneForSelfTest() noexcept;
    void DebugClearDiagnosticsForSelfTest() noexcept;
    void DebugRemoveDiagnosticsForTask(uint64_t taskId) noexcept;
    void DebugAppendCompletedTaskForSelfTest(CompletedTaskSummary summary) noexcept;
    bool DebugValidateRetainedSourceActionForSelfTest(uint64_t taskId, IFileSystem* fileSystem) noexcept;
    bool DebugFlushPendingSettingsSaveForSelfTest(DWORD timeoutMs) noexcept;
    SettingsSaveDebugSnapshot DebugGetSettingsSaveSnapshotForSelfTest() noexcept;
    void DebugForceNextFileOperationCompletedPostFailure() noexcept;
    void DebugForceNextOrphanedCompletionDrainSubmissionFailure() noexcept;
    void DebugClearConcurrentOverlapRelationsForSelfTest(Task& task) noexcept;
    void DebugForceLiveOutputIndexOverflowForSelfTest() noexcept;
    [[nodiscard]] bool DebugLiveOutputIndexOverflowedForSelfTest() noexcept;
    void DebugRecordCompletionApplyTidForSelfTest(DWORD threadId) noexcept;
    [[nodiscard]] bool DebugTakeFallbackCompletedPayloadForSelfTest(uint64_t taskId, TaskCompletedPayload& out) noexcept;
    [[nodiscard]] DWORD DebugLastFileOperationCompletionWorkerTid() const noexcept;
    [[nodiscard]] DWORD DebugLastFileOperationCompletionApplyTid() const noexcept;
    [[nodiscard]] uint64_t DebugLastRestoredCompletionWorkerTaskId() const noexcept;
    void DebugSetTaskPresentationNowTickForSelfTest(std::optional<ULONGLONG> nowTick) noexcept;
    [[nodiscard]] uint64_t DebugInlineRenameSilentCleanCompletionCount() const noexcept;
    [[nodiscard]] uint64_t DebugTaskPresentedCount() const noexcept;
    void DebugLoadInterruptedMoveBreadcrumbsForSelfTest() noexcept
    {
        LoadInterruptedMoveBreadcrumbs();
    }
#endif

    void RecordTaskDiagnostic(uint64_t taskId,
                              FileSystemOperation operation,
                              DiagnosticSeverity severity,
                              HRESULT status,
                              std::wstring_view category,
                              std::wstring_view message,
                              std::wstring_view sourcePath,
                              std::wstring_view destinationPath) noexcept;
    void EnqueueTaskDiagnostic(TaskDiagnosticEntry entry) noexcept;

    bool EnterOperation(Task& task, std::stop_token stopToken) noexcept;
    void LeaveOperation(Task& task) noexcept;
    void PostCompleted(Task& task) noexcept;

    Task* FindTask(uint64_t taskId) noexcept;
    void RemoveTask(uint64_t taskId) noexcept;
    // R4-A02-1: compare a prepared task's scopes with every other published prepared or active task in this host using
    // the interlock predicate (no enumeration, no provider call). taskCount counts the overlapping
    // tasks; problem and firstTaskId describe the first conflicting pair. Only older task IDs
    // are predecessors, preventing two simultaneously prepared tasks from queueing after each other.
    struct SameHostOverlapAdvice final
    {
        FileOperations::SameHostOverlapProblem problem = FileOperations::SameHostOverlapProblem::None;
        uint64_t firstTaskId                           = 0u;
        uint32_t taskCount                             = 0u;
        std::array<uint64_t, FileOperations::kMaxSameHostOverlapRelations> taskIds{};
        size_t relationCount = 0u;
        bool relationOverflow = false;

        [[nodiscard]] bool CanRunConcurrently() const noexcept
        {
            return ! relationOverflow && relationCount == taskCount && relationCount > 0u;
        }
    };
    void PublishPreparedMutationInterlock(Task& task) noexcept;
    void WithdrawPreparedMutationInterlock(Task& task) noexcept;
    [[nodiscard]] SameHostOverlapAdvice FindSameHostOverlap(const Task& task) noexcept;
    struct LiveOutputConflictAdvice final
    {
        uint64_t publisherTaskId = 0u;
        bool indexOverflow       = false;
    };
    void NoteLiveOutputPublished(Task& task, std::wstring_view providerPath) noexcept;
    [[nodiscard]] LiveOutputConflictAdvice FindLiveOutputConflict(
        const Task& task,
        std::wstring_view providerPath,
        FileOperations::MutationInterlockAccess access,
        const std::array<uint64_t, FileOperations::kMaxSameHostOverlapRelations>& ignoredPublisherTaskIds,
        size_t ignoredPublisherTaskIdCount) noexcept;
    [[nodiscard]] bool WaitForLiveOutputPublisher(Task& task, uint64_t publisherTaskId, std::stop_token stopToken) noexcept;
    [[nodiscard]] bool TakeFallbackCompletedPayload(uint64_t taskId, TaskCompletedPayload& out) noexcept;

private:
    struct OrphanedCompletionDrainContext final
    {
        OrphanedCompletionDrainContext()                                                 = default;
        OrphanedCompletionDrainContext(const OrphanedCompletionDrainContext&)            = delete;
        OrphanedCompletionDrainContext& operator=(const OrphanedCompletionDrainContext&) = delete;
        OrphanedCompletionDrainContext(OrphanedCompletionDrainContext&&)                 = delete;
        OrphanedCompletionDrainContext& operator=(OrphanedCompletionDrainContext&&)      = delete;

        FileOperationState* state = nullptr;
        TaskCompletedPayload completed{};
        std::jthread worker;
    };

    [[nodiscard]] bool ScheduleOrphanedCompletionDrain(Task& task, const TaskCompletedPayload& completed) noexcept;
    void DrainOrphanedCompletion(OrphanedCompletionDrainContext& context) noexcept;
    [[nodiscard]] bool TryPostCompletedWakeup(const TaskCompletedPayload& completed) noexcept;
    static void CALLBACK OrphanedCompletionDrainCallback(PTP_CALLBACK_INSTANCE instance, void* context) noexcept;
    void EnsurePopupVisible() noexcept;
    void RequestTaskPresentationRefresh() noexcept;
    void RequestActionablePromptPresentation() noexcept;
    [[nodiscard]] ULONGLONG TaskPresentationNowTick() const noexcept;
    CompletedTaskSummary RecordCompletedTask(Task& task) noexcept;
    bool ExecuteCompletedRetainedSourceAction(uint64_t taskId, bool cutAgain) noexcept;
    void FlushDiagnostics(bool force) noexcept;
    static std::filesystem::path GetDiagnosticsLogDirectory() noexcept;
    static std::filesystem::path GetDiagnosticsLogPathForDate(const SYSTEMTIME& localTime) noexcept;
    std::filesystem::path GetLatestDiagnosticsLogPathUnlocked() const noexcept;
    void RemoveFromQueue(uint64_t taskId) noexcept;
    void UpdateQueuePausedTasks() noexcept;
    void QueueSettingsSave(std::wstring_view context) noexcept;
    void FlushPendingSettingsSave() noexcept;
    void LoadInterruptedMoveBreadcrumbs() noexcept;

    FolderWindow& _owner;
    std::mutex _mutex;
    std::shared_ptr<void> _uiLifetime;
    std::vector<std::unique_ptr<Task>> _tasks;
    std::vector<FolderWindow::InformationalTaskUpdate> _informationalTasks;
    std::deque<CompletedTaskSummary> _completedTasks;
    static constexpr size_t kMaxAcceptedMoveClipboardSequences = 64u;
    std::deque<uint32_t> _acceptedMoveClipboardSequences;
    uint64_t _nextTaskId = 1;
    std::atomic<bool> _interruptedMoveSummariesPendingPresentation{false};
    std::atomic<bool> _actionablePromptPresentationPending{false};

#ifdef ENABLE_TESTS
    static constexpr ULONGLONG kTaskPresentationLiveClock = std::numeric_limits<ULONGLONG>::max();
    std::atomic<ULONGLONG> _debugTaskPresentationNowTick{kTaskPresentationLiveClock};
    std::atomic<uint64_t> _debugInlineRenameSilentCleanCompletionCount{0};
    std::atomic<uint64_t> _debugTaskPresentedCount{0};
#endif

    wil::unique_hwnd _popup;
    wil::unique_hwnd _issuesPane;
    std::optional<RECT> _lastPopupRect;

    std::mutex _queueMutex;
    std::condition_variable _queueCv;
    std::deque<uint64_t> _queue;
    unsigned long _activeOperations = 0;
    struct ActiveMutationInterlock final
    {
        uint64_t taskId                                                   = 0;
        Task* task                                                         = nullptr;
        const std::vector<FileOperations::MutationInterlockScope>* scopes = nullptr;
    };
    // A task publishes its immutable scope vector before asking for overlap consent. Keeping this
    // separate from the active set closes the preparation/preparation race without making a
    // not-yet-admitted task block execution. Activation transfers the entry atomically under
    // _queueMutex.
    std::vector<ActiveMutationInterlock> _preparedMutationInterlocks;
    std::vector<ActiveMutationInterlock> _activeMutationInterlocks;
    struct LivePublishedScope final
    {
        uint64_t taskId                                             = 0u;
        const FileOperations::MutationInterlockScope* plannedScope = nullptr;
        std::wstring providerPath;
    };
    static constexpr size_t kMaxLivePublishedScopes = 256u;
    std::array<LivePublishedScope, kMaxLivePublishedScopes> _livePublishedScopes{};
    size_t _livePublishedScopeCount      = 0u;
    bool _livePublishedIndexOverflow     = false;

    std::mutex _diagnosticsMutex;
    std::deque<TaskDiagnosticEntry> _diagnosticsInMemory;
    std::vector<TaskDiagnosticEntry> _diagnosticsPendingFlush;
    std::unordered_map<uint64_t, std::pair<unsigned long, unsigned long>> _taskDiagnosticCounts;
    std::unordered_map<uint64_t, std::wstring> _taskLastDiagnosticMessage;
    std::unordered_map<uint64_t, std::deque<TaskDiagnosticEntry>> _taskIssueDiagnostics;
    ULONGLONG _lastDiagnosticsFlushTick   = 0;
    ULONGLONG _lastDiagnosticsCleanupTick = 0;

    std::mutex _fallbackCompletedMutex;
    std::vector<TaskCompletedPayload> _fallbackCompletedPayloads;
    std::mutex _completionWorkerTransferMutex;
    std::mutex _orphanedCompletionDrainMutex;
    std::condition_variable _orphanedCompletionDrainCv;
    uint32_t _orphanedCompletionDrainsOutstanding = 0u;
    std::atomic_bool _completionShutdown{false};
#ifdef ENABLE_TESTS
    std::atomic_bool _debugForceNextCompletedPostFailure{false};
    std::atomic_bool _debugForceNextOrphanedCompletionDrainSubmissionFailure{false};
    std::atomic<DWORD> _debugLastCompletionWorkerTid{0};
    std::atomic<DWORD> _debugLastCompletionApplyTid{0};
    std::atomic<uint64_t> _debugLastRestoredCompletionWorkerTaskId{0};
#endif

    std::atomic<bool> _queueNewTasks{true};
};

// pluginId is diagnostic-only; flags select the Permanent Delete versus Recycle capability gate.
[[nodiscard]] bool CanSameFileSystemOperation(const wil::com_ptr<IFileSystem>& fileSystem,
                                              std::wstring_view providerPath,
                                              FileSystemOperation operation,
                                              std::wstring_view pluginId = {},
                                              FileSystemFlags flags = FILESYSTEM_FLAG_NONE) noexcept;
[[nodiscard]] bool IsAutoDismissableFileOperationCompletion(HRESULT resultHr, unsigned long warningCount, unsigned long errorCount) noexcept;
[[nodiscard]] bool ShouldRetryPublishedDestinationVerification(HRESULT hr) noexcept;

#ifdef ENABLE_TESTS
[[nodiscard]] bool CanCrossFileSystemCopyMoveForSelfTest(const wil::com_ptr<IFileSystem>& sourceFileSystem,
                                                         std::wstring_view sourceProviderPath,
                                                         std::wstring_view sourcePluginId,
                                                         const wil::com_ptr<IFileSystem>& destinationFileSystem,
                                                         std::wstring_view destinationProviderPath,
                                                         std::wstring_view destinationPluginId,
                                                         FileSystemOperation operation) noexcept;

[[nodiscard]] bool ResolveFileOpsAtomicWriterRouteForSelfTest(IFileSystemAtomicWriter* atomicWriter,
                                                              const wchar_t* destinationPath,
                                                              FileSystemFlags taskFlags,
                                                              bool overwriteGranted,
                                                              bool replaceReadOnlyGranted,
                                                              FileSystemFlags& effectiveFinalFlags,
                                                              FileSystemFlags& fallbackSiblingFlags) noexcept;

enum class FileOpsBridgePipelineMode : unsigned char
{
    Default,
    Disabled,
    Enabled,
};

enum class FileOpsBridgeReparsePolicyOverride : int
{
    None     = -1,
    Preserve = 0,
    Skip     = 1,
};

void SetFileOpsBridgePipelineModeForSelfTest(FileOpsBridgePipelineMode mode) noexcept;
FileOpsBridgePipelineMode GetFileOpsBridgePipelineModeForSelfTest() noexcept;
void SetFileOpsBridgeProducerDelayForSelfTest(unsigned int delayMs) noexcept;

void SetFileOpsBridgeTraversalDepthLimitForSelfTest(uint64_t maxDepth) noexcept;
unsigned int GetFileOpsBridgeProducerDelayForSelfTest() noexcept;
void SetFileOpsBridgeFailNextFileCopiesForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsBridgeFailNextFileCopyAttemptsForSelfTest() noexcept;
void SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsBridgeFailNextSourceGetSizeAttemptsForSelfTest() noexcept;
void SetFileOpsBridgeFailNextDestinationGetSizeForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsBridgeFailNextDestinationGetSizeAttemptsForSelfTest() noexcept;
void SetFileOpsBridgeFailNextDestinationOpenForSelfTest(unsigned long count, HRESULT status) noexcept;
unsigned long TakeFileOpsBridgeFailNextDestinationOpenAttemptsForSelfTest() noexcept;
// R0-RC3: the next destination basic-information read before a replace decision fails (no occupant token).
void SetFileOpsBridgeFailNextDestinationBasicInfoForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsBridgeFailNextDestinationBasicInfoAttemptsForSelfTest() noexcept;
void SetFileOpsBridgeNullNextSourceReaderForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsBridgeNullNextSourceReaderAttemptsForSelfTest() noexcept;
void SetFileOpsBridgeReportWrongDestinationSizeForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsBridgeReportWrongDestinationSizeAttemptsForSelfTest() noexcept;
void SetFileOpsBridgeFailNextStageEntropyForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsBridgeFailNextStageEntropyAttemptsForSelfTest() noexcept;
void SetFileOpsBridgeOverReportNextReadForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsBridgeOverReportNextReadAttemptsForSelfTest() noexcept;
void SetFileOpsBridgePrematureEofNextReadForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsBridgePrematureEofNextReadAttemptsForSelfTest() noexcept;
void SetFileOpsBridgeUnderConsumeNextWriteForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsBridgeUnderConsumeNextWriteAttemptsForSelfTest() noexcept;
void SetFileOpsBridgeOverReportNextWriteForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsBridgeOverReportNextWriteAttemptsForSelfTest() noexcept;
void SetFileOpsBridgeInjectHostileChildNamesForSelfTest(bool enabled) noexcept;
unsigned long TakeFileOpsBridgeInjectHostileChildNameAttemptsForSelfTest() noexcept;
void SetFileOpsBridgeInjectFileReparseForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsBridgeInjectFileReparseAttemptsForSelfTest() noexcept;
void SetFileOpsBridgeReparsePolicyOverrideForSelfTest(FileOpsBridgeReparsePolicyOverride policy) noexcept;
unsigned long TakeFileOpsBridgeReplacePublishedDestinationAttemptsForSelfTest() noexcept;
void SetFileOpsAutoConcurrencyOverrideForSelfTest(bool enabled, unsigned int preferredConcurrency, uint32_t storageKind) noexcept;
void SetFileOpsPostFinishedCompletionPauseForSelfTest(bool enabled) noexcept;
bool HasFileOpsPostFinishedCompletionPauseEnteredForSelfTest() noexcept;
void ReleaseFileOpsPostFinishedCompletionPauseForSelfTest() noexcept;
void SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(bool enabled) noexcept;
bool HasFileOpsBridgeMoveSourceCleanupPauseEnteredForSelfTest() noexcept;
void ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest() noexcept;
void SetFileOpsNativeMoveCreateDirectoryRaceForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsNativeMoveCreateDirectoryRaceAttemptsForSelfTest() noexcept;
void SetFileOpsManagedCleanupKnownNonCommitForSelfTest(HRESULT status, unsigned long count) noexcept;
unsigned long TakeFileOpsManagedCleanupKnownNonCommitAttemptsForSelfTest() noexcept;
void SetFileOpsPermanentDeleteKnownNonCommitForSelfTest(HRESULT status, unsigned long count) noexcept;
unsigned long TakeFileOpsPermanentDeleteKnownNonCommitAttemptsForSelfTest() noexcept;
void SetFileOpsManagedCleanupUnknownOutcomeForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsManagedCleanupUnknownOutcomeAttemptsForSelfTest() noexcept;
void SetFileOpsVerificationForceHostReadbackForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsVerificationForceHostReadbackAttemptsForSelfTest() noexcept;
void SetFileOpsVerificationForceUnavailableForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsVerificationForceUnavailableAttemptsForSelfTest() noexcept;
void SetFileOpsVerificationForceMismatchForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsVerificationForceMismatchAttemptsForSelfTest() noexcept;
void SetFileOpsVerificationCapabilityBudgetExceededForSelfTest(bool enabled) noexcept;
void SetFileOpsVerificationReadbackPauseForSelfTest(bool enabled) noexcept;
bool HasFileOpsVerificationReadbackPauseEnteredForSelfTest() noexcept;
void ReleaseFileOpsVerificationReadbackPauseForSelfTest() noexcept;
void SetFileOpsBridgePublishedDestinationRetryPauseForSelfTest(bool enabled) noexcept;
bool HasFileOpsBridgePublishedDestinationRetryPauseEnteredForSelfTest() noexcept;
void ReleaseFileOpsBridgePublishedDestinationRetryPauseForSelfTest() noexcept;
void SetFileOpsConflictMetadataPauseForSelfTest(bool enabled, ULONGLONG bailoutMs = 5'000ull) noexcept;
bool HasFileOpsConflictMetadataPauseEnteredForSelfTest() noexcept;
void ReleaseFileOpsConflictMetadataPauseForSelfTest() noexcept;
void SetFileOpsKeepBothNestedConflictPauseForSelfTest(bool enabled) noexcept;
bool HasFileOpsKeepBothNestedConflictPauseEnteredForSelfTest() noexcept;
void ReleaseFileOpsKeepBothNestedConflictPauseForSelfTest() noexcept;
void SetFileOpsInlineRenameBeforeMutationPauseForSelfTest(bool enabled) noexcept;
bool HasFileOpsInlineRenameBeforeMutationPauseEnteredForSelfTest() noexcept;
void ReleaseFileOpsInlineRenameBeforeMutationPauseForSelfTest() noexcept;
void SetFileOpsBatchRenameBeforeExecutionPauseForSelfTest(bool enabled) noexcept;
bool HasFileOpsBatchRenameBeforeExecutionPauseEnteredForSelfTest() noexcept;
void ReleaseFileOpsBatchRenameBeforeExecutionPauseForSelfTest() noexcept;
void SetFileOpsRecycleEscalationBeforeBindPauseForSelfTest(bool enabled) noexcept;
bool HasFileOpsRecycleEscalationBeforeBindPauseEnteredForSelfTest() noexcept;
void ReleaseFileOpsRecycleEscalationBeforeBindPauseForSelfTest() noexcept;
void SetFileOpsPermanentDeleteBeforeBindPauseForSelfTest(bool enabled) noexcept;
bool HasFileOpsPermanentDeleteBeforeBindPauseEnteredForSelfTest() noexcept;
void ReleaseFileOpsPermanentDeleteBeforeBindPauseForSelfTest() noexcept;
void SetFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest(bool enabled) noexcept;
bool HasFileOpsPermanentDeleteBeforeRecheckPauseEnteredForSelfTest() noexcept;
void ReleaseFileOpsPermanentDeleteBeforeRecheckPauseForSelfTest() noexcept;
void SetFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest(bool enabled) noexcept;
bool HasFileOpsPermanentDeleteBeforeLiveOutputGuardPauseEnteredForSelfTest() noexcept;
void ReleaseFileOpsPermanentDeleteBeforeLiveOutputGuardPauseForSelfTest() noexcept;
void SetFileOpsLiveOutputPublishedPauseForSelfTest(bool enabled) noexcept;
bool HasFileOpsLiveOutputPublishedPauseEnteredForSelfTest() noexcept;
void ReleaseFileOpsLiveOutputPublishedPauseForSelfTest() noexcept;
DWORD TakeFileOpsInlineRenameAdmissionThreadIdForSelfTest() noexcept;
std::optional<FileOperations::OperationStrategy> DebugGetPreparedTransferStrategyForSelfTest(uint64_t taskId) noexcept;
bool DebugGetPreparedTransferPlanShapeForSelfTest(uint64_t taskId, uint32_t* planCount, uint32_t* strategyMask) noexcept;
DWORD TakeFileOpsInlineRenameExecutionThreadIdForSelfTest() noexcept;
unsigned long TakeFileOpsInlineRenameExecutionAttemptsForSelfTest() noexcept;
struct FileOpsConflictMetadataDebugResult
{
    bool available           = false;
    bool isDirectory         = false;
    unsigned long attributes = 0;
    __int64 lastWriteTime    = 0;
};
bool DebugReadFileOpsConflictMetadataForSelfTest(IFileSystemIO* io, std::wstring_view path, FileOpsConflictMetadataDebugResult& out) noexcept;
bool DebugFileOpsUnboundCollisionWithholdsDestructiveActionsForSelfTest() noexcept;
bool DebugFileOpsConflictActionPolicyCoverageForSelfTest() noexcept;
bool RunFileOpsPerItemSchedulerShutdownQuietPointSelfTestForSelfTest(FolderWindow::FileOperationState& state) noexcept;
bool RunFileOpsPerItemSchedulerNestedSaturationSelfTestForSelfTest(FolderWindow::FileOperationState& state) noexcept;
bool RunFileOpsPerItemSchedulerFailurePolicySelfTestForSelfTest(FolderWindow::FileOperationState& state) noexcept;
bool RunFileOpsWorkerStartGateCancellationSelfTestForSelfTest(FolderWindow::FileOperationState& state) noexcept;
bool RunFileOpsBridgePausedReaderStopSelfTestForSelfTest(FolderWindow::FileOperationState& state) noexcept;
bool RunFileOpsBridgeDirectoryBufferValidationSelfTestForSelfTest() noexcept;
void ResetFileOpsBridgeBufferBudgetPeakForSelfTest() noexcept;
uint64_t GetFileOpsBridgeBufferBudgetPeakForSelfTest() noexcept;
uint64_t GetFileOpsBridgeBufferBudgetInUseForSelfTest() noexcept;
bool IsFileOpsCircuitBreakerTransientErrorForSelfTest(DWORD error) noexcept;
#endif
