#pragma once

// Internal implementation header for FolderWindow file operations.
// Keep this header private to the FolderWindow file-operation translation units.

#include "BatchRenameExecutionEngine.h"
#include "FileOperationArtifactRegistry.h"
#include "FileOperationMoveBreadcrumb.h"
#include "FolderWindowInternal.h"

#include <algorithm>
#include <array>
#include <chrono>
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

enum class ConflictActionPlacement : uint8_t
{
    Published,
    Primary,
    Overflow,
    Escape,
};

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
    LeftoverCleanup, // I35 R7.5: Clean Up Leftovers of Interrupted Operations (an ordinary pane Delete)
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
    bool verificationHostReadback                 = false;
    bool verificationProviderBlake3Proof          = false;
    bool verificationWriterDigestProof            = false; // R3-2: the destination writer proves the published content
    bool verificationCapabilityCheckDeferred      = false;
    bool cancellationAbort                        = false;
    bool cancellationDeadline                     = false;
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
    // Diagnostic provenance only; neither this nor an existence probe authorizes replacement.
    bool destinationIdentityUnsupported = false;
    bool samePathText                   = false;
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
    // I35 R7.3 (D-7.3): the root lives on a location that opens it but exposes no file identity (its
    // binding is Unsupported, never Indeterminate). pinnedDeleteIdentity then holds corroboration
    // facts (kind, attributes, size, times, the 64-bit index when the location reports one), not an
    // identity: the scope is conservative, never exact authority, and the root is deleted by name
    // only after its own consent and only while those facts still match on the deleting handle.
    bool byNameCorroborated = false;
    // I35 R7.3: why a Recycle root's best-effort by-name pin failed while the task prepared (its refused
    // Recycle then says so instead of offering a delete by name); S_OK otherwise.
    HRESULT byNamePinStatus = S_OK;
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
    MutationInterlockAccess access  = MutationInterlockAccess::WriteSource;
    bool exactDeleteAuthority       = false;
    bool conservativeIdentityDomain = false;
    bool byNameCorroborated         = false; // I35 R7.3: deleted by name after its own consent
};

struct PreparationStrategyFact final
{
    OperationStrategy strategy = OperationStrategy::Copy;
    uint64_t selectedRootCount = 0u;
};

// Immutable, authority-free observer facts published after worker-owned preparation. Bound
// objects, provider interfaces, descendant discoveries, and mutable execution state never enter
// this snapshot.
struct PreparationSnapshot final
{
    uint64_t taskId                  = 0u;
    FileSystemOperation operation    = FILESYSTEM_COPY;
    uint64_t selectedRootCount       = 0u;
    uint64_t copyOnlyCount           = 0u;
    uint64_t buildSnapshotUs         = 0u;
    uint64_t selectedRootReadinessUs = 0u;
    uint64_t retainedBytes           = 0u;
    HRESULT status                   = E_PENDING;
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

// Immutable inspection facts, not navigation success or mutation authority. A held
// context is the prompt-generation token; equal path strings do not make it current.
struct ConflictInspectionLocation
{
    QualifiedEndpoint endpoint;
    std::wstring providerPath;
    std::wstring parentPath;
    std::wstring itemName;
};

struct ConflictInspectionContext
{
    FileSystemOperation operation = FILESYSTEM_COPY;
    size_t sourceIndex            = 0u;
    std::optional<ConflictInspectionLocation> source;
    std::optional<ConflictInspectionLocation> destination;
};

// A retained adapter for the admitted role, independent of whichever provider a pane
// currently displays. Qualification is read-only and confers no operation consent.
struct ConflictInspectionTarget
{
    std::shared_ptr<const ConflictInspectionContext> context;
    uint64_t promptGeneration = 0u;
    bool destination          = false;
    wil::com_ptr<IFileSystem> fileSystem;
    std::wstring pluginId;
    std::wstring pluginShortId;
    std::wstring instanceContext; // Raw pane context, never the canonical endpoint instanceId.
    // The live prompt's no-follow kind of the role's item, when known; breaks a same-name tie only.
    std::optional<bool> itemIsDirectory;
};

[[nodiscard]] HRESULT QualifyConflictInspectionLocation(const ConflictInspectionLocation& location,
                                                        FileSystemOperation operation,
                                                        IFileSystem* fileSystem,
                                                        std::wstring_view pluginId,
                                                        std::wstring_view instanceContext) noexcept;

enum class ConflictInspectionNavigationStatus : uint8_t
{
    Selected,
    MissingOrFiltered,
    Ambiguous,
    Unavailable,
    Stale,
    NavigationFailed,
    RollbackFailed,
    Canceled, // The user canceled the pane's enumeration; the previous location was restored.
};
struct ConflictInspectionNavigationResult
{
    ConflictInspectionNavigationStatus status = ConflictInspectionNavigationStatus::Unavailable;
    HRESULT error                             = S_OK;
    bool focusUnchanged                       = false; // Caller still owns popup hide/return and final focus commit.
    uint64_t focusTicket                      = 0u;    // One-shot UI-thread commit; invalidated by a newer request.
};
using ConflictInspectionCompletion = std::function<void(const ConflictInspectionNavigationResult&)>;

// Empty name/parent pairs require the original exact-path fallback. Display-only
// qualification is weaker than the admitted endpoint required for inspection.
struct ConflictPathDisplay
{
    std::wstring sourceName;
    std::wstring sourceParentPath;
    std::wstring destinationName;
    std::wstring destinationParentPath;
    std::shared_ptr<const ConflictInspectionContext> inspection;
};

[[nodiscard]] ConflictPathDisplay ResolveConflictPathDisplay(const FileOperationPlanGroup& plans,
                                                             FileSystemOperation operation,
                                                             std::optional<size_t> sourceIndex,
                                                             std::wstring_view sourcePath,
                                                             std::wstring_view destinationPath) noexcept;

// Observation identity only, carried explicitly by the executing item. Neither
// paths nor a selected-root index can establish this association or grant consent.
struct ConflictObservationReference final
{
    uint64_t taskId           = 0u;
    uint64_t promptGeneration = 0u;

    bool operator==(const ConflictObservationReference&) const noexcept = default;
};

struct ConflictResultObservation final
{
    ConflictObservationReference reference;
    PublicationState publication                = PublicationState::Unknown;
    VerificationState verification              = VerificationState::NotRequested;
    SourceDisposition sourceDisposition         = SourceDisposition::Unknown;
    ItemCompletion completion                   = ItemCompletion::Indeterminate;
    OwnedStageDisposition ownedStageDisposition = OwnedStageDisposition::Unknown;
    HRESULT status                              = E_PENDING;

    bool operator==(const ConflictResultObservation&) const noexcept = default;
};

[[nodiscard]] std::wstring FormatConflictResultDiagnostic(const ConflictResultObservation& result);

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
    // Populated only by an executor carrying the exact accepted prompt through
    // this item's mutation. Absent means no conflict-outcome association is known.
    std::optional<ConflictObservationReference> conflictObservation;
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
    // I35 R7.3 (D-7.3): the by-name consents of a location without file identities. Distinct classes
    // keep their decisions, records and the task-scoped grant apart from the exact-identity ones.
    PermanentDeleteByNameConfirmation, // the permanent-delete card when a root is deleted by name
    RecycleFailedByName,               // a Recycle escalation to a delete by name
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
    // Recycle escalation only: Delete permanently plus the task-scoped grant for later items of the
    // same task whose Recycle also fails with a known non-commit on the same route (Q2). Accepted as
    // PermanentDelete with Apply-to-all, as SkipAll is accepted as Skip.
    PermanentDeleteAll,
    // Explicit permanent-delete retry for one exact retained selected root/subtree only.
    DeleteReadOnly,
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

enum class ConflictRepeatContext : uint8_t
{
    None,
    SelectedRoot,
    Descendant,
};

enum class ConflictRepeatUnavailableReason : uint8_t
{
    None,
    CacheFull,
    UnqualifiedDestination,
};

struct ConflictDecisionScope
{
    ConflictClass bucket                 = ConflictClass::Unknown;
    ConflictItemKind sourceKind          = ConflictItemKind::Unknown;
    ConflictItemKind destinationKind     = ConflictItemKind::Unknown;
    ConflictLinkKind destinationLinkKind = ConflictLinkKind::Unknown;
    std::wstring sourceProfileId;
    std::wstring destinationProfileId;
    std::wstring destinationRootId;
    bool applyToAllEligible = false;
    // Only Keep Both uses this action-specific key dimension. Other choices retain their
    // task-wide scope across selected roots and descendants. Non-mutating Skip for the
    // access/sharing/path/network failure families also ignores observed object shape.
    ConflictRepeatContext keepBothContext = ConflictRepeatContext::None;

    bool operator==(const ConflictDecisionScope&) const noexcept = default;
};

enum class ConflictDecisionOrigin : uint8_t
{
    User,
    CachedScope,
};

// A problem/choice join key is the containing diagnostic task ID plus this generation.
// The optional selected root is descriptive, never a descendant result join.
struct ConflictPromptObservation final
{
    uint64_t generation  = 0u;
    ConflictClass bucket = ConflictClass::Unknown;
    std::optional<size_t> sourceItemIndex;
    // Provider-qualified display facts only; no retained inspection/navigation capability.
    // Empty pairs require the original exact-path fallback, never reparsing here.
    std::wstring sourceName{};
    std::wstring sourceParentPath{};
    std::wstring destinationName{};
    std::wstring destinationParentPath{};
};

// Observational record only: this is never consulted to authorize a mutation. The selected
// root index is not a descendant identity and must not be used to infer a result correlation.
struct AcceptedConflictDecision final
{
    uint64_t taskId           = 0u;
    uint64_t promptGeneration = 0u;
    std::optional<size_t> sourceItemIndex;
    ConflictAction action = ConflictAction::None;
    ConflictDecisionScope scope{};
    bool applyToAll                      = false;
    ConflictDecisionOrigin origin        = ConflictDecisionOrigin::User;
    uint64_t originatingPromptGeneration = 0u;
    // Q10: a cached apply-to-all decision is one record per originating prompt generation, never
    // one per item. The record keeps its first application's prompt and item; appliedCount counts
    // every item the cache decided, and the last application closes the range. Zero for a user choice.
    uint64_t appliedCount         = 0u;
    uint64_t lastPromptGeneration = 0u;
    std::optional<size_t> lastSourceItemIndex;

    bool operator==(const AcceptedConflictDecision&) const noexcept = default;
};

[[nodiscard]] std::wstring FormatConflictActionLabel(ConflictAction action, ConflictClass bucket, FileSystemOperation operation) noexcept;
// The localized problem sentence of a conflict class, shared by the popup question and Issues.
[[nodiscard]] UINT ConflictClassMessageId(ConflictClass bucket) noexcept;
// Typed conflict text takes priority; an otherwise unclassified failure keeps its actual error
// explanation and code. The optional system text reuses an Issues view's cached OS lookup.
[[nodiscard]] std::wstring FormatConflictProblemText(ConflictClass bucket, HRESULT status, std::optional<std::wstring_view> systemText = std::nullopt) noexcept;
[[nodiscard]] std::wstring FormatAcceptedConflictDecisionDiagnostic(const AcceptedConflictDecision& decision);

enum class DeferredConsentRisk : uint8_t
{
    InsufficientSpace,
    SpaceUnknown,
    EfsPlaintext,
    SparseInflation,
    PlaceholderHydration,
    MetadataLoss,
    RecycleEscalation,
    SameHostOverlap,    // R4-A02-1: another live or queued task in this host overlaps this task's scopes
    SameHostLiveOutput, // R4-A02-2: this mutation would invalidate another concurrently-live task's output
    PermanentDelete,    // C1: the initial permanent-delete confirmation, asked on the card after Preparing pinned the roots
    // I35 R7.3 (D-7.3): one explicit consent before a root on a location without file identities is
    // deleted by name with corroboration (the card), or before a failed Recycle escalates to that.
    PermanentDeleteByName,
    RecycleEscalationByName,
};

// R4-A02-1: the concrete problem an overlap advisory names (roles of the first conflicting pair).
enum class SameHostOverlapProblem : uint8_t
{
    None,
    RemovesRead,      // this task writes/removes what the other task reads
    SameNames,        // both publish into the same destination names
    SameMembers,      // both remove or rename the same members
    RemovesPublished, // this task removes/renames what the other task publishes
    ReplacesRead,     // this task publishes over what the other task reads
    // I35 R7.3: the other task deletes a folder by name and this task writes or changes items there
    // (DeletedByNameWritten) or reads from there (DeletedByNameRead).
    DeletedByNameWritten,
    DeletedByNameRead,
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
// I35 R7.1: the touch guard's reason facts for one binding result (see FileOperationArtifacts::IdentityProbe).
[[nodiscard]] FileOperationArtifacts::IdentityProbe ToArtifactIdentityProbe(const ObjectBindingResult& bound) noexcept;
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
[[nodiscard]] bool QualifiedEndpointsHaveSamePathContract(const QualifiedEndpoint& left, const QualifiedEndpoint& right) noexcept;
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

    // I35 R10.1: what a transfer stream is doing between two of its progress reports. The progress
    // ABI carries no phase. The host derives it from each report's item bytes (a file announced at
    // zero bytes is Opening, a file whose bytes are all written is Finalizing until the stream names
    // another item), and the host's own cross-file-system bridge names its phases explicitly.
    enum class ProgressPhase : uint8_t
    {
        Unknown,
        Opening,
        Transferring,
        Finalizing,
        Publishing,
        Verifying,
    };
    static constexpr size_t kProgressPhaseCount = 6u;

    // I35 R10.2: the side a stalled transfer is most likely waiting on, from the storage it reads and
    // writes (a network or high-latency side is the one that stalls); Storage when neither or both are.
    enum class ProgressStallSide : uint8_t
    {
        Storage,
        Destination,
        Source,
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
        uint64_t taskId    = 0u;
        uint64_t requestId = 0u;
        FileOperationArtifacts::TouchGuardRequest request;
    };

    struct ClipboardMoveReadyPayload final
    {
        uint64_t taskId = 0u;
    };

    // Value-only construction inputs. UI ingress captures identities and options once; the
    // canonical builder may then qualify a proposed admission without publishing a task, reading
    // panes/settings, consuming a clipboard receipt or granting mutation authority.
    struct OperationConstruction
    {
        FileSystemOperation operation = FILESYSTEM_COPY;
        FolderWindow::Pane sourcePane = FolderWindow::Pane::Left;
        std::optional<FolderWindow::Pane> destinationPane;
        wil::com_ptr<IFileSystem> fileSystem;
        wil::com_ptr<IFileSystem> destinationFileSystem;
        std::vector<std::filesystem::path> sourcePaths;
        std::filesystem::path destinationFolder;
        FileSystemFlags flags = FILESYSTEM_FLAG_NONE;
        FileOperations::OperationOptions options;
        ExecutionMode executionMode = ExecutionMode::PerItem;
        std::vector<FolderWindow::ResolvedFileOperationItem> resolvedItems;
        std::wstring sourcePluginId;
        std::wstring sourcePluginShortId;
        std::wstring sourceInstanceContext;
        std::wstring destinationPluginId;
        std::wstring destinationPluginShortId;
        std::wstring destinationInstanceContext;
        FileOperations::DeleteOrigin deleteOrigin = FileOperations::DeleteOrigin::PaneCommand;
        std::optional<FileOperations::ConsentKind> capturedConsentKind;
        std::optional<uint32_t> moveClipboardSequence;
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
        // I35 D-5.2: skipped roots or a Move that kept sources (Partial without a counted warning)
        // keep the card; auto-dismiss never removes it.
        bool partialOutcome = false;
    };

    // Q9: a producer message that is a localized resource. Issues, the Problem detail and the
    // card's Last note render it in the live language; the text rendered when it was recorded is
    // the technical Message kept in logs and exports.
    struct LocalizedDiagnosticText final
    {
        UINT messageId = 0u; // 0: the producer's message is technical text only
        std::variant<std::monostate, std::wstring, unsigned long, std::pair<unsigned long, unsigned long>> argument;

        [[nodiscard]] std::wstring Render() const noexcept;
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
        std::optional<FileOperations::ConflictPromptObservation> conflictPrompt;
        std::optional<FileOperations::AcceptedConflictDecision> acceptedDecision;
        std::optional<FileOperations::ConflictResultObservation> conflictResult;
        std::wstring category;
        std::wstring message;
        LocalizedDiagnosticText userText;
        std::wstring sourcePath;
        std::wstring destinationPath;
        std::wstring concurrencyMode;
        std::wstring storageType;
        std::wstring destinationStorageType;
        unsigned long autoTunedConcurrency       = 0;
        unsigned long effectiveConcurrencyBudget = 0;
    };

    // Borrowed facts from one retained diagnostic. Do not retain these pointers past
    // the entry's lifetime; the pane and action gate must use the same qualification.
    struct QualifiedIssueFacts final
    {
        bool originalProblem = false;
        // I35 R5.2: the original problem is an Information note (a question the user was asked, or a
        // producer's localized note about an expected outcome). It is retained, joined and exported
        // like any original row, but shown only when fileOperations.showInformationalNotes is on.
        bool informational                                               = false;
        const FileOperations::AcceptedConflictDecision* acceptedDecision = nullptr;
        const FileOperations::ConflictResultObservation* conflictResult  = nullptr;

        [[nodiscard]] bool ProjectsRow() const noexcept
        {
            return originalProblem || acceptedDecision || conflictResult;
        }
        // A Warning/Error original row: visible whether or not informational notes are shown.
        [[nodiscard]] bool ProjectsProblemRow() const noexcept
        {
            return originalProblem && ! informational;
        }
    };

    [[nodiscard]] static QualifiedIssueFacts QualifyIssueFacts(const TaskDiagnosticEntry& issue) noexcept;
    // I35 R5.2: an Info entry that carries typed user-facing facts (a prompt observation or localized
    // producer text) and is neither a decision nor a result record. It is recorded whatever the
    // Advanced log verbosity, enters the task reservoir and export, and never counts as a problem.
    [[nodiscard]] static bool IsUserNote(const TaskDiagnosticEntry& issue) noexcept;
    // Q9: the localized primary problem text of an Issues row or a Last note, derived from typed
    // facts (producer resource text, conflict class, stable producer code, HRESULT). The English
    // engine message is technical detail only (Technical details, logs and exports). systemText,
    // when present, is the caller's already looked-up system message for issue.status (empty: the
    // status has none, so it is not looked up again); a status without one never shows an
    // "HRESULT 0x..." literal.
    [[nodiscard]] static std::wstring FormatIssueProblemText(const TaskDiagnosticEntry& issue,
                                                             std::optional<std::wstring_view> systemText = std::nullopt) noexcept;

    // The bounded per-task facts of a completed summary: everything except its per-root path
    // lists and issue reservoir. The popup copies only these facts (x-perf#3).
    struct CompletedTaskFacts
    {
        uint64_t taskId = 0;
        FileOperations::TaskSelectionSummary selection;
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
        // The verification half of a finished card's bar (Verify On): bytes verified and their total.
        uint64_t verificationTotalBytes     = 0;
        uint64_t verificationCompletedBytes = 0;

        // While the one traversal remains open, totals may be unknown; keep a best-effort
        // top-level type breakdown for UI even after discovery-ahead was skipped.
        bool discoverySkipped             = false;
        bool discoveryClosed              = false;
        unsigned long completedFiles      = 0;
        unsigned long completedFolders    = 0;
        bool copiedFileCountKnown         = false;
        uint64_t copiedFileCount          = 0u;
        bool copiedFolderCountKnown       = false;
        uint64_t copiedFolderCount        = 0u;
        bool completedWorkCountsKnown     = false;
        uint64_t completedWorkFileCount   = 0u;
        uint64_t completedWorkFolderCount = 0u;
        uint64_t completedWorkBytes       = 0u;
        std::wstring sourcePath;
        std::wstring destinationPath;

        bool autoConcurrencyUsed                       = false;
        uint32_t autoConcurrencyStorageKind            = FILESYSTEM_STORAGE_UNKNOWN;
        uint32_t autoConcurrencyDestinationStorageKind = FILESYSTEM_STORAGE_UNKNOWN;
        unsigned int autoTunedConcurrency              = 0;
        unsigned int effectiveConcurrencyBudget        = 0;

        unsigned long warningCount = 0;
        unsigned long errorCount   = 0;
        // I35 R5.3: Information notes the task recorded (questions asked, expected outcomes). They
        // never count as problems; the card shows them only when showInformationalNotes is on.
        unsigned long noteCount = 0;
        std::wstring lastDiagnosticMessage;
        std::wstring lastInfoNoteMessage;
        // I35 R5.6 / D-5.2: typed root outcomes. Skipped roots make the task Partial and keep its
        // card without a counted warning when the user's skips alone explain the partial result.
        unsigned long skippedItemCount   = 0;
        unsigned long failedItemCount    = 0;
        bool partialExplainedByUserSkips = false;
        // I35 R9.4 (P3): false only when no selected root crossed the mutation gate, no owned stage
        // was created, no cut list was used up and no item has an unknown outcome. A task canceled in
        // that state changed nothing; its result line says so instead of listing untouched sources.
        bool startedMutation = true;
        std::wstring resultSummary;
        unsigned long publishedItemCount               = 0;
        unsigned long unknownPublicationItemCount      = 0;
        unsigned long removedSourceCount               = 0;
        unsigned long retainedSourceCount              = 0;
        unsigned long retainedSourceNativeCount        = 0; // rename merges whose emptied source folder stayed
        unsigned long canceledRetainedSourceCount      = 0; // canceled before publishing: kept, never copied
        unsigned long unknownSourceCount               = 0;
        unsigned long missingSourceCount               = 0;
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

    struct CompletedTaskSummary : CompletedTaskFacts
    {
        struct RetainedSourceActionItem final
        {
            std::filesystem::path providerPath;
            FileOperations::ProviderIdentitySnapshot identity;
        };

        std::vector<std::filesystem::path> retainedSourcePaths;
        std::vector<std::filesystem::path> unknownSourcePaths;
        std::vector<RetainedSourceActionItem> exactRetainedSourceItems;
        std::vector<TaskDiagnosticEntry> issueDiagnostics;

        // Q9: the typed source of lastDiagnosticMessage, which CollectCompletedTaskPresentations
        // renders again when the live language changes. lastNoteText, a completion fallback, wins
        // over lastProblem; a summary with neither keeps its text as built.
        std::optional<TaskDiagnosticEntry> lastProblem;
        LocalizedDiagnosticText lastNoteText;
        uint64_t lastNoteLanguageRevision = 0u;
        // I35 R5.3: the last Information note, the card's Last note when no problem exists and
        // informational notes are shown; formatted into lastInfoNoteMessage like lastProblem.
        std::optional<TaskDiagnosticEntry> lastInfoNote;
    };

    // One retained summary as the popup presents it: the bounded facts plus what it derives
    // from the per-root lists. Shared immutably per issues revision (x-perf#3).
    struct CompletedTaskPresentation final
    {
        CompletedTaskFacts facts;
        bool hasSourceActionPaths           = false;
        size_t exactRetainedSourceItemCount = 0u;
        // I35 R5.3: any retained diagnostic (notes included); Show log and Export issues stay
        // reachable whatever the informational-notes setting.
        bool hasRetainedDiagnostics = false;
    };

    // The completed card's Last note in the current language: lastNoteText, else the typed text of
    // lastProblem, else empty.
    [[nodiscard]] static std::wstring FormatCompletedLastNote(const CompletedTaskSummary& summary) noexcept;
    // I35 R5.3: the typed text of the last Information note, else empty.
    [[nodiscard]] static std::wstring FormatCompletedLastInfoNote(const CompletedTaskSummary& summary) noexcept;

    // What the popup telemetry timer has to follow (x-perf#5). Waiting means every open task is
    // paused, queued, waiting for others or waiting on an unanswered prompt.
    enum class PopupTelemetryWork : uint8_t
    {
        None,
        Waiting,
        Active,
    };

    using MutationInterlockScopes = std::vector<FileOperations::MutationInterlockScope>;

    struct MutationInterlockRelations final
    {
        std::array<uint64_t, FileOperations::kMaxSameHostOverlapRelations> predecessorKeys{};
        size_t predecessorCount = 0u;
        std::array<uint64_t, FileOperations::kMaxSameHostOverlapRelations> concurrentKeys{};
        size_t concurrentCount = 0u;
        bool advisoryDecided   = false;

        [[nodiscard]] bool Allows(uint64_t key) const noexcept
        {
            const size_t count = (std::min)(concurrentCount, concurrentKeys.size());
            return std::ranges::find(concurrentKeys.begin(), concurrentKeys.begin() + static_cast<std::ptrdiff_t>(count), key) !=
                   concurrentKeys.begin() + static_cast<std::ptrdiff_t>(count);
        }
    };

    // A registry entry owns one immutable admission. Advisory completion replaces only
    // this small snapshot; all revisions retain the same immutable authority vector.
    struct MutationInterlockAdmission final
    {
        uint64_t key = 0u;
        std::shared_ptr<const MutationInterlockScopes> scopes;
        MutationInterlockRelations relations;
        bool publishesUnderWriteSourceScopes = false;
    };

    struct AdmissionDestinationInputs final
    {
        // Null means use the source provider. An explicitly supplied identical provider
        // remains non-null because this participates in native/bridge selection.
        wil::com_ptr<IFileSystem> fileSystem;
        std::filesystem::path folder;
        std::optional<FolderWindow::Pane> pane;
        std::wstring pluginId;
        std::wstring pluginShortId;
        std::wstring instanceContext;
        std::vector<FolderWindow::ResolvedFileOperationItem> resolvedItems;
    };

    struct SelectedAdmissionView final
    {
        std::shared_ptr<const AdmissionDestinationInputs> destination;
        std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans;
        std::shared_ptr<const FileOperationArtifacts::TouchGuardReceipt> artifactTouchReceipt;
        std::shared_ptr<const FileOperations::PreparationSnapshot> preparation;
        std::shared_ptr<const MutationInterlockAdmission> interlock;
    };

    enum class OperationEntryOutcome : uint8_t
    {
        Cancelled,
        Entered
    };
    struct OperationEntryResult final
    {
        OperationEntryOutcome outcome = OperationEntryOutcome::Cancelled;
        std::shared_ptr<const SelectedAdmissionView> accepted;
    };

    struct Task final : public IFileSystemCallback, public IFileSystemOperationControl, public IFileSystemFileWorkCallback
    {
        // Maximum number of in-flight file lines the popup can display for a single task.
        // This should be >= the Copy/Move worker concurrency cap so parallel file copies can be represented.
        static constexpr size_t kMaxInFlightFiles = 16u;

        using ConflictBucket      = FileOperations::ConflictClass;
        using ConflictAction      = FileOperations::ConflictAction;
        using TaskLifecyclePhase  = FileOperations::TaskLifecyclePhase;
        using PreparationSnapshot = FileOperations::PreparationSnapshot;

        enum class TaskPresentationState : uint8_t
        {
            NotApplicable,
            Hidden,
            RevealRequested,
            Presented,
            SuppressedCleanSuccess,
        };

        using ConflictItemKind      = FileOperations::ConflictItemKind;
        using ConflictLinkKind      = FileOperations::ConflictLinkKind;
        using ConflictDecisionScope = FileOperations::ConflictDecisionScope;

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

        struct CachedConflictDecision
        {
            ConflictDecisionScope scope{};
            ConflictAction action                = ConflictAction::None;
            uint64_t originatingPromptGeneration = 0u;
        };

        // Q10: the one retained Issues record of a cached apply-to-all grant (task + originating
        // prompt generation). Each cached application folds into record.acceptedDecision, and the
        // typed result of each such item folds into commonResult; nothing is retained per item.
        struct CachedConflictAggregate
        {
            TaskDiagnosticEntry record; // first application; its prompt generation is the row identity
            std::optional<FileOperations::ConflictResultObservation> commonResult;
            uint64_t resultCount    = 0u;
            bool mixedResults       = false; // contradictory item results leave the aggregate result unknown
            ULONGLONG publishedTick = 0u;
            bool publicationPending = false; // folded since the last publication; a user prompt publishes it
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
            // Observational refusal only. The host must still obtain a fresh explicit
            // one-root/subtree DeleteReadOnly confirmation and retain exact authority.
            std::atomic<bool> readonlyDeleteBlocked{false};
            // Discovery scope. A selected root's traversal owner reports that root's cumulative
            // totals and closes its scope exactly once. A provider call the traversal owner makes
            // inside its own walk (exact source cleanup, a child's Native rename, an owned-stage
            // removal) is a different discovery scope: it carries the same cancellation, deadline
            // and bandwidth controls, but it must not publish child-local totals or closure into
            // the root. Such calls receive a separate cookie with suppression set.
            bool suppressDiscoveryReports = false;
            // A provisional provider mutation whose own report cannot yet know whether it resolved
            // the selected root: a Native directory rename that may continue as a host rename merge.
            // Cumulative counts are kept; only the one-way closure is withheld until the host knows
            // which traversal owns the rest of the root.
            std::atomic<bool> deferDiscoveryClosure{false};
            // A BulkItems Recycle resume call, or a BulkItems folder escalation: a further provider
            // call, made after the host closed discovery, for work the first call never reached (the
            // selected items it never attempted, or the escalated folder's descendants). Its cumulative
            // totals start at zero in this cookie and add to the task's totals; they are not growth
            // after close, because that closure covered only what the first call reached.
            bool resumedDiscoveryScope = false;
            bool sourceIsDirectory     = false;
            // I35 R7.3: discovery reports through this cookie that counted an object, suppressed or not. A
            // delete by name reports its root only after the facts matched on the deleting handle, so a
            // change of this count across one call tells a stop after that check from a refused open.
            std::atomic<uint32_t> objectDiscoveryReports{0u};
            uint64_t lastDiscoveredBytes       = 0;
            uint64_t lastDiscoveredFiles       = 0;
            uint64_t lastDiscoveredDirectories = 0;
            std::array<unsigned int, static_cast<size_t>(ConflictBucket::Count)> issueRetryCounts{};
        };

        struct ConflictPromptState
        {
            uint64_t generation = 0u;
            std::optional<size_t> sourceItemIndex;
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
            FileOperations::ConflictPathDisplay pathDisplay;
            std::array<ConflictAction, kMaxActions> actions{};
            size_t actionCount = 0;
            std::array<ConflictAction, kMaxActions> primaryActions{};
            size_t primaryActionCount = 0;
            std::array<ConflictAction, kMaxActions> overflowActions{};
            size_t overflowActionCount                                              = 0;
            ConflictAction defaultAction                                            = ConflictAction::None;
            ConflictAction escapeAction                                             = ConflictAction::None;
            bool applyToAllEligible                                                 = false;
            bool applyToAllChecked                                                  = false;
            FileOperations::ConflictRepeatUnavailableReason repeatUnavailableReason = FileOperations::ConflictRepeatUnavailableReason::None;
            bool skipAllEligible                                                    = false;
            bool buttonsPublishable                                                 = false;
            bool retryFailed                                                        = false;
            unsigned int attemptCount = 0u; // C2: attempts already made for this item and bucket; each explicit Retry is one
            bool factItemCountKnown   = false;
            uint64_t factItemCount    = 0;
            bool factBytesKnown       = false;
            uint64_t factBytes        = 0;
            uint8_t overlapProblem    = 0u; // FileOperations::SameHostOverlapProblem (R4-A02-1)
            uint64_t overlapTaskId    = 0u; // the first overlapping task named by the advisory
            std::wstring consentDetail;     // C1: what a permanent-delete confirmation deletes (the counts text)
            // A paragraph the question of a delete consent ends with: how a location without object
            // binding or file identities deletes (C10, I35 R7.3) and the leftover warning of items the
            // touch guard handed to this consent (I35 R7.6). Empty for every other prompt.
            std::wstring consentNote;
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
            uint64_t nextPromptGeneration = 0u;
            bool decisionWasOffered       = false;
            uint64_t decisionGeneration   = 0u;
            DWORD ownerThreadId           = 0;
            std::optional<ConflictAction> decisionAction;
            bool decisionApplyToAll = false;
            wil::unique_event_nothrow decisionEvent;
            // Q10: one entry per originating apply-to-all generation this task applied from its cache.
            std::vector<CachedConflictAggregate> cachedAggregates;
            bool cachedAggregatesFinal = false; // taken by RecordCompletedTask; later folds are ignored
        };

        struct InFlightFileProgress
        {
            const void* cookieKey     = nullptr;
            uint64_t progressStreamId = 0;
            std::wstring sourcePath;
            uint64_t totalBytes                      = 0;
            uint64_t completedBytes                  = 0;
            ULONGLONG lastUpdateTick                 = 0;
            bool completionCountedWhileDiscoveryOpen = false;
            size_t scopeIndex                        = 0u;
            uint64_t logicalLeafId                   = 0u;
            FileSystemFileWorkPhase phase            = FileSystemFileWorkPhase::Begin;
            FileSystemFileWorkKind kind              = FileSystemFileWorkKind::RegularFile;
            bool explicitLifecycle                   = false;
            bool sizeKnown                           = false;
            bool sizeCounted                         = false;
            bool publicationCounted                  = false;
            bool terminal                            = false;
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
            // I35 R10.5: the phase and item of this stream's last report, so the next gap is
            // attributed to the phase it was spent in and to the item it belonged to.
            ProgressPhase lastPhase   = ProgressPhase::Unknown;
            uint64_t lastItemPathHash = 0;
            uint64_t lastUpdateUs     = 0;
        };

        // I35 R10.5: the cadence of the progress reports, measured where every route's reports arrive.
        // A gap is the time between two consecutive reports of one stream about the same item (the
        // file's own cadence; a stream's idle time between two items is no gap). A stall is such a
        // gap longer than kStallThresholdUs. A silence is a time longer than kStallThresholdUs in
        // which no stream reported at all: what the speed display sees. Updated under
        // _progressStreamPerfMutex.
        struct ProgressCadenceStats
        {
            // 1 ms buckets below 200 ms, 5 ms below 1 s, 50 ms below 10 s, 1 s below 120 s, then overflow.
            static constexpr size_t kGapBucketCount     = 651u;
            static constexpr uint64_t kStallThresholdUs = 600'000u;

            std::array<uint32_t, kGapBucketCount> gapHistogram{};
            uint64_t gapCount = 0;
            uint64_t gapSumUs = 0;
            uint64_t maxGapUs = 0;
            // Stalls by the phase the stream had announced when the gap began.
            std::array<uint32_t, kProgressPhaseCount> stallCountByPhase{};
            std::array<uint64_t, kProgressPhaseCount> stallUsByPhase{};
            // A stream's time between two items (its worker finished one file and later started another).
            uint64_t betweenItemGapCount = 0;
            uint64_t maxBetweenItemGapUs = 0;
            // Silences of the whole task, by the phase the streams reporting just before had announced.
            std::array<uint32_t, kProgressPhaseCount> silenceCountByPhase{};
            std::array<uint64_t, kProgressPhaseCount> silenceUsByPhase{};
            uint64_t maxSilenceUs             = 0;
            uint64_t lastTaskReportUs         = 0;
            ProgressPhase lastTaskReportPhase = ProgressPhase::Unknown;
            // Reports that repeat a stream's item in Finalizing without new bytes (a phase boundary after
            // the commit or the metadata, or the item's closing report): the R10.1 boundaries are seen.
            uint64_t finalizingRepeatCount = 0;
            // Throughput of every whole wall-clock second, from the first report that carried bytes.
            uint64_t throughputBucketCount     = 0;
            double throughputBucketSum         = 0.0;
            double throughputBucketSumSquares  = 0.0;
            uint64_t throughputBucketStartUs   = 0;
            uint64_t throughputBucketBaseBytes = 0;
            uint64_t firstByteUs               = 0;
            uint64_t firstBytes                = 0;
            uint64_t lastByteUs                = 0;
            uint64_t lastBytes                 = 0;

            [[nodiscard]] static size_t GapBucketIndex(uint64_t gapUs) noexcept;
            [[nodiscard]] static uint64_t GapBucketUpperUs(size_t index) noexcept;
            // The smallest bucket bound at or below which at least `fraction` of the gaps fall.
            [[nodiscard]] uint64_t GapPercentileUs(double fraction) const noexcept;
            [[nodiscard]] uint32_t StallCount() const noexcept;
            [[nodiscard]] uint64_t StallUs() const noexcept;
            [[nodiscard]] uint32_t SilenceCount() const noexcept;
            [[nodiscard]] uint64_t SilenceUs() const noexcept;
            // Coefficient of variation of the whole-second throughput buckets (0 without two buckets).
            [[nodiscard]] double ThroughputCoefficientOfVariation() const noexcept;
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
            std::atomic<uint64_t> bridgeDirectoryRestoreUs{0}; // created folders bound again and restored (I35)
            std::atomic<uint64_t> bridgeDirectoryRestoreCount{0};
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
            std::atomic<uint64_t> fileWorkCallbackUs{0};
            std::atomic<uint64_t> fileWorkCallbackCount{0};
            std::atomic<uint64_t> fileWorkLockWaitUs{0};
            std::atomic<uint64_t> fileWorkLockHoldUs{0};
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
        HRESULT STDMETHODCALLTYPE FileSystemReportFileWork(const FileSystemFileWorkProgress* progress, void* cookie) noexcept override;
        void PublishFileWorkAvailabilityLocked() noexcept; // _inFlightFilesMutex held

        // completedDeltaBytes: the bytes this item verified since its previous report. Each selected item
        // verifies through its own context; the task's verified-byte count is the sum over its items.
        HRESULT ReportVerificationProgress(const wchar_t* sourcePath,
                                           const wchar_t* destinationPath,
                                           uint64_t itemTotalBytes,
                                           uint64_t itemCompletedBytes,
                                           uint64_t completedDeltaBytes,
                                           bool active) noexcept;

        // I35 R10.5: a consistent copy of the progress-report cadence measured so far.
        [[nodiscard]] ProgressCadenceStats SnapshotProgressCadence() noexcept;

        // I35 R10.1: FileSystemProgress with the phase the caller names (the host's own bridge, which
        // knows its phases); Unknown derives the phase from the item bytes as for any provider report.
        HRESULT ReportProgressWithPhase(ProgressPhase phase,
                                        FileSystemOperation operationType,
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
                                        void* cookie) noexcept;

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
        // allowLiveInvalidation false withholds "Delete it anyway" (I35 R7.3: a delete by name never
        // invalidates another task's live output; it waits or skips).
        [[nodiscard]] LiveOutputGuardDisposition GuardLiveOutputBeforeInvalidation(std::wstring_view providerPath,
                                                                                   FileOperations::MutationInterlockAccess access,
                                                                                   bool allowLiveInvalidation = true) noexcept;
        void SetWaitForOthers(bool wait) noexcept;
        void SetWaitingInQueue(bool waiting) noexcept;
        void SetQueuePaused(bool paused) noexcept;
        void ToggleConflictApplyToAllChecked(uint64_t expectedGeneration) noexcept;
        bool SubmitConflictDecision(uint64_t expectedGeneration,
                                    ConflictAction action,
                                    std::optional<bool> applyToAllChecked,
                                    FileOperations::ConflictActionPlacement placement = FileOperations::ConflictActionPlacement::Published) noexcept;
#ifdef ENABLE_TESTS
        void DebugSubmitCurrentConflictDecision(ConflictAction action, bool applyToAllChecked) noexcept;
#endif
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
        [[nodiscard]] bool AllowsConcurrentOverlapWith(uint64_t admissionKey) const noexcept;
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
        // A paused, queued or prompt-blocked task runs again: re-evaluate the telemetry cadence now.
        void NotifyTelemetryWorkResumed() noexcept;

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
        HRESULT ExecuteOperation(std::shared_ptr<const SelectedAdmissionView> selectedAdmission) noexcept;
        [[nodiscard]] std::shared_ptr<const SelectedAdmissionView> LoadSelectedAdmission() const noexcept;
        // Initial construction only: supplies no scope authority and is forbidden once
        // selected-root preparation exists. The Task is not yet publicly runnable.
        void InitializeAdmissionInputs(std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans,
                                       AdmissionDestinationInputs inputs,
                                       std::optional<FileOperationArtifacts::TouchGuardReceipt> receipt = std::nullopt) noexcept;
        void StoreInitialArtifactTouchReceipt(std::optional<FileOperationArtifacts::TouchGuardReceipt> receipt) noexcept;
        [[nodiscard]] std::shared_ptr<const FileOperations::FileOperationPlanGroup> LoadPlans() const noexcept;
        void StorePlans(std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans) noexcept;
        HRESULT PrepareBatchRenameAdmission() noexcept;
        HRESULT PrepareBatchRenameArtifactGuard() noexcept;
        HRESULT PrepareArtifactTouchGuard(const FileOperations::FileOperationPlanGroup& plans,
                                          IFileSystem* sourceFileSystem,
                                          IFileSystem* destinationFileSystem,
                                          std::optional<FileOperationArtifacts::TouchGuardReceipt>& out) noexcept;
        [[nodiscard]] bool IsBatchRenameArtifactPromptPending(uint64_t requestId) noexcept;
        void CompleteBatchRenameArtifactPrompt(uint64_t requestId, HRESULT status, std::optional<FileOperationArtifacts::TouchGuardReceipt> receipt) noexcept;
        void FailBatchRenameArtifactPromptWithoutPayload(HRESULT status) noexcept;
        HRESULT RevalidateArtifactTouchGuard(const SelectedAdmissionView& admission) noexcept;
        HRESULT ExecuteInlineRename() noexcept;
        HRESULT ExecuteBatchRename() noexcept;
        HRESULT ExecuteBatchRenameMutation(size_t operationIndex,
                                           const std::filesystem::path& sourcePath,
                                           const std::filesystem::path& destinationPath,
                                           std::vector<FileOperations::BoundObjectAuthority>& authorities) noexcept;
        void InitializeSourceItemResultBuilders() noexcept;
        void MarkSourceItemsMutationPossible() noexcept;
        void MarkSourceItemMutationPossible(size_t sourceIndex) noexcept;
        void EnterSourceItemsMutationGate() noexcept;
        // I35 R8.1: a selected root crossed the governed-change boundary (per-root dispatch or the bulk
        // gate): from then on a task-wide prompt is no longer "before the first mutation".
        [[nodiscard]] bool HasMutated() noexcept;
        [[nodiscard]] static bool ClipboardMutationGateAllows(HRESULT readinessStatus, HRESULT consumptionStatus, bool consumed) noexcept;
        [[nodiscard]] bool StoreTypedItemResult(FileOperations::FileOperationItemResult result) noexcept;
        // Q10: the final record of each cached apply-to-all aggregate. Called once by
        // RecordCompletedTask; later applications and results are no longer folded.
        [[nodiscard]] std::vector<TaskDiagnosticEntry> TakeFinalCachedConflictAggregates() noexcept;
        HRESULT FinalizeTypedItemResults(HRESULT operationStatus) noexcept;
        // I35 R5.6 / D-5.2: selected roots by typed terminal completion (a missing terminal counts as
        // indeterminate). The skips alone explain a non-complete result when every root that did not
        // complete was skipped and nothing failed, was canceled or has an unknown outcome.
        struct RootOutcomeCounts final
        {
            unsigned long selected      = 0u;
            unsigned long completed     = 0u;
            unsigned long skipped       = 0u;
            unsigned long failed        = 0u;
            unsigned long canceled      = 0u;
            unsigned long indeterminate = 0u;

            [[nodiscard]] bool OnlySkipsExplainIncompleteRoots() const noexcept
            {
                return skipped > 0u && failed == 0u && canceled == 0u && indeterminate == 0u && completed + skipped == selected;
            }
        };
        [[nodiscard]] RootOutcomeCounts CountRootOutcomes() noexcept;
        // The task.result text for a partial Move, built from the finalized typed source dispositions.
        [[nodiscard]] std::wstring FormatPartialMoveResultMessage(unsigned long completedItems,
                                                                  unsigned long totalItems,
                                                                  uint64_t completedBytes,
                                                                  uint64_t totalBytes) noexcept;
        HRESULT PrepareForExecution() noexcept;
        HRESULT PrepareMutationInterlockScopes() noexcept;
        void InstallPreparedMutationInterlockScopes(MutationInterlockScopes scopes, bool publishesUnderWriteSourceScopes) noexcept;
        [[nodiscard]] std::shared_ptr<const MutationInterlockAdmission> LoadMutationInterlockAdmission() const noexcept;
        [[nodiscard]] std::shared_ptr<const MutationInterlockScopes> LoadMutationInterlockScopes() const noexcept;
        // R1d-OR1: the destination child-name contract for every transfer leaf, on the task thread.
        [[nodiscard]] HRESULT PrepareTransferDestinationNames(const FileOperations::FileOperationPlanGroup& plans) noexcept;
        HRESULT RunSameHostOverlapAdvisory() noexcept;
        HRESULT RunPermanentDeleteConfirmation() noexcept;
        // C10: pins the identity of a native-authority delete root through IFileSystemIdentityDelete
        // (S_OK without a pin when the provider has no such contract). I35 R7.3: requireContract pins
        // the corroboration facts of an identity-less Local root and fails without the contract.
        HRESULT PinNativeDeleteIdentity(const FileOperations::QualifiedEndpoint& endpoint,
                                        std::wstring_view providerPath,
                                        std::vector<FileOperations::MutationInterlockScope>& scopes,
                                        bool requireContract = false) noexcept;
        [[nodiscard]] bool PermanentDeleteRunsByNameOnly() const noexcept;
        // I35 R7.3: a permanent delete root this task deletes by name with corroboration.
        [[nodiscard]] bool PermanentDeleteHasCorroboratedByNameRoot() const noexcept;
        // I35 R7.3: refuses a by-name root that another live task holds in a prepared or active scope on
        // the same path, an ancestor or a descendant (S_OK when none does). Logs the refusal.
        [[nodiscard]] HRESULT RefuseByNameRootHeldByAnotherTask(const FileOperations::QualifiedEndpoint& endpoint, std::wstring_view providerPath) noexcept;
        // I35 R7.6: the paragraph a delete consent ends with (how the location deletes, and the
        // leftover warning when a handed-off root has a possible-artifact name).
        [[nodiscard]] std::wstring BuildPermanentDeleteConsentNote() const;
        [[nodiscard]] bool IsInlineRenameTask() const noexcept;
        HRESULT BuildPreparationSnapshot(uint64_t selectedRootReadinessUs) noexcept;
        HRESULT RunPreConsumptionDecisionGate() noexcept;
        HRESULT CreateMoveBreadcrumbAfterAcceptance() noexcept;
        void PublishLifecyclePhase(TaskLifecyclePhase phase) noexcept;
        void CompleteClipboardMoveAdmission(HRESULT status, bool consumed) noexcept;
        // I35 R6.10: who closed a task's discovery, recorded with the Debug `discovery.closed` line.
        enum class DiscoveryCloser : uint8_t
        {
            Producer,     // the provider's walker reported the root's traversal closed
            PerItemGuard, // the per-item executor closed a root whose call returned
            ExecutionEnd, // the task's execution ended (or was canceled) with discovery still open
        };
        void BeginDiscovery() noexcept;
        void MarkDiscoveryItemClosed(size_t sourceIndex, DiscoveryCloser closer) noexcept;
        // I35 R6.10: a root whose provider call started; with its discovery still open it is the only
        // kind of root that can be silent.
        void MarkDiscoveryItemStarted(size_t sourceIndex) noexcept;
        // Records the listing-paused note decided under the progress lock, after the caller left it.
        void FlushDiscoverySilenceNote() noexcept;
        void CloseDiscovery(DiscoveryCloser closer = DiscoveryCloser::ExecutionEnd) noexcept;
        void NoteDiscoveryCompletionWhileOpen(uint64_t observedAtPerfUs, uint64_t completedBytes, uint64_t completedMutations) noexcept;
        void InitializeFileSystemOptions(FileSystemOptions& options, void* operationControlCookie = nullptr) const noexcept;
        void LogDiagnostic(DiagnosticSeverity severity,
                           HRESULT status,
                           std::wstring_view category,
                           std::wstring_view message,
                           std::wstring_view sourcePath      = {},
                           std::wstring_view destinationPath = {}) noexcept;
        // A producer whose message is a localized resource passes it typed (Q9), so Issues and the
        // Last note show exactly that text in the live language.
        void LogDiagnostic(DiagnosticSeverity severity,
                           HRESULT status,
                           std::wstring_view category,
                           LocalizedDiagnosticText userText,
                           std::wstring_view sourcePath      = {},
                           std::wstring_view destinationPath = {}) noexcept;
        void RecordDiagnostic(DiagnosticSeverity severity,
                              HRESULT status,
                              std::wstring_view category,
                              std::wstring_view message,
                              LocalizedDiagnosticText userText,
                              std::wstring_view sourcePath,
                              std::wstring_view destinationPath) noexcept;

        static HRESULT BuildPathArrayArena(const std::vector<std::filesystem::path>& paths,
                                           FileSystemArenaOwner& arenaOwner,
                                           const wchar_t*** outPaths,
                                           unsigned long* outCount) noexcept;

        FileOperationState* _state  = nullptr;
        FolderWindow* _folderWindow = nullptr;
        uint64_t _taskId            = 0;
        // The admitted operation keeps its task identity and destination for its lifetime.
        std::atomic<uint64_t> _admissionKey{0u};
        std::atomic<std::shared_ptr<const SelectedAdmissionView>> _selectedAdmission;
        std::atomic<TaskLifecyclePhase> _lifecyclePhase{TaskLifecyclePhase::Preparing};
        std::optional<BatchRenameAdmissionInput> _pendingBatchRenameAdmission;
        FileSystemOperation _operation = FILESYSTEM_COPY;
        ExecutionMode _executionMode   = ExecutionMode::PerItem;
        FolderWindow::Pane _sourcePane = FolderWindow::Pane::Left;
        std::wstring _sourcePluginId;
        std::wstring _sourcePluginShortId;
        std::wstring _sourceInstanceContext;
        wil::com_ptr<IFileSystem> _fileSystem;
        std::vector<std::filesystem::path> _sourcePaths;
        // Initialized before publication and never changed by child progress.
        FileOperations::TaskSelectionSummary _selection;
        // C1: the permanent-delete confirmation this task still owes on its card, in the modal's words.
        bool _permanentDeleteConfirmationPending = false;
        std::wstring _permanentDeleteConsentDetail;
        std::wstring _permanentDeleteConsentFrom;
        int _permanentDeleteTestOverride = 0;
        std::vector<DWORD> _sourcePathAttributesHint;
        FileSystemFlags _flags                  = FILESYSTEM_FLAG_NONE;
        unsigned long _crossFsBridgeBufferBytes = 4096u * 1024u;
        std::atomic<unsigned long> _resolvedCrossFsBridgeBufferBytes{0};
        std::atomic<bool> _autoConcurrencyUsed{false};
        std::atomic<uint32_t> _autoConcurrencyStorageKind{FILESYSTEM_STORAGE_UNKNOWN};
        std::atomic<uint32_t> _autoConcurrencyDestinationStorageKind{FILESYSTEM_STORAGE_UNKNOWN};
        std::atomic<unsigned int> _autoTunedConcurrency{0};
        std::atomic<unsigned int> _effectiveConcurrencyBudget{0};
        bool _clipboardMoveAdmission = false;
        // I35 R5.6: the confirmation named a known Copy-only Move, so "Copied; source kept" is an expected
        // outcome (one Information note per task) rather than a deviation the user must act on.
        bool _copyOnlyMoveDisclosed = false;
        std::atomic<bool> _copyOnlyMoveNoteRecorded{false};
        // I35 XP-10: one note per task names a source that reports no file sizes.
        std::atomic<bool> _sizelessSourceNoteRecorded{false};
        // I35 R7.3: one note per task says items were deleted by name (the location exposes no identities).
        std::atomic<bool> _byNameDeleteNoteRecorded{false};
        // I35 R5.4/R5.5: one metadata classification for the bridge and the provider-native Copy
        // routes (FileSystemMetadataTransferResult meanings). Inherited permissions and a destination
        // that keeps no timestamps or permissions are counted for one Information note per task.
        struct MetadataOutcome final
        {
            uint32_t lostFeatures           = 0u;
            uint32_t changedFeatures        = 0u;
            uint32_t informationalFeatures  = 0u;
            uint32_t informationalObjects   = 1u;
            HRESULT firstFailure            = S_OK;
            bool destinationKeepsNoMetadata = false; // the route has no metadata authority at all
        };
        void ReportMetadataOutcome(std::wstring_view sourcePath, std::wstring_view destinationPath, const MetadataOutcome& outcome) noexcept;
        void NoteDestinationKeepsNoMetadata() noexcept;
        // Called once when execution ends, before the result is recorded.
        void RecordMetadataNotes() noexcept;
        std::atomic<uint64_t> _metadataInheritedObjects{0u};
        std::atomic<uint64_t> _metadataUnsupportedObjects{0u};
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
        std::mutex _batchRenameArtifactPromptMutex;
        std::condition_variable _batchRenameArtifactPromptCv;
        uint64_t _batchRenameArtifactPromptNextRequestId    = 0u;
        uint64_t _batchRenameArtifactPromptPendingRequestId = 0u;
        std::optional<FileOperationArtifacts::TouchGuardReceipt> _batchRenameArtifactPromptResultReceipt;
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
        // BulkItems Recycle resume only: while a further provider call recycles the items an earlier
        // call never reached, maps that call's item index to the task's source index. Guarded by
        // _sourceItemStatusMutex; empty otherwise.
        std::vector<size_t> _bulkResumeSourceIndices;
        std::atomic<uint64_t> _duplicateTerminalStoreRejectCount{0u};
        std::atomic<uint64_t> _invalidTerminalStoreRejectCount{0u};

        std::atomic<bool> _waitForOthers{false};
        std::atomic<bool> _overlapQueued{false};
        std::atomic<bool> _hasPublishedLiveOutput{false};
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
        ULONGLONG _presentationDeadlineTick  = 0;
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
        // Observation only: how many discovery reports added work after the task's totals were
        // already presented as final. Closure is one-way, so a correct traversal never grows after
        // it; any nonzero value means some scope closed a root it did not finish owning.
        std::atomic<uint64_t> _discoveryGrowthAfterCloseCount{0};
        // I35 R6.10: the last discovery report, and whether the one "listing paused" note was recorded.
        std::atomic<ULONGLONG> _discoveryLastReportTick{0};
        std::atomic<bool> _discoverySilenceRecorded{false};
        std::atomic<unsigned long> _discoverySilenceNoteSeconds{0};
        std::atomic<uint64_t> _discoveryOpenStartedRoots{0};
        std::mutex _discoveryMutex;
        std::vector<uint8_t> _discoveryItemClosed;
        std::vector<uint8_t> _discoveryItemStarted;
        // Fresh metadata counted before transfer admission. Written only by the task worker before
        // it launches item workers; later producer reports for those same leaf scopes are redundant.
        std::vector<std::optional<uint64_t>> _freshSelectedFileBytes;
        std::atomic<bool> _freshSelectedFileMetadataChanged{false};
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
        // A traversal producer can report links while every ordinary worker owns a file row.
        // Its reserved stream never consumes or replaces one of the sixteen visible file slots.
        std::array<InFlightFileProgress, kMaxInFlightFiles> _producerFileWork{};
        struct FileWorkScope
        {
            bool supported                = false;
            bool bytesKnown               = false;
            bool freshlyCountedFile       = false;
            uint64_t publishedFiles       = 0u;
            uint64_t knownSizeFiles       = 0u;
            bool directoriesKnown         = false;
            uint64_t completedDirectories = 0u;
            uint64_t completedObjects     = 0u;
        };
        std::vector<FileWorkScope> _fileWorkScopes; // guarded by _inFlightFilesMutex
        size_t _fileWorkSupportedScopes      = 0u;
        size_t _fileWorkBytesKnownScopes     = 0u;
        size_t _fileWorkDirectoryKnownScopes = 0u;
        uint64_t _fileWorkPublishedFiles     = 0u;
        uint64_t _fileWorkKnownSizeFiles     = 0u;
        uint64_t _fileWorkKnownSizeBytes     = 0u;
        std::atomic<uint64_t> _exactDiscoveredBytes{0u};
        bool _fileWorkCountsUnavailable = false; // one-way for the selected task
        std::atomic<bool> _copiedFileCountKnown{false};
        std::atomic<uint64_t> _copiedFileCount{0u};
        std::atomic<bool> _copiedFolderCountKnown{false};
        std::atomic<uint64_t> _copiedFolderCount{0u};
        std::atomic<bool> _completedWorkCountsKnown{false};
        std::atomic<uint64_t> _completedWorkFileCount{0u};
        std::atomic<uint64_t> _completedWorkFolderCount{0u};
        std::atomic<uint64_t> _completedWorkBytes{0u};
        std::atomic<bool> _discoveredBytesKnown{false};
        // The files whose completion was counted while discovery was open, by stream (guarded by
        // _inFlightFilesMutex). A file's later reports (a phase boundary after a commit over the 300 ms
        // in-flight expiry, its closing report) may add its entry again; they never count it again.
        struct CountedDiscoveryCompletion
        {
            const void* cookieKey     = nullptr;
            uint64_t progressStreamId = 0;
            uint64_t pathHash         = 0;
        };
        std::array<CountedDiscoveryCompletion, kMaxInFlightFiles * 2u> _countedDiscoveryCompletions{};
        size_t _countedDiscoveryCompletionCount = 0;
        size_t _countedDiscoveryCompletionNext  = 0;
        std::mutex _progressStreamPerfMutex;
        std::array<ProgressStreamPerf, kMaxInFlightFiles> _progressStreamPerf{};
        size_t _progressStreamPerfCount = 0;
        ProgressCadenceStats _progressCadence{}; // guarded by _progressStreamPerfMutex
        // I35 R10.2: the last report that advanced the transfer bytes, and what the streams that were
        // reporting had announced at the last report (the cause a stall is named by).
        std::atomic<ULONGLONG> _lastByteProgressTick{0};
        std::atomic<ProgressPhase> _progressStallPhase{ProgressPhase::Unknown};
        std::atomic<ProgressStallSide> _progressStallSide{ProgressStallSide::Storage};
        std::array<ConflictWorkerPerf, kMaxInFlightFiles> _conflictWorkerPerf{};
        size_t _conflictWorkerPerfCount = 0;

        PerfStats _perf{};
        std::atomic<uint64_t> _bridgeDirectoryEnsureCount{0};
        std::atomic<uint64_t> _bridgeSourceDirectoryEnumerationCount{0};
        std::atomic<uint64_t> _bridgeFileAdmissionCount{0};
        std::atomic<uint64_t> _bridgeActiveFileTransfers{0};
        std::atomic<uint64_t> _bridgeMaxActiveFileTransfers{0};
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

    // Failure leaves admissionOut untouched. This constructs an envelope only; StartOperation or
    // its worker still owns publication and every required decision.
    // I35 XP-04: whether the planned transfers can verify at all (a writer that proves its own digest
    // counts as a route).
    [[nodiscard]] static HostFileOperationVerificationAvailability PlannedVerificationAvailability(
        const FileOperations::FileOperationPlanGroup& plans) noexcept;
    static HRESULT BuildOperationAdmission(OperationConstruction construction,
                                           OperationAdmission& admissionOut,
                                           std::chrono::steady_clock::time_point startedAt = std::chrono::steady_clock::now());

    // Worker-side private preparation. Shares the canonical scope/ancestor binding algorithm and
    // leaves scopesOut unchanged on refusal. Never publishes a queue registration or task state.
    static HRESULT BuildTransferInterlockScopes(const FileOperations::FileOperationPlanGroup& plans,
                                                IFileSystem* sourceFileSystem,
                                                IFileSystem* destinationFileSystem,
                                                FileSystemOperation operation,
                                                std::vector<FileOperations::MutationInterlockScope>& scopesOut) noexcept;

    // Authority-free facts from an explicit private preparation; publication/observer delivery is
    // a separate caller action. A failure leaves snapshotOut unchanged.
    static HRESULT BuildPreparationFacts(uint64_t taskId,
                                         FileSystemOperation operation,
                                         const FileOperations::FileOperationPlanGroup& plans,
                                         std::span<const FileOperations::MutationInterlockScope> scopes,
                                         uint64_t selectedRootReadinessUs,
                                         std::shared_ptr<const FileOperations::PreparationSnapshot>& snapshotOut) noexcept;

    struct TransferDestinationNameFailure
    {
        std::wstring rejectedLeaf;
        std::wstring sourcePath;
        std::wstring destinationFolder;
    };
    // Provider calls run on the worker under its cancel watch. Rejection facts are private: the
    // caller logs a task error before execution when a destination name is rejected.
    static HRESULT ValidateTransferDestinationNames(const FileOperations::FileOperationPlanGroup& plans,
                                                    IFileSystem* destinationFileSystem,
                                                    FileSystemOperation operation,
                                                    std::stop_token stopToken,
                                                    const std::atomic<bool>& cancelled,
                                                    TransferDestinationNameFailure& failureOut) noexcept;

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
                             std::function<void(uint64_t completedItems, uint64_t totalItems)> progressCallback                  = {},
                             std::function<void(uint64_t taskId)> publishedCallback                                              = {},
                             uint64_t* taskIdOut                                                                                 = nullptr,
                             std::function<HRESULT()> preConsumptionDecisionGate                                                 = {},
                             std::function<void(std::shared_ptr<const FileOperations::PreparationSnapshot>)> preparationObserver = {});

    HRESULT AdmitScheduledRename(FolderWindow::Pane sourcePane,
                                 const wil::com_ptr<IFileSystem>& fileSystem,
                                 FileOperations::RenameOrigin origin,
                                 std::vector<BatchRenameExecutionOp> operations,
                                 std::function<void(uint64_t completedItems, uint64_t totalItems)> progressCallback                  = {},
                                 std::function<void(uint64_t taskId)> publishedCallback                                              = {},
                                 uint64_t* taskIdOut                                                                                 = nullptr,
                                 std::function<HRESULT()> preConsumptionDecisionGate                                                 = {},
                                 std::function<void(std::shared_ptr<const FileOperations::PreparationSnapshot>)> preparationObserver = {});

    void OnBatchRenameArtifactPrompt(std::unique_ptr<BatchRenameArtifactPromptPayload> payload) noexcept;
    void CompleteBatchRenameArtifactPromptByTaskId(uint64_t taskId,
                                                   uint64_t requestId,
                                                   HRESULT status,
                                                   std::optional<FileOperationArtifacts::TouchGuardReceipt> receipt = std::nullopt) noexcept;
    // The explicit zero-lParam fallback has no request identity and may only wake a failure.
    void CompleteBatchRenameArtifactPromptByTaskId(uint64_t taskId, HRESULT status) noexcept;
    void OnClipboardMoveReady(std::unique_ptr<ClipboardMoveReadyPayload> payload) noexcept;
    void CompleteClipboardMoveAdmissionByTaskId(uint64_t taskId, HRESULT status) noexcept;

    void ShowPopup() noexcept;
    void OnPopupHiddenByUser() noexcept;

    // I35 R8: the decision mirror, a modeless decision overlay over the pane that shows a pending prompt's
    // folder, and the main-window pending-decision indicator (FolderWindow.FileOperations.DecisionMirror.cpp).
    // Owner UI thread only. Reconcile runs on every presentation refresh before the popup refresh, so a
    // mirror that takes the focus claims the prompt's announcement first; the indicator update runs after
    // it and returns when the indicator must be looked at again (a run of decisions settling), if ever.
    void ReconcileDecisionMirror() noexcept;
    [[nodiscard]] std::optional<UINT> UpdateDecisionPendingIndicator() noexcept;
    // A pane's location or visibility changed: re-evaluate the mirror now (posts a presentation refresh).
    void NotifyDecisionMirrorPaneChanged() noexcept;
    // R8.12: one announcement per published prompt. The mirror announces when it takes the focus for it; the
    // card announces otherwise, but only after a reconcile decided (Defer: ask again at the next sync).
    enum class PromptAnnouncementClaim : uint8_t
    {
        Announce,         // the card announces it now
        AlreadyAnnounced, // the mirror (or the card before) announced it
        Defer,            // the mirror has not decided yet
    };
    [[nodiscard]] PromptAnnouncementClaim ClaimPromptAnnouncementForCard(uint64_t taskId, uint64_t generation) noexcept;
    // R8.11: fileOperations.mirrorDecisionsInPane (default on).
    [[nodiscard]] bool GetMirrorDecisionsInPane() const noexcept;
    // R8.10: the window a pre-task prompt anchors on: the originating pane's view while it is shown,
    // otherwise the Folder window.
    [[nodiscard]] HWND IngressPromptAnchor(FolderWindow::Pane pane) const noexcept;
#ifdef ENABLE_TESTS
    struct DecisionMirrorDebugSnapshot
    {
        bool visible               = false;
        bool windowVisible         = false;
        bool tookFocus             = false;
        bool announced             = false;
        uint64_t taskId            = 0u;
        uint64_t generation        = 0u;
        FolderWindow::Pane pane    = FolderWindow::Pane::Left;
        HWND overlay               = nullptr;
        HWND anchor                = nullptr;
        unsigned long pendingCount = 0u;
        uint64_t showCount         = 0u;
        uint64_t submitCount       = 0u;
        uint64_t acceptedCount     = 0u;
        uint64_t flashCount        = 0u;
        uint64_t reconcileCount    = 0u;
        bool indicatorLatched      = false;
        HWND foreground            = nullptr;
        // Shows that failed, F6 focus returns, card announcements deferred to the mirror's reconcile, the
        // prompts the last reconcile evaluated, and whether the shown prompt's focus was declined (F6).
        uint64_t showFailureCount = 0u;
        uint64_t focusReturnCount = 0u;
        uint64_t cardDeferCount   = 0u;
        size_t evaluatedCount     = 0u;
        bool focusDeclined        = false;
    };
    [[nodiscard]] bool DebugGetDecisionMirrorSnapshot(DecisionMirrorDebugSnapshot& out) noexcept;
    // Stands in the foreground window for the focus and flash rules (their owner-chain test still runs);
    // when the mirror takes the focus the stand-in moves to the overlay, and back to the owner when it hides
    // or F6 sends the focus back to the pane.
    void DebugSetDecisionMirrorForegroundOverride(std::optional<HWND> foreground) noexcept;
    // The next `count` mirror shows fail as a failed overlay show would.
    void DebugFailNextDecisionMirrorShows(uint32_t count) noexcept;
    // Holds the mirror's reconcile back (the presentation pass and the popup still run) until released.
    void DebugSetDecisionMirrorReconcilePaused(bool paused) noexcept;
    // Who announced the prompt: true the mirror, false the card, nullopt nobody yet.
    [[nodiscard]] std::optional<bool> DebugGetPromptAnnouncer(uint64_t taskId, uint64_t generation) noexcept;
    // The button ids the mirror gives a published primary action, the More... button and Escape.
    [[nodiscard]] static uint32_t DebugDecisionMirrorPrimaryButtonId(Task::ConflictAction action) noexcept;
    [[nodiscard]] static uint32_t DebugDecisionMirrorMoreButtonId() noexcept;
    [[nodiscard]] static uint32_t DebugDecisionMirrorEscapeButtonId() noexcept;
    [[nodiscard]] static uint32_t DebugDecisionMirrorApplyToAllOptionId() noexcept;
#endif

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
    // anchorPane: the pane the request came from, when the caller knows it (the prompt covers that pane).
    HRESULT ConfirmExternalArtifactTouch(const wil::com_ptr<IFileSystem>& fileSystem,
                                         std::wstring_view pluginId,
                                         std::wstring_view instanceContext,
                                         std::span<const std::filesystem::path> providerPaths,
                                         FileOperationArtifacts::TouchGuardReceipt* receiptOut = nullptr,
                                         std::optional<FolderWindow::Pane> anchorPane          = std::nullopt) noexcept;

    void ApplyTheme(const AppTheme& theme);
    void Shutdown() noexcept;
    void NotifyQueueChanged();
    bool HasActiveOperations() noexcept;
    // Owned terminal workers still require telemetry until their UI completion is
    // consumed. Retained completed summaries alone never require a timer. Owner thread only.
    [[nodiscard]] PopupTelemetryWork GetPopupTelemetryWork() noexcept;
    bool HasCancelableOperations() noexcept;
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
    // Requests cancellation of every unfinished task not already canceled and returns their IDs
    // (I35 R9.5: the popup's one-time completed-group collapse follows exactly this cohort).
    std::vector<uint64_t> CancelAll() noexcept;
    void CollectTasks(std::vector<Task*>& outTasks) noexcept;
    void CollectInformationalTasks(std::vector<FolderWindow::InformationalTaskUpdate>& outTasks) noexcept;
    void CollectCompletedTasks(std::vector<CompletedTaskSummary>& outTasks) noexcept;
    // The retained summaries as the popup presents them, built once per issues revision and
    // shared: a popup tick copies a reference, not the per-root paths or issue entries (x-perf#3).
    [[nodiscard]] std::shared_ptr<const std::vector<CompletedTaskPresentation>> CollectCompletedTaskPresentations() noexcept;
    void CollectDiagnostics(std::vector<TaskDiagnosticEntry>& outEntries) noexcept;
    // outCompletedIssues receives only the retained summaries' issue entries, in summary order (model#3).
    bool CollectIssuesSnapshotIfChanged(uint64_t lastRevision,
                                        std::vector<TaskDiagnosticEntry>& outDiagnostics,
                                        std::vector<TaskDiagnosticEntry>& outCompletedIssues,
                                        uint64_t& outCapturedRevision) noexcept;
    void CollectProjectableIssueTaskIds(std::vector<uint64_t>& outTaskIds) noexcept;
    void CollectTaskDiagnosticSnapshot(uint64_t taskId,
                                       unsigned long& warningCount,
                                       unsigned long& errorCount,
                                       std::wstring& lastDiagnosticMessage,
                                       unsigned long* noteCount          = nullptr,
                                       std::wstring* lastInfoNoteMessage = nullptr) noexcept;
    // I35 R5.3: display rule for Information notes (cards, Last note, Issues, direct Issues action).
    [[nodiscard]] bool GetShowInformationalNotes() const noexcept;
    // The running task's Warning/Error counters so far (no text formatting).
    void GetTaskProblemCounts(uint64_t taskId, unsigned long& warningCount, unsigned long& errorCount) noexcept;
    [[nodiscard]] std::optional<UINT> RefreshDeferredTaskPresentation() noexcept;
    void RevealPendingTasksImmediately() noexcept;
    void DismissCompletedTask(uint64_t sessionTaskId) noexcept;
    // Bulk "Clear completed": removes only the listed summaries and never one backed by an
    // interrupted-Move breadcrumb, which only its own per-card Dismiss acknowledges.
    void ClearCompletedTasks(std::span<const uint64_t> sessionTaskIds) noexcept;
    uint64_t CreateOrUpdateInformationalTask(const FolderWindow::InformationalTaskUpdate& update) noexcept;
    void DismissInformationalTask(uint64_t taskId) noexcept;
    bool GetAutoDismissSuccess() const noexcept;
    void SetAutoDismissSuccess(bool enabled) noexcept;
    bool GetPopupFooterOnly() const noexcept;
    void SetPopupFooterOnly(bool footerOnly) noexcept;
    bool GetPopupCompactDensity() const noexcept;
    void SetPopupCompactDensity(bool compactDensity) noexcept;
    bool GetPopupThroughputExpanded() const noexcept;
    void SetPopupThroughputExpanded(bool expanded) noexcept;
    bool OpenDiagnosticsLogForTask(uint64_t taskId) noexcept;
    bool ExportTaskIssuesReport(uint64_t taskId, std::filesystem::path* reportPathOut = nullptr, bool openAfterExport = true) noexcept;
    bool OpenCompletedTaskSource(uint64_t taskId) noexcept;
    bool SelectCompletedTaskRetainedSources(uint64_t taskId) noexcept;
    bool CutCompletedTaskRetainedSourcesAgain(uint64_t taskId) noexcept;
    void ToggleIssuesPane() noexcept;
    [[nodiscard]] bool OpenIssuesForTask(uint64_t taskId) noexcept;
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
    void DebugResetPopupForSelfTest() noexcept;
    void DebugEnsurePopupVisibleForSelfTest() noexcept;
    HWND GetIssuesPaneHwndForSelfTest() noexcept;
    void DebugResetIssuesPaneForSelfTest(bool clearViewState = true) noexcept;
    void DebugClearDiagnosticsForSelfTest() noexcept;
    // The real completed More callback uses openAfterExport=true. Keep its
    // report under the case root and suppress external application launch.
    struct IssuesExportRouteDebugSnapshot
    {
        uint32_t attempts           = 0u;
        uint32_t suppressedLaunches = 0u;
        uint32_t refusedCreates     = 0u;
    };
    void DebugConfigureIssuesExportRouteForSelfTest(std::filesystem::path reportDirectory, bool refuseCreate) noexcept;
    [[nodiscard]] IssuesExportRouteDebugSnapshot DebugGetIssuesExportRouteForSelfTest() noexcept;
    // Route the real task log into a case-owned directory and observe its
    // completed More callback without opening an external application.
    struct DiagnosticsLogRouteDebugSnapshot
    {
        uint32_t attempts           = 0u;
        uint32_t suppressedLaunches = 0u;
        uint64_t openedTaskId       = 0u;
        std::filesystem::path openedPath;
    };
    void DebugConfigureDiagnosticsLogRouteForSelfTest(std::filesystem::path logDirectory) noexcept;
    [[nodiscard]] DiagnosticsLogRouteDebugSnapshot DebugGetDiagnosticsLogRouteForSelfTest() noexcept;
    void DebugFlushDiagnosticsForSelfTest() noexcept;
    void DebugRemoveDiagnosticsForTask(uint64_t taskId) noexcept;
    void DebugAppendCompletedTaskForSelfTest(CompletedTaskSummary summary) noexcept;
    // Records a finalized direct-executor fixture task through the production completed-summary
    // path; the caller removes the recorded summary with DebugRemoveDiagnosticsForTask.
    [[nodiscard]] CompletedTaskSummary DebugRecordCompletedTaskForSelfTest(Task& task) noexcept;
    bool DebugValidateRetainedSourceActionForSelfTest(uint64_t taskId, IFileSystem* fileSystem) noexcept;
    bool DebugFlushPendingSettingsSaveForSelfTest(DWORD timeoutMs) noexcept;
    SettingsSaveDebugSnapshot DebugGetSettingsSaveSnapshotForSelfTest() noexcept;
    void DebugForceNextFileOperationCompletedPostFailure() noexcept;
    void DebugForceAllFileOperationCompletedPostsFailure(bool enabled) noexcept;
    void DebugSuppressTaskPresentationPostsForSelfTest(bool enabled) noexcept;
    // Fails the recovery-heartbeat arm that follows passingArms successful ones, once;
    // a negative value disarms the fault.
    void DebugFailRecoveryTimerArmForSelfTest(int32_t passingArms) noexcept;
    [[nodiscard]] size_t DebugFallbackCompletionCountForSelfTest() noexcept;
    [[nodiscard]] uint64_t DebugCompletionRecoveryTickCountForSelfTest() const noexcept;
    // The owner heartbeat interval RefreshDeferredTaskPresentation last chose (0 = stopped).
    [[nodiscard]] UINT DebugLastPresentationHeartbeatMsForSelfTest() const noexcept;
    // How often the shared completed presentation and the projectable-issue task set were rebuilt.
    [[nodiscard]] uint64_t DebugCompletedPresentationBuildCountForSelfTest() const noexcept;
    [[nodiscard]] uint64_t DebugProjectableIssueScanCountForSelfTest() const noexcept;
    void DebugForceNextOrphanedCompletionDrainSubmissionFailure() noexcept;
    void DebugForceAllOrphanedCompletionDrainSubmissionFailures(bool enabled) noexcept;
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
                              std::wstring_view destinationPath,
                              LocalizedDiagnosticText userText = {}) noexcept;
    void EnqueueTaskDiagnostic(TaskDiagnosticEntry entry) noexcept;
    // Q10: replaces the one retained record of a cached apply-to-all aggregate (task + originating
    // generation + first application) in both Issues reservoirs, or adds it. Only the final
    // record, with its complete count, is written to the JSONL log.
    void PublishCachedConflictAggregate(TaskDiagnosticEntry entry, bool final) noexcept;

    OperationEntryResult EnterOperation(Task& task, std::stop_token stopToken) noexcept;
    void LeaveOperation(Task& task, const std::shared_ptr<const SelectedAdmissionView>& accepted) noexcept;
    void PostCompleted(Task& task) noexcept;

    Task* FindTask(uint64_t taskId) noexcept;
    // UI-thread freshness check. A navigation completion must also revalidate its
    // pane generation, live endpoint, item selection and foreground ownership.
    [[nodiscard]] bool IsPendingConflictInspectionCurrent(uint64_t taskId,
                                                          uint64_t promptGeneration,
                                                          const std::shared_ptr<const FileOperations::ConflictInspectionContext>& context) noexcept;
    // UI-thread only; clears out on every failure and rechecks prompt freshness after
    // provider queries. Does not navigate, focus, enumerate or mutate the provider.
    [[nodiscard]] HRESULT ResolvePendingConflictInspectionTarget(uint64_t taskId,
                                                                 uint64_t promptGeneration,
                                                                 std::shared_ptr<const FileOperations::ConflictInspectionContext> context,
                                                                 bool destination,
                                                                 FileOperations::ConflictInspectionTarget& out) noexcept;
    [[nodiscard]] HRESULT BeginPendingConflictInspection(uint64_t taskId,
                                                         uint64_t promptGeneration,
                                                         std::shared_ptr<const FileOperations::ConflictInspectionContext> context,
                                                         bool destination,
                                                         FileOperations::ConflictInspectionCompletion completion) noexcept;
    void CancelPendingConflictInspection() noexcept;
    // UI thread: the task's prompt was answered, the task canceled or removed. Ends its pending
    // inspection as Stale and drops its unconsumed focus ticket; other tasks' requests are kept.
    void CancelConflictInspectionForTask(uint64_t taskId) noexcept;
    // UI thread: the pane's enumeration was canceled with no settlement to wait for (path cleared,
    // network drive disconnect). Ends that pane's pending inspection as Stale.
    void CancelConflictInspectionForPane(FolderWindow::Pane pane) noexcept;
    [[nodiscard]] bool CommitPendingConflictInspectionFocus(uint64_t ticket, HWND popup) noexcept;
    [[nodiscard]] FolderView::EnumerationSettledCallback GetConflictInspectionEnumerationCallback(FolderWindow::Pane pane);
    void RemoveTask(uint64_t taskId) noexcept;
    // R4-A02-1: compare a prepared task's scopes with every other published prepared or active task in this host using
    // the interlock predicate (no enumeration, no provider call). taskCount counts the overlapping
    // tasks; problem and firstTaskId describe the first conflicting pair. Older admission keys
    // are predecessors, preventing two simultaneously prepared tasks from queueing after each other.
    // admissionKeys holds one entry per relation; taskIds names each overlapping task for display.
    struct SameHostOverlapAdvice final
    {
        FileOperations::SameHostOverlapProblem problem = FileOperations::SameHostOverlapProblem::None;
        uint64_t firstTaskId                           = 0u;
        uint32_t taskCount                             = 0u;
        std::array<uint64_t, FileOperations::kMaxSameHostOverlapRelations> taskIds{};
        size_t taskIdCount = 0u;
        std::array<uint64_t, FileOperations::kMaxSameHostOverlapRelations> admissionKeys{};
        size_t relationCount  = 0u;
        bool relationOverflow = false;
        // The latest queue place among the overlapping tasks (queue order keys, not admission keys:
        // a reordered or late-queued peer holds a queue place unrelated to its admission key).
        uint64_t latestQueueOrderKey = 0u;
        // I35 R7.3: an overlapping pair involves a delete by name. That delete deletes exactly the tree
        // it pinned, so the two tasks never run at the same time.
        bool byNameDelete = false;

        [[nodiscard]] bool CanRunConcurrently() const noexcept
        {
            return ! byNameDelete && ! relationOverflow && relationCount > 0u;
        }
    };
    struct ActiveMutationInterlock final
    {
        uint64_t taskId = 0u;
        Task* task      = nullptr; // lifetime/status only; admission authority never comes from Task
        std::shared_ptr<const MutationInterlockAdmission> admission;
    };
    [[nodiscard]] static bool HasPublishedOverlapPredecessor(const MutationInterlockAdmission& candidate,
                                                             std::span<const ActiveMutationInterlock> prepared,
                                                             std::span<const ActiveMutationInterlock> active) noexcept;
    [[nodiscard]] static bool MutationInterlocksPermit(uint64_t taskId,
                                                       const MutationInterlockAdmission& candidate,
                                                       std::span<const ActiveMutationInterlock> active,
                                                       uint64_t& comparisons) noexcept;
    [[nodiscard]] static bool HasConcurrentPublisherConsent(const MutationInterlockAdmission& candidate, const ActiveMutationInterlock& publisher) noexcept;
    void PublishPreparedMutationInterlock(Task& task) noexcept;
    void PublishMutationInterlockDecision(Task& task, MutationInterlockRelations relations) noexcept;
    static void EraseMutationInterlockAdmission(std::vector<ActiveMutationInterlock>& registrations, uint64_t admissionKey) noexcept;
    void WithdrawPreparedMutationInterlock(uint64_t admissionKey) noexcept;
    [[nodiscard]] SameHostOverlapAdvice FindSameHostOverlap(const Task& task) noexcept;
    [[nodiscard]] SameHostOverlapAdvice FindSameHostOverlap(uint64_t taskId, const MutationInterlockAdmission& candidate) noexcept;
    // I35 R7.3: the first other live task whose prepared or active scope names this path, one of its
    // ancestors or a descendant at the same provider root, or 0. A delete by name refuses that path;
    // another root of the same identity domain cannot be compared by path and stays serialized by the
    // conservative interlock (the deleting handle's corroboration then sees any replaced object). A
    // prepared task that decided to queue after `admissionKey` (this task's admission) waits for it and
    // does not count.
    [[nodiscard]] uint64_t FindLiveTaskHoldingPath(uint64_t taskId,
                                                   const FileOperations::QualifiedEndpoint& endpoint,
                                                   std::wstring_view providerPath,
                                                   uint64_t admissionKey = 0u) noexcept;
    // I35 R7.5: the first live task (prepared or running) one of whose scopes on this plugin names
    // `folderPath`, one of its ancestors or anything below it, compared with that scope's own path rules;
    // 0 when none. A running Copy's temporary file is a sibling of its destination, so the leftover
    // listing asks about the folder, not the item. No provider call.
    [[nodiscard]] uint64_t FindLiveTaskWorkingIn(std::wstring_view pluginId, std::wstring_view folderPath) noexcept;
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
    void RecoverFallbackCompletions() noexcept;
    // A completion is waiting in the fallback queue; the owner heartbeat stays fast (x-lifetime#7).
    [[nodiscard]] bool HasPendingFallbackCompletions() noexcept;
    static constexpr UINT kCompletionRecoveryHeartbeatMs     = 100u;
    static constexpr UINT kCompletionRecoveryIdleHeartbeatMs = 1000u;
    [[nodiscard]] HRESULT ArmCompletionRecoveryTimer() noexcept;

private:
    struct PendingConflictInspection
    {
        uint64_t generation     = 0u;
        uint64_t taskId         = 0u;
        FolderWindow::Pane pane = FolderWindow::Pane::Left;
        FileOperations::ConflictInspectionTarget target;
        FileOperations::ConflictInspectionCompletion completion;
        std::optional<FolderView::InspectionVisit> visit;
        wil::com_ptr<IFileSystem> selectedFileSystem;
        std::wstring selectedPluginId;
        std::wstring selectedInstanceContext;
        std::filesystem::path previousQualifiedPath;
        std::filesystem::path previousProviderPath;
        std::wstring previousPluginId;
        std::wstring previousInstanceContext;
        HWND foreground                                 = nullptr;
        HWND focus                                      = nullptr;
        DWORD inputTick                                 = 0u;
        bool inputKnown                                 = false;
        bool rollingBack                                = false;
        HRESULT navigationError                         = S_OK;
        std::chrono::steady_clock::time_point startedAt = std::chrono::steady_clock::now();
    };
    std::unique_ptr<PendingConflictInspection> _pendingConflictInspection;
    std::unique_ptr<PendingConflictInspection> _completedConflictInspection;
    uint64_t _inspectionRequestGeneration = 0u;
    void TryCompletePendingConflictInspection(std::optional<HRESULT> settlement = std::nullopt) noexcept;
    void FinishPendingConflictInspection(FileOperations::ConflictInspectionNavigationResult result) noexcept;

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
    std::filesystem::path GetDiagnosticsLogDirectory() const noexcept;
    std::filesystem::path GetDiagnosticsLogPathForDate(const SYSTEMTIME& localTime) const noexcept;
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
    // Guarded by _mutex; rebuilt when _issuesRevision moves (every _completedTasks change bumps it).
    std::shared_ptr<const std::vector<CompletedTaskPresentation>> _completedPresentation;
    uint64_t _completedPresentationRevision                    = 0u;
    uint64_t _completedPresentationLanguageRevision            = 0u;
    static constexpr size_t kMaxAcceptedMoveClipboardSequences = 64u;
    std::deque<uint32_t> _acceptedMoveClipboardSequences;
    uint64_t _nextTaskId = 1;
    std::atomic<bool> _interruptedMoveSummariesPendingPresentation{false};
    std::atomic<bool> _actionablePromptPresentationPending{false};

    // I35 R8 decision mirror state (FolderWindow.FileOperations.DecisionMirror.cpp); owner UI thread only.
    struct DecisionMirror;
    struct DecisionMirrorDeleter
    {
        void operator()(DecisionMirror* mirror) const noexcept;
    };
    std::unique_ptr<DecisionMirror, DecisionMirrorDeleter> _decisionMirror;
    [[nodiscard]] DecisionMirror& EnsureDecisionMirror() noexcept;
    // First step of Shutdown: hides the mirror without handing back focus and drops every callback.
    void ShutdownDecisionMirror() noexcept;
    // The user opened Inspect from the card for this prompt: the mirror leaves the pane to the inspection.
    void NoteDecisionMirrorInspection(uint64_t taskId, uint64_t generation) noexcept;

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
    // A task publishes its immutable scope vector before asking for overlap consent. Keeping this
    // separate from the active set closes the preparation/preparation race without making a
    // not-yet-admitted task block execution. Activation transfers the entry atomically under
    // _queueMutex.
    std::vector<ActiveMutationInterlock> _preparedMutationInterlocks;
    std::vector<ActiveMutationInterlock> _activeMutationInterlocks;
    struct LivePublishedScope final
    {
        uint64_t taskId                                            = 0u;
        uint64_t admissionKey                                      = 0u;
        const FileOperations::MutationInterlockScope* plannedScope = nullptr;
        std::wstring providerPath;
    };
    static constexpr size_t kMaxLivePublishedScopes = 256u;
    std::array<LivePublishedScope, kMaxLivePublishedScopes> _livePublishedScopes{};
    size_t _livePublishedScopeCount  = 0u;
    bool _livePublishedIndexOverflow = false;

#ifdef ENABLE_TESTS
    std::filesystem::path _debugIssuesExportReportDirectory;
    bool _debugIssuesExportRefuseCreate = false;
    IssuesExportRouteDebugSnapshot _debugIssuesExportRoute{};
    mutable std::mutex _debugDiagnosticsLogRouteMutex;
    std::filesystem::path _debugDiagnosticsLogDirectory;
    DiagnosticsLogRouteDebugSnapshot _debugDiagnosticsLogRoute{};
#endif
    std::mutex _diagnosticsMutex;
    std::atomic<uint64_t> _issuesRevision{1u};
    std::mutex _projectableIssueTaskIdsMutex; // leaf lock, never held with _mutex or _diagnosticsMutex
    std::vector<uint64_t> _projectableIssueTaskIds;
    uint64_t _projectableIssueTaskIdsRevision = 0u;
    bool _projectableIssueTaskIdsShowNotes    = false; // the visibility rule the cached ids were built with
    std::deque<TaskDiagnosticEntry> _diagnosticsInMemory;
    // model#2: _diagnosticsInMemory[0, prefix) are known original Warning/Error problems, so the
    // eviction search for the oldest observation resumes there (amortized O(1) per push).
    size_t _diagnosticsInMemoryOriginalPrefix = 0u;
    std::vector<TaskDiagnosticEntry> _diagnosticsPendingFlush;
    std::unordered_map<uint64_t, std::pair<unsigned long, unsigned long>> _taskDiagnosticCounts;
    // The last Warning/Error of each running task; its Last note text is formatted only when read.
    std::unordered_map<uint64_t, TaskDiagnosticEntry> _taskLastProblem;
    // I35 R5.3: Information notes of each running task (count and the last one), kept apart from the
    // Warning/Error counters so they never become problems.
    std::unordered_map<uint64_t, unsigned long> _taskNoteCounts;
    std::unordered_map<uint64_t, TaskDiagnosticEntry> _taskLastNote;
    // model#2: one running task's bounded Issues reservoir; entries[0, originalPrefix) are known
    // original Warning/Error problems, as for _diagnosticsInMemory.
    struct TaskIssueReservoir
    {
        std::deque<TaskDiagnosticEntry> entries;
        size_t originalPrefix = 0u;
    };
    std::unordered_map<uint64_t, TaskIssueReservoir> _taskIssueDiagnostics;
    ULONGLONG _lastDiagnosticsFlushTick   = 0;
    ULONGLONG _lastDiagnosticsCleanupTick = 0;

    std::mutex _fallbackCompletedMutex;
    struct FallbackCompletion final
    {
        TaskCompletedPayload payload{};
        // A transferred handle is not a quiet point. Only the reaper's completed join
        // (or shutdown's completed drain) permits destruction of the matching Task.
        bool workerJoined = false;
    };
    std::vector<FallbackCompletion> _fallbackCompletedPayloads;
    std::mutex _completionWorkerTransferMutex;
    std::mutex _orphanedCompletionDrainMutex;
    std::condition_variable _orphanedCompletionDrainCv;
    uint32_t _orphanedCompletionDrainsOutstanding = 0u;
    std::atomic_bool _completionShutdown{false};
    // Sticky admission history: preserve the existing conservative flush for this
    // state's saves, without waiting for unrelated global work in an unused state.
    std::atomic_bool _hasQueuedSettingsSave{false};
#ifdef ENABLE_TESTS
    std::atomic_bool _debugForceNextCompletedPostFailure{false};
    std::atomic_bool _debugForceAllCompletedPostsFailure{false};
    std::atomic_bool _debugSuppressTaskPresentationPosts{false};
    std::atomic<int32_t> _debugRecoveryTimerArmPassesBeforeFailure{-1};
    std::atomic<uint64_t> _debugCompletionRecoveryTicks{0u};
    std::atomic<UINT> _debugLastPresentationHeartbeatMs{0u};
    std::atomic<uint64_t> _debugCompletedPresentationBuilds{0u};
    std::atomic<uint64_t> _debugProjectableIssueScans{0u};
    std::atomic_bool _debugForceNextOrphanedCompletionDrainSubmissionFailure{false};
    std::atomic_bool _debugForceAllOrphanedCompletionDrainSubmissionFailures{false};
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
                                              FileSystemFlags flags      = FILESYSTEM_FLAG_NONE) noexcept;
[[nodiscard]] bool IsAutoDismissableFileOperationCompletion(HRESULT resultHr,
                                                            unsigned long warningCount,
                                                            unsigned long errorCount,
                                                            bool partialOutcome = false) noexcept;
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
void SetFileOpsBridgeFailNextSourceGetSizeForSelfTest(unsigned long count, HRESULT status) noexcept;
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
bool SetFileOpsBridgeCommittedErrorPublicationForSelfTest(std::wstring_view sourceProviderPath, const std::filesystem::path& destinationPath) noexcept;
bool TakeAndClearFileOpsBridgeCommittedErrorPublicationForSelfTest() noexcept;
void SetFileOpsAutoConcurrencyOverrideForSelfTest(bool enabled, unsigned int preferredConcurrency, uint32_t storageKind) noexcept;
// I35 R10.3: replaces the adopted FileSystemTransferTuning defaults of every task started from now on
// (std::nullopt restores them), so one build can run each transfer candidate for A/B evidence.
void SetFileOpsTransferTuningOverrideForSelfTest(std::optional<uint32_t> tuning) noexcept;
// I35 R10.5: cancels the task whose progress report is the `occurrence`th one announcing `phase`
// (std::nullopt disarms); Take returns how many cancellations fired since arming.
void SetFileOpsCancelAtProgressPhaseForSelfTest(std::optional<FolderWindow::FileOperationState::ProgressPhase> phase, unsigned int occurrence) noexcept;
[[nodiscard]] unsigned int TakeFileOpsCancelAtProgressPhaseFiredForSelfTest() noexcept;
// I35 R10.5: every storage path under `prefix` reports this preferred concurrency and storage kind to
// the automatic concurrency budget, so a test can tell the two sides of one route apart.
void AddFileOpsStoragePathConcurrencyOverrideForSelfTest(std::wstring_view prefix, unsigned int preferredConcurrency, uint32_t storageKind);
void ClearFileOpsStoragePathConcurrencyOverridesForSelfTest() noexcept;
// I35 S6: the host's progress accounting driven with synthetic reports on a task that never runs.
struct FileOpsProgressAccountingDebugResult final
{
    // The phase a report without item bytes (0 of 0) announces.
    FolderWindow::FileOperationState::ProgressPhase unsizedReportPhase = FolderWindow::FileOperationState::ProgressPhase::Unknown;
    // A 700 ms silence after that report, by the phase it is attributed to.
    uint32_t unknownSilences = 0;
    uint32_t unknownStalls   = 0;
    // A zero-delta Finalizing report after the item's last bytes.
    uint64_t finalizingRepeats = 0;
    // One file completed while discovery was open, reported again after its in-flight entry expired
    // (a commit longer than the 300 ms expiry), then a second file completed: completions counted.
    uint64_t completionsCountedWhileOpen = 0;
};
[[nodiscard]] FileOpsProgressAccountingDebugResult DebugRunFileOpsProgressAccountingForSelfTest(FolderWindow::FileOperationState& state) noexcept;
void SetFileOpsPostFinishedCompletionPauseForSelfTest(bool enabled) noexcept;
bool HasFileOpsPostFinishedCompletionPauseEnteredForSelfTest() noexcept;
void ReleaseFileOpsPostFinishedCompletionPauseForSelfTest() noexcept;
void SetFileOpsFallbackWorkerReturnPauseForSelfTest(bool enabled) noexcept;
bool HasFileOpsFallbackWorkerReturnPauseEnteredForSelfTest() noexcept;
void ReleaseFileOpsFallbackWorkerReturnPauseForSelfTest() noexcept;
void SetFileOpsBridgeMoveSourceCleanupPauseForSelfTest(bool enabled, ULONGLONG bailoutMs = 5'000ull) noexcept;
unsigned int GetFileOpsBridgeMoveSourceCleanupPausedCountForSelfTest() noexcept;
bool HasFileOpsBridgeMoveSourceCleanupPauseEnteredForSelfTest() noexcept;
void ReleaseFileOpsBridgeMoveSourceCleanupPauseForSelfTest() noexcept;
// Holds the bridge just before it restores the basic information of the created destination folder at
// `destinationPath` (empty: never), once that folder's children have finished.
void SetFileOpsBridgeDirectoryRestorePauseForSelfTest(std::wstring_view destinationPath) noexcept;
bool HasFileOpsBridgeDirectoryRestorePauseEnteredForSelfTest() noexcept;
void ReleaseFileOpsBridgeDirectoryRestorePauseForSelfTest() noexcept;
void SetFileOpsSelectedLeafDiscoveryPublishedPauseForSelfTest(bool enabled) noexcept;
bool HasFileOpsSelectedLeafDiscoveryPublishedPauseEnteredForSelfTest() noexcept;
void ReleaseFileOpsSelectedLeafDiscoveryPublishedPauseForSelfTest() noexcept;
void SetFileOpsNativeMoveCreateDirectoryRaceForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsNativeMoveCreateDirectoryRaceAttemptsForSelfTest() noexcept;
void SetFileOpsManagedCleanupKnownNonCommitForSelfTest(HRESULT status, unsigned long count) noexcept;
unsigned long TakeFileOpsManagedCleanupKnownNonCommitAttemptsForSelfTest() noexcept;
void SetFileOpsPermanentDeleteKnownNonCommitForSelfTest(HRESULT status, unsigned long count) noexcept;
unsigned long TakeFileOpsPermanentDeleteKnownNonCommitAttemptsForSelfTest() noexcept;
// I35 S4: the last per-item receipt the host kept for (task, item), taken once; empty when none was kept.
std::optional<FileSystemItemMutationResult> TakeFileOpsItemReceiptForSelfTest(uint64_t taskId, size_t itemIndex) noexcept;
void SetFileOpsManagedCleanupUnknownOutcomeForSelfTest(unsigned long count) noexcept;
unsigned long TakeFileOpsManagedCleanupUnknownOutcomeAttemptsForSelfTest() noexcept;
// Provider DeleteItems calls the BulkItems Delete executor made since the last take.
uint64_t TakeFileOpsBulkDeleteProviderCallsForSelfTest() noexcept;
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
enum class FileOpsBatchRenameReceiptFaultForSelfTest : uint8_t
{
    None,
    UnknownOutcome,
    NoCommitNoOriginal,
    CommittedOriginalRetained,
    InvalidBoolean,
    ShortPrefix,
};
void SetFileOpsBatchRenameReceiptFaultForSelfTest(FileOpsBatchRenameReceiptFaultForSelfTest fault) noexcept;
enum class FileOpsRecycleDeleteReceiptFaultForSelfTest : uint8_t
{
    None,
    ShortPrefix,
    InvalidBoolean,
    CommittedOriginalRetained,
};
void SetFileOpsRecycleDeleteReceiptFaultForSelfTest(FileOpsRecycleDeleteReceiptFaultForSelfTest fault) noexcept;
void SetFileOpsBatchRenameBeforeMutationPauseForSelfTest(bool enabled) noexcept;
bool HasFileOpsBatchRenameBeforeMutationPauseEnteredForSelfTest() noexcept;
void ReleaseFileOpsBatchRenameBeforeMutationPauseForSelfTest() noexcept;
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
void DebugBuildFileOpsVisualConflictPolicyForSelfTest(FolderWindow::FileOperationState::Task::ConflictPromptState& prompt,
                                                      bool allowReplacement,
                                                      bool allowReadOnlyDelete      = false,
                                                      FileSystemOperation operation = FILESYSTEM_COPY) noexcept;
bool RunFileOpsPerItemSchedulerShutdownQuietPointSelfTestForSelfTest(FolderWindow::FileOperationState& state) noexcept;
bool RunFileOpsPerItemSchedulerNestedSaturationSelfTestForSelfTest(FolderWindow::FileOperationState& state) noexcept;
bool RunFileOpsPerItemSchedulerFailurePolicySelfTestForSelfTest(FolderWindow::FileOperationState& state) noexcept;
bool RunFileOpsWorkerStartGateCancellationSelfTestForSelfTest(FolderWindow::FileOperationState& state) noexcept;
bool RunFileOpsBridgePausedReaderStopSelfTestForSelfTest(FolderWindow::FileOperationState& state) noexcept;
bool RunFileOpsBridgeDirectoryBufferValidationSelfTestForSelfTest() noexcept;
bool RunFileOpsBridgeStageDebtSelfTestForSelfTest() noexcept;
void ResetFileOpsBridgeBufferBudgetPeakForSelfTest() noexcept;
uint64_t GetFileOpsBridgeBufferBudgetPeakForSelfTest() noexcept;
uint64_t GetFileOpsBridgeBufferBudgetInUseForSelfTest() noexcept;
bool IsFileOpsCircuitBreakerTransientErrorForSelfTest(DWORD error) noexcept;
#endif
