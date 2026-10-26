#pragma once

#include "FileOperationsCompletedGroupLayout.h"
#include "FolderWindow.h"
#include <DxUi/DxUi.h>

#include <array>
#include <filesystem>
#include <memory>
#include <optional>
#include <shobjidl.h>
#include <span>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <vector>

#ifdef ENABLE_TESTS
#include "SelfTest/FileOperations/FileOperations.Tests.FileOperationsSurfaceResources.h"
#endif

namespace FileOperations
{
struct ConflictInspectionContext;
}

namespace FileOperationsPopupInternal
{
struct PopupHitTest
{
    enum class Kind : uint8_t
    {
        None,
        FooterCancelAll,
        FooterPauseResumeAll,
        FooterAutoDismiss,
        FooterQueueMode,
        FooterDensity,
        FooterToggleDetails,
        CompletedGroupToggle,
        CompletedGroupClear,
        TaskToggleCollapse,
        TaskStartNow,
        TaskQueueMoveUp,
        TaskQueueMoveDown,
        TaskPause,
        TaskCancel,
        TaskSkip,
        TaskSpeedLimit,
        TaskShowLog,
        TaskExportIssues,
        TaskCompletedMore,
        TaskConflictToggleApplyToAll,
        TaskConflictAction,
        TaskConflictMore,
        TaskDismiss,
        FooterOptions,
        TaskToggleThroughput,
        TaskConflictInspectSource,
        TaskConflictInspectDestination,
        TaskIssues,
    };

    Kind kind                 = Kind::None;
    uint64_t taskId           = 0;
    uint32_t data             = 0;
    uint64_t promptGeneration = 0u;
};

#ifdef ENABLE_TESTS
inline constexpr uint32_t kPopupSelfTestDestroyOnNextShowData = 0xD1570001u;

struct PopupSelfTestInvoke
{
    PopupHitTest::Kind kind   = PopupHitTest::Kind::None;
    uint64_t taskId           = 0;
    uint32_t data             = 0;
    uint64_t promptGeneration = 0u;
};

#endif

struct PopupButton
{
    D2D1_RECT_F bounds{};
    PopupHitTest hit{};
};

// What Space does on a focused task-owned hosted control (popup-b#4).
enum class HostedTaskSpaceKeyRoute : uint8_t
{
    ControlActivation, // The focused control handles Space as its own activation.
    PauseOrResume,     // The task's pause shortcut.
    Ignore,            // A Ctrl/Alt chord: neither pauses nor activates.
};

[[nodiscard]] HostedTaskSpaceKeyRoute ResolveHostedTaskSpaceKeyRoute(PopupHitTest::Kind kind, bool controlEnabled, bool modified) noexcept;

struct PopupMenuAnchor
{
    POINT screenPoint{};
    DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
};

enum class PopupStatusVisualTone : uint8_t
{
    None,
    Neutral,
    Accent,
    Muted,
    Warning,
    Error,
    Ok,
};

struct TaskSnapshot
{
    static constexpr size_t kMaxInFlightFiles   = 16u;
    static constexpr size_t kMaxConflictActions = 8u;

    enum class StatusKind : uint8_t
    {
        None,
        Waiting,
        Discovering,
        Preparing,
        Running,
        Verifying,
        Paused,
        Conflict,
        Done,
        Partial,
        Indeterminate,
        Failed,
        Canceled,
        Stopping,
    };

    enum class Kind : uint8_t
    {
        FileOperation,
        Informational,
    };

    Kind kind                  = Kind::FileOperation;
    bool cancellationRequested = false;
    FolderWindow::InformationalTaskUpdate informational{};

    struct InFlightFileSnapshot
    {
        const void* cookieKey     = nullptr;
        uint64_t progressStreamId = 0;
        std::wstring sourcePath;
        uint64_t totalBytes           = 0;
        uint64_t completedBytes       = 0;
        ULONGLONG lastUpdateTick      = 0;
        uint64_t logicalLeafId        = 0u;
        FileSystemFileWorkPhase phase = FileSystemFileWorkPhase::Begin;
        bool explicitLifecycle        = false;
        bool sizeKnown                = false;
        bool terminal                 = false;
    };

    uint64_t taskId               = 0;
    FileSystemOperation operation = FILESYSTEM_COPY;
    FileOperations::TaskSelectionSummary selection;
    uint8_t lifecyclePhase = 0xFFu; // FileOperations::TaskLifecyclePhase; unknown for informational rows.

    unsigned long totalItems                       = 0;
    unsigned long completedItems                   = 0;
    uint64_t totalBytes                            = 0;
    uint64_t completedBytes                        = 0;
    uint64_t itemTotalBytes                        = 0;
    uint64_t itemCompletedBytes                    = 0;
    bool verificationRequested                     = false;
    bool verificationActive                        = false;
    uint64_t verificationTotalBytes                = 0;
    uint64_t verificationCompletedBytes            = 0;
    uint64_t verificationItemTotalBytes            = 0;
    uint64_t verificationItemCompletedBytes        = 0;
    unsigned long verifiedItemCount                = 0;
    unsigned long verificationProblemItemCount     = 0;
    unsigned long verificationFailedItemCount      = 0;
    unsigned long verificationUnavailableItemCount = 0;
    unsigned long verificationCanceledItemCount    = 0;
    uint8_t verificationState                      = 0u; // FileOperations::VerificationState

    // Best-effort top-level type breakdown while traversal totals are still open.
    unsigned long completedFiles      = 0;
    unsigned long completedFolders    = 0;
    uint64_t copiedFileCount          = 0u;
    bool copiedFileCountKnown         = false;
    uint64_t copiedFolderCount        = 0u;
    bool copiedFolderCountKnown       = false;
    uint64_t completedWorkFileCount   = 0u;
    uint64_t completedWorkFolderCount = 0u;
    uint64_t completedWorkBytes       = 0u;
    bool completedWorkCountsKnown     = false;
    bool discoveredBytesKnown         = false;
    uint64_t exactDiscoveredBytes     = 0u;

    std::wstring currentSourcePath;
    std::wstring currentDestinationPath;

    std::array<InFlightFileSnapshot, kMaxInFlightFiles> inFlightFiles{};
    size_t inFlightFileCount = 0;

    struct ConflictPromptSnapshot
    {
        uint64_t generation = 0u;
        std::shared_ptr<const FileOperations::ConflictInspectionContext> inspection;
        struct ItemMetadata
        {
            bool available           = false;
            bool isDirectory         = false;
            bool isLink              = false;
            bool sizeKnown           = false;
            uint64_t sizeBytes       = 0;
            __int64 lastWriteTime    = 0;
            unsigned long attributes = 0;
        };

        bool active          = false;
        bool metadataLoading = false;
        bool deferredConsent = false;
        uint8_t bucket       = 0;
        HRESULT status       = S_OK;
        std::wstring sourcePath;
        std::wstring destinationPath;
        std::wstring sourceName;
        std::wstring sourceParentPath;
        std::wstring destinationName;
        std::wstring destinationParentPath;
        ItemMetadata sourceMetadata;
        ItemMetadata destinationMetadata;
        std::array<uint8_t, kMaxConflictActions> actions{};
        size_t actionCount = 0;
        std::array<uint8_t, kMaxConflictActions> primaryActions{};
        size_t primaryActionCount = 0;
        std::array<uint8_t, kMaxConflictActions> overflowActions{};
        size_t overflowActionCount      = 0;
        uint8_t defaultAction           = 0u;
        uint8_t escapeAction            = 0u;
        bool applyToAllEligible         = false;
        bool applyToAllChecked          = false;
        bool skipAllEligible            = false;
        uint8_t repeatUnavailableReason = 0u; // FileOperations::ConflictRepeatUnavailableReason
        bool sourceIsDirectory          = false;
        bool buttonsPublishable         = false;
        bool retryFailed                = false;
        unsigned int attemptCount       = 0u; // C2: attempts already made; the question says so
        bool factItemCountKnown         = false;
        uint64_t factItemCount          = 0;
        bool factBytesKnown             = false;
        uint64_t factBytes              = 0;
        uint8_t overlapProblem          = 0u; // R4-A02-1
        uint64_t overlapTaskId          = 0;
        std::wstring consentDetail; // C1: what a permanent-delete confirmation deletes
        std::wstring consentNote;   // C10/I35 R7.3: how the location deletes; a handed-off leftover's warning
    };

    ConflictPromptSnapshot conflict{};

    uint64_t desiredSpeedLimitBytesPerSecond       = 0;
    uint64_t effectiveSpeedLimitBytesPerSecond     = 0;
    bool autoConcurrencyUsed                       = false;
    uint32_t autoConcurrencyStorageKind            = FILESYSTEM_STORAGE_UNKNOWN;
    uint32_t autoConcurrencyDestinationStorageKind = FILESYSTEM_STORAGE_UNKNOWN;
    unsigned int autoTunedConcurrency              = 0;
    unsigned int effectiveConcurrencyBudget        = 0;

    bool finished    = false;
    HRESULT resultHr = S_OK;
    // Actual retained projectable Issues rows, independent of severity counters.
    bool hasProjectableIssues  = false;
    unsigned long warningCount = 0;
    unsigned long errorCount   = 0;
    std::wstring lastDiagnosticMessage;
    // I35 R5.3: Information notes, shown only when showInformationalNotes is on (captured per snapshot).
    unsigned long noteCount = 0;
    std::wstring lastInfoNoteMessage;
    bool showInformationalNotes = false;
    // Any retained diagnostic, notes included: Show log and Export issues stay reachable.
    bool hasRetainedDiagnostics = false;
    // I35 R5.6 / D-5.2: typed skipped roots keep the task Partial without a counted warning.
    unsigned long skippedItemCount   = 0;
    bool partialExplainedByUserSkips = false;
    std::wstring resultSummary;
    unsigned long publishedItemCount          = 0;
    unsigned long unknownPublicationItemCount = 0;
    unsigned long removedSourceCount          = 0;
    bool hasIndeterminateResult               = false;
    bool interruptedMoveNotice                = false;
    uint64_t interruptedOperationId           = 0u;
    bool clipboardMoveAdmission               = false;
    bool clipboardMoveConsumed                = false;
    bool hasSourceActionPaths                 = false;
    bool exactRetainedSourceActionAvailable   = false;
    unsigned long retainedSourceCount         = 0;
    unsigned long retainedSourceNativeCount   = 0;
    unsigned long canceledRetainedSourceCount = 0;
    unsigned long unknownSourceCount          = 0;
    unsigned long missingSourceCount          = 0;
    StatusKind statusKind                     = StatusKind::None;
    // I35 R9.4 (P3): false when the completed task changed nothing (CompletedTaskFacts).
    bool startedMutation = true;
    // The completed summary's final facts are projected, not the finished task's interim counters.
    bool completedSummaryProjected = false;

    bool started                       = false;
    bool paused                        = false;
    bool hasProgressCallbacks          = false;
    ULONGLONG lastProgressCallbackTick = 0;
    ULONGLONG operationStartTick       = 0;

    bool waitingForOthers  = false;
    bool waitingInQueue    = false;
    bool overlapQueued     = false; // R4-A02-1: queued behind named overlapping tasks; no Start now
    bool queuePaused       = false;
    uint64_t queueOrderKey = 0;
    bool canMoveQueueUp    = false;
    bool canMoveQueueDown  = false;

    // Single-pass discovery knowledge state.
    bool discoveryAheadActive               = false;
    bool discoverySkipped                   = false;
    bool discoveryClosed                    = false;
    bool firstMutationBeforeDiscoveryClosed = false;
    uint64_t discoveredTotalBytes           = 0;
    unsigned long discoveredFileCount       = 0;
    unsigned long discoveredDirectoryCount  = 0;
    ULONGLONG discoveryElapsedMs            = 0;

    unsigned long plannedItems = 0;
    std::filesystem::path destinationFolder;
    std::optional<FolderWindow::Pane> destinationPane;
    std::wstring destinationPluginId;
    std::wstring destinationPluginShortId;
    std::wstring destinationInstanceContext;
};

// Cumulative retained-tree work (popup-b#1, x-perf#4). A rebuild replaces the whole hosted tree; a
// reconcile adds, keeps or retires individual controls by identity. Production code keeps these
// for the perf metrics; self-tests read them.
struct PopupHostedTreeDebugCounters final
{
    uint64_t syncs                   = 0u;
    uint64_t rebuilds                = 0u;
    uint64_t reconciles              = 0u;
    uint64_t controlsCreated         = 0u;
    uint64_t controlsRetired         = 0u;
    uint64_t accessibleTextUpdates   = 0u;
    uint64_t accessibleTextDeferrals = 0u;
    // Retired during a capture and published again before it ended: shown again, not recreated.
    uint64_t controlsRevived = 0u;
    // Whole synchronizations (tree work, control updates and publication), by kind.
    uint64_t rebuildUs   = 0u;
    uint64_t reconcileUs = 0u;
};

#ifdef ENABLE_TESTS
struct PopupTaskSnapshotRequest
{
    uint64_t taskId = 0;
    bool found      = false;
    TaskSnapshot snapshot{};
};

struct CaptionGlyphDebugSnapshot
{
    bool statusVisible                 = false;
    bool usesDirectWriteGlyphRendering = false;
    bool usesGdiTextFallback           = true;
    bool highContrastSuppressed        = false;
};

// UI-thread scalar observation only: reading these counters never builds a task
// snapshot, paints, or calls taskbar COM. Baseline and candidate use one fixture.
struct PopupTelemetryDebugSnapshot final
{
    uint64_t timerTicks       = 0u;
    uint64_t taskSnapshots    = 0u;
    uint64_t rateSnapshots    = 0u;
    uint64_t rateUpdates      = 0u;
    uint64_t taskbarAttempts  = 0u;
    uint64_t taskbarSuccesses = 0u;
    uint32_t lastTaskbarState = 0u;
    UINT timerIntervalMs      = 0u;
    bool idlePolicyEnabled    = false;
};

struct PopupLayoutDebugSnapshot
{
    // Facts recorded at the real paint sites. These observations never select or render a row.
    struct PaintedFileProgress final
    {
        std::wstring sourcePath;
        uint64_t progressStreamId = 0u;
        uint64_t totalBytes       = 0u;
        uint64_t completedBytes   = 0u;
        D2D1_RECT_F track{};
        D2D1_RECT_F fill{};
        D2D1_COLOR_F trackColor{};
        D2D1_COLOR_F fillColor{};
        bool determinate = false;
    };
    std::array<PaintedFileProgress, TaskSnapshot::kMaxInFlightFiles> taskPaintedFileProgress{};
    size_t taskPaintedFileProgressCount = 0u;
    uint64_t taskPaintGeneration        = 0u;
    std::wstring taskPaintedDiscoveryText;
    std::wstring taskPaintedTransferText;
    D2D1_RECT_F taskPaintedDiscoveryRect{};
    D2D1_COLOR_F taskPaintedDiscoveryColor{};
    D2D1_RECT_F taskPaintedTransferRect{};
    D2D1_COLOR_F taskPaintedTransferColor{};
    uint64_t taskPaintedTransferCompletedBytes = 0u;
    FileOperationsSelfTest::SurfaceResources surfaceResources{};
    uint64_t taskId = 0;
    bool found      = false;

    size_t visibleButtonCount                      = 0u;
    size_t footerVisibleButtonCount                = 0u;
    bool hasVisibleButtonOverlap                   = false;
    bool taskHasVisibleButtonOverlap               = false;
    bool conflictApplyToAllVisible                 = false;
    bool conflictApplyToAllIsCheckbox              = false;
    bool conflictApplyToAllChecked                 = false;
    bool conflictMoreVisible                       = false;
    bool completedDiagnosticsMoreVisible           = false;
    bool completedDiagnosticsMoreButtonRectVisible = false;
    bool completedIssuesVisible                    = false;
    bool completedIssuesHitTargetActive            = false;
    bool completedIssuesLabelFits                  = false;
    bool completedActionRowFits                    = true;
    D2D1_RECT_F completedIssuesButtonRect{};
    std::wstring completedIssuesHostedText;
    std::wstring completedIssuesHostedName;
    bool completedShowLogVisible                 = false;
    bool completedExportIssuesVisible            = false;
    bool completedFailedItemsActionVisible       = false;
    bool completedOpenDestinationActionVisible   = false;
    bool completedRevealDestinationActionVisible = false;
    bool completedDismissVisible                 = false;
    bool taskToggleCollapseVisible               = false;
    bool taskStartNowVisible                     = false;
    bool taskQueueMoveUpVisible                  = false;
    bool taskQueueMoveDownVisible                = false;
    bool taskPauseVisible                        = false;
    bool taskCancelVisible                       = false;
    bool taskSkipVisible                         = false;
    bool taskSpeedLimitVisible                   = false;
    bool taskSpeedLimitUsesSelectorChrome        = false;
    bool taskSpeedLimitShowsCurrentValue         = false;
    float taskSpeedLimitLabelWidthPx             = 0.0f;
    float taskSpeedLimitTextWidthPx              = 0.0f;
    bool taskSpeedLimitLabelFits                 = false;
    bool taskPauseHitTargetActive                = false;
    bool taskCancelHitTargetActive               = false;
    bool taskSkipHitTargetActive                 = false;
    float taskSkipLabelWidthPx                   = 0.0f;
    float taskSkipControlWidthPx                 = 0.0f;
    bool taskSpeedLimitHitTargetActive           = false;
    bool taskDiscoveryAheadActive                = false;
    bool taskDiscoveryIndicatorVisible           = false;
    size_t taskPaintedPathLabelCount             = 0u;
    bool taskPaintedPathLabelsFit                = true;
    size_t taskPaintedPathCount                  = 0u;
    bool taskPaintedPathsFit                     = true;
    bool taskPaintedPathsValidUnicode            = true;
    bool taskPaintedPathsRetainRawExtension      = true;
    bool taskScanStatusVisible                   = false;
    bool taskScanStatusFits                      = false;
    bool taskScanStatusValidUnicode              = false;
    bool taskScanStatusAbbreviated               = false;
    bool taskScanStatusNeedsAbbreviation         = false;
    bool taskScanStatusTooltipComplete           = false;
    std::wstring taskScanStatusText;
    std::wstring taskScanStatusCopyPath;
    uint64_t scanStatusMeasurementCount         = 0u;
    size_t scanStatusRetainedLayouts            = 0u;
    size_t scanStatusLayoutSlots                = 0u;
    float taskPaintedPathLabelMinimumSlackPx    = 0.0f;
    bool taskTransferProgressVisible            = false;
    bool taskTransferProgressDeterminate        = false;
    bool taskTransferProgressProvisional        = false;
    bool taskTransferProgressMarqueeVisible     = false;
    bool footerPauseResumeAllVisible            = false;
    bool footerPauseResumeAllPauses             = false;
    bool footerQueueModeSegmentedVisible        = false;
    bool footerQueueModeSelectorVisible         = false;
    bool footerQueueModeUsesSelectorChrome      = false;
    bool footerQueueModeSelectorHitTargetActive = false;
    bool footerQueueModeIsParallel              = false;
    bool footerQueueHitTargetActive             = false;
    bool footerParallelHitTargetActive          = false;
    bool footerAutoDismissVisible               = false;
    bool footerAutoDismissLabelVisible          = false;
    bool footerAutoDismissEnabled               = false;
    wchar_t checkboxGlyph                       = L'\0';
    bool footerDensityToggleVisible             = false;
    bool footerDensityHitTargetActive           = false;
    bool footerOptionsVisible                   = false;
    bool footerOptionsHitTargetActive           = false;
    size_t footerMeasuredActionCount            = 0u;
    float footerActionMinimumPaddingPx          = 0.0f;
    bool footerActionHitTargetsActive           = true;
    bool popupCompactDensity                    = false;
    bool taskThroughputDisclosureVisible        = false;
    bool taskThroughputExpanded                 = true;
    bool taskThroughputSummaryVisible           = false;
    D2D1_RECT_F taskThroughputDisclosureRect{};
    bool footerAggregateProgressVisible             = false;
    bool footerAggregateProgressDeterminate         = false;
    double footerAggregateProgressFraction          = 0.0;
    uint32_t footerOpenDiscoveryTasks               = 0u;
    bool footerOnly                                 = false;
    bool footerDetailsToggleVisible                 = false;
    bool footerDetailsToggleRightAligned            = false;
    bool footerDetailsToggleUsesDisclosureChrome    = false;
    bool highContrastEnabled                        = false;
    bool reducedMotionEnabled                       = false;
    bool layoutMotionActive                         = false;
    bool autoResizeMotionActive                     = false;
    bool selectedTaskHeightMotionActive             = false;
    bool completedGroupLayoutMotionActive           = false;
    bool autoResizeAnimationEnabled                 = false;
    bool footerQueueModeAnimationEnabled            = false;
    bool graphStatusAnimationEnabled                = false;
    bool conflictStackedPathRows                    = false;
    bool conflictSourceMetadataVisible              = false;
    bool conflictDestinationMetadataVisible         = false;
    bool conflictMetadataSizeCompareVisible         = false;
    bool conflictMetadataDateCompareVisible         = false;
    bool conflictPromptPrecedesActions              = false;
    float conflictPromptContentBottomPx             = 0.0f;
    float conflictPromptVisibleBottomPx             = 0.0f;
    float conflictActionsTopPx                      = 0.0f;
    bool conflictWaitingForDecisionVisible          = false;
    bool conflictDecisionDetailsLoadingVisible      = false;
    bool conflictPromptVisibleWithoutClipping       = false;
    bool conflictContextVisibleWithoutClipping      = false;
    bool conflictDiscoveryIndicatorVisible          = false;
    bool conflictTransferProgressVisible            = false;
    bool lastNoteVisibleWithoutClipping             = false;
    bool taskDuplicateUnderGraphItemBarVisible      = false;
    bool taskStatusStripeVisible                    = false;
    bool taskStatusChipVisible                      = false;
    bool taskStatusChipActuallyPainted              = false;
    bool taskStatusChipPaintedTextFits              = false;
    float taskStatusChipPaintedTextWidthPx          = 0.0f;
    float taskStatusChipPaintedTextAvailableWidthPx = 0.0f;
    D2D1_RECT_F taskStatusChipPaintedRect{};
    D2D1_RECT_F taskStatusChipPaintedTextRect{};
    std::wstring taskStatusChipPaintedText;
    std::wstring taskStatusHeaderPaintedText;
    std::wstring taskSubjectPaintedText;
    D2D1_RECT_F taskSubjectRect{};
    D2D1_RECT_F taskAmountRect{};
    D2D1_RECT_F taskPercentRect{};
    D2D1_RECT_F taskRateRect{};
    D2D1_RECT_F taskEtaRect{};
    D2D1_RECT_F taskCountsRect{};
    D2D1_RECT_F taskThroughputLabelRect{};
    D2D1_RECT_F taskConcurrencyRect{};
    std::wstring taskCountsText;
    std::wstring taskConcurrencyText;
    struct PaintedRecursiveCounts final
    {
        unsigned long filesFound   = 0u;
        unsigned long foldersFound = 0u;
        uint64_t filesCopied       = 0u;
        uint64_t foldersCopied     = 0u;
        bool filesCopiedKnown      = false;
        bool foldersCopiedKnown    = false;
        bool discoveryClosed       = false;
    };
    PaintedRecursiveCounts taskPaintedRecursiveCounts{};
    bool taskStatusGlyphSignalVisible       = false;
    bool taskStatusTextSignalVisible        = false;
    bool taskStatusColorBlindSafeEncoding   = false;
    bool taskCollapsed                      = false;
    bool taskCompactRow                     = false;
    bool taskAutoCollapsedOnCompletion      = false;
    bool taskCompactProgressVisible         = false;
    bool taskHiddenByCompletedGroup         = false;
    bool completedGroupVisible              = false;
    bool completedGroupExpanded             = false;
    bool completedGroupToggleVisible        = false;
    bool completedGroupClearVisible         = false;
    bool completedGroupClearTooltipEmpty    = false;
    bool completedGroupClearLabelFits       = false;
    size_t completedGroupExpectedBadges     = 0u;
    size_t completedGroupPaintedBadges      = 0u;
    uint32_t completedGroupPaintedBadgeMask = 0u;
    std::array<wchar_t, 5> completedGroupPaintedGlyphs{};
    std::array<std::wstring, 5> completedGroupHostedNames{};
    bool completedGroupBadgeCountsFit       = false;
    bool taskTerminalSummaryVisible         = false;
    bool taskTerminalSummaryFits            = false;
    bool taskTerminalSummaryPrecedesActions = false;
    // Actual painted semantic region and synchronized retained DxUi::Label, not a projection of TaskSnapshot counts.
    bool taskTerminalFactsVisible         = false;
    bool taskTerminalFactsFits            = false;
    bool taskTerminalFactsFollowsSummary  = false;
    bool taskTerminalFactsPrecedesActions = false;
    std::wstring taskTerminalFactsText;
    std::wstring taskTerminalFactsHostedName;
    D2D1_RECT_F taskTerminalFactsPaintedRect{};
    bool completedGroupAnimationEnabled    = false;
    bool taskCardAnimationEnabled          = false;
    bool hostedProgressAnimationEnabled    = false;
    bool taskHostedProgressPresent         = false;
    double taskHostedProgressValue         = 0.0;
    bool taskHostedProgressIndeterminate   = false;
    bool taskHostedProgressAnimationPaused = false;
    bool taskHostedProgressEnabled         = false;
    bool taskHostedProgressSegmented       = false;
    std::wstring taskHostedProgressHelpText;
    // I35 (owner report 2026-10-13): the verification-half marks over the task's bar, the text beside its compact meter,
    // and whether its card height has not reached its target yet.
    bool taskVerificationMarksVisible   = false;
    double taskVerificationMarksPercent = 0.0;
    bool taskVerificationMarksOutlined  = false;
    std::wstring taskCompactMeterLabel;
    bool taskHeightAnimationUnsettled             = false;
    bool usesDxUiHost                             = false;
    size_t hostedActionControlCount               = 0u;
    size_t hostedProgressControlCount             = 0u;
    size_t hostedGraphControlCount                = 0u;
    size_t hostedGraphStallBandCount              = 0u;
    size_t hostedTooltipRegionCount               = 0u;
    bool completedGroupBadgeTooltipsInformative   = false;
    uint64_t hostedAccessibilityNotificationCount = 0u;
    int scrollPosition                            = 0;
    D2D1_RECT_F listViewportRect{};
    bool hostedKeyboardFocusVisible = false;
    bool hostedFocusedHasFocus      = false;
    bool hostedReducedMotion        = false;
    std::wstring hostedFocusedAutomationId;
    uint64_t hostedNativeTabCount   = 0u;
    uint64_t hostedFocusChangeCount = 0u;
    std::wstring hostedPreviousFocusTargetId;
    std::wstring hostedLastFocusTargetId;
    std::wstring hostedDefaultAutomationId;
    std::wstring hostedCancelAutomationId;
    uint32_t completedGroupCount            = 0u;
    uint32_t completedGroupCompletedCount   = 0u;
    uint32_t completedGroupPartialCount     = 0u;
    uint32_t completedGroupFailedCount      = 0u;
    uint32_t completedGroupCanceledCount    = 0u;
    uint32_t completedGroupVisibleTaskCount = 0u;
    uint32_t taskStatusVisualTone           = 0u;
    uint32_t taskStatusVisualColorRef       = 0u;
    uint32_t taskUnderGraphProgressBarCount = 0u;
    uint64_t footerAggregateCompletedBytes  = 0;
    uint64_t footerAggregateTotalBytes      = 0;
    uint64_t footerAggregateCompletedItems  = 0;
    uint64_t footerAggregateTotalItems      = 0;
    double footerAggregateBytesPerSecond    = 0.0;
    bool footerAggregateEtaVisible          = false;
    uint64_t footerAggregateEtaSeconds      = 0;
    uint32_t taskbarProgressState           = 0u;
    uint64_t taskbarProgressCompleted       = 0;
    uint64_t taskbarProgressTotal           = 0;
    uint64_t taskbarUpdateCount             = 0;
    uint64_t telemetryTimerTickCount        = 0;
    uint64_t taskSnapshotBuildCount         = 0;
    bool taskbarButtonReady                 = false;
    bool taskbarListAvailable               = false;
    bool taskbarListRetryPending            = false;
    uint32_t taskbarListAttemptCount        = 0u;
    uint64_t taskbarListRetryDelayMs        = 0;
    D2D1_RECT_F footerQueueSegmentRect{};
    D2D1_RECT_F footerParallelSegmentRect{};
    D2D1_RECT_F footerSummaryRect{};
    float footerSummaryRequiredHeightPx = 0.0f;
    uint32_t footerSummaryLineCount     = 0u;
    bool footerSummaryFits              = false;
    D2D1_RECT_F completedDiagnosticsMoreButtonRect{};
    size_t conflictPrimaryActionCount   = 0u;
    size_t conflictOverflowActionCount  = 0u;
    uint8_t conflictDefaultAction       = 0u;
    uint8_t conflictEscapeAction        = 0u;
    bool conflictButtonsPublishable     = false;
    bool conflictPrimaryLabelsFit       = true;
    bool conflictScopeLabelFits         = false;
    size_t conflictMeasuredPrimaryCount = 0u;
    bool conflictSkipAllEligible        = false;
    size_t completedVisibleActionCount  = 0u;
    std::array<uint8_t, TaskSnapshot::kMaxConflictActions> conflictPrimaryActions{};
    std::array<uint8_t, TaskSnapshot::kMaxConflictActions> conflictOverflowActions{};
    TaskSnapshot::StatusKind taskStatusKind = TaskSnapshot::StatusKind::None;
    size_t taskStatusActiveStateCount       = 0u;
    uint32_t globalRunningCount             = 0u;
    uint32_t globalWaitingCount             = 0u;
    uint32_t globalNeedAttentionCount       = 0u;
    uint32_t completedAutoCollapsedCount    = 0u;
    bool globalSummaryVisible               = false;
    std::wstring globalSummaryText;

    // Graph hue fairness (Fairstream 4D): aggregated over the task's live rate-history buckets.
    uint32_t graphMultiHueBucketCount      = 0u;
    uint32_t graphSingleHueBucketCount     = 0u;
    uint32_t graphDistinctHueCount         = 0u;
    double graphMinHueShare                = 0.0; // per-hue share of summed multi-hue bucket weight
    double graphMaxHueShare                = 0.0;
    uint32_t graphDebugAccumulateCalls     = 0u;
    uint32_t graphDebugLastPending         = 0u;
    uint32_t graphDebugMaxStreams          = 0u;
    uint32_t graphRowColorMatchCount       = 0u;
    uint32_t graphRowColorMismatchCount    = 0u;
    bool graphCurrentBandwidthLineVisible  = false;
    bool graphCurrentBandwidthLabelVisible = false;
    bool graphEtaLabelVisible              = false;
    bool graphEtaLabelRightAligned         = false;
    bool graphLabelsOutsidePlot            = false;
    bool graphLabelsFit                    = false;
    D2D1_RECT_F graphPlotRect{};
    float graphLabelVisibleOverhangPx    = 0.0f;
    float graphLabelWhitespaceOverhangPx = 0.0f;
    D2D1_RECT_F graphRateLabelRect{};
    D2D1_RECT_F graphEtaLabelRect{};
    size_t graphHistoryCount       = 0u;
    size_t graphHistoryWriteIndex  = 0u;
    float graphHistoryOldestSample = 0.0f;
    float graphHistoryNewestSample = 0.0f;
    size_t graphStalledSampleCount = 0u; // I35 R10.2: samples the stall band hatches
    D2D1_RECT_F graphStallBandRect{};    // where the stall band may hatch
    float graphLabelRowBottomPx = 0.0f;  // bottom of the rate/ETA labels, above or inside the plot
    D2D1_RECT_F conflictBodyBounds{};
    D2D1_RECT_F conflictPanelBounds{};
    float conflictBodyContentHeightPx = 0.0f;
    float conflictBodyOffsetDip       = 0.0f;
    size_t taskInspectionControlCount = 0u;
    bool taskInspectionResultFits     = false;
    bool taskInspectionLabelsFit      = true;
    bool taskInlineSpeedRowVisible    = false;
    bool taskInlineEtaRowVisible      = false;
    float taskExpandedBaseHeightDip   = 0.0f;
    bool taskPaintedBodyBoundsValid   = false;
    D2D1_RECT_F taskPaintedCardRect{};
    float taskPaintedBodyBottomPx = 0.0f;
    D2D1_RECT_F taskPaintedActionRowRect{};
    std::wstring taskLastNoteText;
    D2D1_RECT_F taskLastNoteRect{};
    uint32_t taskLastNoteLineCount             = 0u;
    bool taskLastNoteFits                      = false;
    bool taskLastNotePrecedesProgress          = false;
    double graphCurrentBandwidthBytesPerSecond = 0.0;
    bool failureSurfaceActive                  = false; // C4: native text + Cancel all + Close, no D2D/DxUi
    bool failureSurfaceReady                   = false;
    size_t failureSurfaceChildCount            = 0u;
    uint64_t failureSurfaceTabCount            = 0u;
    uint64_t failureSurfaceReentrantEntryCount = 0u;
    std::wstring failureSurfaceText;
};

struct GraphHueWeightDebugSnapshot
{
    static constexpr size_t kMaxHues = 16u;

    size_t hueCount    = 0u;
    double totalWeight = 0.0;
    std::array<float, kMaxHues> hues{};
    std::array<double, kMaxHues> weights{};
};

struct PopupPointerDebugSnapshot
{
    uint64_t mouseMoveCount = 0u;
    POINT lastMouseMove{};
    float hostDpi      = 0.0f;
    HWND hostHwnd      = nullptr;
    HWND nativeCapture = nullptr;
    std::wstring capturedAutomationId;
    std::wstring hitAutomationId;
    D2D1_RECT_F hitBoundsDip{};
    bool hitVisible = false;
    bool hitEnabled = false;
    D2D1_RECT_F targetBoundsDip{};
    bool targetVisible = false;
    bool targetEnabled = false;
    bool targetHovered = false;
};
#endif

// I35 R10.2: what a transfer that delivered no bytes for longer than the stall threshold is waiting
// on, named next to its speed.
enum class RateStallKind : uint8_t
{
    None,
    WaitingForDestination,
    WaitingForSource,
    WaitingForStorage,
    Finalizing,
};

struct RateSnapshot
{
    struct InFlightStreamSnapshot
    {
        const void* cookieKey     = nullptr;
        uint64_t progressStreamId = 0;
        std::wstring sourcePath;
        uint64_t totalBytes      = 0;
        uint64_t completedBytes  = 0;
        ULONGLONG lastUpdateTick = 0;
    };

    uint64_t taskId               = 0;
    FileSystemOperation operation = FILESYSTEM_COPY;

    unsigned long completedItems        = 0;
    uint64_t totalBytes                 = 0;
    uint64_t completedBytes             = 0;
    uint64_t discoveredTotalBytes       = 0; // D2-A08 provisional ETA input while discovery is open
    bool discoveredBytesKnown           = false;
    uint64_t verificationTotalBytes     = 0;
    uint64_t verificationCompletedBytes = 0;
    uint64_t verificationReadBytes      = 0;
    std::wstring currentSourcePath;
    ULONGLONG lastProgressCallbackTick = 0;
    ULONGLONG progressStateChangeTick  = 0;
    // R10.2: the last report that advanced the bytes (0 before the first byte), and the stall name
    // the host derives from the phase the streams had announced and the slower side of the transfer.
    ULONGLONG lastByteProgressTick = 0;
    RateStallKind stallKind        = RateStallKind::WaitingForStorage;
    // The task's effective speed limit (0: none). A transfer held at its limit is not stalled.
    uint64_t speedLimitBytesPerSecond = 0;
    std::array<InFlightStreamSnapshot, TaskSnapshot::kMaxInFlightFiles> inFlightFiles{};
    size_t inFlightFileCount = 0;
    bool started             = false;
    bool paused              = false;
    bool waitingForOthers    = false;
    bool waitingInQueue      = false;
    bool queuePaused         = false;
    bool waitingForDecision  = false; // a question waits for the user: no bytes is not a stall
    bool discoveryClosed     = false;
    bool finished            = false;
    // Cancellation was requested: what the task waits on is its own abort, never a named stall.
    bool stopping = false;
};

struct RateHistory
{
    static constexpr size_t kMaxSamples             = 180u; // ~18s @ 100ms
    static constexpr size_t kMaxHueWeightsPerSample = TaskSnapshot::kMaxInFlightFiles;

    struct HueWeight
    {
        float hue         = -1.0f;
        double weight     = 0.0;
        uint8_t colorSlot = DxUi::ThroughputGraphSample::kInvalidColorSlot;
    };

    struct StreamProgress
    {
        const void* cookieKey     = nullptr;
        uint64_t progressStreamId = 0;
        std::wstring sourcePath;
        uint64_t completedBytes  = 0;
        ULONGLONG lastUpdateTick = 0;
        // Golden-angle palette hue assigned when the stream first appears; stable for the
        // stream's lifetime and well-separated from its concurrent neighbors (a raw path-hash
        // hue gives no minimum separation, so equal streams often looked identical).
        float assignedHue         = -1.0f;
        uint8_t assignedColorSlot = DxUi::ThroughputGraphSample::kInvalidColorSlot;
    };

    // I35 R10.2: the task's cumulative values at one telemetry tick; the displayed rates are the
    // deltas over the last few seconds of these samples.
    struct WindowSample
    {
        ULONGLONG tick             = 0;
        uint64_t bytes             = 0;
        uint64_t verificationBytes = 0;
        unsigned long items        = 0;
    };
    static constexpr size_t kMaxWindowSamples = 128u; // > 5 s at the 50 ms minimum spacing

    std::array<float, kMaxSamples> samples{};
    std::array<float, kMaxSamples> verificationSamples{};
    std::array<float, kMaxSamples> hues{}; // Per-sample hue (0-360) for rainbow mode
    std::array<std::array<HueWeight, kMaxHueWeightsPerSample>, kMaxSamples> hueWeights{};
    std::array<uint8_t, kMaxSamples> hueWeightCounts{};
    // R10.2: 1 when the bucket was spent stalled; the graph hatches it as a stall band.
    std::array<uint8_t, kMaxSamples> stalledSamples{};
    size_t count      = 0;
    size_t writeIndex = 0;
    bool initialized  = false;

    ULONGLONG lastStateChangeTick      = 0;
    ULONGLONG lastProgressCallbackTick = 0;
    ULONGLONG lastDisplaySampleTick    = 0;

    // R10.2 window: a ring of WindowSample, oldest at windowFirst. The window never reaches before
    // windowFloorTick: the first sighting, the end of a pause or queue wait, or the end of a stall
    // longer than the window (measuring restarts there, as a resume restarts it).
    std::array<WindowSample, kMaxWindowSamples> window{};
    size_t windowFirst            = 0;
    size_t windowCount            = 0;
    ULONGLONG windowFloorTick     = 0;
    ULONGLONG lastByteAdvanceTick = 0; // the telemetry tick at which the cumulative values last rose
    RateStallKind stallKind       = RateStallKind::None;
    // The task's report cadence: the recent gaps between its byte-advancing reports (a ring). A stall
    // must outlast 1.5 x their median, so a steady cadence coarser than 600 ms (a throttled or slow
    // destination reporting its chunks) is never named a stall.
    static constexpr size_t kMaxByteReportGaps = 8u;
    std::array<ULONGLONG, kMaxByteReportGaps> byteReportGapsMs{};
    size_t byteReportGapCount    = 0;
    size_t byteReportGapNext     = 0;
    ULONGLONG lastByteReportTick = 0; // the task's last byte report seen, or where measuring restarted

    ULONGLONG pendingBucketMs                  = 0;
    double pendingWeightedSampleMs             = 0.0;
    double pendingWeightedVerificationSampleMs = 0.0;
    ULONGLONG pendingStalledMs                 = 0;
    float pendingHue                           = -1.0f;
    std::array<HueWeight, kMaxHueWeightsPerSample> pendingHueWeights{};
    size_t pendingHueWeightCount = 0;
    std::array<StreamProgress, TaskSnapshot::kMaxInFlightFiles> streamProgress{};
    size_t streamProgressCount = 0;

    // Diagnostics for the fairness selftest (cheap counters, always maintained).
    uint32_t debugAccumulateCalls  = 0;
    uint32_t debugLastPendingCount = 0;
    uint32_t debugMaxStreamsSeen   = 0;

    double displayedBytesPerSec             = 0.0;
    double displayedVerificationBytesPerSec = 0.0;
    double displayedItemsPerSec             = 0.0;
    double smoothedEtaSeconds               = 0.0;
    bool hasSmoothedEta                     = false;
    // D2-A08: while discovery is open, the remaining time over the workload discovered so far.
    double provisionalEtaSeconds = 0.0;
    bool hasProvisionalEta       = false;
};

#ifdef ENABLE_TESTS
// I35 R10.2: one synthetic telemetry tick for the deterministic rate-model gates. Each tick drives the
// production per-task step of UpdateRates with these inputs.
struct RateModelDebugTick
{
    ULONGLONG tick                     = 0;
    uint64_t completedBytes            = 0;
    uint64_t totalBytes                = 0;
    uint64_t verificationReadBytes     = 0;
    unsigned long completedItems       = 0;
    ULONGLONG lastProgressCallbackTick = 0;
    ULONGLONG lastByteProgressTick     = 0;
    ULONGLONG progressStateChangeTick  = 0;
    RateStallKind stallKind            = RateStallKind::WaitingForStorage;
    uint64_t speedLimitBytesPerSecond  = 0;
    bool paused                        = false;
    bool waitingForDecision            = false;
    bool stopping                      = false;
};

// What the popup would show after one RateModelDebugTick.
struct RateModelDebugObservation
{
    double displayedBytesPerSec             = 0.0;
    double displayedVerificationBytesPerSec = 0.0;
    double displayedItemsPerSec             = 0.0;
    RateStallKind stallKind                 = RateStallKind::None;
    std::wstring rateLabelText;
    bool hasSmoothedEta       = false;
    double smoothedEtaSeconds = 0.0;
    size_t sampleCount        = 0;
    size_t stalledSampleCount = 0;
    float newestSample        = 0.0f;
    bool newestSampleStalled  = false;
};
#endif

class FileOperationsPopupState final
{
public:
    FolderWindow::FileOperationState* fileOps = nullptr;
    FolderWindow* folderWindow                = nullptr;
    std::weak_ptr<void> hostLifetime;

#ifdef ENABLE_TESTS
    // UI-thread-only F6 counterfactual; no product setting or process-global policy.
    void DebugSetIdleSuspensionEnabled(HWND hwnd, bool enabled) noexcept
    {
        _debugIdleSuspensionEnabled = enabled;
        RefreshTelemetryTimer(hwnd);
    }
    [[nodiscard]] bool DebugHasVisibleUiMotion(HWND hwnd) const noexcept
    {
        return HasVisibleUiMotion(hwnd, GetTickCount64());
    }
    // Leaves an auto-resize pending, as a content change just before opening path detail does.
    void DebugArmAutoResizeMotionForSelfTest() noexcept
    {
        _autoResizePending        = true;
        _autoResizePendingDueTick = GetTickCount64();
    }
    [[nodiscard]] PopupTelemetryDebugSnapshot DebugTelemetrySnapshot() const noexcept
    {
        return {_telemetryTimerTickCount,
                _taskSnapshotBuildCount,
                _rateSnapshotBuildCount,
                _rateUpdateCount,
                _taskbarUpdateCount,
                _taskbarSuccessCount,
                _lastTaskbarState,
                _telemetryTimerIntervalMs,
                _debugIdleSuspensionEnabled};
    }
    // UI-thread-only documentation fixture; never starts workers or admits operations.
    void DebugSetVisualScenario(std::vector<TaskSnapshot> tasks, bool collapsed, bool completedGroupExpanded);
    // A live fixture transition keeps geometry/choices. Publish its text immediately so a deferred
    // accessibility-text timer cannot accidentally supply the missing height-animation frames.
    void DebugUpdateLiveVisualScenario(std::vector<TaskSnapshot> tasks)
    {
        _hostedAccessibleTextFlush = true;
        debugVisualScenario        = std::move(tasks);
    }
    // Runs one telemetry timer tick as production does when nothing is pending: no deferred accessible text
    // and no explicit motion window (the state in which the last tick turns the popup idle).
    void DebugRunIdleTimerTickForSelfTest(HWND hwnd) noexcept;
    // Leaves a task's card height away from its target, its motion started `ageMs` ago, as when every timer tick of the
    // motion came after its deadline (a busy UI thread).
    void DebugAgeTaskHeightMotionForSelfTest(uint64_t taskId, double displacementPx, ULONGLONG ageMs) noexcept
    {
        const auto found = _taskHeightAnimations.find(taskId);
        if (found == _taskHeightAnimations.end() || ! found->second.height.initialized)
            return;
        ScalarAnimation& height = found->second.height;
        const ULONGLONG now     = GetTickCount64();
        height.start            = height.target + displacementPx;
        height.displayed        = height.start;
        height.startTick        = now > ageMs ? now - ageMs : 0u;
    }
    // I35 R10.2: names a stall on one task of the visual fixture (None clears it).
    void DebugSetVisualStallKind(uint64_t taskId, RateStallKind stall) noexcept
    {
        if (const auto it = _rates.find(taskId); it != _rates.end())
            it->second.stallKind = stall;
    }
    // Same UI thread as the synthetic fixture. Exercise host recovery without clearing
    // consumer resources here: the normal render path must replace the old generation.
    [[nodiscard]] bool DebugSimulateDeviceLoss() noexcept
    {
        _controlsHost.DebugSimulateDeviceLoss();
        return _controlsHost.GetDeviceContext() == nullptr;
    }
    [[nodiscard]] uintptr_t DebugPreparedContextIdentity() const noexcept
    {
        return reinterpret_cast<uintptr_t>(_target.get());
    }
    [[nodiscard]] bool DebugHasCurrentPreparedFrame() const noexcept
    {
        return _preparedBackground && _target && _target.get() == _controlsHost.GetDeviceContext();
    }
    // Pins the surface as if the host could not attach again: neither reopening nor the
    // backoff timer re-attaches it, so tests can drive the native children deterministically.
    [[nodiscard]] bool DebugEnterFailureSurface(HWND hwnd) noexcept
    {
        _debugFailureSurfacePinned = true;
        return EnterFailureSurface(hwnd);
    }
    [[nodiscard]] size_t DebugHostedHitCountForSelfTest(PopupHitTest::Kind kind, uint64_t taskId) const noexcept
    {
        size_t count = 0u;
        for (const PopupHitTest& hit : _hostedButtonHits)
            count += hit.kind == kind && hit.taskId == taskId ? 1u : 0u;
        return count;
    }
    [[nodiscard]] uint64_t DebugHostedRenderCount() const noexcept
    {
        return _controlsHost.DebugGetRenderCount();
    }
    [[nodiscard]] uint64_t DebugHostedActionCountForSelfTest() const noexcept
    {
        return _debugHostedActionCount;
    }
    [[nodiscard]] PopupHostedTreeDebugCounters DebugHostedTreeCountersForSelfTest() const noexcept
    {
        return _hostedTreeCounters;
    }
    // Perf baseline only: replace the whole hosted tree on every structural change, the strategy
    // before popup-b#1, so one fixed workload can be timed both ways in the same build.
    bool debugRebuildOnStructuralChange = false;
    // Keeps the controls a reconcile removes alive and detached (never in the tree) until reset, so a
    // test can prove no surviving panel still routes pointer input to them.
    void DebugSetRetainRetiredControlsForSelfTest(bool retain) noexcept
    {
        _debugRetainRetiredControls = retain;
        if (! retain)
            _debugRetiredControlQuarantine.clear();
    }
    [[nodiscard]] bool DebugRetiredControlRetainedForSelfTest(uintptr_t identity) const noexcept
    {
        return std::ranges::any_of(_debugRetiredControlQuarantine, [identity](const std::unique_ptr<DxUi::Control>& control) noexcept {
            return reinterpret_cast<uintptr_t>(control.get()) == identity;
        });
    }
    [[nodiscard]] uint64_t DebugDetachedControlInputCountForSelfTest() const noexcept
    {
        return _debugDetachedControlInputCount;
    }
    // Keeps a hosted path region's own hover flag unchanged from now on (false until it was hovered),
    // while its ScrollPanel still records it as the hovered child. False when no such region exists.
    bool DebugFreezeHostedRegionHoverFlagForSelfTest(std::wstring_view automationId) noexcept;
    [[nodiscard]] size_t DebugRetiredHostedControlCountForSelfTest() const noexcept
    {
        return _retiredHostedControls.size();
    }
    [[nodiscard]] bool DebugHostedCaptureActiveForSelfTest() const noexcept
    {
        return _controlsHost.GetCapturedControl() != nullptr;
    }
    // Drops the host capture through the public DxUi call that does not notify the captured control.
    void DebugReleaseHostCaptureSilentlyForSelfTest() noexcept
    {
        _controlsHost.ReleaseMouseCapture();
    }
    [[nodiscard]] size_t DebugHostedBodyChildCountForSelfTest() const noexcept
    {
        return _controlsBody ? _controlsBody->DebugChildCount() : 0u;
    }
    // A hosted region's published bounds, in client pixels.
    [[nodiscard]] std::optional<D2D1_RECT_F> DebugHostedRegionBoundsForSelfTest(std::wstring_view automationId) const
    {
        for (const HostedTooltipDescriptor& descriptor : _hostedTooltipDescriptors)
            if (descriptor.automationId == automationId)
                return descriptor.bounds;
        return std::nullopt;
    }
    // The retained control currently published under an automation ID (0 when none): the same
    // value across paints proves the control survived instead of being recreated.
    [[nodiscard]] uintptr_t DebugHostedControlIdentityForSelfTest(std::wstring_view automationId) const noexcept
    {
        const auto identity = [automationId](const DxUi::Control* control) noexcept
        { return control && control->GetAccessibleAutomationId() == automationId ? reinterpret_cast<uintptr_t>(control) : uintptr_t{0}; };
        for (const DxUi::Button* control : _hostedButtons)
            if (const uintptr_t found = identity(control))
                return found;
        for (const DxUi::Control* control : _hostedTooltipRegions)
            if (const uintptr_t found = identity(control))
                return found;
        for (const DxUi::ProgressBar* control : _hostedProgressBars)
            if (const uintptr_t found = identity(control))
                return found;
        for (const DxUi::ThroughputGraph* control : _hostedGraphs)
            if (const uintptr_t found = identity(control))
                return found;
        return 0u;
    }
    // Accessible Name/HelpText actually published on a hosted region, not its descriptor.
    [[nodiscard]] bool DebugHostedRegionAccessibleTextForSelfTest(std::wstring_view automationId, std::wstring& name, std::wstring& helpText) const
    {
        for (const DxUi::Control* region : _hostedTooltipRegions)
        {
            if (region && region->GetAccessibleAutomationId() == automationId)
            {
                name     = std::wstring(region->GetAccessibleName());
                helpText = std::wstring(region->GetAccessibleHelpText());
                return true;
            }
        }
        return false;
    }
    [[nodiscard]] bool DebugHostedProgressAccessibleNameForSelfTest(uint64_t taskId, std::wstring& name) const
    {
        for (size_t index = 0u; index < _hostedProgressDescriptors.size() && index < _hostedProgressBars.size(); ++index)
        {
            if (_hostedProgressBars[index] && _hostedProgressDescriptors[index].taskId == taskId)
            {
                name = std::wstring(_hostedProgressBars[index]->GetAccessibleName());
                return true;
            }
        }
        return false;
    }
    [[nodiscard]] bool DebugConflictBodyVisibleForSelfTest(uint64_t taskId) const noexcept
    {
        const auto* body = FindConflictBody(taskId);
        return body && body->visible;
    }
    [[nodiscard]] size_t DebugHostedRegionCountWithPrefixForSelfTest(std::wstring_view prefix) const noexcept
    {
        size_t count = 0u;
        for (const HostedTooltipDescriptor& descriptor : _hostedTooltipDescriptors)
            count += descriptor.automationId.starts_with(prefix) ? 1u : 0u;
        return count;
    }
    [[nodiscard]] uint64_t DebugConflictEscapeTaskIdForSelfTest() const noexcept
    {
        return _conflictEscapeTaskId;
    }
    [[nodiscard]] bool DebugHostedButtonTextForSelfTest(std::wstring_view automationId, std::wstring& text, std::wstring& accessibleName) const
    {
        for (const DxUi::Button* button : _hostedButtons)
        {
            if (button && button->GetAccessibleAutomationId() == automationId)
            {
                text           = std::wstring(button->GetText());
                accessibleName = std::wstring(button->GetAccessibleName());
                return true;
            }
        }
        return false;
    }
    [[nodiscard]] std::wstring DebugHostedRegionTextForSelfTest(std::wstring_view automationId) const
    {
        for (const HostedTooltipDescriptor& descriptor : _hostedTooltipDescriptors)
        {
            if (descriptor.automationId == automationId)
                return descriptor.text;
        }
        return {};
    }
    // The text a region actually painted (a path or notice may be abbreviated; DebugHostedRegionText
    // returns the exact text) and whether it fits that region's bounds.
    [[nodiscard]] std::wstring DebugHostedPaintedTextForSelfTest(std::wstring_view automationId, bool& fits) const;
    [[nodiscard]] bool DebugSetHostedButtonFocusForSelfTest(HWND hwnd, std::wstring_view automationId) noexcept
    {
        if (hwnd == nullptr || hwnd != _controlsHost.GetHwnd() || ! _pathDetailId.empty())
            return false;
        if (automationId.empty())
        {
            _controlsHost.SetFocusControl(nullptr, false);
            return _controlsHost.GetFocusControl() == nullptr;
        }
        for (DxUi::Button* button : _hostedButtons)
        {
            if (button && button->IsFocusable() && button->GetAccessibleAutomationId() == automationId)
            {
                _controlsHost.SetFocusControl(button, false);
                return _controlsHost.GetFocusControl() == button;
            }
        }
        return false;
    }
    [[nodiscard]] bool DebugSetHostedPathRegionFocusForSelfTest(HWND hwnd, std::wstring_view automationId) noexcept
    {
        if (hwnd == nullptr || hwnd != _controlsHost.GetHwnd() || ! _pathDetailId.empty())
            return false;
        for (DxUi::Control* region : _hostedTooltipRegions)
        {
            if (region && region->IsFocusable() && region->GetAccessibleAutomationId() == automationId)
            {
                _controlsHost.SetFocusControl(region, false);
                return _controlsHost.GetFocusControl() == region;
            }
        }
        return false;
    }
    std::optional<std::vector<TaskSnapshot>> debugVisualScenario;
    // Synthetic fixed-viewport benchmarks keep their declared client geometry;
    // ordinary operation/card resize scenarios retain the normal sizing policy.
    bool debugFixedGeometry = false;
    // Counterfactual for the completion-to-idle race; never changes task outcomes or product settings.
    bool debugAutoCollapseSnapEnabled = true;
    std::array<std::wstring, 2> debugGraphLabelText;
    [[nodiscard]] std::wstring_view DebugFocusedPathTooltip() const noexcept
    {
        return _controlsHost.GetTooltipText();
    }
    void DebugOpenPathDetail(std::wstring_view automationId) noexcept
    {
        OpenPathDetail(_controlsHost.GetHwnd(), automationId);
    }
    [[nodiscard]] bool DebugSetPathDetailLogicalFocusForSelfTest(HWND hwnd, std::wstring_view automationId) noexcept
    {
        if (hwnd == nullptr || hwnd != _controlsHost.GetHwnd() || ! _pathDetailField || ! _pathDetailCopy || ! _pathDetailClose ||
            (automationId != L"FileOperations.PathDetail.Text" && automationId != L"FileOperations.PathDetail.Copy" &&
             automationId != L"FileOperations.PathDetail.Close"))
            return false;
        _pathDetailLogicalFocusId.assign(automationId);
        _pathDetailNeedsFocus = true;
        if (automationId == L"FileOperations.PathDetail.Text")
            _controlsHost.ResetInteractionState();
        SyncPathDetail(hwnd);
        return _pathDetailLogicalFocusId == automationId;
    }
    [[nodiscard]] std::wstring_view DebugPathDetailLogicalFocusForSelfTest() const noexcept
    {
        return _pathDetailLogicalFocusId;
    }
    [[nodiscard]] DxUi::TextField* DebugPathDetailField() const noexcept
    {
        return _pathDetailField;
    }
    [[nodiscard]] bool DebugPathDetailMultilineState(DxUi::TextFieldDebugMultilineState& state) const noexcept
    {
        return _pathDetailField && _pathDetailField->DebugGetMultilineState(_controlsHost, state);
    }
    [[nodiscard]] bool DebugScrollPathDetailPageDownForSelfTest() noexcept
    {
        // Static capture drives the real control handler without acquiring native
        // text focus. Protected Commands and Gallery Interaction cover routed input.
        return ! _pathDetailId.empty() && _pathDetailField && _pathDetailField->OnKeyDown(_controlsHost, VK_NEXT, 0u);
    }
    [[nodiscard]] const std::vector<PopupButton>& DebugVisualButtons() const noexcept
    {
        return _buttons;
    }
    [[nodiscard]] bool DebugVisualButtonHovered(PopupHitTest::Kind kind) const noexcept
    {
        for (size_t index = 0u; index < _hostedButtonHits.size() && index < _hostedButtons.size(); ++index)
            if (_hostedButtonHits[index].kind == kind && _hostedButtons[index])
                return _hostedButtons[index]->IsHovered();
        return false;
    }
    // Synchronous scalar/string copies only; no control survives a message pump.
    [[nodiscard]] PopupPointerDebugSnapshot DebugPointerSnapshotForSelfTest(POINT clientPoint, PopupHitTest::Kind kind) noexcept;
    [[nodiscard]] uint64_t DebugNativeMouseMoveCountForSelfTest() const noexcept
    {
        return _debugMouseMoveCount;
    }
#endif

    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    void SetExpandedPlacementForRestore(const RECT& rect, bool preserveInitialSize = false) noexcept
    {
        _footerOnlyRestoreWindowRect = rect;
        // A restored expanded window owns its first measured geometry, just as
        // returning from footer-only mode does. Content changes may grow it later.
        _footerOnlyRestorePending = preserveInitialSize;
    }

    FileOperationsPopupState()                                           = default;
    FileOperationsPopupState(const FileOperationsPopupState&)            = delete;
    FileOperationsPopupState(FileOperationsPopupState&&)                 = delete;
    FileOperationsPopupState& operator=(const FileOperationsPopupState&) = delete;
    FileOperationsPopupState& operator=(FileOperationsPopupState&&)      = delete;

private:
    struct HostedProgressDescriptor final
    {
        D2D1_RECT_F bounds{};
        uint64_t taskId              = 0u;
        double value                 = 0.0;
        bool indeterminate           = false;
        bool segmented               = false;
        double primarySegmentValue   = 0.0;
        double secondarySegmentValue = 0.0;
        // A finished task: the part of its verification half not verified is outlined (I35, owner decision 2026-10-13).
        bool verificationSettled = false;
    };

    struct HostedGraphDescriptor final
    {
        D2D1_RECT_F bounds{};
        D2D1_RECT_F rateLabelBounds{};
        D2D1_RECT_F etaLabelBounds{};
        bool separateLabelRow    = false;
        bool labelsInTaskSummary = false;
        uint64_t taskId          = 0u;
        std::vector<DxUi::ThroughputGraphSample> samples;
        std::vector<double> verificationSamples;
        std::wstring overlayText;
        std::wstring currentBandwidthText;
        std::wstring currentBandwidthTrailingText;
        double currentBandwidthBytesPerSecond = 0.0;
        bool rainbowMode                      = false;
        bool perStreamBands                   = false;
        // I35 R10.2: one flag per sample, 1 where the stall band hatches the plot.
        std::vector<uint8_t> stalledSamples;
        // Where the stall band may hatch (pixels), and the bottom of the rate/ETA label row: above the plot
        // in the separate-row layout, inside it when there is no room for a row.
        D2D1_RECT_F stallBandBounds{};
        float labelRowBottom = 0.0f;
    };

    struct HostedTooltipDescriptor final
    {
        D2D1_RECT_F bounds{};
        std::wstring text;
        std::wstring automationId;
        std::wstring copyPath;
        uint64_t comparisonTaskId = 0u;
        // Owning task of a task-scoped region; focus on it is focus on that task.
        uint64_t taskId = 0u;
        // UIA HelpText and tooltip when they say more than the name (Q7: the other active streams).
        std::wstring helpText;
        // Text that changes with progress (rates, current paths, counters): its UIA Name/HelpText is
        // republished at most once per interval (x-perf#4). Other text is published at once.
        bool liveText = false;
#ifdef ENABLE_TESTS
        D2D1_RECT_F paintedLabelBounds{};
        D2D1_RECT_F paintedContainerBounds{};
        std::wstring paintedLabel;
        std::wstring paintedPath;
        wchar_t paintedGlyph = L'\0';
#endif
    };

    struct GraphLabelLayout final
    {
        std::wstring rateText;
        std::wstring etaText;
        wil::com_ptr<IDWriteTextFormat> format;
        wil::com_ptr<IDWriteTextLayout> rate;
        wil::com_ptr<IDWriteTextLayout> eta;
        float width       = 0.0f;
        float rateWidth   = 0.0f;
        float etaWidth    = 0.0f;
        float rateHeight  = 0.0f;
        float etaHeight   = 0.0f;
        float leadingSlot = 0.0f;
        bool stacked      = false;
    };

    // Bounded by snapshot rows, with text/layout work only when the inputs change.
    struct CardInformationLayout final
    {
        GraphLabelLayout progress;
        GraphLabelLayout counts;
        GraphLabelLayout throughput;
        GraphLabelLayout concurrency;
        GraphLabelLayout lastNote;
        std::wstring headerText;
        std::wstring paintedHeader;
        float headerWidth = 0.0f;
        wil::com_ptr<IDWriteTextFormat> headerFormat;
    };
    void PrepareCardInformation(const std::vector<TaskSnapshot>& snapshot, float width);
    void PrepareFactPair(GraphLabelLayout& pair, std::wstring leading, std::wstring trailing, IDWriteTextFormat* format, float width, float leadingSlot = 0.0f);
    [[nodiscard]] float FactPairHeight(const GraphLabelLayout& pair) const noexcept;
    void DrawFactPair(const GraphLabelLayout& pair, const D2D1_RECT_F& bounds, uint64_t taskId, std::wstring_view leadingId, std::wstring_view trailingId);

    struct ScalarAnimation final
    {
        double displayed    = 0.0;
        double start        = 0.0;
        double target       = 0.0;
        ULONGLONG startTick = 0u;
        bool initialized    = false;
    };

    struct TaskHeightAnimation final
    {
        ScalarAnimation height;
        bool expanded = false;
    };

    enum class CaptionStatus : uint8_t
    {
        None,
        Ok,
        Warning,
        Error,
    };

    void ApplyScrollBarTheme(HWND hwnd) const noexcept;
    void RefreshLocalizedFooterText();
    [[nodiscard]] bool IsReducedMotionEnabled() const noexcept;

    bool IsTaskCollapsed(uint64_t taskId) const noexcept;
    bool IsTaskThroughputExpanded(uint64_t taskId) const noexcept;
    bool IsTaskCollapsedForDisplay(uint64_t taskId, bool compactDensity) const noexcept;
    void ToggleTaskCollapsed(uint64_t taskId, bool compactDensity) noexcept;
    void AutoCollapseCompletedTasks(const std::vector<TaskSnapshot>& snapshot) noexcept;
    void ResolveBulkCancelGroupCollapse(HWND hwnd, const std::vector<TaskSnapshot>& snapshot, bool showCompletedGroup) noexcept;
    void CleanupCollapsedTasks(const std::vector<TaskSnapshot>& snapshot) noexcept;
    [[nodiscard]] float ResolveAnimatedTaskHeight(uint64_t taskId, float targetHeight, ULONGLONG tick, bool reducedMotion, bool expanded = false) noexcept;
    [[nodiscard]] float ResolveCompletedGroupVisibility(ULONGLONG tick, bool reducedMotion) noexcept;
    [[nodiscard]] float ResolveDisclosureProgress(ScalarAnimation& animation, bool expanded, ULONGLONG tick, bool reducedMotion) noexcept;
    [[nodiscard]] double ResolveHostedProgressValue(size_t index, double targetValue, ULONGLONG tick, bool reducedMotion) noexcept;

    void DiscardDeviceResources() noexcept;
    void EnsureFactories() noexcept;
    void EnsureTextFormats() noexcept;
    void EnsureCaptionGlyphTextFormats(UINT dpi) noexcept;
    void EnsureTarget(HWND hwnd) noexcept;
    bool EnsureCaptionGlyphTarget(UINT dpi) noexcept;
    void EnsureBrushes() noexcept;

    std::vector<TaskSnapshot> BuildSnapshot() const;
    std::vector<RateSnapshot> BuildRateSnapshot() const;
    void UpdateRates() noexcept;

    void LayoutChrome(float width, float height, bool showPauseResumeAll, bool showCancelAll, std::wstring_view summaryText) noexcept;
    void PrepareFooterSummary(float widthPx, std::wstring_view text) noexcept;
    void MeasureFooterActionWidths() noexcept;
    float MeasureSpeedLimitWidthDip(size_t slot, std::wstring_view caption) noexcept;
    float FooterHeightPixels() const noexcept;
    void UpdateScrollBar(HWND hwnd, float viewH, float contentH) noexcept;
    void AutoResizeWindow(HWND hwnd, float desiredContentHeight, size_t taskCount, bool footerOnly, bool reducedMotion) noexcept;
    void ArmUiMotionTimer(HWND hwnd) noexcept;
    [[nodiscard]] bool HasVisibleUiMotion(HWND hwnd, ULONGLONG nowTick) const noexcept;
    void RefreshTelemetryTimer(HWND hwnd) noexcept;

    void DrawDxUiButtonChrome(const PopupButton& button, IDWriteTextFormat* format, std::wstring_view text, DxUi::ButtonVariant variant) noexcept;
    void DrawButton(const PopupButton& button, IDWriteTextFormat* format, std::wstring_view text) noexcept;
    void DrawFooterQueueModeControl(const PopupButton& button, bool queueMode) noexcept;
    void DrawFooterAutoDismissControl(const PopupButton& button, bool enabled) noexcept;
    wchar_t DrawCenteredGlyph(const D2D1_RECT_F& rc, wchar_t fluentGlyph, wchar_t fallbackGlyph, ID2D1Brush* glyphBrush = nullptr) noexcept;
    void DrawMenuButton(const PopupButton& button, IDWriteTextFormat* format, std::wstring_view text) noexcept;
    void DrawSelectorButton(const PopupButton& button, IDWriteTextFormat* format, std::wstring_view text) noexcept;
    void DrawCheckboxBox(const D2D1_RECT_F& rect, bool checked) noexcept;
    void DrawDisclosureChevron(const D2D1_RECT_F& rc, float expandedProgress) noexcept;
    void DrawDiscoveryActivityIndicator(const D2D1_RECT_F& rc, ULONGLONG tick, bool reducedMotion) noexcept;
    void DrawBandwidthGraph(const D2D1_RECT_F& rect,
                            uint64_t taskId,
                            const RateHistory& history,
                            std::wstring_view currentValueTrailingText,
                            std::wstring_view overlayText,
                            bool rainbowMode,
                            bool perStreamBands,
                            bool showLabels               = true,
                            const D2D1_RECT_F* rateBounds = nullptr,
                            const D2D1_RECT_F* etaBounds  = nullptr) noexcept;
    // I35 R10.2: the bottom of the labels the hosted graph draws inside its plot when they have no row
    // above it (DxUi::ThroughputGraph: from 3 DIP below the top, each in its half of the width).
    [[nodiscard]] float InPlotGraphLabelRowBottom(const D2D1_RECT_F& plot, std::wstring_view leading, std::wstring_view trailing) noexcept;
    // I35 R10.2: the width of the widest rate label carrying a stall name (measured once per small text
    // format). The graph's label rows are laid out as if it were shown, so a stall name never moves them.
    [[nodiscard]] float StalledRateLabelSlotWidth() noexcept;
    void Render(HWND hwnd) noexcept;
    void SyncHostedControls(HWND hwnd, const std::vector<TaskSnapshot>& snapshot) noexcept;
    // popup-b#1: adds, keeps or retires hosted controls by identity; never resets interaction state.
    void ReconcileHostedControls(HWND hwnd, bool rightToLeft) noexcept;
    // Task owning the focused hosted control; 0 for no focus or a global control.
    [[nodiscard]] uint64_t FocusedHostedTaskId() const noexcept;
    // x-perf#4: every changed Name/HelpText republishes the whole UIA snapshot, so live text is
    // published at most once per interval per control; a new control or fixture publishes at once.
    void PublishHostedAccessibleText(DxUi::Control& control, std::wstring name, std::wstring helpText, bool liveText, ULONGLONG now);
    void OpenPathDetail(HWND hwnd, std::wstring_view automationId) noexcept;
    void ClosePathDetail(HWND hwnd, bool restoreFocus) noexcept;
    void SyncPathDetail(HWND hwnd) noexcept;
    [[nodiscard]] std::wstring HostedControlText(const PopupHitTest& hit) const;
    [[nodiscard]] std::wstring HostedControlAccessibleName(const PopupHitTest& hit) const;
    [[nodiscard]] bool HandleHostedKeyDown(HWND hwnd, UINT virtualKey, LPARAM keyData, bool systemKey) noexcept;
    [[nodiscard]] bool SubmitPublishedConflictEscapeAction(HWND hwnd) noexcept;
    void UpdateLastPopupRect(HWND hwnd) noexcept;
    void UpdateCaptionStatus(HWND hwnd, const std::vector<TaskSnapshot>& snapshot) noexcept;
    void PaintCaptionStatusGlyph(HWND hwnd) noexcept;
    bool EnsureTaskbarList() noexcept;
    void UpdateTaskbarProgress(HWND hwnd) noexcept;
    void ApplyTaskbarProgress(HWND hwnd, uint32_t state, uint64_t completed, uint64_t total) noexcept;
    void ClearTaskbarProgress(HWND hwnd) noexcept;

    PopupHitTest HitTest(float x, float y) const noexcept;
    std::optional<PopupMenuAnchor> ResolveButtonMenuAnchor(HWND hwnd, const PopupHitTest& hit, DxUi::ContextMenuRootVerticalPlacement placement) const noexcept;
    void Invalidate(HWND hwnd) const noexcept;

    LRESULT OnActivatedHit(HWND hwnd, const PopupHitTest& hit) noexcept;
    [[nodiscard]] bool IsTaskDecisionAvailable(uint64_t taskId) const noexcept;
    [[nodiscard]] bool HasCancelableTasks() const noexcept;
    void BeginConflictInspection(HWND hwnd, const PopupHitTest& hit) noexcept;
    [[nodiscard]] std::wstring ConflictInspectionResultText(const TaskSnapshot& task) const;

    bool ConfirmCancelAll(HWND hwnd) noexcept;
    void ShowQueueModeMenu(HWND hwnd) noexcept;
    void ShowFooterOptionsMenu(HWND hwnd) noexcept;
    void ShowSpeedLimitMenu(HWND hwnd, uint64_t taskId) noexcept;
    bool ShowCustomSpeedLimitPromptForTask(HWND hwnd, uint64_t requestedTaskId) noexcept;
    bool SubmitCompletedOverflowAction(HWND hwnd, uint64_t taskId, uint32_t action, bool openExportAfterWrite) noexcept;
    void ShowCompletedOverflowMenu(HWND hwnd, uint64_t taskId) noexcept;
    bool SubmitConflictPrimaryAction(HWND hwnd, uint64_t taskId, uint32_t rawAction, uint64_t generation) noexcept;
    bool SubmitConflictOverflowAction(HWND hwnd, uint64_t taskId, uint32_t rawAction, uint64_t generation) noexcept;
    void ShowConflictOverflowMenu(HWND hwnd, uint64_t taskId, uint64_t generation) noexcept;

    LRESULT OnCreate(HWND hwnd) noexcept;
    LRESULT OnPaint(HWND hwnd) noexcept;
    LRESULT OnThemeChanged(HWND hwnd) noexcept;
    LRESULT OnNcDestroy(HWND hwnd) noexcept;
    LRESULT OnSize(HWND hwnd, WPARAM sizeType, UINT width, UINT height) noexcept;
    LRESULT OnDpiChanged(HWND hwnd, UINT newDpi, const RECT& suggested) noexcept;
    LRESULT OnGetMinMaxInfo(HWND hwnd, MINMAXINFO* info) noexcept;
    LRESULT OnMove(HWND hwnd) noexcept;
    LRESULT OnTimer(HWND hwnd, UINT_PTR timerId) noexcept;
    LRESULT OnEnterSizeMove(HWND hwnd) noexcept;
    LRESULT OnExitSizeMove(HWND hwnd) noexcept;
    LRESULT OnVScroll(HWND hwnd, UINT request) noexcept;
    LRESULT OnMouseWheel(HWND hwnd, int delta, POINT screenPoint) noexcept;
    LRESULT OnClose(HWND hwnd) noexcept;
    LRESULT OnShowWindow(HWND hwnd, bool showing) noexcept;
    void EnsureHostedFocusVisible(HWND hwnd, DxUi::Control* control) noexcept;
    // C4: graphics-independent failure surface (native children; no Direct2D, DirectWrite, or DxUi).
    bool EnterFailureSurface(HWND hwnd) noexcept;
    // Q6: attaches the DxUi host, theme, root and callbacks; false leaves the host detached.
    bool TryAttachHost(HWND hwnd) noexcept;
    // Re-attaches the host from the failure surface; on success the native children are gone.
    bool TryLeaveFailureSurface(HWND hwnd) noexcept;
    // An attached host whose frame cannot be prepared is retried with bounded backoff; true
    // when repeated failures (or a detached host) selected the failure surface instead.
    bool OnFramePrepareFailed(HWND hwnd, HRESULT hr) noexcept;
    // x-arch#2: measurement and fit failures degrade the layout; they never select the failure surface.
    void ReportDegradedLayout(uint32_t step, HRESULT hr) noexcept;
    void LayoutFailureSurface(HWND hwnd) noexcept;
    void RefreshFailureSurfaceText() noexcept;
    bool HandleFailureSurfaceKeyDown(HWND hwnd, HWND focused, UINT virtualKey, LPARAM keyData) noexcept;
    bool HandleFailureSurfaceKeyUp(HWND focused, UINT virtualKey) noexcept;
    static LRESULT CALLBACK FailureSurfaceButtonProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT OnNcPaint(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept;
    LRESULT OnNcActivate(HWND hwnd, WPARAM wParam, LPARAM lParam) noexcept;
#ifdef ENABLE_TESTS
    LRESULT OnSelfTestInvoke(HWND hwnd, const PopupSelfTestInvoke* payload) noexcept;
    LRESULT OnTaskSnapshotRequest(const PopupTaskSnapshotRequest* request) const noexcept;
    LRESULT OnCaptionGlyphSnapshotRequest(CaptionGlyphDebugSnapshot* snapshot) const noexcept;
    LRESULT OnLayoutSnapshotRequest(PopupLayoutDebugSnapshot* snapshot) const noexcept;
#endif

    UINT _dpi = USER_DEFAULT_SCREEN_DPI;
    SIZE _clientSize{};
    size_t _dispatchDepth = 0u;
    bool _deletePending   = false;

    bool _inSizeMove    = false;
    bool _inThemeChange = false;
#ifdef ENABLE_TESTS
    struct PaintedTaskBodyForSelfTest final
    {
        D2D1_RECT_F card{};
        float bodyBottom = 0.0f;
        D2D1_RECT_F actions{};
    };
    std::unordered_map<uint64_t, PaintedTaskBodyForSelfTest> _paintedTaskBodiesForSelfTest;
    // Outer paint bounds also exist for compact cards, whose body has already been retired.
    std::unordered_map<uint64_t, D2D1_RECT_F> _paintedTaskCardsForSelfTest;
    struct PaintedLiveProgressForSelfTest final
    {
        std::array<PopupLayoutDebugSnapshot::PaintedFileProgress, TaskSnapshot::kMaxInFlightFiles> files{};
        size_t fileCount = 0u;
        std::wstring discoveryText;
        PopupLayoutDebugSnapshot::PaintedRecursiveCounts recursiveCounts{};
        std::wstring transferText;
        D2D1_RECT_F discoveryRect{};
        D2D1_COLOR_F discoveryColor{};
        D2D1_RECT_F transferRect{};
        D2D1_COLOR_F transferColor{};
        uint64_t transferCompletedBytes = 0u;
    };
    std::unordered_map<uint64_t, PaintedLiveProgressForSelfTest> _paintedLiveProgressForSelfTest;
    uint64_t _paintGenerationForSelfTest = 0u;
    struct PaintedStatusChipForSelfTest final
    {
        bool painted = false;
        D2D1_RECT_F chrome{};
        D2D1_RECT_F textRect{};
        std::wstring text;
        std::wstring headerText;
        std::wstring paintedHeaderText;
        D2D1_RECT_F headerRect{};
        wil::com_ptr<IDWriteTextFormat> paintedFormat;
    };
    std::unordered_map<uint64_t, PaintedStatusChipForSelfTest> _paintedStatusChipsForSelfTest;
    // The text painted beside each task's compact meter: its percent, or its verification label.
    std::unordered_map<uint64_t, std::wstring> _paintedCompactMeterLabelsForSelfTest;
    bool _destroyOnNextShowForSelfTest = false;
    std::unordered_map<uint64_t, uint32_t> _graphRowColorMatchCountsForSelfTest;
    std::unordered_map<uint64_t, uint32_t> _graphRowColorMismatchCountsForSelfTest;
#endif

    CaptionStatus _captionStatus = CaptionStatus::None;

    float _scrollY                    = 0.0f;
    float _contentHeight              = 0.0f;
    float _lastAutoSizedContentHeight = 0.0f; // For auto-resize tracking
    float _lastAutoSizedFooterHeight  = 0.0f;
    float _footerHeightDip            = 88.0f;
    size_t _lastTaskCount             = 0; // For auto-resize tracking
    int _maxAutoSizedWindowHeight     = 0; // Sticky max window height (prevents resize "dancing")
    int _scrollPos                    = 0;
    bool _scrollBarVisible            = false;

    D2D1_RECT_F _footerCancelAllRect{};
    D2D1_RECT_F _footerPauseResumeAllRect{};
    D2D1_RECT_F _footerAutoDismissRect{};
    D2D1_RECT_F _footerQueueModeRect{};
    D2D1_RECT_F _footerQueueSegmentRect{};
    D2D1_RECT_F _footerParallelSegmentRect{};
    D2D1_RECT_F _footerDensityRect{};
    D2D1_RECT_F _footerOptionsRect{};
    D2D1_RECT_F _footerDetailsToggleRect{};
    D2D1_RECT_F _footerAggregateProgressRect{};
    D2D1_RECT_F _footerSummaryRect{};
    D2D1_RECT_F _listViewportRect{};

    std::vector<PopupButton> _buttons;

    // Bound to the painted immutable prompt, never a later same-task decision.
    struct PublishedInspectionBinding
    {
        uint64_t taskId           = 0u;
        uint64_t promptGeneration = 0u;
        std::shared_ptr<const FileOperations::ConflictInspectionContext> context;
    };
    std::vector<PublishedInspectionBinding> _publishedInspectionContexts;
    struct PublishedConflictKeyboardBinding
    {
        uint64_t taskId       = 0u;
        uint64_t generation   = 0u;
        uint8_t defaultAction = 0u;
        uint8_t escapeAction  = 0u;
        uint8_t bucket        = 0u;
        bool actionable       = false;
    };
    // One entry per current task, including tasks without a decision. Focus changes
    // consult only the last published generation and never rebuild a task snapshot.
    std::vector<PublishedConflictKeyboardBinding> _publishedConflictKeyboardBindings;
    [[nodiscard]] const PublishedConflictKeyboardBinding* SyncHostedConflictKeyboardBindings(uint64_t fallbackTaskId = 0u) noexcept;
    bool _syncingHostedControls = false;
    struct InspectionPresentation
    {
        PopupHitTest invokingHit;
        std::shared_ptr<const FileOperations::ConflictInspectionContext> context;
        UINT messageId = 0u;
        bool pending   = false;
    };
    std::optional<InspectionPresentation> _inspectionPresentation;
    struct InspectionReturn
    {
        PopupHitTest invokingHit;
        std::shared_ptr<const FileOperations::ConflictInspectionContext> context;
        int scrollPos     = 0;
        bool restoreFocus = false;
    };
    std::optional<InspectionReturn> _inspectionReturn;
    // Completion callbacks outlive HWNDs; reset before detach to reject HWND reuse.
    std::shared_ptr<void> _inspectionLifetime;

    DxUi::WindowHost _controlsHost;
    DxUi::Panel* _controlsRoot        = nullptr;
    DxUi::ScrollPanel* _controlsBody  = nullptr;
    DxUi::Panel* _pathDetailPanel     = nullptr;
    DxUi::Label* _pathDetailTitle     = nullptr;
    DxUi::TextField* _pathDetailField = nullptr;
    DxUi::Button* _pathDetailCopy     = nullptr;
    DxUi::Button* _pathDetailClose    = nullptr;
    std::wstring _pathDetailId;
    std::wstring _pathDetailReturnId;
    std::wstring _pathDetailLogicalFocusId;
    // Owning task of the open detail's region; closing keeps focus on that task.
    uint64_t _pathDetailTaskId          = 0u;
    bool _pathDetailNeedsFocus          = false;
    bool _failureSurface                = false;
    bool _failureSurfaceReady           = false;
    bool _failureSurfaceInitializing    = false;
    bool _failureSurfaceFailureReported = false;
    bool _failureCancelAllEnterArmed    = false;
    HWND _failureText                   = nullptr;
    HWND _failureCancelAll              = nullptr;
    HWND _failureClose                  = nullptr;
    UINT _failureFontDpi                = 0u;
    wil::unique_hfont _failureFont;
    std::wstring _failureTextValue;
    // Q6 recovery state: consecutive unprepared frames and the backoff for the next attempt;
    // re-attach attempts from the failure surface back off until a frame is prepared again.
    uint32_t _framePrepareFailureCount      = 0u;
    ULONGLONG _framePrepareFirstFailureTick = 0u;
    ULONGLONG _framePrepareRetryTick        = 0u;
    uint32_t _hostReattachAttemptCount      = 0u;
    ULONGLONG _hostReattachRetryTick        = 0u;
    uint32_t _degradedLayoutReportedSteps   = 0u;
    std::vector<DxUi::Button*> _hostedButtons;
    std::vector<PopupHitTest> _hostedButtonHits;
    // Parent panel of each hosted button and region (the root, the body or a comparison panel).
    std::vector<DxUi::Panel*> _hostedButtonParents;
    std::vector<DxUi::Panel*> _hostedTooltipParents;
    struct RetiredHostedControl final
    {
        enum class Role : uint8_t
        {
            Panel,
            Button,
            Progress,
            Graph,
            GraphStallBand,
            ProgressVerificationMarks,
            Region,
            PathRegion,
        };
        DxUi::Control* control = nullptr;
        DxUi::Panel* parent    = nullptr;
        Role role              = Role::Panel;
        PopupHitTest hit{}; // Buttons only.
    };
    // Retired while a mouse capture was active: hidden (a retired button never acts), shown again when
    // its identity is published again, removed once no capture is active.
    std::vector<RetiredHostedControl> _retiredHostedControls;
    std::unordered_map<const DxUi::Control*, ULONGLONG> _hostedAccessibleTextTicks;
    bool _hostedAccessibleTextFlush   = false;
    bool _hostedAccessibleTextPending = false;
    // A focus restore after controls were retired must not scroll the list (popup-b#3).
    bool _restoringHostedFocus                   = false;
    uint64_t _recycleEscalationFocusedTaskId     = 0u;
    uint64_t _recycleEscalationFocusedGeneration = 0u;
    uint64_t _conflictEscapeTaskId               = 0u;
    uint64_t _conflictEscapeGeneration           = 0u;
    uint8_t _conflictEscapeAction                = 0u;
    std::vector<HostedProgressDescriptor> _hostedProgressDescriptors;
    std::vector<DxUi::ProgressBar*> _hostedProgressBars;
    std::vector<ScalarAnimation> _hostedProgressAnimations;
    std::vector<HostedGraphDescriptor> _hostedGraphDescriptors;
    std::vector<GraphLabelLayout> _graphLabelLayouts;
    std::vector<CardInformationLayout> _cardInformationLayouts;
    wil::com_ptr<IDWriteTextFormat> _stalledRateSlotFormat; // the format _stalledRateSlotWidth was measured with
    float _stalledRateSlotWidth = 0.0f;
    std::vector<DxUi::ThroughputGraph*> _hostedGraphs;
    // I35 R10.2: the stall-band overlay painted over each graph (same index as _hostedGraphs).
    std::vector<DxUi::Control*> _hostedGraphStallBands;
    // I35 (owner report 2026-10-13): the verification-half marks painted over each bar (same index as _hostedProgressBars).
    std::vector<DxUi::Control*> _hostedProgressVerificationMarks;
    std::vector<HostedTooltipDescriptor> _hostedTooltipDescriptors;
    std::vector<DxUi::Control*> _hostedTooltipRegions;
    std::vector<uint64_t> _hostedTooltipComparisonIds;
    std::vector<uint64_t> _hostedTooltipTaskIds;
    std::unordered_map<uint64_t, TaskSnapshot::StatusKind> _announcedTaskStatuses;
    std::unordered_map<uint64_t, bool> _announcedConflictMetadataLoading;
    bool _hostedRightToLeft                        = false;
    bool _hostedAccessibilityInitialized           = false;
    uint64_t _hostedAccessibilityNotificationCount = 0u;
    PopupHostedTreeDebugCounters _hostedTreeCounters;

    std::unordered_map<uint64_t, RateHistory> _rates;
    std::unordered_map<uint64_t, bool> _collapsedTasks;
    std::unordered_map<uint64_t, bool> _throughputExpandedTasks;
    bool _preserveWindowSizeForThroughputChange = false;
    std::unordered_set<uint64_t> _compactExpandedTasks;
    std::unordered_map<uint64_t, TaskHeightAnimation> _taskHeightAnimations;
    std::unordered_map<uint64_t, ScalarAnimation> _taskDisclosureAnimations;
    ScalarAnimation _completedGroupVisibilityAnimation;
    ScalarAnimation _footerDisclosureAnimation;
    bool _completedGroupExpanded = true;
    // Task IDs the last published Clear completed header covered; a later completion stays.
    std::vector<uint64_t> _completedGroupClearTaskIds;
    // I35 R9.5: the tasks the last confirmed Cancel all requested; cleared once they all finished.
    std::vector<uint64_t> _bulkCancelCohort;
    bool _reducedMotion = false;
    std::wstring _footerNewTasksText;
    std::wstring _footerQueueText;
    std::wstring _footerParallelText;
    std::wstring _footerAutoDismissOnText;
    std::wstring _footerAutoDismissOffText;
    bool _footerAutoDismissLabelVisible = false;
#ifdef ENABLE_TESTS
    wchar_t _debugLastDrawnCheckboxGlyph = L'\0';
    uint64_t _debugNativeTabCount        = 0u;
    uint64_t _debugMouseMoveCount        = 0u;
    POINT _debugLastMouseMove{};
    uint64_t _debugHostedActionCount                 = 0u;
    uint64_t _debugFailureSurfaceTabCount            = 0u;
    uint64_t _debugFailureSurfaceReentrantEntryCount = 0u;
    bool _debugFailureSurfacePinned                  = false;
    uint64_t _debugFocusChangeCount                  = 0u;
    std::wstring _debugPreviousFocusTargetId;
    std::wstring _debugLastFocusTargetId;
    bool _debugRetainRetiredControls = false;
    std::vector<std::unique_ptr<DxUi::Control>> _debugRetiredControlQuarantine;
    uint64_t _debugDetachedControlInputCount = 0u;
#endif

    std::optional<RECT> _footerOnlyRestoreWindowRect;
    bool _footerOnlyRestorePending = false;
    bool _autoResizePending        = false;
    bool _autoResizeAnimating      = false;
    RECT _autoResizePendingTargetRect{};
    RECT _autoResizeAnimationStartRect{};
    RECT _autoResizeAnimationTargetRect{};
    ULONGLONG _autoResizePendingDueTick     = 0;
    ULONGLONG _autoResizeAnimationStartTick = 0;
    ULONGLONG _uiMotionTimerDueTick         = 0;
    UINT _telemetryTimerIntervalMs          = 0u;
    wil::com_ptr<ID2D1Factory> _d2dFactory;
    wil::com_ptr<IDWriteFactory> _dwriteFactory;
    wil::com_ptr<ID2D1DeviceContext> _target;
    wil::com_ptr<ID2D1CommandList> _preparedBackground;
    wil::com_ptr<ID2D1DCRenderTarget> _captionGlyphTarget;

    wil::com_ptr<IDWriteTextFormat> _headerFormat;
    wil::com_ptr<IDWriteTextFormat> _bodyFormat;
    wil::com_ptr<IDWriteTextFormat> _smallFormat;
    wil::com_ptr<IDWriteTextFormat> _graphTrailingLabelFormat;
    wil::com_ptr<IDWriteTextFormat> _buttonFormat;
    wil::com_ptr<IDWriteTextFormat> _buttonSmallFormat;
    wil::com_ptr<IDWriteTextFormat> _graphOverlayFormat;
    wil::com_ptr<IDWriteTextFormat> _statusIconFormat;
    wil::com_ptr<IDWriteTextFormat> _statusIconFallbackFormat;
    wil::com_ptr<IDWriteTextFormat> _captionGlyphFormat;
    wil::com_ptr<IDWriteTextFormat> _captionGlyphFallbackFormat;

    struct ConflictBodyLayout final
    {
        uint64_t taskId = 0u;
        std::shared_ptr<const FileOperations::ConflictInspectionContext> context;
        D2D1_RECT_F bounds{};
        D2D1_RECT_F viewport{};
        float contentHeightPx    = 0.0f;
        float offsetDip          = 0.0f;
        bool visible             = false;
        DxUi::ScrollPanel* panel = nullptr; // Borrowed from the retained control tree.
    };
    std::vector<ConflictBodyLayout> _conflictBodies;
    std::vector<uint64_t> _hostedConflictBodyIds;
    std::vector<DxUi::ScrollPanel*> _hostedConflictBodyPanels;
    const float _conflictBodyScrollbarDip = DxUi::ScrollPanel{}.GetScrollbarThickness();
    bool _syncingConflictBodies           = false;
    ConflictBodyLayout* FindConflictBody(uint64_t taskId) noexcept;
    const ConflictBodyLayout* FindConflictBody(uint64_t taskId) const noexcept;

    struct ConflictActionsLayout final
    {
        std::array<std::wstring, TaskSnapshot::kMaxConflictActions + 1u> labels;
        std::array<D2D1_RECT_F, TaskSnapshot::kMaxConflictActions + 1u> bounds{};
        wil::com_ptr<IDWriteTextFormat> format;
        std::wstring scopeLabel;
        float scopeHeightDip = 0.0f;
        float widthDip       = 0.0f;
        float heightDip      = 0.0f;
        UINT dpi             = 0u;
        size_t count         = 0u;
        bool rightToLeft     = false;
    };
    // One bounded slot per current snapshot item. Removed tasks release their text/format references.
    std::vector<ConflictActionsLayout> _conflictActionMeasurements;
    HRESULT MeasureConflictActions(const TaskSnapshot& task, float widthPx, bool rightToLeft, ConflictActionsLayout& layout) noexcept;

    // One application-owned discovery action; retain the format identity to avoid stale
    // cache hits after host recreation. Width is logical, so monitor scaling is applied at use.
    wil::com_ptr<IDWriteTextFormat> _discoverySkipMeasuredFormat;
    std::wstring _discoverySkipMeasuredText;
    float _discoverySkipWidthDip = 84.0f;

    struct SpeedLimitWidthMeasurement final
    {
        std::wstring caption;
        float widthDip = 0.0f;
    };
    // Bounded by the current snapshot, never retained by historical task identity.
    // Caption and retained format identity cover limit, locale and font changes.
    std::vector<SpeedLimitWidthMeasurement> _speedLimitWidthMeasurements;
    wil::com_ptr<IDWriteTextFormat> _speedLimitMeasuredFormat;

    struct TerminalSummaryLayout final
    {
        uint64_t taskId = 0u;
        std::wstring text;
        wil::com_ptr<IDWriteTextFormat> format;
        wil::com_ptr<IDWriteTextLayout> layout;
        float widthPx  = 0.0f;
        float heightPx = 0.0f;
    };
    // At most one text layout per current expanded terminal snapshot slot.
    // Hidden/removed tasks release layouts; unchanged paints reuse their metrics.
    std::vector<TerminalSummaryLayout> _terminalSummaries;
    HRESULT PrepareTerminalSummary(size_t slot, const TaskSnapshot& task, float widthPx) noexcept;

    struct ScanStatusLayout final
    {
        TerminalSummaryLayout prepared;
        std::wstring paintedText;
        // x-perf#7: the measured path truncation, keyed on the path and the scaffold with its
        // counter digits normalized, so changing counters only re-format the row.
        std::wstring measuredPath;
        std::wstring measuredScaffold;
        std::wstring paintedPath;
        bool wholeLineAbbreviated = false;
        bool usedThisPaint        = false;
    };
    // One slot per current informational snapshot; never accumulated by historical ID.
    // Full text, width/height and retained format identity invalidate the layout together; the
    // path truncation is measured again only when the path, row or counter digit count changes.
    std::vector<ScanStatusLayout> _scanStatusLayouts;
    uint64_t _scanStatusMeasurementCount = 0u;

    struct TerminalFactsLayout final
    {
        TerminalSummaryLayout prepared;
        std::array<unsigned long, 10> counts{};
        uint64_t languageRevision = 0u;
        uint8_t verificationState = 0u;
        bool startedMutation      = true;
    };
    // Numeric facts are bounded by current expanded snapshots. Zero is absence of
    // recorded evidence, never a complement or a claim of complete success.
    std::vector<TerminalFactsLayout> _terminalFacts;
    HRESULT PrepareTerminalFacts(size_t slot, const TaskSnapshot& task, float widthPx, uint64_t languageRevision) noexcept;

    struct TerminalActionsLayout final
    {
        std::array<std::wstring, 3> labels;
        std::array<D2D1_RECT_F, 3> bounds{};
        wil::com_ptr<IDWriteTextFormat> format;
        float widthDip   = 0.0f;
        float heightDip  = 0.0f;
        UINT dpi         = 0u;
        bool rightToLeft = false;
    };
    // One bounded action layout per currently visible expanded terminal card.
    std::vector<TerminalActionsLayout> _terminalActions;
    HRESULT PrepareTerminalActions(size_t slot, const TaskSnapshot& task, float widthPx, bool rightToLeft) noexcept;

    struct CompletedGroupMeasurement final
    {
        std::wstring clearText;
        std::wstring headingText;
        std::array<uint32_t, 5> counts{};
        wil::com_ptr<IDWriteTextFormat> actionFormat;
        wil::com_ptr<IDWriteTextFormat> headingFormat;
        wil::com_ptr<IDWriteTextFormat> countFormat;
        wil::com_ptr<IDWriteTextLayout> headingLayout;
        FileOperationsCompletedGroupLayout::Layout layout{};
        float widthPx = 0.0f;
        UINT dpi      = 0u;
        bool expanded = false;
    };
    CompletedGroupMeasurement _completedGroupMeasurement;
    HRESULT PrepareCompletedGroup(float widthPx, uint32_t total, const std::array<uint32_t, 5>& counts) noexcept;

    // Fixed Source/Destination/Delete/Verify vocabulary; retain the format identity
    // so DPI/font recreation cannot reuse measurements from a destroyed format.
    wil::com_ptr<IDWriteTextFormat> _pathLabelsMeasuredFormat;
    std::array<std::wstring, 6> _pathLabelsMeasuredText;
    std::array<float, 6> _pathLabelWidthsPx{};
    float MeasurePathLabelWidth(size_t slot, std::wstring_view text) noexcept;

    // Fixed footer vocabulary: use the larger Pause/Resume width to avoid action movement
    // when the engine changes state. No per-task or per-frame measurement storage.
    wil::com_ptr<IDWriteTextFormat> _footerMeasuredFormat;
    std::array<std::wstring, 3> _footerMeasuredLabels;
    std::array<float, 2> _footerActionWidthsDip{{112.0f, 112.0f}};

    // One current summary layout, shared by measurement and paint. Unlike the
    // generic immediate wrapped-text helper, unchanged frames allocate nothing.
    wil::com_ptr<IDWriteTextLayout> _footerSummaryLayout;
    wil::com_ptr<IDWriteTextFormat> _footerSummaryFormat;
    std::wstring _footerSummaryText;
    float _footerSummaryWidthPx         = 0.0f;
    float _footerSummaryHeightPx        = 0.0f;
    float _footerMinimumSummaryHeightPx = 0.0f;

    wil::com_ptr<ID2D1SolidColorBrush> _bgBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _textBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _subTextBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _borderBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _progressBgBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _progressGlobalBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _progressItemBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _progressVerifyBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _checkboxFillBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _checkboxCheckBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _statusOkBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _statusWarningBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _statusErrorBrush;
    D2D1::ColorF _progressItemBaseColor = D2D1::ColorF(D2D1::ColorF::Black);
    wil::com_ptr<ID2D1SolidColorBrush> _graphDynamicBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _buttonBgBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _buttonChromeBrush;
    wil::com_ptr<ID2D1SolidColorBrush> _captionGlyphBrush;
    wil::com_ptr<ITaskbarList3> _taskbarList;
    UINT _captionGlyphDpi                = 0;
    ULONGLONG _taskbarListRetryAfterTick = 0;
    uint32_t _taskbarListAttemptCount    = 0u;
    bool _taskbarButtonReady             = false;
    uint64_t _taskbarUpdateCount         = 0;
    uint32_t _cachedTaskbarState         = 0u;
    uint64_t _cachedTaskbarCompleted     = 0u;
    uint64_t _cachedTaskbarTotal         = 0u;
#ifdef ENABLE_TESTS
    bool _debugIdleSuspensionEnabled         = true;
    uint64_t _telemetryTimerTickCount        = 0u;
    mutable uint64_t _taskSnapshotBuildCount = 0u;
    mutable uint64_t _rateSnapshotBuildCount = 0u;
    uint64_t _rateUpdateCount                = 0u;
    uint64_t _taskbarSuccessCount            = 0u;
    uint32_t _lastTaskbarState               = 0u;
#endif

    int _mouseWheelRemainder = 0;
};

#ifdef ENABLE_TESTS
// The whole-task presentation the card actually renders, exposed so a test can assert the real
// projection of a live task instead of recomputing the rules it is meant to be checking.
struct TaskProgressPresentationDebugSnapshot final
{
    double fraction               = 0.0;
    bool determinate              = false;
    bool provisional              = false;
    bool transferActive           = false;
    bool showDiscoveryActivity    = false;
    bool showTransferProgress     = false;
    bool showTransferMarquee      = false;
    bool showThroughputGraph      = false;
    bool includeInAggregateCohort = false;
};
#endif
} // namespace FileOperationsPopupInternal

class FileOperationsPopup final
{
public:
    static HWND Create(FolderWindow::FileOperationState* fileOps, FolderWindow* folderWindow, HWND ownerWindow, std::weak_ptr<void> hostLifetime) noexcept;
    // UI-thread-only, nonactivating notification after a task/worker or
    // informational transition, always sent outside the state ownership lock.
    static void RefreshTelemetry(HWND popup) noexcept;

private:
    FileOperationsPopup() = delete;
};

#ifdef ENABLE_TESTS
struct FileOperationsSpeedLimitPromptDebugSnapshot
{
    D2D1_RECT_F bodyBounds{};
    float bodyScrollOffsetDip  = 0.0f;
    float bodyContentHeightDip = 0.0f;
    bool bodyScrollable        = false;
    bool fieldFocused          = false;
    bool helpFocused           = false;
    bool hasEditState          = false;
    uintptr_t fieldIdentity    = 0u;
    size_t caretIndex          = 0u;
    std::optional<size_t> selectionAnchorIndex;
    std::wstring labelText;
    std::wstring cancelText;
    D2D1_RECT_F clientBounds{};
    D2D1_RECT_F labelBounds{};
    D2D1_RECT_F fieldBounds{};
    D2D1_RECT_F hintBounds{};
    D2D1_RECT_F okBounds{};
    D2D1_RECT_F cancelBounds{};
    float requiredLabelHeightDip        = 0.0f;
    float requiredHintHeightDip         = 0.0f;
    bool usesDxUiHost                   = false;
    size_t visibleChildWindowCount      = 0u;
    uint64_t initialLimitBytesPerSecond = 0;
    std::wstring text;
    std::wstring hintText;
    std::wstring validationText;
};

[[nodiscard]] bool DebugInvokeFileOperationsPopup(HWND popup, const FileOperationsPopupInternal::PopupSelfTestInvoke& invoke) noexcept;
[[nodiscard]] bool DebugGetFileOperationsPopupTaskSnapshot(HWND popup, uint64_t taskId, FileOperationsPopupInternal::TaskSnapshot& out) noexcept;
[[nodiscard]] bool DebugGetFileOperationsPopupCaptionGlyphSnapshot(HWND popup, FileOperationsPopupInternal::CaptionGlyphDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugGetFileOperationsPopupLayoutSnapshot(HWND popup, FileOperationsPopupInternal::PopupLayoutDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugGetFileOperationsPopupTelemetrySnapshot(HWND popup, FileOperationsPopupInternal::PopupTelemetryDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugSetFileOperationsPopupIdleSuspensionForF6(HWND popup, FolderWindow::FileOperationState* expectedOwner, bool enabled) noexcept;
[[nodiscard]] bool DebugGetFileOperationsPopupUiMotionForF6(HWND popup, FolderWindow::FileOperationState* expectedOwner, bool& visibleMotion) noexcept;
[[nodiscard]] bool DebugUpdateFileOperationsPopupTaskbarProgress(HWND popup) noexcept;
[[nodiscard]] bool DebugBuildFileOperationsPopupGlobalSummarySnapshot(const std::vector<FileOperationsPopupInternal::TaskSnapshot>& tasks,
                                                                      FileOperationsPopupInternal::PopupLayoutDebugSnapshot& out,
                                                                      double displayedBytesPerSecOverride = -1.0,
                                                                      double aggregateEtaSecondsOverride  = -1.0) noexcept;
void DebugFailNextFileOperationsTaskbarListAttempts(unsigned int attempts) noexcept;
void DebugFailNextFileOperationsDxUiHostAttachAttempts(unsigned int attempts) noexcept;
// Fails the next frame preparations (the host's PrimeForShow), one per render attempt.
void DebugFailNextFileOperationsD2DTargetAttempts(unsigned int attempts) noexcept;
// Fails the next popup text measurements (conflict, completed-group and terminal actions).
void DebugFailNextFileOperationsPopupLayoutMeasurements(unsigned int attempts) noexcept;
void DebugFailNextFileOperationsFailureSurfaceChildCreation(unsigned int ordinal) noexcept;
void DebugFailNextFileOperationsFailureSurfaceButtonSetup(unsigned int ordinal) noexcept;
[[nodiscard]] bool DebugEnterFileOperationsPopupFailureSurface(HWND popup) noexcept;
[[nodiscard]] bool DebugCanRestoreFileOperationsPathDetailTextFocus(HWND popup, HWND focus, HWND foreground) noexcept;
[[nodiscard]] bool DebugBuildFileOperationsGraphFairColorWeightSnapshot(FileOperationsPopupInternal::GraphHueWeightDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugBuildFileOperationsGraphFairnessHistorySnapshot(FileOperationsPopupInternal::PopupLayoutDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugFileOperationsStreamColorSlotsRemainUniqueUnderPermutation() noexcept;
[[nodiscard]] bool DebugValidateFileOperationsVerificationPresentation() noexcept;
[[nodiscard]] bool DebugValidateFileOperationsMissingSourcePresentation(std::wstring& failure) noexcept;
[[nodiscard]] bool DebugValidateFileOperationsStoppingPresentation() noexcept;
[[nodiscard]] float DebugComputeFileOperationsTaskCompleteFraction(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept;
void DebugPublishFileOperationsPlannedItemTotalAfterDiscovery(FileOperationsPopupInternal::TaskSnapshot& task) noexcept;
[[nodiscard]] bool DebugFileOperationsTaskHasKnownCompactProgress(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept;
void DebugResolveFileOperationsWholeTaskProgress(const FileOperationsPopupInternal::TaskSnapshot& task,
                                                 FileOperationsPopupInternal::TaskProgressPresentationDebugSnapshot& out) noexcept;
[[nodiscard]] std::wstring DebugBuildFileOperationsTaskHeaderText(const FileOperationsPopupInternal::TaskSnapshot& task, ULONGLONG nowTick);
[[nodiscard]] std::wstring DebugFormatFileOperationsConflictTimestamp(__int64 fileTime) noexcept;
[[nodiscard]] std::wstring DebugFormatFileOperationsDiscoveryProgress(const FileOperationsPopupInternal::TaskSnapshot& task);
[[nodiscard]] std::wstring DebugFormatFileOperationsDiscoveredSoFar(const FileOperationsPopupInternal::TaskSnapshot& task);
[[nodiscard]] DxUi::ButtonVariant DebugResolveFileOperationsHostedButtonVariant(FileOperationsPopupInternal::PopupHitTest::Kind kind) noexcept;
[[nodiscard]] std::wstring DebugFormatFileOperationsSpeedLimitSelectorText(uint64_t bytesPerSecond);
// I35 R10.2: drives the production per-task rate step with synthetic ticks (one observation per tick).
[[nodiscard]] bool DebugRunFileOperationsRateModel(FileSystemOperation operation,
                                                   std::span<const FileOperationsPopupInternal::RateModelDebugTick> ticks,
                                                   std::vector<FileOperationsPopupInternal::RateModelDebugObservation>& out) noexcept;
// The displayed byte rate after 10 s at a steady rate followed by silenceMs without any progress.
[[nodiscard]] double DebugWindowedRateAfterSilence(double steadyBytesPerSecond, ULONGLONG silenceMs) noexcept;
[[nodiscard]] std::wstring DebugFormatFileOperationsRateLabelText(double bytesPerSecond, FileOperationsPopupInternal::RateStallKind stall);
[[nodiscard]] double DebugSmoothEtaSecondsForDisplay(double previousEtaSeconds, double sampleEtaSeconds, ULONGLONG elapsedMs) noexcept;
[[nodiscard]] float DebugEaseFileOperationsGraphLatestPointYForDisplay(float previousY, float targetY, ULONGLONG elapsedMs) noexcept;
[[nodiscard]] float DebugEaseFileOperationsAutoResizeFraction(ULONGLONG elapsedMs, ULONGLONG durationMs) noexcept;
[[nodiscard]] D2D1_RECT_F DebugComputeFileOperationsIndeterminateBarFill(const D2D1_RECT_F& bar, ULONGLONG tick, bool reducedMotion) noexcept;
[[nodiscard]] HWND GetFileOperationsSpeedLimitPromptHandle() noexcept;
void DebugShowFileOperationsSpeedLimitPromptForGallery(HWND owner, const AppTheme& theme) noexcept;
[[nodiscard]] bool DebugGetFileOperationsSpeedLimitPromptSnapshot(FileOperationsSpeedLimitPromptDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugSetFileOperationsSpeedLimitPromptText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugSetFileOperationsSpeedLimitPromptClientHeightLimit(std::optional<float> heightDip) noexcept;
[[nodiscard]] bool DebugSetFileOperationsSpeedLimitPromptSelection(size_t start, size_t end) noexcept;
[[nodiscard]] bool DebugConfirmFileOperationsSpeedLimitPrompt() noexcept;
[[nodiscard]] bool DebugCancelFileOperationsSpeedLimitPrompt() noexcept;
#endif
