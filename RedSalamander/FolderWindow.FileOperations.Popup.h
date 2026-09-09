#pragma once

#include "DxUi/DxUi.h"
#include "FolderWindow.h"

#include <array>
#include <filesystem>
#include <memory>
#include <optional>
#include <shobjidl.h>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <vector>

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
        TaskDestination,
        TaskSpeedLimit,
        TaskShowLog,
        TaskExportIssues,
        TaskCompletedMore,
        TaskConflictToggleApplyToAll,
        TaskConflictAction,
        TaskConflictMore,
        TaskDismiss,
        FooterOptions,
    };

    Kind kind       = Kind::None;
    uint64_t taskId = 0;
    uint32_t data   = 0;
};

#ifdef ENABLE_TESTS
inline constexpr uint32_t kPopupSelfTestDestroyOnNextShowData = 0xD1570001u;

struct PopupSelfTestInvoke
{
    PopupHitTest::Kind kind = PopupHitTest::Kind::None;
    uint64_t taskId         = 0;
    uint32_t data           = 0;
};

#endif

struct PopupButton
{
    D2D1_RECT_F bounds{};
    PopupHitTest hit{};
};

struct PopupMenuAnchor
{
    POINT screenPoint{};
    RedSalamander::DxUi::ContextMenuSessionCallbacks sessionCallbacks{};
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
    };

    enum class Kind : uint8_t
    {
        FileOperation,
        Informational,
    };

    Kind kind = Kind::FileOperation;
    FolderWindow::InformationalTaskUpdate informational{};

    struct InFlightFileSnapshot
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
    uint8_t lifecyclePhase        = 0xFFu; // FileOperations::TaskLifecyclePhase; unknown for informational rows.

    unsigned long totalItems                   = 0;
    unsigned long completedItems               = 0;
    uint64_t totalBytes                        = 0;
    uint64_t completedBytes                    = 0;
    uint64_t itemTotalBytes                    = 0;
    uint64_t itemCompletedBytes                = 0;
    bool verificationRequested                 = false;
    bool verificationActive                    = false;
    uint64_t verificationTotalBytes            = 0;
    uint64_t verificationCompletedBytes        = 0;
    uint64_t verificationItemTotalBytes        = 0;
    uint64_t verificationItemCompletedBytes    = 0;
    unsigned long verifiedItemCount            = 0;
    unsigned long verificationProblemItemCount = 0;
    uint8_t verificationState                  = 0u; // FileOperations::VerificationState

    // Best-effort top-level type breakdown while traversal totals are still open.
    unsigned long completedFiles   = 0;
    unsigned long completedFolders = 0;

    std::wstring currentSourcePath;
    std::wstring currentDestinationPath;

    std::array<InFlightFileSnapshot, kMaxInFlightFiles> inFlightFiles{};
    size_t inFlightFileCount = 0;

    struct ConflictPromptSnapshot
    {
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
        ItemMetadata sourceMetadata;
        ItemMetadata destinationMetadata;
        std::array<uint8_t, kMaxConflictActions> actions{};
        size_t actionCount = 0;
        std::array<uint8_t, kMaxConflictActions> primaryActions{};
        size_t primaryActionCount = 0;
        std::array<uint8_t, kMaxConflictActions> overflowActions{};
        size_t overflowActionCount = 0;
        uint8_t defaultAction      = 0u;
        uint8_t escapeAction       = 0u;
        bool applyToAllEligible    = false;
        bool applyToAllChecked     = false;
        bool skipAllEligible       = false;
        bool buttonsPublishable    = false;
        bool retryFailed           = false;
        unsigned int attemptCount  = 0u; // C2: attempts already made; the question says so
        bool factItemCountKnown    = false;
        uint64_t factItemCount     = 0;
        bool factBytesKnown        = false;
        uint64_t factBytes         = 0;
        uint8_t overlapProblem     = 0u; // R4-A02-1
        uint64_t overlapTaskId     = 0;
        std::wstring consentDetail; // C1: what a permanent-delete confirmation deletes
    };

    ConflictPromptSnapshot conflict{};

    uint64_t desiredSpeedLimitBytesPerSecond       = 0;
    uint64_t effectiveSpeedLimitBytesPerSecond     = 0;
    bool autoConcurrencyUsed                       = false;
    uint32_t autoConcurrencyStorageKind            = FILESYSTEM_STORAGE_UNKNOWN;
    uint32_t autoConcurrencyDestinationStorageKind = FILESYSTEM_STORAGE_UNKNOWN;
    unsigned int autoTunedConcurrency              = 0;
    unsigned int effectiveConcurrencyBudget        = 0;

    bool finished              = false;
    HRESULT resultHr           = S_OK;
    unsigned long warningCount = 0;
    unsigned long errorCount   = 0;
    std::wstring lastDiagnosticMessage;
    std::wstring resultSummary;
    bool hasIndeterminateResult             = false;
    bool interruptedMoveNotice              = false;
    uint64_t interruptedOperationId         = 0u;
    bool clipboardMoveAdmission             = false;
    bool clipboardMoveConsumed              = false;
    bool hasSourceActionPaths               = false;
    bool exactRetainedSourceActionAvailable = false;
    unsigned long retainedSourceCount       = 0;
    unsigned long retainedSourceNativeCount = 0;
    unsigned long unknownSourceCount        = 0;
    StatusKind statusKind                   = StatusKind::None;

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

struct PopupLayoutDebugSnapshot
{
    uint64_t taskId = 0;
    bool found      = false;

    size_t visibleButtonCount                      = 0u;
    size_t footerVisibleButtonCount                = 0u;
    bool hasVisibleButtonOverlap                   = false;
    bool taskHasVisibleButtonOverlap               = false;
    bool conflictApplyToAllVisible                 = false;
    bool conflictMoreVisible                       = false;
    bool completedDiagnosticsMoreVisible           = false;
    bool completedDiagnosticsMoreButtonRectVisible = false;
    bool completedShowLogVisible                   = false;
    bool completedExportIssuesVisible              = false;
    bool completedFailedItemsActionVisible         = false;
    bool completedOpenDestinationActionVisible     = false;
    bool completedRevealDestinationActionVisible   = false;
    bool completedDismissVisible                   = false;
    bool taskToggleCollapseVisible                 = false;
    bool taskStartNowVisible                       = false;
    bool taskQueueMoveUpVisible                    = false;
    bool taskQueueMoveDownVisible                  = false;
    bool taskPauseVisible                          = false;
    bool taskCancelVisible                         = false;
    bool taskSkipVisible                           = false;
    bool taskSpeedLimitVisible                     = false;
    bool taskSpeedLimitUsesSelectorChrome          = false;
    bool taskSpeedLimitShowsCurrentValue           = false;
    bool taskDiscoveryAheadActive                  = false;
    bool taskDiscoveryIndicatorVisible             = false;
    bool taskTransferProgressVisible               = false;
    bool taskTransferProgressDeterminate           = false;
    bool taskTransferProgressProvisional           = false;
    bool taskTransferProgressMarqueeVisible        = false;
    bool footerPauseResumeAllVisible               = false;
    bool footerPauseResumeAllPauses                = false;
    bool footerQueueModeSegmentedVisible           = false;
    bool footerQueueModeSelectorVisible            = false;
    bool footerQueueModeUsesSelectorChrome         = false;
    bool footerQueueModeSelectorHitTargetActive    = false;
    bool footerQueueModeIsParallel                 = false;
    bool footerQueueHitTargetActive                = false;
    bool footerParallelHitTargetActive             = false;
    bool footerAutoDismissVisible                  = false;
    bool footerAutoDismissLabelVisible             = false;
    bool footerAutoDismissEnabled                  = false;
    wchar_t checkboxGlyph                          = L'\0';
    bool footerDensityToggleVisible                = false;
    bool footerDensityHitTargetActive              = false;
    bool footerOptionsVisible                      = false;
    bool footerOptionsHitTargetActive              = false;
    bool popupCompactDensity                       = false;
    bool footerAggregateProgressVisible            = false;
    bool footerAggregateProgressDeterminate        = false;
    uint32_t footerOpenDiscoveryTasks              = 0u;
    bool footerOnly                                = false;
    bool footerDetailsToggleVisible                = false;
    bool footerDetailsToggleRightAligned           = false;
    bool footerDetailsToggleUsesDisclosureChrome   = false;
    bool highContrastEnabled                       = false;
    bool reducedMotionEnabled                      = false;
    bool layoutMotionActive                        = false;
    bool autoResizeMotionActive                    = false;
    bool selectedTaskHeightMotionActive            = false;
    bool completedGroupLayoutMotionActive          = false;
    bool autoResizeAnimationEnabled                = false;
    bool footerQueueModeAnimationEnabled           = false;
    bool graphStatusAnimationEnabled               = false;
    bool conflictStackedPathRows                   = false;
    bool conflictSourceMetadataVisible             = false;
    bool conflictDestinationMetadataVisible        = false;
    bool conflictMetadataSizeCompareVisible        = false;
    bool conflictMetadataDateCompareVisible        = false;
    bool conflictPromptPrecedesActions             = false;
    bool conflictWaitingForDecisionVisible         = false;
    bool conflictDecisionDetailsLoadingVisible     = false;
    bool conflictPromptVisibleWithoutClipping      = false;
    bool conflictContextVisibleWithoutClipping     = false;
    bool conflictDiscoveryIndicatorVisible         = false;
    bool conflictTransferProgressVisible           = false;
    bool lastNoteVisibleWithoutClipping            = false;
    bool taskDuplicateUnderGraphItemBarVisible     = false;
    bool taskStatusStripeVisible                   = false;
    bool taskStatusChipVisible                     = false;
    bool taskVerificationBadgeVisible              = false;
    bool taskRetainedSourceBadgeVisible            = false;
    bool taskUnknownSourceBadgeVisible             = false;
    bool taskTerminalAxisBadgesColorBlindSafe      = false;
    bool taskStatusGlyphSignalVisible              = false;
    bool taskStatusTextSignalVisible               = false;
    bool taskStatusColorBlindSafeEncoding          = false;
    bool taskCollapsed                             = false;
    bool taskCompactRow                            = false;
    bool taskAutoCollapsedOnCompletion             = false;
    bool taskCompactProgressVisible                = false;
    bool taskHiddenByCompletedGroup                = false;
    bool completedGroupVisible                     = false;
    bool completedGroupExpanded                    = false;
    bool completedGroupToggleVisible               = false;
    bool completedGroupClearVisible                = false;
    bool completedGroupClearTooltipEmpty           = false;
    bool completedGroupAnimationEnabled            = false;
    bool taskCardAnimationEnabled                  = false;
    bool hostedProgressAnimationEnabled            = false;
    bool usesDxUiHost                              = false;
    size_t hostedActionControlCount                = 0u;
    size_t hostedProgressControlCount              = 0u;
    size_t hostedGraphControlCount                 = 0u;
    size_t hostedTooltipRegionCount                = 0u;
    bool completedGroupBadgeTooltipsInformative    = false;
    uint64_t hostedAccessibilityNotificationCount  = 0u;
    std::wstring hostedFocusedAutomationId;
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
    bool taskbarButtonReady                 = false;
    bool taskbarListAvailable               = false;
    bool taskbarListRetryPending            = false;
    uint32_t taskbarListAttemptCount        = 0u;
    uint64_t taskbarListRetryDelayMs        = 0;
    D2D1_RECT_F footerQueueSegmentRect{};
    D2D1_RECT_F footerParallelSegmentRect{};
    D2D1_RECT_F footerSummaryRect{};
    D2D1_RECT_F completedDiagnosticsMoreButtonRect{};
    size_t conflictPrimaryActionCount  = 0u;
    size_t conflictOverflowActionCount = 0u;
    uint8_t conflictDefaultAction      = 0u;
    uint8_t conflictEscapeAction       = 0u;
    bool conflictButtonsPublishable    = false;
    bool conflictSkipAllEligible       = false;
    size_t completedVisibleActionCount = 0u;
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
    uint32_t graphMultiHueBucketCount          = 0u;
    uint32_t graphSingleHueBucketCount         = 0u;
    uint32_t graphDistinctHueCount             = 0u;
    double graphMinHueShare                    = 0.0; // per-hue share of summed multi-hue bucket weight
    double graphMaxHueShare                    = 0.0;
    uint32_t graphDebugAccumulateCalls         = 0u;
    uint32_t graphDebugLastPending             = 0u;
    uint32_t graphDebugMaxStreams              = 0u;
    uint32_t graphRowColorMatchCount           = 0u;
    uint32_t graphRowColorMismatchCount        = 0u;
    bool graphCurrentBandwidthLineVisible      = false;
    bool graphCurrentBandwidthLabelVisible     = false;
    bool graphEtaLabelVisible                  = false;
    bool graphEtaLabelRightAligned             = false;
    bool taskInlineSpeedRowVisible             = false;
    bool taskInlineEtaRowVisible               = false;
    float taskExpandedBaseHeightDip            = 0.0f;
    double graphCurrentBandwidthBytesPerSecond = 0.0;
    bool failureSurfaceActive                  = false; // C4: native text + Cancel all + Close, no D2D/DxUi
    size_t failureSurfaceChildCount            = 0u;
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
#endif

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
    uint64_t verificationTotalBytes     = 0;
    uint64_t verificationCompletedBytes = 0;
    uint64_t verificationReadBytes      = 0;
    std::wstring currentSourcePath;
    ULONGLONG lastProgressCallbackTick = 0;
    ULONGLONG progressStateChangeTick  = 0;
    std::array<InFlightStreamSnapshot, TaskSnapshot::kMaxInFlightFiles> inFlightFiles{};
    size_t inFlightFileCount = 0;
    bool started             = false;
    bool paused              = false;
    bool waitingForOthers    = false;
    bool waitingInQueue      = false;
    bool queuePaused         = false;
    bool discoveryClosed     = false;
    bool finished            = false;
};

struct RateHistory
{
    static constexpr size_t kMaxSamples             = 180u; // ~18s @ 100ms
    static constexpr size_t kMaxHueWeightsPerSample = TaskSnapshot::kMaxInFlightFiles;

    struct HueWeight
    {
        float hue         = -1.0f;
        double weight     = 0.0;
        uint8_t colorSlot = RedSalamander::DxUi::ThroughputGraphSample::kInvalidColorSlot;
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
        uint8_t assignedColorSlot = RedSalamander::DxUi::ThroughputGraphSample::kInvalidColorSlot;
    };

    std::array<float, kMaxSamples> samples{};
    std::array<float, kMaxSamples> verificationSamples{};
    std::array<float, kMaxSamples> hues{}; // Per-sample hue (0-360) for rainbow mode
    std::array<std::array<HueWeight, kMaxHueWeightsPerSample>, kMaxSamples> hueWeights{};
    std::array<uint8_t, kMaxSamples> hueWeightCounts{};
    size_t count      = 0;
    size_t writeIndex = 0;
    bool initialized  = false;

    ULONGLONG lastStateChangeTick      = 0;
    ULONGLONG resumeTick               = 0;
    ULONGLONG lastProgressCallbackTick = 0;
    ULONGLONG lastDisplaySampleTick    = 0;
    uint64_t lastBytes                 = 0;
    uint64_t lastVerificationBytes     = 0;
    unsigned long lastItems            = 0;

    ULONGLONG pendingBucketMs                  = 0;
    double pendingWeightedSampleMs             = 0.0;
    double pendingWeightedVerificationSampleMs = 0.0;
    float pendingHue                           = -1.0f;
    std::array<HueWeight, kMaxHueWeightsPerSample> pendingHueWeights{};
    size_t pendingHueWeightCount = 0;
    std::array<StreamProgress, TaskSnapshot::kMaxInFlightFiles> streamProgress{};
    size_t streamProgressCount = 0;

    // Diagnostics for the fairness selftest (cheap counters, always maintained).
    uint32_t debugAccumulateCalls  = 0;
    uint32_t debugLastPendingCount = 0;
    uint32_t debugMaxStreamsSeen   = 0;

    double smoothedBytesPerSec              = 0.0;
    double smoothedItemsPerSec              = 0.0;
    double displayedBytesPerSec             = 0.0;
    double smoothedVerificationBytesPerSec  = 0.0;
    double displayedVerificationBytesPerSec = 0.0;
    double displayedItemsPerSec             = 0.0;
    double smoothedEtaSeconds               = 0.0;
    bool hasSmoothedEta                     = false;
    // D2-A08: while discovery is open, the remaining time over the workload discovered so far.
    double provisionalEtaSeconds = 0.0;
    bool hasProvisionalEta       = false;
};

class FileOperationsPopupState final
{
public:
    FolderWindow::FileOperationState* fileOps = nullptr;
    FolderWindow* folderWindow                = nullptr;
    std::weak_ptr<void> hostLifetime;

    static LRESULT CALLBACK WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    LRESULT WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
    void SetExpandedPlacementForRestore(const RECT& rect) noexcept
    {
        _footerOnlyRestoreWindowRect = rect;
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
    };

    struct HostedGraphDescriptor final
    {
        D2D1_RECT_F bounds{};
        uint64_t taskId = 0u;
        std::vector<RedSalamander::DxUi::ThroughputGraphSample> samples;
        std::vector<double> verificationSamples;
        std::wstring overlayText;
        std::wstring currentBandwidthText;
        std::wstring currentBandwidthTrailingText;
        double currentBandwidthBytesPerSecond = 0.0;
        bool rainbowMode                      = false;
        bool perStreamBands                   = false;
    };

    struct HostedTooltipDescriptor final
    {
        D2D1_RECT_F bounds{};
        std::wstring text;
        std::wstring automationId;
    };

    struct ScalarAnimation final
    {
        double displayed    = 0.0;
        double start        = 0.0;
        double target       = 0.0;
        ULONGLONG startTick = 0u;
        bool initialized    = false;
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
    bool IsTaskCollapsedForDisplay(uint64_t taskId, bool compactDensity) const noexcept;
    void ToggleTaskCollapsed(uint64_t taskId, bool compactDensity) noexcept;
    void AutoCollapseCompletedTasks(const std::vector<TaskSnapshot>& snapshot) noexcept;
    void CleanupCollapsedTasks(const std::vector<TaskSnapshot>& snapshot) noexcept;
    [[nodiscard]] float ResolveAnimatedTaskHeight(uint64_t taskId, float targetHeight, ULONGLONG tick, bool reducedMotion) noexcept;
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

    void LayoutChrome(float width, float height, bool showPauseResumeAll, bool showCancelAll) noexcept;
    void UpdateScrollBar(HWND hwnd, float viewH, float contentH) noexcept;
    void AutoResizeWindow(HWND hwnd, float desiredContentHeight, size_t taskCount, bool footerOnly, bool reducedMotion) noexcept;
    void ArmUiMotionTimer(HWND hwnd) noexcept;

    void DrawDxUiButtonChrome(const PopupButton& button,
                              IDWriteTextFormat* format,
                              std::wstring_view text,
                              RedSalamander::DxUi::ButtonVariant variant) noexcept;
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
                            bool perStreamBands) noexcept;
    void Render(HWND hwnd) noexcept;
    void SyncHostedControls(HWND hwnd, const std::vector<TaskSnapshot>& snapshot) noexcept;
    [[nodiscard]] std::wstring HostedControlText(const PopupHitTest& hit) const;
    [[nodiscard]] std::wstring HostedControlAccessibleName(const PopupHitTest& hit) const;
    [[nodiscard]] bool HandleHostedKeyDown(HWND hwnd, UINT virtualKey) noexcept;
    [[nodiscard]] bool SubmitPublishedConflictEscapeAction(HWND hwnd) noexcept;
    void UpdateLastPopupRect(HWND hwnd) noexcept;
    void UpdateCaptionStatus(HWND hwnd, const std::vector<TaskSnapshot>& snapshot) noexcept;
    void PaintCaptionStatusGlyph(HWND hwnd) noexcept;
    bool EnsureTaskbarList() noexcept;
    void UpdateTaskbarProgress(HWND hwnd) noexcept;
    void ApplyTaskbarProgress(HWND hwnd, uint32_t state, uint64_t completed, uint64_t total) noexcept;
    void ClearTaskbarProgress(HWND hwnd) noexcept;

    PopupHitTest HitTest(float x, float y) const noexcept;
    std::optional<PopupMenuAnchor> ResolveButtonMenuAnchor(HWND hwnd,
                                                           const PopupHitTest& hit,
                                                           RedSalamander::DxUi::ContextMenuRootVerticalPlacement placement) const noexcept;
    void Invalidate(HWND hwnd) const noexcept;

    LRESULT OnActivatedHit(HWND hwnd, const PopupHitTest& hit) noexcept;

    bool ConfirmCancelAll(HWND hwnd) noexcept;
    void ShowQueueModeMenu(HWND hwnd) noexcept;
    void ShowFooterOptionsMenu(HWND hwnd) noexcept;
    void ShowSpeedLimitMenu(HWND hwnd, uint64_t taskId) noexcept;
    bool ShowCustomSpeedLimitPromptForTask(HWND hwnd, uint64_t requestedTaskId) noexcept;
    void ShowDestinationMenu(HWND hwnd, uint64_t taskId) noexcept;
    bool SubmitCompletedOverflowAction(HWND hwnd, uint64_t taskId, uint32_t action, bool openExportAfterWrite) noexcept;
    void ShowCompletedOverflowMenu(HWND hwnd, uint64_t taskId) noexcept;
    bool SubmitConflictPrimaryAction(HWND hwnd, uint64_t taskId, uint32_t rawAction) noexcept;
    bool SubmitConflictOverflowAction(HWND hwnd, uint64_t taskId, uint32_t rawAction) noexcept;
    void ShowConflictOverflowMenu(HWND hwnd, uint64_t taskId) noexcept;

    LRESULT OnCreate(HWND hwnd) noexcept;
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
    LRESULT OnMouseWheel(HWND hwnd, int delta) noexcept;
    LRESULT OnClose(HWND hwnd) noexcept;
    // C4: graphics-independent failure surface (native children; no Direct2D, DirectWrite, or DxUi).
    bool EnterFailureSurface(HWND hwnd) noexcept;
    void LayoutFailureSurface(HWND hwnd) noexcept;
    void RefreshFailureSurfaceText() noexcept;
    bool HandleFailureSurfaceKeyDown(HWND hwnd, HWND focused, UINT virtualKey) noexcept;
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
    bool _destroyOnNextShowForSelfTest = false;
    std::unordered_map<uint64_t, uint32_t> _graphRowColorMatchCountsForSelfTest;
    std::unordered_map<uint64_t, uint32_t> _graphRowColorMismatchCountsForSelfTest;
#endif

    CaptionStatus _captionStatus = CaptionStatus::None;

    float _scrollY                    = 0.0f;
    float _contentHeight              = 0.0f;
    float _lastAutoSizedContentHeight = 0.0f; // For auto-resize tracking
    size_t _lastTaskCount             = 0;    // For auto-resize tracking
    int _maxAutoSizedWindowHeight     = 0;    // Sticky max window height (prevents resize "dancing")
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

    RedSalamander::DxUi::WindowHost _controlsHost;
    RedSalamander::DxUi::Panel* _controlsRoot = nullptr;
    bool _failureSurface                      = false;
    HWND _failureText                         = nullptr;
    HWND _failureCancelAll                    = nullptr;
    HWND _failureClose                        = nullptr;
    UINT _failureFontDpi                      = 0u;
    wil::unique_hfont _failureFont;
    std::wstring _failureTextValue;
    std::vector<RedSalamander::DxUi::Button*> _hostedButtons;
    std::vector<PopupHitTest> _hostedButtonHits;
    uint64_t _recycleEscalationFocusedTaskId = 0u;
    uint64_t _conflictEscapeTaskId           = 0u;
    uint8_t _conflictEscapeAction            = 0u;
    std::vector<HostedProgressDescriptor> _hostedProgressDescriptors;
    std::vector<RedSalamander::DxUi::ProgressBar*> _hostedProgressBars;
    std::vector<ScalarAnimation> _hostedProgressAnimations;
    std::vector<HostedGraphDescriptor> _hostedGraphDescriptors;
    std::vector<RedSalamander::DxUi::ThroughputGraph*> _hostedGraphs;
    std::vector<HostedTooltipDescriptor> _hostedTooltipDescriptors;
    std::vector<RedSalamander::DxUi::Label*> _hostedTooltipRegions;
    std::unordered_map<uint64_t, TaskSnapshot::StatusKind> _announcedTaskStatuses;
    std::unordered_map<uint64_t, bool> _announcedConflictMetadataLoading;
    bool _hostedRightToLeft                        = false;
    bool _hostedAccessibilityInitialized           = false;
    uint64_t _hostedAccessibilityNotificationCount = 0u;

    std::unordered_map<uint64_t, RateHistory> _rates;
    std::unordered_map<uint64_t, bool> _collapsedTasks;
    std::unordered_set<uint64_t> _compactExpandedTasks;
    std::unordered_map<uint64_t, ScalarAnimation> _taskHeightAnimations;
    std::unordered_map<uint64_t, ScalarAnimation> _taskDisclosureAnimations;
    ScalarAnimation _completedGroupVisibilityAnimation;
    ScalarAnimation _footerDisclosureAnimation;
    bool _completedGroupExpanded = true;
    bool _reducedMotion          = false;
    std::wstring _footerNewTasksText;
    std::wstring _footerQueueText;
    std::wstring _footerParallelText;
    std::wstring _footerAutoDismissOnText;
    std::wstring _footerAutoDismissOffText;
    bool _footerAutoDismissLabelVisible = false;
#ifdef ENABLE_TESTS
    wchar_t _debugLastDrawnCheckboxGlyph = L'\0';
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
    wil::com_ptr<ID2D1Factory> _d2dFactory;
    wil::com_ptr<IDWriteFactory> _dwriteFactory;
    wil::com_ptr<ID2D1HwndRenderTarget> _target;
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

    int _mouseWheelRemainder = 0;
};
} // namespace FileOperationsPopupInternal

class FileOperationsPopup final
{
public:
    static HWND Create(FolderWindow::FileOperationState* fileOps, FolderWindow* folderWindow, HWND ownerWindow, std::weak_ptr<void> hostLifetime) noexcept;

private:
    FileOperationsPopup() = delete;
};

#ifdef ENABLE_TESTS
struct FileOperationsSpeedLimitPromptDebugSnapshot
{
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
[[nodiscard]] bool DebugUpdateFileOperationsPopupTaskbarProgress(HWND popup) noexcept;
[[nodiscard]] bool DebugBuildFileOperationsPopupGlobalSummarySnapshot(const std::vector<FileOperationsPopupInternal::TaskSnapshot>& tasks,
                                                                      FileOperationsPopupInternal::PopupLayoutDebugSnapshot& out,
                                                                      double displayedBytesPerSecOverride = -1.0,
                                                                      double aggregateEtaSecondsOverride  = -1.0) noexcept;
void DebugFailNextFileOperationsTaskbarListAttempts(unsigned int attempts) noexcept;
void DebugFailNextFileOperationsDxUiHostAttachAttempts(unsigned int attempts) noexcept;
void DebugFailNextFileOperationsD2DTargetAttempts(unsigned int attempts) noexcept;
[[nodiscard]] bool DebugBuildFileOperationsGraphFairColorWeightSnapshot(FileOperationsPopupInternal::GraphHueWeightDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugBuildFileOperationsGraphFairnessHistorySnapshot(FileOperationsPopupInternal::PopupLayoutDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugFileOperationsStreamColorSlotsRemainUniqueUnderPermutation() noexcept;
[[nodiscard]] bool DebugValidateFileOperationsVerificationPresentation() noexcept;
[[nodiscard]] float DebugComputeFileOperationsTaskCompleteFraction(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept;
void DebugPublishFileOperationsPlannedItemTotalAfterDiscovery(FileOperationsPopupInternal::TaskSnapshot& task) noexcept;
[[nodiscard]] bool DebugFileOperationsTaskHasKnownCompactProgress(const FileOperationsPopupInternal::TaskSnapshot& task) noexcept;
[[nodiscard]] std::wstring DebugFormatFileOperationsConflictTimestamp(__int64 fileTime) noexcept;
[[nodiscard]] RedSalamander::DxUi::ButtonVariant DebugResolveFileOperationsHostedButtonVariant(FileOperationsPopupInternal::PopupHitTest::Kind kind) noexcept;
[[nodiscard]] std::wstring DebugFormatFileOperationsSpeedLimitSelectorText(uint64_t bytesPerSecond);
[[nodiscard]] double DebugSmoothRateForDisplay(double previousRate, double sampleRate, ULONGLONG elapsedMs) noexcept;
[[nodiscard]] double DebugDecayRateForCallbackSilence(double smoothedRate, ULONGLONG silenceMs) noexcept;
[[nodiscard]] double DebugSmoothEtaSecondsForDisplay(double previousEtaSeconds, double sampleEtaSeconds, ULONGLONG elapsedMs) noexcept;
[[nodiscard]] float DebugEaseFileOperationsGraphLatestPointYForDisplay(float previousY, float targetY, ULONGLONG elapsedMs) noexcept;
[[nodiscard]] float DebugEaseFileOperationsAutoResizeFraction(ULONGLONG elapsedMs, ULONGLONG durationMs) noexcept;
[[nodiscard]] D2D1_RECT_F DebugComputeFileOperationsIndeterminateBarFill(const D2D1_RECT_F& bar, ULONGLONG tick, bool reducedMotion) noexcept;
[[nodiscard]] HWND GetFileOperationsSpeedLimitPromptHandle() noexcept;
[[nodiscard]] bool DebugGetFileOperationsSpeedLimitPromptSnapshot(FileOperationsSpeedLimitPromptDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugSetFileOperationsSpeedLimitPromptText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugConfirmFileOperationsSpeedLimitPrompt() noexcept;
[[nodiscard]] bool DebugCancelFileOperationsSpeedLimitPrompt() noexcept;
#endif
