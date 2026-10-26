// I35 S5 (Specs/Plans/Done/FileOperations_CopyMoveUxAndReliabilityReview_2026-10-11.md, section 8): the
// decision mirror. A published File Operations decision is rendered on its task card and, while a pane
// shows the folder the decision is about, mirrored as a modeless overlay over that pane. One engine
// decision, two renderers: either answer goes through Task::SubmitConflictDecision with the prompt's
// generation, the first accepted answer wins, and the other renderer disappears. The mirror never decides
// on its own (no timer, focus loss, close or teardown submits anything), never blocks the window and is
// never a nested message loop. The main window also shows a pending-decision segment in the pane status
// bars and flashes the popup's taskbar button once when decisions start to wait.

#include "FolderWindow.FileOperationsInternal.h"

#include "DxUiThemePalette.h"
#include "FolderWindow.FileOperations.PromptText.h"
#include "HostServices.h"
#include "NavigationLocation.h"
#include "Ui/AlertOverlayTheme.h"
#include "Ui/AlertOverlayWindow.h"

#include <DxUi/DxUi.h>

#include <algorithm>
#include <deque>
#include <format>
#include <unordered_map>

namespace
{
using Task           = FolderWindow::FileOperationState::Task;
using ConflictAction = Task::ConflictAction;
using ConflictBucket = Task::ConflictBucket;
using Placement      = FileOperations::ConflictActionPlacement;

// A published primary action is shown as kPrimaryButtonBase plus its raw value. Escape is bound to an id no
// button carries, so it always submits the prompt's own escape action with the Escape placement.
constexpr uint32_t kPrimaryButtonBase   = 0x100u;
constexpr uint32_t kMoreButtonId        = 0x200u;
constexpr uint32_t kEscapeButtonId      = 0x201u;
constexpr uint32_t kApplyToAllOptionId  = 0x202u;
constexpr size_t kMaxRememberedPrompts  = 64u;  // closed, inspected or announced {taskId, generation} keys
constexpr size_t kMaxOriginFolders      = 8u;   // distinct folders a task-wide decision can be about
constexpr size_t kMaxOriginItemsScanned = 256u; // selected roots examined for their parent folders
constexpr size_t kMaxActions            = Task::ConflictPromptState::kMaxActions;

struct PromptKey final
{
    uint64_t taskId     = 0u;
    uint64_t generation = 0u;

    [[nodiscard]] bool operator==(const PromptKey&) const noexcept = default;
};

// A folder a decision is about: a provider path under its endpoint's path contract, with the plugin and
// the raw pane instance context of its role.
struct MirrorLocation final
{
    FileSystemPathIdentity identity;
    std::wstring pluginId;
    std::wstring instanceContext;
    std::wstring folder;
};

// The published prompt the mirror renders, copied under the task's arbiter lock and used without it.
struct MirrorPrompt final
{
    PromptKey key;
    uint64_t orderKey               = 0u;
    FileSystemOperation operation   = FILESYSTEM_COPY;
    ConflictBucket bucket           = ConflictBucket::Unknown;
    HRESULT status                  = S_OK;
    unsigned int attemptCount       = 0u;
    uint8_t overlapProblem          = 0u;
    uint64_t overlapTaskId          = 0u;
    uint8_t repeatUnavailableReason = 0u;
    bool readOnlyDeleteOffered      = false;
    bool sourceIsDirectory          = false;
    bool skipAllOffered             = false;
    std::wstring sourcePath;
    std::wstring destinationPath;
    std::wstring consentDetail;
    std::wstring consentNote;
    std::wstring sourceName;
    std::wstring sourceParentPath;
    std::wstring destinationName;
    std::wstring destinationParentPath;
    std::shared_ptr<const FileOperations::ConflictInspectionContext> inspection;
    std::array<ConflictAction, kMaxActions> primaryActions{};
    size_t primaryActionCount = 0u;
    std::array<ConflictAction, kMaxActions> overflowActions{};
    size_t overflowActionCount   = 0u;
    ConflictAction defaultAction = ConflictAction::None;
    ConflictAction escapeAction  = ConflictAction::None;
    bool applyToAllEligible      = false;
    bool applyToAllChecked       = false;
    bool factItemCountKnown      = false;
    uint64_t factItemCount       = 0u;
    bool factBytesKnown          = false;
    uint64_t factBytes           = 0u;
    FolderWindow::Pane pane      = FolderWindow::Pane::Left;
};

[[nodiscard]] FileOperations::ConflictPromptTextFacts PromptTextFacts(const MirrorPrompt& prompt) noexcept
{
    FileOperations::ConflictPromptTextFacts facts{};
    facts.operation               = prompt.operation;
    facts.bucket                  = static_cast<uint8_t>(prompt.bucket);
    facts.status                  = prompt.status;
    facts.attemptCount            = prompt.attemptCount;
    facts.overlapProblem          = prompt.overlapProblem;
    facts.overlapTaskId           = prompt.overlapTaskId;
    facts.repeatUnavailableReason = prompt.repeatUnavailableReason;
    facts.readOnlyDeleteOffered   = prompt.readOnlyDeleteOffered;
    facts.sourceIsDirectory       = prompt.sourceIsDirectory;
    facts.skipAllOffered          = prompt.skipAllOffered;
    facts.sourcePath              = prompt.sourcePath;
    facts.destinationPath         = prompt.destinationPath;
    facts.sourceName              = prompt.sourceName;
    facts.sourceParentPath        = prompt.sourceParentPath;
    facts.destinationName         = prompt.destinationName;
    facts.destinationParentPath   = prompt.destinationParentPath;
    facts.consentDetail           = prompt.consentDetail;
    facts.consentNote             = prompt.consentNote;
    facts.factItemCountKnown      = prompt.factItemCountKnown;
    facts.factItemCount           = prompt.factItemCount;
    facts.factBytesKnown          = prompt.factBytesKnown;
    facts.factBytes               = prompt.factBytes;
    return facts;
}

[[nodiscard]] bool SameColor(const D2D1::ColorF& left, const D2D1::ColorF& right) noexcept
{
    return left.r == right.r && left.g == right.g && left.b == right.b && left.a == right.a;
}

[[nodiscard]] bool SameAlertTheme(const RedSalamander::Ui::AlertTheme& left, const RedSalamander::Ui::AlertTheme& right) noexcept
{
    return SameColor(left.background, right.background) && SameColor(left.text, right.text) && SameColor(left.accent, right.accent) &&
           SameColor(left.selectionBackground, right.selectionBackground) && SameColor(left.selectionText, right.selectionText) &&
           SameColor(left.errorBackground, right.errorBackground) && SameColor(left.errorText, right.errorText) &&
           SameColor(left.warningBackground, right.warningBackground) && SameColor(left.warningText, right.warningText) &&
           SameColor(left.infoBackground, right.infoBackground) && SameColor(left.infoText, right.infoText) && left.darkBase == right.darkBase &&
           left.highContrast == right.highContrast;
}

template <typename Keys> [[nodiscard]] bool Remembered(const Keys& keys, PromptKey key) noexcept
{
    return std::ranges::find(keys, key) != keys.end();
}

void Remember(std::deque<PromptKey>& keys, PromptKey key)
{
    if (Remembered(keys, key))
        return;
    keys.push_back(key);
    while (keys.size() > kMaxRememberedPrompts)
        keys.pop_front();
}
} // namespace

struct FolderWindow::FileOperationState::DecisionMirror final
{
    explicit DecisionMirror(FileOperationState& owner) noexcept : state(owner)
    {
    }

    DecisionMirror(const DecisionMirror&)            = delete;
    DecisionMirror(DecisionMirror&&)                 = delete;
    DecisionMirror& operator=(const DecisionMirror&) = delete;
    DecisionMirror& operator=(DecisionMirror&&)      = delete;
    ~DecisionMirror()                                = default;

    FileOperationState& state;
    std::unique_ptr<RedSalamander::Ui::AlertOverlayWindow> overlay;

    // The prompt the overlay shows ({0, 0}: none) and what it was shown with.
    PromptKey shown;
    FolderWindow::Pane shownPane = FolderWindow::Pane::Left;
    HWND shownAnchor             = nullptr;
    std::wstring shownText;
    RedSalamander::Ui::AlertTheme shownTheme{};
    ConflictAction shownEscapeAction = ConflictAction::None;
    std::array<ConflictAction, kMaxActions> shownOverflowActions{};
    size_t shownOverflowActionCount    = 0u;
    FileSystemOperation shownOperation = FILESYSTEM_COPY;
    ConflictBucket shownBucket         = ConflictBucket::Unknown;
    bool shownApplyToAll               = false;
    // The engine's Apply-to-all value the overlay shows (the card's checkbox and the mirror's option are one
    // engine flag).
    bool shownApplyToAllChecked = false;
    bool tookFocus              = false;
    bool announced              = false;

    // Prompts the user closed on the mirror or is inspecting from the card: never mirrored again.
    std::deque<PromptKey> suppressed;
    // Prompts the user sent the focus away from (F6): still mirrored, never with the focus taken again.
    std::deque<PromptKey> focusDeclined;
    // R8.12: who announced a prompt (exactly one surface), and the prompts the last reconcile decided about
    // (rebuilt from the live pending set at every pass, so no live prompt is ever evicted), so the card waits
    // for that decision instead of racing the mirror to the announcement.
    struct AnnouncementClaim final
    {
        PromptKey key;
        bool byMirror = false;
    };
    std::deque<AnnouncementClaim> announcements;
    std::vector<PromptKey> evaluated;

    // The folders a task-wide decision is about, kept per selected admission (immutable once published).
    struct CachedOrigin final
    {
        std::shared_ptr<const FolderWindow::FileOperationState::SelectedAdmissionView> admission;
        std::vector<MirrorLocation> locations;
    };
    std::unordered_map<uint64_t, CachedOrigin> origins;
    std::vector<Task*> scratchTasks;

    unsigned long pendingCount = 0u;
    // D-8.1 debounce: a run of decisions (the pending set with gaps shorter than kIndicatorSettleMs between
    // two prompts) flashes once and keeps the status segment; the latch drops once the set stayed empty.
    bool indicatorLatched             = false;
    ULONGLONG indicatorEmptySinceTick = 0u;

    // A self-test stands in the foreground window (and moves it to the overlay when the mirror takes the
    // focus, as the activation that follows SetFocus does).
    std::optional<HWND> debugForegroundOverride;
#ifdef ENABLE_TESTS
    // A self-test makes the next shows fail, or holds the reconcile back while the popup synchronizes.
    uint32_t debugFailShows   = 0u;
    bool debugReconcilePaused = false;
#endif
    uint64_t showCount        = 0u;
    uint64_t showFailureCount = 0u;
    uint64_t submitCount      = 0u;
    uint64_t acceptedCount    = 0u;
    uint64_t flashCount       = 0u;
    uint64_t reconcileCount   = 0u;
    uint64_t focusReturnCount = 0u;
    uint64_t cardDeferCount   = 0u;

    [[nodiscard]] FolderWindow& Owner() const noexcept
    {
        return state._owner;
    }

    [[nodiscard]] bool OverlayShown() const noexcept
    {
        return overlay && overlay->IsVisible() && shown.taskId != 0u;
    }

    [[nodiscard]] HWND OwnerRoot() const noexcept
    {
        const HWND owner = Owner().GetHwnd();
        return owner ? GetAncestor(owner, GA_ROOT) : nullptr;
    }

    [[nodiscard]] HWND ForegroundWindow() const noexcept
    {
        return debugForegroundOverride.has_value() ? debugForegroundOverride.value() : GetForegroundWindow();
    }

    // The user works in the owner window: the foreground window is the owner's top-level window or a window
    // that window owns, directly or through an owner chain (the mirror once it took the focus, the File
    // Operations popup, a dialog of the owner).
    [[nodiscard]] bool OwnerRootIsForeground(HWND anyOwnedWindow) const noexcept
    {
        const HWND root = anyOwnedWindow ? GetAncestor(anyOwnedWindow, GA_ROOT) : nullptr;
        if (! root)
            return false;
        for (HWND window = ForegroundWindow(); window != nullptr; window = GetWindow(window, GW_OWNER))
        {
            if (window == root)
                return true;
        }
        return false;
    }

    // A dialog of the owner runs its own modal loop (Change attributes, Make file list, credentials): the
    // owner window is disabled and nothing may sit above that dialog.
    [[nodiscard]] bool OwnerRootEnabled() const noexcept
    {
        const HWND root = OwnerRoot();
        return root != nullptr && IsWindowEnabled(root) != FALSE;
    }

    // ---- prompt facts ----------------------------------------------------------------------------

    [[nodiscard]] static bool CopyPendingPrompt(Task& task, MirrorPrompt& out) noexcept
    {
        std::scoped_lock lock(task._conflictArbiter.mutex);
        const auto& prompt = task._conflictArbiter.prompt;
        // R8.1: an actionable published prompt that nobody has answered yet. A decision the card already
        // gave stays "active" until the worker consumes it; it is no longer pending here.
        if (! prompt.active || prompt.metadataLoading || ! prompt.buttonsPublishable || prompt.generation == 0u || prompt.primaryActionCount == 0u ||
            task._conflictArbiter.decisionAction.has_value())
            return false;
        out.key                     = {task._taskId, prompt.generation};
        out.bucket                  = prompt.bucket;
        out.status                  = prompt.status;
        out.attemptCount            = prompt.attemptCount;
        out.overlapProblem          = prompt.overlapProblem;
        out.overlapTaskId           = prompt.overlapTaskId;
        out.repeatUnavailableReason = static_cast<uint8_t>(prompt.repeatUnavailableReason);
        out.sourceIsDirectory       = prompt.sourceMetadata.objectKindKnown ? prompt.sourceMetadata.isDirectory
                                                                            : prompt.decisionScope.sourceKind == FileOperations::ConflictItemKind::Directory;
        const auto offered          = std::span(prompt.actions.data(), std::min(prompt.actionCount, prompt.actions.size()));
        out.readOnlyDeleteOffered   = std::ranges::find(offered, ConflictAction::DeleteReadOnly) != offered.end();
        out.skipAllOffered          = std::ranges::find(offered, ConflictAction::SkipAll) != offered.end();
        out.sourcePath              = prompt.sourcePath;
        out.destinationPath         = prompt.destinationPath;
        out.consentDetail           = prompt.consentDetail;
        out.consentNote             = prompt.consentNote;
        out.sourceName              = prompt.pathDisplay.sourceName;
        out.sourceParentPath        = prompt.pathDisplay.sourceParentPath;
        out.destinationName         = prompt.pathDisplay.destinationName;
        out.destinationParentPath   = prompt.pathDisplay.destinationParentPath;
        out.inspection              = prompt.pathDisplay.inspection;
        out.primaryActionCount      = std::min(prompt.primaryActionCount, out.primaryActions.size());
        std::copy_n(prompt.primaryActions.begin(), out.primaryActionCount, out.primaryActions.begin());
        out.overflowActionCount = std::min(prompt.overflowActionCount, out.overflowActions.size());
        std::copy_n(prompt.overflowActions.begin(), out.overflowActionCount, out.overflowActions.begin());
        out.defaultAction      = prompt.defaultAction;
        out.escapeAction       = prompt.escapeAction;
        out.applyToAllEligible = prompt.applyToAllEligible;
        out.applyToAllChecked  = prompt.applyToAllChecked;
        out.factItemCountKnown = prompt.factItemCountKnown;
        out.factItemCount      = prompt.factItemCount;
        out.factBytesKnown     = prompt.factBytesKnown;
        out.factBytes          = prompt.factBytes;
        return true;
    }

    // ---- which pane shows the decision's folder (R8.1) -------------------------------------------

    static void AddLocation(std::vector<MirrorLocation>& locations,
                            const FileOperations::QualifiedEndpoint& endpoint,
                            std::wstring_view pluginId,
                            std::wstring_view instanceContext,
                            std::wstring_view folder)
    {
        if (locations.size() >= kMaxOriginFolders || folder.empty() || ! endpoint.pathIdentity.has_value() || pluginId.empty())
            return;
        const FileSystemPathIdentity& identity = endpoint.pathIdentity.value();
        for (const MirrorLocation& existing : locations)
        {
            if (NavigationLocation::EqualsNoCase(existing.pluginId, pluginId) && existing.instanceContext == instanceContext &&
                EquivalentPath(identity, existing.folder, folder))
                return;
        }
        locations.push_back(MirrorLocation{identity, std::wstring(pluginId), std::wstring(instanceContext), std::wstring(folder)});
    }

    static void AddParentLocation(std::vector<MirrorLocation>& locations,
                                  const FileOperations::QualifiedEndpoint& endpoint,
                                  std::wstring_view pluginId,
                                  std::wstring_view instanceContext,
                                  std::wstring_view itemPath)
    {
        if (! endpoint.pathIdentity.has_value())
            return;
        std::wstring parent;
        if (TryGetFileSystemParentPath(endpoint.pathIdentity.value(), itemPath, parent))
            AddLocation(locations, endpoint, pluginId, instanceContext, parent);
    }

    // The folders a task-wide decision is about: Delete and Rename the source parents, Copy the
    // destination folder, Move both. Bounded: a selection normally shares one parent.
    [[nodiscard]] static std::vector<MirrorLocation> ComputeOriginLocations(const Task& task, const SelectedAdmissionView& admission)
    {
        std::vector<MirrorLocation> locations;
        const FileSystemOperation operation = task._operation;
        const bool sourceRole               = operation == FILESYSTEM_DELETE || operation == FILESYSTEM_MOVE || operation == FILESYSTEM_RENAME;
        const bool destinationRole          = operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE;
        const bool hasDestinationContext    = admission.destination && ! admission.destination->pluginId.empty();
        const std::wstring_view destinationPluginId =
            hasDestinationContext ? std::wstring_view(admission.destination->pluginId) : std::wstring_view(task._sourcePluginId);
        const std::wstring_view destinationContext =
            hasDestinationContext ? std::wstring_view(admission.destination->instanceContext) : std::wstring_view(task._sourceInstanceContext);
        if (! admission.plans)
            return locations;
        for (const FileOperations::FileOperationPlan& plan : *admission.plans)
        {
            if (const auto* transfer = std::get_if<FileOperations::TransferPlan>(&plan))
            {
                if (sourceRole)
                {
                    const size_t count = std::min(transfer->selectedItems.size(), kMaxOriginItemsScanned);
                    for (size_t index = 0u; index < count; ++index)
                        AddParentLocation(locations,
                                          transfer->sourceEndpoint,
                                          task._sourcePluginId,
                                          task._sourceInstanceContext,
                                          transfer->selectedItems[index].providerPath);
                }
                if (destinationRole)
                    AddLocation(locations, transfer->destinationEndpoint, destinationPluginId, destinationContext, transfer->destination.providerFolderPath);
            }
            else if (const auto* deletion = std::get_if<FileOperations::DeletePlan>(&plan))
            {
                const size_t count = std::min(deletion->selectedItems.size(), kMaxOriginItemsScanned);
                for (size_t index = 0u; index < count; ++index)
                    AddParentLocation(
                        locations, deletion->endpoint, task._sourcePluginId, task._sourceInstanceContext, deletion->selectedItems[index].providerPath);
            }
            else if (const auto* rename = std::get_if<FileOperations::RenamePlan>(&plan))
            {
                const size_t count = std::min(rename->finalMappings.size(), kMaxOriginItemsScanned);
                for (size_t index = 0u; index < count; ++index)
                    AddParentLocation(
                        locations, rename->endpoint, task._sourcePluginId, task._sourceInstanceContext, rename->finalMappings[index].source.providerPath);
            }
        }
        return locations;
    }

    // The folders this decision is about. A per-item prompt names its item's role parents; a task-wide
    // prompt names the task's origin folders, and only before the task's first governed change.
    [[nodiscard]] std::vector<MirrorLocation> RelevantLocations(Task& task, const MirrorPrompt& prompt)
    {
        std::vector<MirrorLocation> locations;
        const FileSystemOperation operation = task._operation;
        if (prompt.inspection)
        {
            const auto admission             = task.LoadSelectedAdmission();
            const bool hasDestinationContext = admission && admission->destination && ! admission->destination->pluginId.empty();
            const std::wstring_view destinationPluginId =
                hasDestinationContext ? std::wstring_view(admission->destination->pluginId) : std::wstring_view(task._sourcePluginId);
            const std::wstring_view destinationContext =
                hasDestinationContext ? std::wstring_view(admission->destination->instanceContext) : std::wstring_view(task._sourceInstanceContext);
            const bool sourceRole      = operation == FILESYSTEM_DELETE || operation == FILESYSTEM_MOVE || operation == FILESYSTEM_RENAME;
            const bool destinationRole = operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE;
            if (sourceRole && prompt.inspection->source.has_value())
            {
                const auto& source = prompt.inspection->source.value();
                AddLocation(locations, source.endpoint, task._sourcePluginId, task._sourceInstanceContext, source.parentPath);
            }
            if (destinationRole && prompt.inspection->destination.has_value())
            {
                const auto& destination = prompt.inspection->destination.value();
                AddLocation(locations, destination.endpoint, destinationPluginId, destinationContext, destination.parentPath);
            }
            if (! locations.empty())
                return locations;
        }
        if (task.HasMutated())
            return locations;
        const auto admission = task.LoadSelectedAdmission();
        if (! admission)
            return locations;
        auto& cached = origins[prompt.key.taskId];
        if (cached.admission != admission)
        {
            cached.admission = admission;
            cached.locations = ComputeOriginLocations(task, *admission);
        }
        return cached.locations;
    }

    [[nodiscard]] bool PaneShows(FolderWindow::Pane pane, const MirrorLocation& location) const noexcept
    {
        const FolderWindow& owner        = Owner();
        const FolderWindow::PaneState& p = pane == FolderWindow::Pane::Left ? owner._leftPane : owner._rightPane;
        // A preview or terminal tab hides the pane's view; the mirror never sits over another surface.
        if (! p.hFolderView || IsWindowVisible(p.hFolderView.get()) == FALSE)
            return false;
        if (! NavigationLocation::EqualsNoCase(p.pluginId, location.pluginId) || p.instanceContext != location.instanceContext)
            return false;
        const std::optional<std::filesystem::path> folder = p.folderView.GetFolderPath();
        return folder.has_value() && EquivalentPath(location.identity, folder->native(), location.folder);
    }

    [[nodiscard]] bool PaneShowsAny(FolderWindow::Pane pane, const std::vector<MirrorLocation>& locations) const noexcept
    {
        return std::ranges::any_of(locations, [&](const MirrorLocation& location) noexcept { return PaneShows(pane, location); });
    }

    // The pane that shows the decision's folder: the pane already showing it, else the active pane,
    // else the other one.
    [[nodiscard]] std::optional<FolderWindow::Pane> ResolvePane(const std::vector<MirrorLocation>& locations,
                                                                std::optional<FolderWindow::Pane> keep) const noexcept
    {
        if (locations.empty())
            return std::nullopt;
        if (keep.has_value() && PaneShowsAny(keep.value(), locations))
            return keep;
        const FolderWindow::Pane active = Owner().GetActivePane();
        if (PaneShowsAny(active, locations))
            return active;
        const FolderWindow::Pane other = active == FolderWindow::Pane::Left ? FolderWindow::Pane::Right : FolderWindow::Pane::Left;
        if (PaneShowsAny(other, locations))
            return other;
        return std::nullopt;
    }

    // ---- reconcile (R8.1, R8.6, R8.7) --------------------------------------------------------------

    void Reconcile() noexcept
    {
        Debug::Perf::Scope perf(L"FileOps.DecisionMirror.ReconcileUs");
        ++reconcileCount;
        if (state._completionShutdown.load(std::memory_order_acquire))
        {
            Hide(false);
            pendingCount = 0u;
            return;
        }

        // Borrow the tasks outside _mutex like the popup does: only this owner thread removes them, and
        // the prompt lock is never taken under _mutex. Nothing below pumps messages before the last use.
        scratchTasks.clear();
        {
            std::scoped_lock lock(state._mutex);
            scratchTasks.reserve(state._tasks.size());
            for (const auto& task : state._tasks)
                if (task)
                    scratchTasks.push_back(task.get());
        }
        std::erase_if(origins, [&](const auto& entry) noexcept {
            return std::ranges::none_of(scratchTasks, [&](const Task* task) noexcept { return task->_taskId == entry.first; });
        });

        const bool enabled           = state.GetMirrorDecisionsInPane();
        const bool shownStillVisible = OverlayShown();
        unsigned long pending        = 0u;
        std::optional<MirrorPrompt> current;
        std::optional<MirrorPrompt> first;
        evaluated.clear();
        for (Task* task : scratchTasks)
        {
            if (task->_taskFinished.load(std::memory_order_acquire) || task->CancelRequestedDurable())
                continue;
            MirrorPrompt prompt;
            if (! CopyPendingPrompt(*task, prompt))
                continue;
            // R8.12: this pass decides whether the mirror announces the prompt; the card may announce it after.
            evaluated.push_back(prompt.key);
            if (task->IsPresentationHidden())
                continue;
            ++pending;
            if (! enabled || Remembered(suppressed, prompt.key))
                continue;
            prompt.operation                     = task->_operation;
            prompt.orderKey                      = task->GetQueueOrderKey();
            const bool isShown                   = shownStillVisible && prompt.key == shown;
            const std::vector<MirrorLocation> at = RelevantLocations(*task, prompt);
            const auto pane                      = ResolvePane(at, isShown ? std::optional<FolderWindow::Pane>(shownPane) : std::nullopt);
            if (! pane.has_value())
                continue;
            prompt.pane = pane.value();
            if (isShown)
            {
                current = std::move(prompt);
                continue;
            }
            // R8.6: FIFO by queue order, then by prompt generation.
            if (! first.has_value() || prompt.orderKey < first->orderKey ||
                (prompt.orderKey == first->orderKey && prompt.key.generation < first->key.generation))
                first = std::move(prompt);
        }
        pendingCount = pending;
        perf.SetValue0(pending);

        // A dialog of the owner runs a modal loop: the mirror leaves the screen (without taking the focus
        // back) and returns at the first pass after the dialog closed.
        const bool ownerEnabled = OwnerRootEnabled();

        // Keep a prompt on screen while it stays eligible (no jump between decisions); the next one is
        // shown once it is answered, closed or no longer shown in a pane.
        if (current.has_value() && ownerEnabled)
        {
            perf.SetValue1(1u);
            if (current->pane == shownPane && Owner().GetFolderViewHwnd(current->pane) == shownAnchor)
            {
                Refresh(current.value());
                return;
            }
        }
        Hide(ownerEnabled);
        if (! ownerEnabled)
            return;
        if (current.has_value())
            first = std::move(current);
        // A blocking host prompt's nested loop is running: the mirror never appears above it.
        if (first.has_value() && ! HostIsBlockingPromptActive())
            Show(first.value());
    }

    // ---- presentation (R8.3, R8.4, R8.12, R8.13) -----------------------------------------------------

    [[nodiscard]] RedSalamander::Ui::AlertModel BuildModel(const MirrorPrompt& prompt, std::wstring& textOut) const
    {
        using namespace RedSalamander::Ui;
        const FileOperations::ConflictPromptTextFacts facts = PromptTextFacts(prompt);
        AlertModel model{};
        model.severity = AlertSeverity::Warning;
        switch (prompt.operation)
        {
            case FILESYSTEM_COPY: model.presentation = AlertPresentation::Copy; break;
            case FILESYSTEM_MOVE: model.presentation = AlertPresentation::Move; break;
            case FILESYSTEM_DELETE: model.presentation = AlertPresentation::Delete; break;
            case FILESYSTEM_RENAME:
            case FILESYSTEM_CREATE_DIRECTORY:
            default: model.presentation = AlertPresentation::Severity; break;
        }
        // R8.12: the overlay's UI Automation name is the card's title.
        model.title = LoadStringResource(nullptr, FileOperations::ConflictPromptTitleStringId(prompt.operation));

        std::wstring message;
        if (const auto name = FileOperations::CommonConflictPromptName(facts); ! name.empty())
        {
            message.append(name);
            message.append(L"\n");
        }
        message.append(FileOperations::BuildConflictPromptQuestion(facts));
        if (const std::wstring factText = FileOperations::BuildConflictPromptFactText(facts); ! factText.empty())
        {
            message.append(L"\n");
            message.append(factText);
        }
        if (! prompt.sourcePath.empty())
        {
            message.append(L"\n\n");
            message.append(std::format(L"{0} {1}",
                                       LoadStringResource(nullptr, prompt.operation == FILESYSTEM_DELETE ? IDS_FILEOPS_LABEL_DELETING : IDS_FILEOPS_LABEL_FROM),
                                       prompt.sourcePath));
        }
        if (prompt.operation != FILESYSTEM_DELETE && ! prompt.destinationPath.empty())
        {
            message.append(prompt.sourcePath.empty() ? L"\n\n" : L"\n");
            message.append(std::format(L"{0} {1}", LoadStringResource(nullptr, IDS_FILEOPS_LABEL_TO), prompt.destinationPath));
        }
        model.message  = std::move(message);
        model.closable = true;

        // The engine's published primary actions in order (the popup's presentation-only rule): the
        // published default carries the initial focus ring, so Enter submits it.
        for (size_t index = 0u; index < prompt.primaryActionCount; ++index)
        {
            const ConflictAction action = prompt.primaryActions[index];
            model.buttons.push_back(AlertButton{
                .id      = kPrimaryButtonBase + static_cast<uint32_t>(action),
                .label   = FileOperations::FormatConflictActionLabel(action, prompt.bucket, prompt.operation),
                .primary = action == prompt.defaultAction,
            });
        }
        if (prompt.overflowActionCount > 0u)
            model.buttons.push_back(AlertButton{.id = kMoreButtonId, .label = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_BTN_MORE), .primary = false});
        if (prompt.applyToAllEligible)
        {
            model.options.push_back(AlertOption{
                .id            = kApplyToAllOptionId,
                .label         = LoadStringResource(nullptr, IDS_FILEOPS_CONFLICT_APPLY_TO_ALL),
                .choices       = {{0u, LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF)}, {1u, LoadStringResource(nullptr, IDS_PREFS_COMMON_ON)}},
                .selectedIndex = prompt.applyToAllChecked ? 1u : 0u,
            });
        }

        textOut = model.title;
        textOut.append(L"\x1F");
        textOut.append(model.message);
        for (const AlertButton& button : model.buttons)
        {
            textOut.append(L"\x1F");
            textOut.append(button.label);
        }
        textOut.append(prompt.applyToAllEligible ? L"\x1F+" : L"\x1F-");
        return model;
    }

    [[nodiscard]] std::optional<uint32_t> DefaultButtonId(const MirrorPrompt& prompt) const noexcept
    {
        const auto begin = prompt.primaryActions.begin();
        const auto end   = begin + static_cast<std::ptrdiff_t>(prompt.primaryActionCount);
        if (prompt.defaultAction == ConflictAction::None || std::find(begin, end, prompt.defaultAction) == end)
            return std::nullopt;
        return kPrimaryButtonBase + static_cast<uint32_t>(prompt.defaultAction);
    }

    void Show(const MirrorPrompt& prompt) noexcept
    {
        FolderWindow& owner = Owner();
        const HWND anchor   = owner.GetFolderViewHwnd(prompt.pane);
        if (! anchor || IsWindowVisible(anchor) == FALSE)
            return;
        if (! overlay)
            overlay = std::make_unique<RedSalamander::Ui::AlertOverlayWindow>();

        std::wstring text;
        RedSalamander::Ui::AlertModel model       = BuildModel(prompt, text);
        const RedSalamander::Ui::AlertTheme theme = RedSalamander::Ui::MakeAlertTheme(owner.GetTheme());

        // A new decision in the foreground application is immediately keyboard-operable, even when the
        // other pane held focus. F6 still opts out for this prompt; never request another app's foreground.
        // The published safe default and the overlay's repeat suppression prevent implicit destructive consent.
        const bool takeFocus = OwnerRootIsForeground(anchor) && ! Remembered(focusDeclined, prompt.key);
        // R8.12: the mirror announces when it takes the focus (the card waits for this decision and announces
        // otherwise); a prompt is announced once. The claim is recorded only once the show succeeded.
        const bool announceHere = takeFocus && FindClaim(prompt.key) == nullptr;

        RedSalamander::Ui::AlertModelessShowOptions options{};
        options.takeFocus              = takeFocus;
        options.keysRequireFocus       = true;
        options.followAnchorVisibility = true;
        if (announceHere)
            options.announcement = FileOperations::BuildConflictPromptAnnouncement(PromptTextFacts(prompt));

        overlay->SetCallbacks(RedSalamander::Ui::AlertOverlayWindowCallbacks{
            this, &DecisionMirror::OnButtonThunk, &DecisionMirror::OnCloseThunk, &DecisionMirror::OnOptionChangedThunk, &DecisionMirror::OnFocusReturnedThunk});
        overlay->SetKeyBindings(DefaultButtonId(prompt), prompt.escapeAction != ConflictAction::None ? std::optional<uint32_t>(kEscapeButtonId) : std::nullopt);
        overlay->SetTrackedWindowDestroyedCallback(this, &DecisionMirror::OnAnchorDestroyedThunk);
        bool failShow = false;
#ifdef ENABLE_TESTS
        if (debugFailShows > 0u)
        {
            --debugFailShows;
            failShow = true;
        }
#endif
        const HRESULT hr = failShow ? E_FAIL : overlay->ShowModelessForAnchor(anchor, theme, std::move(model), std::move(options));
        if (FAILED(hr))
        {
            ++showFailureCount;
            Debug::Warning(L"File Operations could not show the decision mirror for task {} (hr=0x{:08X}).", prompt.key.taskId, static_cast<unsigned long>(hr));
            overlay->ClearCallbacks();
            overlay->SetKeyBindings(std::nullopt, std::nullopt);
            return;
        }
        if (announceHere)
            RecordClaim(prompt.key, true);

        shown                    = prompt.key;
        shownPane                = prompt.pane;
        shownAnchor              = anchor;
        shownText                = std::move(text);
        shownTheme               = theme;
        shownEscapeAction        = prompt.escapeAction;
        shownOverflowActions     = prompt.overflowActions;
        shownOverflowActionCount = prompt.overflowActionCount;
        shownOperation           = prompt.operation;
        shownBucket              = prompt.bucket;
        shownApplyToAll          = prompt.applyToAllEligible;
        shownApplyToAllChecked   = prompt.applyToAllChecked;
        tookFocus                = takeFocus;
        announced                = announceHere;
        ++showCount;
        // The activation that follows the focus moves the foreground to the overlay; a self-test that
        // stands in the foreground follows it, so the foreground-chain rules run as they do for real.
        if (takeFocus && debugForegroundOverride.has_value())
            debugForegroundOverride = overlay->GetHwnd();
        Debug::Perf::Emit(L"FileOps.DecisionMirror.Show", L"", prompt.key.taskId, prompt.key.generation, takeFocus ? 1u : 0u, S_OK);
    }

    // The same prompt is still shown: re-show only when its text (language) or theme changed, keeping the
    // focus it took. The Apply-to-all option shows the engine's flag (the card's checkbox),
    // updated in place when only that flag changed.
    void Refresh(const MirrorPrompt& prompt) noexcept
    {
        std::wstring text;
        RedSalamander::Ui::AlertModel model       = BuildModel(prompt, text);
        const RedSalamander::Ui::AlertTheme theme = RedSalamander::Ui::MakeAlertTheme(Owner().GetTheme());
        if (text == shownText && SameAlertTheme(theme, shownTheme))
        {
            if (prompt.applyToAllEligible && prompt.applyToAllChecked != shownApplyToAllChecked && overlay)
            {
                overlay->SetOptionValue(kApplyToAllOptionId, prompt.applyToAllChecked ? 1u : 0u);
                shownApplyToAllChecked = prompt.applyToAllChecked;
            }
            return;
        }
        RedSalamander::Ui::AlertModelessShowOptions options{};
        options.keysRequireFocus       = true;
        options.followAnchorVisibility = true;
        if (FAILED(overlay->ShowModelessForAnchor(shownAnchor, theme, std::move(model), std::move(options))))
        {
            Hide(true);
            return;
        }
        shownText              = std::move(text);
        shownTheme             = theme;
        shownApplyToAllChecked = prompt.applyToAllChecked;
    }

    void Hide(bool restoreFocus) noexcept
    {
        // A hidden overlay is left alone: a reconcile that finds nothing to show must not churn it.
        const HWND overlayWindow = overlay ? overlay->GetHwnd() : nullptr;
        if (overlay && (overlay->IsVisible() || shown.taskId != 0u))
        {
            if (restoreFocus)
                overlay->Hide();
            else
                overlay->HideWithoutRestoringFocus();
        }
        // Hiding the active overlay activates its owner again (a self-test's stand-in follows).
        if (overlayWindow && debugForegroundOverride == std::optional<HWND>{overlayWindow})
            debugForegroundOverride = OwnerRoot();
        shown       = {};
        shownAnchor = nullptr;
        shownText.clear();
        tookFocus = false;
        announced = false;
    }

    [[nodiscard]] const AnnouncementClaim* FindClaim(PromptKey key) const noexcept
    {
        const auto found = std::ranges::find_if(announcements, [&](const AnnouncementClaim& claim) noexcept { return claim.key == key; });
        return found != announcements.end() ? &*found : nullptr;
    }

    void RecordClaim(PromptKey key, bool byMirror)
    {
        announcements.push_back(AnnouncementClaim{key, byMirror});
        while (announcements.size() > kMaxRememberedPrompts)
            announcements.pop_front();
    }

    // R8.12: the card announces a published decision only once the mirror decided not to (a reconcile
    // evaluated the prompt and the mirror did not take the focus for it); a prompt the mirror will never
    // evaluate (not mirror-pending, or a task this state does not run) is the card's at once.
    [[nodiscard]] FileOperationState::PromptAnnouncementClaim ClaimForCard(PromptKey key)
    {
        if (FindClaim(key) != nullptr)
            return PromptAnnouncementClaim::AlreadyAnnounced;
        if (! Remembered(evaluated, key))
        {
            Task* task = state.FindTask(key.taskId);
            MirrorPrompt pending;
            if (task != nullptr && ! task->_taskFinished.load(std::memory_order_acquire) && ! task->CancelRequestedDurable() &&
                CopyPendingPrompt(*task, pending) && pending.key == key)
            {
                ++cardDeferCount;
                return PromptAnnouncementClaim::Defer;
            }
        }
        RecordClaim(key, false);
        return PromptAnnouncementClaim::Announce;
    }

    // ---- resolution (R8.5, R8.6) ----------------------------------------------------------------------

    void Submit(ConflictAction action, Placement placement) noexcept
    {
        const PromptKey key = shown;
        if (key.taskId == 0u || action == ConflictAction::None)
            return;
        ++submitCount;
        // Like the card: Apply to all is the engine's flag of this prompt (both surfaces edit that flag), so
        // the answer never carries a copy that could disagree with the card.
        Task* task          = state.FindTask(key.taskId);
        const bool accepted = task != nullptr && task->SubmitConflictDecision(key.generation, action, std::nullopt, placement);
        if (accepted)
            ++acceptedCount;
        Debug::Perf::Emit(L"FileOps.DecisionMirror.Submit", L"", key.taskId, static_cast<uint64_t>(action), accepted ? 1u : 0u, accepted ? S_OK : S_FALSE);
        // Accepted, already decided, canceled or replaced: this prompt leaves the mirror now, and the
        // next presentation refresh shows whatever is current.
        Hide(true);
        state.RequestTaskPresentationRefresh();
    }

    void OnButton(uint32_t buttonId) noexcept
    {
        const FolderWindow::FileOperationPromptDispatchScope dispatch(Owner());
        if (state._completionShutdown.load(std::memory_order_acquire) || shown.taskId == 0u)
        {
            Hide(false);
            return;
        }
        if (buttonId == kMoreButtonId)
        {
            ShowOverflowMenu();
            return;
        }
        if (buttonId == kEscapeButtonId)
        {
            Submit(shownEscapeAction, Placement::Escape);
            return;
        }
        if (buttonId > kPrimaryButtonBase && buttonId < kPrimaryButtonBase + 0x100u)
            Submit(static_cast<ConflictAction>(buttonId - kPrimaryButtonBase), Placement::Primary);
    }

    // Close only hides this surface; the decision stays on the card, and this prompt is not mirrored again.
    void OnClose() noexcept
    {
        const FolderWindow::FileOperationPromptDispatchScope dispatch(Owner());
        if (shown.taskId != 0u)
            Remember(suppressed, shown);
        Hide(true);
    }

    // The engine's Apply-to-all flag of the published prompt: the flag the card's checkbox toggles.
    [[nodiscard]] static bool SetApplyToAllChecked(Task& task, uint64_t generation, bool checked) noexcept
    {
        std::scoped_lock lock(task._conflictArbiter.mutex);
        auto& prompt = task._conflictArbiter.prompt;
        if (generation == 0u || prompt.generation != generation || ! prompt.active || prompt.metadataLoading || ! prompt.buttonsPublishable ||
            ! prompt.applyToAllEligible || task._conflictArbiter.decisionAction)
            return false;
        prompt.applyToAllChecked = checked;
        return true;
    }

    // Apply to all on the mirror edits the engine flag, so the card and the answer see one value.
    void OnOptionChanged(uint32_t optionId, uint64_t value) noexcept
    {
        const FolderWindow::FileOperationPromptDispatchScope dispatch(Owner());
        if (optionId != kApplyToAllOptionId || shown.taskId == 0u)
            return;
        const bool checked = value != 0u;
        Task* task         = state.FindTask(shown.taskId);
        if (task != nullptr && SetApplyToAllChecked(*task, shown.generation, checked))
            shownApplyToAllChecked = checked;
        // The card re-renders from the flag; a refused change (the prompt moved on) re-presents the mirror.
        state.RequestTaskPresentationRefresh();
    }

    // F6 sent the focus back to the pane: nothing was answered and the mirror stays, but it never takes the
    // focus again for this prompt (a click on it still gives it the focus).
    void OnFocusReturned() noexcept
    {
        if (shown.taskId == 0u)
            return;
        Remember(focusDeclined, shown);
        ++focusReturnCount;
        // The pane's top-level window is active again (a self-test's stand-in follows).
        if (overlay && debugForegroundOverride == std::optional<HWND>{overlay->GetHwnd()})
            debugForegroundOverride = OwnerRoot();
    }

    void ShowOverflowMenu() noexcept
    {
        if (! overlay || ! overlay->GetHwnd() || shownOverflowActionCount == 0u)
            return;
        std::weak_ptr<void> lifetime;
        {
            std::scoped_lock lock(state._mutex);
            lifetime = state._uiLifetime;
        }
        constexpr int kCommandBase = 1;
        std::vector<DxUi::MenuFlyoutItem> items;
        std::vector<ConflictAction> actions;
        items.reserve(shownOverflowActionCount);
        actions.reserve(shownOverflowActionCount);
        for (size_t index = 0u; index < shownOverflowActionCount; ++index)
        {
            DxUi::MenuFlyoutItem item{};
            item.text      = FileOperations::FormatConflictActionLabel(shownOverflowActions[index], shownBucket, shownOperation);
            item.commandId = kCommandBase + static_cast<int>(index);
            items.push_back(std::move(item));
            actions.push_back(shownOverflowActions[index]);
        }
        POINT point{};
        if (const auto rect = overlay->GetButtonScreenRect(kMoreButtonId); rect.has_value())
        {
            point = {rect->left, rect->bottom};
        }
        else
        {
            // No laid-out button: the point the click or key was delivered at, never the live cursor.
            const LPARAM messagePos = static_cast<LPARAM>(GetMessagePos());
            point                   = {GET_X_LPARAM(messagePos), GET_Y_LPARAM(messagePos)};
        }
        const PromptKey key = shown;
        static_cast<void>(DxUi::ContextMenu::ShowAsync(overlay->GetHwnd(),
                                                       point,
                                                       items,
                                                       MakeAppThemeDxPalette(Owner().GetTheme()),
                                                       [this, lifetime, key, actions = std::move(actions)](std::optional<int> chosen) noexcept
        {
            // The same prompt must still be on the mirror; the broker rejects anything stale anyway.
            if (! chosen.has_value() || lifetime.expired() || shown != key || ! OverlayShown())
                return;
            const int index = chosen.value() - kCommandBase;
            if (index < 0 || static_cast<size_t>(index) >= actions.size())
                return;
            const FolderWindow::FileOperationPromptDispatchScope dispatch(Owner());
            Submit(actions[static_cast<size_t>(index)], Placement::Overflow);
        }));
    }

    static void OnButtonThunk(void* context, uint32_t buttonId) noexcept
    {
        if (auto* mirror = static_cast<DecisionMirror*>(context))
            mirror->OnButton(buttonId);
    }

    static void OnCloseThunk(void* context) noexcept
    {
        if (auto* mirror = static_cast<DecisionMirror*>(context))
            mirror->OnClose();
    }

    static void OnOptionChangedThunk(void* context, uint32_t optionId, uint64_t value) noexcept
    {
        if (auto* mirror = static_cast<DecisionMirror*>(context))
            mirror->OnOptionChanged(optionId, value);
    }

    static void OnFocusReturnedThunk(void* context) noexcept
    {
        if (auto* mirror = static_cast<DecisionMirror*>(context))
            mirror->OnFocusReturned();
    }

    // The pane's view is gone: the overlay already hid itself; the next reconcile anchors anew.
    static void OnAnchorDestroyedThunk(void* context, RedSalamander::Ui::AlertOverlayWindow& /*overlay*/) noexcept
    {
        if (auto* mirror = static_cast<DecisionMirror*>(context))
        {
            mirror->shown       = {};
            mirror->shownAnchor = nullptr;
            mirror->shownText.clear();
            mirror->tookFocus = false;
            mirror->announced = false;
        }
    }

    // ---- main-window indicator (R8.9) -------------------------------------------------------------------

    // The pending set briefly empties between two prompts of one run (an answer is consumed before the
    // next prompt's details are read): the flash and the status segment follow the run, not each prompt.
    static constexpr ULONGLONG kIndicatorSettleMs = 1000u;

    // Returns the delay after which the indicator must be looked at again (the set is empty but has not
    // stayed empty long enough to end the run), or nullopt.
    [[nodiscard]] std::optional<UINT> UpdateIndicator() noexcept
    {
        const unsigned long count = pendingCount;
        if (count > 0u)
        {
            // D-8.1: once per run of decisions, never per prompt.
            if (! indicatorLatched)
                FlashPopup();
            indicatorLatched        = true;
            indicatorEmptySinceTick = 0u;
            Owner().SetPaneStatusDecisionText(FormatStringResource(nullptr, IDS_FMT_FILEOPS_DECISION_PENDING, count));
            return std::nullopt;
        }
        if (! indicatorLatched)
        {
            Owner().SetPaneStatusDecisionText({});
            return std::nullopt;
        }
        const ULONGLONG now = GetTickCount64();
        if (indicatorEmptySinceTick == 0u)
            indicatorEmptySinceTick = now;
        const ULONGLONG emptyFor = now - indicatorEmptySinceTick;
        if (emptyFor >= kIndicatorSettleMs)
        {
            indicatorLatched        = false;
            indicatorEmptySinceTick = 0u;
            Owner().SetPaneStatusDecisionText({});
            return std::nullopt;
        }
        // Between two prompts: the segment keeps its text until the set stayed empty for the settle time.
        return static_cast<UINT>(kIndicatorSettleMs - emptyFor);
    }

    void FlashPopup() noexcept
    {
        HWND popup = nullptr;
        {
            std::scoped_lock lock(state._mutex);
            popup = state._popup.get();
        }
        if (! popup || IsWindowVisible(popup) == FALSE)
            return;
        // Only while the user works in the owner window (the mirror that took the focus counts) and the
        // popup is not the foreground window.
        if (! OwnerRootIsForeground(Owner().GetHwnd()) || ForegroundWindow() == popup)
            return;
        ++flashCount;
        if (debugForegroundOverride.has_value())
            return; // a self-test stands in for the foreground; it never flashes the real taskbar
        FLASHWINFO info{};
        info.cbSize  = sizeof(info);
        info.hwnd    = popup;
        info.dwFlags = FLASHW_TRAY | FLASHW_TIMERNOFG;
        static_cast<void>(FlashWindowEx(&info));
    }
};

void FolderWindow::FileOperationState::DecisionMirrorDeleter::operator()(DecisionMirror* mirror) const noexcept
{
    std::default_delete<DecisionMirror>{}(mirror);
}

FolderWindow::FileOperationState::DecisionMirror& FolderWindow::FileOperationState::EnsureDecisionMirror() noexcept
{
    if (! _decisionMirror)
    {
        auto created = std::make_unique<DecisionMirror>(*this);
        _decisionMirror.reset(created.release());
    }
    return *_decisionMirror;
}

void FolderWindow::FileOperationState::ReconcileDecisionMirror() noexcept
{
    if (GetWindowThreadProcessId(_owner.GetHwnd(), nullptr) != GetCurrentThreadId())
        return;
#ifdef ENABLE_TESTS
    if (_decisionMirror && _decisionMirror->debugReconcilePaused)
        return;
#endif
    EnsureDecisionMirror().Reconcile();
}

std::optional<UINT> FolderWindow::FileOperationState::UpdateDecisionPendingIndicator() noexcept
{
    if (GetWindowThreadProcessId(_owner.GetHwnd(), nullptr) != GetCurrentThreadId())
        return std::nullopt;
    return EnsureDecisionMirror().UpdateIndicator();
}

void FolderWindow::FileOperationState::NotifyDecisionMirrorPaneChanged() noexcept
{
    // Only while a decision waits: an ordinary navigation posts nothing.
    if (_decisionMirror && ! _completionShutdown.load(std::memory_order_acquire) && (_decisionMirror->pendingCount > 0u || _decisionMirror->OverlayShown()))
        RequestTaskPresentationRefresh();
}

FolderWindow::FileOperationState::PromptAnnouncementClaim FolderWindow::FileOperationState::ClaimPromptAnnouncementForCard(uint64_t taskId,
                                                                                                                           uint64_t generation) noexcept
{
    // The mirror reconciles only on the owner thread and never during completion shutdown: the card owns
    // the announcement whenever the mirror will not decide about it.
    if (GetWindowThreadProcessId(_owner.GetHwnd(), nullptr) != GetCurrentThreadId() || _completionShutdown.load(std::memory_order_acquire))
        return PromptAnnouncementClaim::Announce;
    return EnsureDecisionMirror().ClaimForCard({taskId, generation});
}

bool FolderWindow::FileOperationState::GetMirrorDecisionsInPane() const noexcept
{
    return ! _owner._settings || ! _owner._settings->fileOperations.has_value() || _owner._settings->fileOperations->mirrorDecisionsInPane;
}

HWND FolderWindow::FileOperationState::IngressPromptAnchor(FolderWindow::Pane pane) const noexcept
{
    // R8.10: a pre-task prompt (no task exists yet) stays a blocking prompt, anchored on the pane the
    // request came from, so it covers that pane instead of the whole window. HOST_ALERT_SCOPE_PANE would
    // resolve through the application Folder window and miss a Compare window's panes.
    const HWND folderView = _owner.GetFolderViewHwnd(pane);
    if (folderView && IsWindowVisible(folderView) != FALSE)
        return folderView;
    return _owner.GetHwnd();
}

void FolderWindow::FileOperationState::ShutdownDecisionMirror() noexcept
{
    if (! _decisionMirror)
        return;
    DecisionMirror& mirror = *_decisionMirror;
    if (mirror.overlay)
    {
        mirror.overlay->SetTrackedWindowDestroyedCallback(nullptr, nullptr);
        mirror.overlay->ClearCallbacks();
        mirror.overlay->HideWithoutRestoringFocus();
    }
    mirror.shown        = {};
    mirror.shownAnchor  = nullptr;
    mirror.pendingCount = 0u;
    // The pane status bars may outlive this state (a deferred reset): leave them without a segment.
    if (GetWindowThreadProcessId(_owner.GetHwnd(), nullptr) == GetCurrentThreadId())
        _owner.SetPaneStatusDecisionText({});
    mirror.indicatorLatched        = false;
    mirror.indicatorEmptySinceTick = 0u;
}

void FolderWindow::FileOperationState::NoteDecisionMirrorInspection(uint64_t taskId, uint64_t generation) noexcept
{
    DecisionMirror& mirror = EnsureDecisionMirror();
    Remember(mirror.suppressed, {taskId, generation});
    if (mirror.shown == PromptKey{taskId, generation})
        mirror.Hide(true);
}

bool FolderWindow::FileOperationState::Task::HasMutated() noexcept
{
    // Every selected root starts Preparing; a per-item dispatch or the bulk gate moves it to
    // MutationPossible and nothing moves it back.
    std::scoped_lock lock(_sourceItemStatusMutex);
    return std::ranges::any_of(_sourceItemResultBuilders,
                               [](const SourceItemResultBuilder& builder) noexcept { return builder.phase == SourceItemExecutionPhase::MutationPossible; });
}

void FolderWindow::NotifyFileOperationsPaneChanged() noexcept
{
    if (_fileOperations)
        _fileOperations->NotifyDecisionMirrorPaneChanged();
}

#ifdef ENABLE_TESTS
bool FolderWindow::FileOperationState::DebugGetDecisionMirrorSnapshot(DecisionMirrorDebugSnapshot& out) noexcept
{
    out = {};
    if (! _decisionMirror)
        return false;
    const DecisionMirror& mirror = *_decisionMirror;
    out.visible                  = mirror.OverlayShown();
    out.windowVisible            = mirror.overlay && mirror.overlay->GetHwnd() && IsWindowVisible(mirror.overlay->GetHwnd()) != FALSE;
    out.tookFocus                = mirror.tookFocus;
    out.announced                = mirror.announced;
    out.taskId                   = mirror.shown.taskId;
    out.generation               = mirror.shown.generation;
    out.pane                     = mirror.shownPane;
    out.overlay                  = mirror.overlay ? mirror.overlay->GetHwnd() : nullptr;
    out.anchor                   = mirror.shownAnchor;
    out.pendingCount             = mirror.pendingCount;
    out.showCount                = mirror.showCount;
    out.submitCount              = mirror.submitCount;
    out.acceptedCount            = mirror.acceptedCount;
    out.flashCount               = mirror.flashCount;
    out.reconcileCount           = mirror.reconcileCount;
    out.indicatorLatched         = mirror.indicatorLatched;
    out.foreground               = mirror.ForegroundWindow();
    out.showFailureCount         = mirror.showFailureCount;
    out.focusReturnCount         = mirror.focusReturnCount;
    out.cardDeferCount           = mirror.cardDeferCount;
    out.evaluatedCount           = mirror.evaluated.size();
    out.focusDeclined            = mirror.shown.taskId != 0u && Remembered(mirror.focusDeclined, mirror.shown);
    return true;
}

void FolderWindow::FileOperationState::DebugSetDecisionMirrorForegroundOverride(std::optional<HWND> foreground) noexcept
{
    EnsureDecisionMirror().debugForegroundOverride = foreground;
}

void FolderWindow::FileOperationState::DebugFailNextDecisionMirrorShows(uint32_t count) noexcept
{
    EnsureDecisionMirror().debugFailShows = count;
}

void FolderWindow::FileOperationState::DebugSetDecisionMirrorReconcilePaused(bool paused) noexcept
{
    EnsureDecisionMirror().debugReconcilePaused = paused;
}

std::optional<bool> FolderWindow::FileOperationState::DebugGetPromptAnnouncer(uint64_t taskId, uint64_t generation) noexcept
{
    const DecisionMirror::AnnouncementClaim* claim = EnsureDecisionMirror().FindClaim({taskId, generation});
    return claim != nullptr ? std::optional<bool>(claim->byMirror) : std::nullopt;
}

uint32_t FolderWindow::FileOperationState::DebugDecisionMirrorPrimaryButtonId(Task::ConflictAction action) noexcept
{
    return kPrimaryButtonBase + static_cast<uint32_t>(action);
}

uint32_t FolderWindow::FileOperationState::DebugDecisionMirrorMoreButtonId() noexcept
{
    return kMoreButtonId;
}

uint32_t FolderWindow::FileOperationState::DebugDecisionMirrorEscapeButtonId() noexcept
{
    return kEscapeButtonId;
}

uint32_t FolderWindow::FileOperationState::DebugDecisionMirrorApplyToAllOptionId() noexcept
{
    return kApplyToAllOptionId;
}
#endif
