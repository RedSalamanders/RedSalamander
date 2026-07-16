#include "DxUi.Internal.h"
#include "Helpers.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cmath>
#include <format>
#include <limits>
#include <memory>
#include <mutex>
#include <new>
#include <optional>
#include <span>
#include <string>
#include <utility>
#include <unordered_set>
#include <vector>

#include <UIAutomation.h>
#include <oleauto.h>

#pragma comment(lib, "uiautomationcore.lib")

namespace RedSalamander::DxUi
{
namespace
{
constexpr PCWSTR kWindowHostPropName                       = L"RedSalamander.DxUi.WindowHost";
constexpr uint32_t kAccessibilityMaxDepth                  = 16u;
constexpr UINT kWindowHostAccessibilityActionMessage       = WM_APP + 0x6A;
#if defined(ENABLE_TESTS)
constexpr UINT kWindowHostAccessibilityCreateProviderMessage = WM_APP + 0x6B;
#endif
constexpr DWORD kAccessibilityUiActionDispatchTimeoutMs    = 5000u;
constexpr LONG kAccessibilityRuntimeIdTreeItem             = 1'001;
constexpr LONG kAccessibilityRuntimeIdGridRow              = 1'002;
constexpr LONG kAccessibilityRuntimeIdGridCell             = 1'003;
constexpr LONG kAccessibilityRuntimeIdGridHeader           = 1'004;
constexpr LONG kAccessibilityRuntimeIdPasswordRevealButton = 1'005;
constexpr size_t kAccessibilityMaxRuntimeIdValueCount      = kAccessibilityMaxDepth + 6u;
constexpr size_t kAccessibilityMaxMaterializedOffscreenSelectedRows = 256u;

class AccessibilityProvider;
class AccessibilityTextRangeProvider;

enum class AccessibilityUiActionKind : uint8_t
{
    SetFocus,
    Invoke,
    Toggle,
    SetStringValue,
    SetRangeValue,
    Select,
    AddToSelection,
    RemoveFromSelection,
    Expand,
    Collapse,
    MoveTextRangeByVisualLine,
    MoveTextRangeEndpointByVisualLine,
    ResolveTextRangeBounds,
    ResolveTextRangeFromPoint,
};

struct AccessibilityUiActionRequest
{
    AccessibilityProvider* provider                   = nullptr;
    AccessibilityTextRangeProvider* textRangeProvider = nullptr;
    AccessibilityUiActionKind kind                    = AccessibilityUiActionKind::Invoke;
    std::wstring stringValue{};
    double numberValue                         = 0.0;
    size_t textRangeStart                      = 0u;
    size_t textRangeEnd                        = 0u;
    TextPatternRangeEndpoint textRangeEndpoint = TextPatternRangeEndpoint_Start;
    UiaPoint textRangePoint{};
    int textRangeMoveCount                       = 0;
    int textRangeMoved                           = 0;
    size_t textRangeResultStart                  = 0u;
    size_t textRangeResultEnd                    = 0u;
    size_t textRangeTextLength                   = 0u;
    std::vector<D2D1_RECT_F> textRangeBoundsDip;
    float textRangeDipToPixelScale               = 1.0f;
    HRESULT result                               = static_cast<HRESULT>(UIA_E_NOTSUPPORTED);
};

struct AccessibilityUiActionDispatch
{
    AccessibilityUiActionDispatch()                                                 = default;
    AccessibilityUiActionDispatch(const AccessibilityUiActionDispatch&)             = delete;
    AccessibilityUiActionDispatch& operator=(const AccessibilityUiActionDispatch&)  = delete;
    AccessibilityUiActionDispatch(AccessibilityUiActionDispatch&&)                  = delete;
    AccessibilityUiActionDispatch& operator=(AccessibilityUiActionDispatch&&)       = delete;

    AccessibilityUiActionRequest request;
    wil::com_ptr_nothrow<IRawElementProviderSimple> providerKeepAlive;
    wil::com_ptr_nothrow<ITextRangeProvider> textRangeProviderKeepAlive;
    wil::unique_event_nothrow completedEvent;

    enum class State : uint8_t
    {
        Pending,
        Taken,
        Abandoned,
    };
    std::atomic<State> state{State::Pending};
};

struct AccessibilityUiActionPayload
{
    AccessibilityUiActionPayload()                                              = default;
    AccessibilityUiActionPayload(const AccessibilityUiActionPayload&)            = delete;
    AccessibilityUiActionPayload& operator=(const AccessibilityUiActionPayload&) = delete;
    AccessibilityUiActionPayload(AccessibilityUiActionPayload&&)                 = delete;
    AccessibilityUiActionPayload& operator=(AccessibilityUiActionPayload&&)      = delete;

    ~AccessibilityUiActionPayload() noexcept
    {
        if (! dispatch)
        {
            return;
        }

        AccessibilityUiActionDispatch::State expected = AccessibilityUiActionDispatch::State::Pending;
        if (dispatch->state.compare_exchange_strong(
                expected, AccessibilityUiActionDispatch::State::Abandoned, std::memory_order_acq_rel, std::memory_order_acquire))
        {
            dispatch->request.result = HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
        if (dispatch->completedEvent)
        {
            static_cast<void>(::SetEvent(dispatch->completedEvent.get()));
        }
    }

    std::shared_ptr<AccessibilityUiActionDispatch> dispatch;
};

#if defined(ENABLE_TESTS)
std::atomic<HANDLE> g_accessibilityUiActionHandlerEnteredEvent{nullptr};
std::atomic<HANDLE> g_accessibilityUiActionHandlerReleaseEvent{nullptr};
std::atomic<HANDLE> g_accessibilityUiActionHandlerTakenEnteredEvent{nullptr};
std::atomic<HANDLE> g_accessibilityUiActionHandlerTakenReleaseEvent{nullptr};
std::atomic<HANDLE> g_accessibilityUiActionPostedEvent{nullptr};
std::atomic<DWORD> g_accessibilityUiActionDispatchTimeoutOverrideMs{0u};
std::atomic<uint32_t> g_accessibilityUiActionExecutionCount{0u};
std::atomic<size_t> g_accessibilityOffscreenSelectedRowMaterializationLimitOverride{0u};

[[nodiscard]] DWORD AccessibilityUiActionDispatchTimeoutMs() noexcept
{
    const DWORD timeoutOverride = g_accessibilityUiActionDispatchTimeoutOverrideMs.load(std::memory_order_acquire);
    return timeoutOverride != 0u ? timeoutOverride : kAccessibilityUiActionDispatchTimeoutMs;
}

[[nodiscard]] size_t AccessibilityOffscreenSelectedRowMaterializationLimit() noexcept
{
    const size_t limitOverride = g_accessibilityOffscreenSelectedRowMaterializationLimitOverride.load(std::memory_order_acquire);
    return limitOverride != 0u ? limitOverride : kAccessibilityMaxMaterializedOffscreenSelectedRows;
}

[[nodiscard]] std::wstring_view AccessibilityGridSnapshotPerfDetail() noexcept
{
    const size_t limitOverride = g_accessibilityOffscreenSelectedRowMaterializationLimitOverride.load(std::memory_order_acquire);
    if (limitOverride > kAccessibilityMaxMaterializedOffscreenSelectedRows)
    {
        return L"selection-baseline-unbounded";
    }
    if (limitOverride != 0u)
    {
        return L"selection-candidate-capped";
    }
    return L"selection";
}

void MaybeStallAccessibilityUiActionHandlerForTest() noexcept
{
    const HANDLE enteredEvent = g_accessibilityUiActionHandlerEnteredEvent.load(std::memory_order_acquire);
    const HANDLE releaseEvent = g_accessibilityUiActionHandlerReleaseEvent.load(std::memory_order_acquire);
    if (enteredEvent)
    {
        static_cast<void>(::SetEvent(enteredEvent));
    }
    if (releaseEvent)
    {
        static_cast<void>(::WaitForSingleObject(releaseEvent, AccessibilityUiActionDispatchTimeoutMs() + 2000u));
    }
}

void MaybeStallTakenAccessibilityUiActionHandlerForTest() noexcept
{
    const HANDLE enteredEvent = g_accessibilityUiActionHandlerTakenEnteredEvent.load(std::memory_order_acquire);
    const HANDLE releaseEvent = g_accessibilityUiActionHandlerTakenReleaseEvent.load(std::memory_order_acquire);
    if (enteredEvent)
    {
        static_cast<void>(::SetEvent(enteredEvent));
    }
    if (releaseEvent)
    {
        static_cast<void>(::WaitForSingleObject(releaseEvent, AccessibilityUiActionDispatchTimeoutMs() + 2000u));
    }
}
#else
[[nodiscard]] constexpr DWORD AccessibilityUiActionDispatchTimeoutMs() noexcept
{
    return kAccessibilityUiActionDispatchTimeoutMs;
}


[[nodiscard]] constexpr size_t AccessibilityOffscreenSelectedRowMaterializationLimit() noexcept
{
    return kAccessibilityMaxMaterializedOffscreenSelectedRows;
}

[[nodiscard]] constexpr std::wstring_view AccessibilityGridSnapshotPerfDetail() noexcept
{
    return L"selection";
}
#endif

struct ControlPath
{
    uint32_t depth = 0u;
    std::array<uint16_t, kAccessibilityMaxDepth> indices{};
};

[[nodiscard]] bool AreControlPathsEqual(const ControlPath& left, const ControlPath& right) noexcept
{
    if (left.depth != right.depth)
    {
        return false;
    }

    for (uint32_t index = 0u; index < left.depth; ++index)
    {
        if (left.indices[index] != right.indices[index])
        {
            return false;
        }
    }
    return true;
}

enum class AccessibilityFragmentKind : uint8_t
{
    Root,
    Control,
    TreeItem,
    GridHeader,
    GridRow,
    GridCell,
    TextFieldPasswordRevealButton,
};

struct AccessibilityFocusedFragmentSnapshot
{
    AccessibilityFragmentKind kind = AccessibilityFragmentKind::Control;
    ControlPath path{};
    size_t treeVisibleIndex = 0u;
    uint64_t gridRowId      = 0u;
};

struct AccessibilityPointHitSnapshot
{
    AccessibilityFragmentKind kind = AccessibilityFragmentKind::Control;
    ControlPath path{};
    D2D1_RECT_F hitRectDip  = D2D1::RectF();
    size_t treeVisibleIndex = 0u;
    uint64_t gridRowId      = 0u;
    size_t gridColumnIndex  = 0u;
};

struct AccessibilityPointHitBuildContext
{
    D2D1_POINT_2F translationDip = D2D1::Point2F();
    std::optional<D2D1_RECT_F> clipRectDip;
};

struct AccessibilityTreeItemSnapshotRecord
{
    size_t visibleIndex = 0u;
    uint64_t itemId     = 0u;
    std::wstring text;
    size_t depth     = 0u;
    bool hasChildren = false;
    bool expanded    = false;
};

struct AccessibilityGridHeaderSnapshotRecord
{
    size_t columnIndex = 0u;
    std::wstring gridHeaderName;
};

struct AccessibilityGridRowSnapshotRecord
{
    size_t rowIndex = 0u;
    uint64_t rowId  = 0u;
    std::wstring gridRowAccessibleName;
    bool gridRowOffscreen = true;
};

struct AccessibilityGridCellStateSnapshotRecord
{
    size_t rowIndex                     = 0u;
    uint64_t rowId                      = 0u;
    size_t columnIndex                  = 0u;
    CONTROLTYPEID gridCellControlTypeId = UIA_TextControlTypeId;
    std::wstring gridCellAccessibleText;
    std::wstring gridCellHelpText;
    bool gridCellEnabled            = false;
    bool gridCellOffscreen          = true;
    bool gridCellChecked            = false;
    bool gridCellSupportsToggle     = false;
    bool gridCellSupportsValue      = false;
    bool gridCellSupportsRangeValue = false;
    double gridCellRangeValue       = 0.0;
};

struct TextRangeSpan
{
    size_t start = 0u;
    size_t end   = 0u;
};

struct AccessibilityControlNavigationSnapshot
{
    ControlPath path{};
    CONTROLTYPEID controlTypeId = UIA_TextControlTypeId;
    std::wstring controlAccessibleName;
    std::wstring controlAccessibleHelpText;
    std::wstring controlAccessibleValue;
    std::wstring controlAccessibleText;
    bool controlVisible              = false;
    bool controlEnabled              = false;
    bool controlFocusable            = false;
    bool controlHasFocus             = false;
    bool controlIsPassword           = false;
    bool controlSupportsInvoke       = false;
    bool controlSupportsToggle       = false;
    bool controlSupportsValue        = false;
    bool controlSupportsText         = false;
    bool controlSupportsRangeValue   = false;
    bool controlSupportsSelection    = false;
    bool controlSupportsTable        = false;
    bool controlValueReadOnly        = true;
    bool controlToggleChecked        = false;
    size_t controlTextSelectionStart = 0u;
    size_t controlTextSelectionEnd   = 0u;
    std::vector<D2D1_RECT_F> controlTextSelectionBoundsDip;
    std::optional<size_t> controlTextCompositionStart;
    std::optional<size_t> controlTextCompositionEnd;
    std::optional<size_t> controlTextConversionTargetStart;
    std::optional<size_t> controlTextConversionTargetEnd;
    double controlRangeValue         = 0.0;
    double controlRangeMinimum       = 0.0;
    double controlRangeMaximum       = 0.0;
    double controlRangeSmallChange   = 0.0;
    double controlRangeLargeChange   = 0.0;
    bool isGrid                      = false;
    bool isTree                      = false;
    bool gridCanSelectMultiple       = false;
    bool hasPasswordRevealButton     = false;
    bool passwordRevealButtonEnabled = false;
    bool treeIsEnabled               = false;
    bool treeHasFocus                = false;
    bool gridIsEnabled               = false;
    bool gridHasFocus                = false;
    size_t gridRowCount              = 0u;
    size_t gridColumnCount           = 0u;
    size_t treeVisibleItemCount      = 0u;
    std::wstring passwordRevealButtonAccessibleName;
    std::optional<size_t> selectedTreeVisibleIndex;
    std::vector<AccessibilityTreeItemSnapshotRecord> treeItems;
    std::vector<AccessibilityGridHeaderSnapshotRecord> gridHeaders;
    std::vector<AccessibilityGridRowSnapshotRecord> gridRows;
    std::vector<AccessibilityGridCellStateSnapshotRecord> gridCells;
    std::vector<size_t> gridVisibleColumns;
    std::vector<size_t> gridVisibleRows;
    std::vector<uint64_t> gridVisibleRowIds;
    std::vector<uint64_t> selectedGridRowIds;
};

struct AccessibilityNavigationTarget
{
    AccessibilityFragmentKind kind = AccessibilityFragmentKind::Root;
    ControlPath path{};
    size_t treeVisibleIndex = 0u;
    uint64_t gridRowId      = 0u;
    size_t gridColumnIndex  = 0u;
};

struct AccessibilityGridCellSnapshotRecord
{
    const AccessibilityControlNavigationSnapshot* controlRecord = nullptr;
    const AccessibilityGridCellStateSnapshotRecord* cellRecord  = nullptr;
    size_t rowIndex                                             = 0u;
    size_t columnIndex                                          = 0u;
};

struct AccessibilitySnapshot
{
    HWND hwnd              = nullptr;
    DWORD buildThreadId    = 0u;
    DWORD windowThreadId   = 0u;
    bool alive             = false;
    bool hasRetainedRoot   = false;
    bool hasCollapsedSemanticRoot = false;
    float pixelsToDipScale = 1.0f;
    std::wstring windowName;
    std::optional<AccessibilityFocusedFragmentSnapshot> focusedFragment;
    std::vector<AccessibilityPointHitSnapshot> pointHitRecords;
    std::vector<ControlPath> semanticControlOrder;
    std::vector<AccessibilityControlNavigationSnapshot> controlNavigationRecords;
};

[[nodiscard]] bool IsSemanticAccessibilityControl(const Control* control) noexcept;
[[nodiscard]] bool TryResolveSingleSemanticRootControlPath(const Control* root, ControlPath& outPath) noexcept;
[[nodiscard]] bool SnapshotHasCollapsedSemanticRoot(const AccessibilitySnapshot& snapshot) noexcept;
[[nodiscard]] std::wstring_view GetControlAccessibleName(const Control* root, const Control* control) noexcept;
[[nodiscard]] std::wstring_view GetControlAccessibleValue(const Control* control) noexcept;
[[nodiscard]] std::wstring GetControlAccessibleTextRangeText(const Control* control);
[[nodiscard]] TextRangeSpan GetControlAccessibleSelectionRange(const WindowHost* host, const Control* control, size_t textLength) noexcept;
[[nodiscard]] TextRangeSpan GetTextRangeSpanFromState(size_t caretIndex, std::optional<size_t> selectionAnchorIndex, size_t textLength) noexcept;
[[nodiscard]] D2D1_RECT_F ResolveTextPatternViewportRect(const Control* control) noexcept;
[[nodiscard]] std::optional<std::vector<D2D1_RECT_F>> TryResolveTextRangeCaretRects(const WindowHost& host,
                                                                                    const Control& control,
                                                                                    std::wstring_view text,
                                                                                    const TextRangeSpan& range);
[[nodiscard]] std::wstring BuildGridCellAccessibleText(const GridCellData& cellData);
[[nodiscard]] std::wstring_view GetGridHeaderAccessibleName(const GridColumnDesc& columnDesc) noexcept;
[[nodiscard]] CONTROLTYPEID GetGridCellControlTypeId(const GridCellData& cellData) noexcept;
[[nodiscard]] CONTROLTYPEID GetControlTypeId(const Control* control) noexcept;
[[nodiscard]] bool SupportsInvokePattern(const Control* control) noexcept;
[[nodiscard]] bool SupportsTogglePattern(const Control* control) noexcept;
[[nodiscard]] bool SupportsValuePattern(const Control* control) noexcept;
[[nodiscard]] bool SupportsTextPattern(const Control* control) noexcept;
[[nodiscard]] bool SupportsRangeValuePattern(const Control* control) noexcept;
[[nodiscard]] bool IsValueReadOnly(const Control* control) noexcept;
[[nodiscard]] bool GridCellSupportsTogglePattern(const GridCellData& cellData) noexcept;
[[nodiscard]] bool GridCellSupportsValuePattern(const GridCellData& cellData) noexcept;
[[nodiscard]] bool GridCellSupportsRangeValuePattern(const GridCellData& cellData) noexcept;
[[nodiscard]] double GetGridCellRangeValue(const GridCellData& cellData) noexcept;
[[nodiscard]] bool FindAccessibilityPathForTarget(const Control* current, const ControlPath& basePath, const Control* target, ControlPath& outPath) noexcept;
void AppendAccessibilitySnapshotPointHits(WindowHost& host, const Control* current, const ControlPath& basePath, AccessibilitySnapshot& snapshot);
void AppendAccessibilitySnapshotPointHits(WindowHost& host,
                                          const Control* current,
                                          const ControlPath& basePath,
                                          AccessibilitySnapshot& snapshot,
                                          const AccessibilityPointHitBuildContext& context);
void AppendAccessibilitySnapshotNavigation(
    WindowHost& host, const Control* root, const Control* current, const ControlPath& basePath, AccessibilitySnapshot& snapshot);
[[nodiscard]] const AccessibilityPointHitSnapshot* FindSnapshotPointHit(const AccessibilitySnapshot& snapshot, D2D1_POINT_2F pointDip) noexcept;
[[nodiscard]] std::optional<D2D1_RECT_F> FindSnapshotFragmentBounds(const AccessibilitySnapshot& snapshot,
                                                                    AccessibilityFragmentKind kind,
                                                                    const ControlPath& path,
                                                                    size_t treeVisibleIndex,
                                                                    uint64_t gridRowId,
                                                                    size_t gridColumnIndex) noexcept;
[[nodiscard]] std::optional<AccessibilityNavigationTarget> ResolveSnapshotNavigationTarget(const AccessibilitySnapshot& snapshot,
                                                                                           AccessibilityFragmentKind kind,
                                                                                           const ControlPath& path,
                                                                                           size_t treeVisibleIndex,
                                                                                           uint64_t gridRowId,
                                                                                           size_t gridColumnIndex,
                                                                                           NavigateDirection direction) noexcept;

struct WindowHostAccessibilityTarget final
{
    explicit WindowHostAccessibilityTarget(HWND hwnd, WindowHost* host) noexcept : hwnd(hwnd), host(host)
    {
    }

    WindowHostAccessibilityTarget(const WindowHostAccessibilityTarget&)            = delete;
    WindowHostAccessibilityTarget& operator=(const WindowHostAccessibilityTarget&) = delete;
    WindowHostAccessibilityTarget(WindowHostAccessibilityTarget&&)                 = delete;
    WindowHostAccessibilityTarget& operator=(WindowHostAccessibilityTarget&&)      = delete;

    [[nodiscard]] ULONG AddRef() noexcept
    {
        return _referenceCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
    }

    [[nodiscard]] ULONG Release() noexcept
    {
        const ULONG remaining = _referenceCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        if (remaining == 0u)
        {
            delete this;
        }
        return remaining;
    }

    [[nodiscard]] WindowHost* ResolveHost() const noexcept
    {
        if (! hwnd || IsWindow(hwnd) == FALSE)
        {
            return nullptr;
        }

        WindowHost* const resolvedHost = host.load(std::memory_order_acquire);
        if (! resolvedHost)
        {
            return nullptr;
        }

        const DWORD windowThreadId = GetWindowThreadProcessId(hwnd, nullptr);
        if (windowThreadId != 0u && windowThreadId != GetCurrentThreadId())
        {
            return nullptr;
        }

        return resolvedHost;
    }

    std::atomic<ULONG> _referenceCount{1u};
    HWND hwnd = nullptr;
    std::atomic<WindowHost*> host{nullptr};
    std::atomic<std::shared_ptr<const AccessibilitySnapshot>> snapshot;
};

[[nodiscard]] std::recursive_mutex& GetAccessibilityTargetMutex() noexcept
{
    static std::recursive_mutex mutex;
    return mutex;
}

[[nodiscard]] std::shared_ptr<const AccessibilitySnapshot> MakeEmptyAccessibilitySnapshot(HWND hwnd)
{
    auto snapshot            = std::make_shared<AccessibilitySnapshot>();
    snapshot->hwnd           = hwnd;
    snapshot->buildThreadId  = GetCurrentThreadId();
    snapshot->windowThreadId = hwnd ? GetWindowThreadProcessId(hwnd, nullptr) : 0u;
    snapshot->alive          = false;
    return snapshot;
}

void PublishAccessibilitySnapshot(WindowHostAccessibilityTarget& target, std::shared_ptr<const AccessibilitySnapshot> snapshot) noexcept
{
    target.snapshot.store(std::move(snapshot), std::memory_order_release);
}

void PublishEmptyAccessibilitySnapshot(WindowHostAccessibilityTarget& target) noexcept
{
    PublishAccessibilitySnapshot(target, MakeEmptyAccessibilitySnapshot(target.hwnd));
}

void PublishWindowHostAccessibilitySnapshot(WindowHostAccessibilityTarget& target, WindowHost& host)
{
    auto snapshot              = std::make_shared<AccessibilitySnapshot>();
    snapshot->hwnd             = target.hwnd;
    snapshot->buildThreadId    = GetCurrentThreadId();
    snapshot->windowThreadId   = target.hwnd ? GetWindowThreadProcessId(target.hwnd, nullptr) : 0u;
    snapshot->alive            = true;
    snapshot->pixelsToDipScale = USER_DEFAULT_SCREEN_DPI / host.GetDpi();
    const Control* const root  = host.GetRoot();
    snapshot->hasRetainedRoot  = root != nullptr;
    AppendAccessibilitySnapshotNavigation(host, root, root, ControlPath{}, *snapshot);
    if (root)
    {
        ControlPath collapsedRootPath{};
        snapshot->hasCollapsedSemanticRoot =
            TryResolveSingleSemanticRootControlPath(root, collapsedRootPath) && snapshot->semanticControlOrder.size() == 1u &&
            AreControlPathsEqual(snapshot->semanticControlOrder.front(), collapsedRootPath);
    }
    AppendAccessibilitySnapshotPointHits(host, root, ControlPath{}, *snapshot);

    Control* const focused = host.GetFocusControl();
    if (root && focused && IsSemanticAccessibilityControl(focused))
    {
        ControlPath focusedPath{};
        if (FindAccessibilityPathForTarget(root, ControlPath{}, focused, focusedPath))
        {
            AccessibilityFocusedFragmentSnapshot focusedFragment{};
            focusedFragment.path = focusedPath;

            if (const auto* tree = dynamic_cast<const Tree*>(focused))
            {
                if (const auto* model = tree->GetModel(); model && tree->GetSelectedItemId())
                {
                    if (const std::optional<size_t> visibleIndex = model->FindVisibleItemById(tree->GetSelectedItemId().value()))
                    {
                        focusedFragment.kind             = AccessibilityFragmentKind::TreeItem;
                        focusedFragment.treeVisibleIndex = visibleIndex.value();
                    }
                }
            }
            else if (const auto* grid = dynamic_cast<const Grid*>(focused))
            {
                if (const auto* model = grid->GetModel(); model)
                {
                    if (const std::optional<size_t> selectedRow = grid->GetPrimarySelectedRow())
                    {
                        focusedFragment.kind      = AccessibilityFragmentKind::GridRow;
                        focusedFragment.gridRowId = model->GetStableRowId(selectedRow.value());
                    }
                }
            }

            snapshot->focusedFragment = focusedFragment;
        }
    }

    wchar_t windowText[128]{};
    const int length = target.hwnd ? GetWindowTextW(target.hwnd, windowText, static_cast<int>(std::size(windowText))) : 0;
    snapshot->windowName.assign(windowText, static_cast<size_t>((std::max)(0, length)));

    PublishAccessibilitySnapshot(target, std::move(snapshot));
}

[[nodiscard]] std::shared_ptr<const AccessibilitySnapshot> CaptureAccessibilitySnapshot(WindowHostAccessibilityTarget* target, HWND hwnd)
{
    if (! target)
    {
        return MakeEmptyAccessibilitySnapshot(hwnd);
    }

    if (auto snapshot = target->snapshot.load(std::memory_order_acquire))
    {
        return snapshot;
    }

    return MakeEmptyAccessibilitySnapshot(hwnd);
}

[[nodiscard]] WindowHostAccessibilityTarget* AcquireWindowHostAccessibilityTarget(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return nullptr;
    }

    const std::scoped_lock lock(GetAccessibilityTargetMutex());
    auto* target = static_cast<WindowHostAccessibilityTarget*>(GetPropW(hwnd, kWindowHostPropName));
    if (! target)
    {
        return nullptr;
    }

    static_cast<void>(target->AddRef());
    return target;
}

[[nodiscard]] bool TryAppendPathIndex(const ControlPath& source, size_t childIndex, ControlPath& outPath) noexcept
{
    if (source.depth >= source.indices.size() || childIndex > (std::numeric_limits<uint16_t>::max)())
    {
        return false;
    }

    outPath                        = source;
    outPath.indices[outPath.depth] = static_cast<uint16_t>(childIndex);
    ++outPath.depth;
    return true;
}

template <typename TControl> [[nodiscard]] TControl* ResolveControlAtPath(TControl* root, const ControlPath& path) noexcept
{
    TControl* current = root;
    for (uint32_t depth = 0u; depth < path.depth; ++depth)
    {
        auto* panel = dynamic_cast<Panel*>(current);
        if (! panel)
        {
            return nullptr;
        }

        const auto children = panel->GetChildren();
        const size_t index  = path.indices[depth];
        if (index >= children.size() || ! children[index])
        {
            return nullptr;
        }

        current = children[index].get();
    }

    return current;
}

[[nodiscard]] bool IsControlPathVisible(const Control* root, const ControlPath& path) noexcept
{
    const Control* current = root;
    if (! current || ! current->IsVisible())
    {
        return false;
    }

    for (uint32_t depth = 0u; depth < path.depth; ++depth)
    {
        const auto* panel = dynamic_cast<const Panel*>(current);
        if (! panel)
        {
            return false;
        }

        const auto children = panel->GetChildren();
        const size_t index  = path.indices[depth];
        if (index >= children.size() || ! children[index])
        {
            return false;
        }

        current = children[index].get();
        if (! current->IsVisible())
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool IsSemanticAccessibilityControl(const Control* control) noexcept
{
    return dynamic_cast<const Label*>(control) != nullptr || dynamic_cast<const Button*>(control) != nullptr ||
           dynamic_cast<const TextField*>(control) != nullptr || dynamic_cast<const ComboBox*>(control) != nullptr ||
           dynamic_cast<const Tree*>(control) != nullptr || dynamic_cast<const Grid*>(control) != nullptr || dynamic_cast<const Slider*>(control) != nullptr ||
           dynamic_cast<const ColorSwatch*>(control) != nullptr;
}

[[nodiscard]] bool FindFirstSemanticControl(const Control* current, const ControlPath& basePath, ControlPath& outPath) noexcept
{
    if (! current || ! current->IsVisible())
    {
        return false;
    }

    if (IsSemanticAccessibilityControl(current))
    {
        outPath = basePath;
        return true;
    }

    auto* panel = dynamic_cast<const Panel*>(current);
    if (! panel)
    {
        return false;
    }

    const auto children = panel->GetChildren();
    for (size_t index = 0u; index < children.size(); ++index)
    {
        if (! children[index])
        {
            continue;
        }

        ControlPath childPath{};
        if (! TryAppendPathIndex(basePath, index, childPath))
        {
            continue;
        }

        if (FindFirstSemanticControl(children[index].get(), childPath, outPath))
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] bool FindNextSemanticControl(const Control* root, const ControlPath& currentPath, ControlPath& outPath) noexcept
{
    if (! root)
    {
        return false;
    }

    ControlPath cursor = currentPath;
    while (cursor.depth != 0u)
    {
        const uint32_t parentDepth = cursor.depth - 1u;
        ControlPath parentPath     = cursor;
        parentPath.depth           = parentDepth;
        const Control* parent      = ResolveControlAtPath(const_cast<Control*>(root), parentPath);
        const auto* panel          = dynamic_cast<const Panel*>(parent);
        if (! panel)
        {
            return false;
        }

        const auto children       = panel->GetChildren();
        const size_t currentIndex = cursor.indices[parentDepth];
        for (size_t siblingIndex = currentIndex + 1u; siblingIndex < children.size(); ++siblingIndex)
        {
            if (! children[siblingIndex])
            {
                continue;
            }

            ControlPath siblingPath{};
            if (! TryAppendPathIndex(parentPath, siblingIndex, siblingPath))
            {
                continue;
            }

            if (FindFirstSemanticControl(children[siblingIndex].get(), siblingPath, outPath))
            {
                return true;
            }
        }

        cursor.depth = parentDepth;
    }

    return false;
}

[[nodiscard]] bool ShouldExposeSingleSemanticRootControl(const Control* control) noexcept
{
    return dynamic_cast<const Checkbox*>(control) != nullptr || dynamic_cast<const Toggle*>(control) != nullptr ||
           dynamic_cast<const Button*>(control) != nullptr || dynamic_cast<const TextField*>(control) != nullptr ||
           dynamic_cast<const ComboBox*>(control) != nullptr || dynamic_cast<const Tree*>(control) != nullptr ||
           dynamic_cast<const Slider*>(control) != nullptr || dynamic_cast<const Grid*>(control) != nullptr ||
           dynamic_cast<const ColorSwatch*>(control) != nullptr;
}

[[nodiscard]] bool TryResolveSingleSemanticRootControlPath(const Control* root, ControlPath& outPath) noexcept
{
    ControlPath firstPath{};
    if (! FindFirstSemanticControl(root, ControlPath{}, firstPath))
    {
        return false;
    }

    const Control* control = ResolveControlAtPath(const_cast<Control*>(root), firstPath);
    if (! ShouldExposeSingleSemanticRootControl(control))
    {
        return false;
    }

    ControlPath nextPath{};
    if (FindNextSemanticControl(root, firstPath, nextPath))
    {
        return false;
    }

    outPath = firstPath;
    return true;
}

[[nodiscard]] bool IsNonEmptyAccessibilityRect(const D2D1_RECT_F& rect) noexcept
{
    return rect.right > rect.left && rect.bottom > rect.top;
}

[[nodiscard]] D2D1_RECT_F TranslateAccessibilityRect(const D2D1_RECT_F& rect, D2D1_POINT_2F translationDip) noexcept
{
    return D2D1::RectF(rect.left + translationDip.x, rect.top + translationDip.y, rect.right + translationDip.x, rect.bottom + translationDip.y);
}

[[nodiscard]] std::optional<D2D1_RECT_F> IntersectAccessibilityRects(const D2D1_RECT_F& left, const D2D1_RECT_F& right) noexcept
{
    const D2D1_RECT_F intersection = D2D1::RectF(
        (std::max)(left.left, right.left), (std::max)(left.top, right.top), (std::min)(left.right, right.right), (std::min)(left.bottom, right.bottom));
    if (! IsNonEmptyAccessibilityRect(intersection))
    {
        return std::nullopt;
    }
    return intersection;
}

[[nodiscard]] std::optional<D2D1_RECT_F> ApplyAccessibilityPointHitContext(const D2D1_RECT_F& rect,
                                                                           const AccessibilityPointHitBuildContext& context) noexcept
{
    const D2D1_RECT_F translated = TranslateAccessibilityRect(rect, context.translationDip);
    if (! IsNonEmptyAccessibilityRect(translated))
    {
        return std::nullopt;
    }
    if (context.clipRectDip)
    {
        return IntersectAccessibilityRects(translated, context.clipRectDip.value());
    }
    return translated;
}

void AppendAccessibilityPointHit(AccessibilitySnapshot& snapshot,
                                 AccessibilityFragmentKind kind,
                                 const ControlPath& path,
                                 const D2D1_RECT_F& hitRectDip,
                                 size_t treeVisibleIndex = 0u,
                                 uint64_t gridRowId      = 0u,
                                 size_t gridColumnIndex  = 0u)
{
    if (! IsNonEmptyAccessibilityRect(hitRectDip))
    {
        return;
    }

    snapshot.pointHitRecords.push_back(AccessibilityPointHitSnapshot{.kind             = kind,
                                                                     .path             = path,
                                                                     .hitRectDip       = hitRectDip,
                                                                     .treeVisibleIndex = treeVisibleIndex,
                                                                     .gridRowId        = gridRowId,
                                                                     .gridColumnIndex  = gridColumnIndex});
}

void AppendTransformedAccessibilityPointHit(AccessibilitySnapshot& snapshot,
                                            AccessibilityFragmentKind kind,
                                            const ControlPath& path,
                                            const D2D1_RECT_F& hitRectDip,
                                            const AccessibilityPointHitBuildContext& context,
                                            size_t treeVisibleIndex = 0u,
                                            uint64_t gridRowId      = 0u,
                                            size_t gridColumnIndex  = 0u)
{
    if (const std::optional<D2D1_RECT_F> transformed = ApplyAccessibilityPointHitContext(hitRectDip, context))
    {
        AppendAccessibilityPointHit(snapshot, kind, path, transformed.value(), treeVisibleIndex, gridRowId, gridColumnIndex);
    }
}

void AppendTreeAccessibilityPointHits(const Tree& tree,
                                      const ControlPath& path,
                                      AccessibilitySnapshot& snapshot,
                                      const AccessibilityPointHitBuildContext& context)
{
    const auto* model = tree.GetModel();
    if (! model)
    {
        return;
    }

    const D2D1_RECT_F treeHitBounds = tree.GetHitBounds();
    const size_t visibleItemCount   = model->GetVisibleItemCount();
    for (size_t visibleIndex = tree.GetFirstVisibleItemIndex(); visibleIndex < visibleItemCount; ++visibleIndex)
    {
        const std::optional<D2D1_RECT_F> rowRect = tree.GetVisibleItemHitRect(visibleIndex);
        if (! rowRect)
        {
            continue;
        }
        if (rowRect->bottom < treeHitBounds.top)
        {
            continue;
        }
        if (rowRect->top > treeHitBounds.bottom)
        {
            break;
        }

        AppendTransformedAccessibilityPointHit(snapshot, AccessibilityFragmentKind::TreeItem, path, rowRect.value(), context, visibleIndex);
    }
}

void AppendGridAccessibilityPointHits(const Grid& grid,
                                      const ControlPath& path,
                                      AccessibilitySnapshot& snapshot,
                                      const AccessibilityPointHitBuildContext& context)
{
    const auto* model = grid.GetModel();
    if (! model)
    {
        return;
    }

    const size_t visibleColumnCount = grid.GetVisibleColumnCount();
    std::vector<size_t> visibleColumns;
    visibleColumns.reserve(visibleColumnCount);
    for (size_t visibleColumnIndex = 0u; visibleColumnIndex < visibleColumnCount; ++visibleColumnIndex)
    {
        const std::optional<size_t> columnIndex = grid.GetVisibleColumnAt(visibleColumnIndex);
        if (! columnIndex)
        {
            continue;
        }

        visibleColumns.push_back(columnIndex.value());
        if (const std::optional<D2D1_RECT_F> headerRect = grid.GetVisibleColumnHeaderRect(columnIndex.value()))
        {
            AppendTransformedAccessibilityPointHit(
                snapshot, AccessibilityFragmentKind::GridHeader, path, headerRect.value(), context, 0u, 0u, columnIndex.value());
        }
    }

    const size_t visibleRowCount = grid.GetVisibleRowCount();
    for (size_t visibleRowIndex = 0u; visibleRowIndex < visibleRowCount; ++visibleRowIndex)
    {
        const std::optional<size_t> rowIndex = grid.GetVisibleRowAt(visibleRowIndex);
        if (! rowIndex)
        {
            continue;
        }

        const uint64_t rowId = model->GetStableRowId(rowIndex.value());
        for (const size_t columnIndex : visibleColumns)
        {
            if (const std::optional<D2D1_RECT_F> cellRect = grid.GetVisibleCellRect(rowIndex.value(), columnIndex))
            {
                AppendTransformedAccessibilityPointHit(
                    snapshot, AccessibilityFragmentKind::GridCell, path, cellRect.value(), context, 0u, rowId, columnIndex);
            }
        }

        if (const std::optional<D2D1_RECT_F> rowRect = grid.GetVisibleRowRect(rowIndex.value()))
        {
            AppendTransformedAccessibilityPointHit(snapshot, AccessibilityFragmentKind::GridRow, path, rowRect.value(), context, 0u, rowId);
        }
    }
}

void AppendAccessibilitySnapshotPointHits(WindowHost& host, const Control* current, const ControlPath& basePath, AccessibilitySnapshot& snapshot)
{
    AppendAccessibilitySnapshotPointHits(host, current, basePath, snapshot, AccessibilityPointHitBuildContext{});
    AppendAccessibilityPointHit(snapshot, AccessibilityFragmentKind::Root, ControlPath{}, host.GetClientBoundsDip());
}

void AppendAccessibilitySnapshotPointHits(WindowHost& host,
                                          const Control* current,
                                          const ControlPath& basePath,
                                          AccessibilitySnapshot& snapshot,
                                          const AccessibilityPointHitBuildContext& context)
{
    if (! current || ! current->IsVisible())
    {
        return;
    }

    AccessibilityPointHitBuildContext childContext = context;
    if (const auto* scrollPanel = dynamic_cast<const ScrollPanel*>(current))
    {
        const std::optional<D2D1_RECT_F> viewport = ApplyAccessibilityPointHitContext(scrollPanel->GetViewportRect(), context);
        if (! viewport)
        {
            return;
        }

        childContext.clipRectDip = viewport.value();
        childContext.translationDip.y -= scrollPanel->GetScrollOffset();
    }

    if (const auto* panel = dynamic_cast<const Panel*>(current))
    {
        const auto children = panel->GetChildren();
        for (size_t index = children.size(); index-- > 0u;)
        {
            if (! children[index])
            {
                continue;
            }

            ControlPath childPath{};
            if (! TryAppendPathIndex(basePath, index, childPath))
            {
                continue;
            }

            AppendAccessibilitySnapshotPointHits(host, children[index].get(), childPath, snapshot, childContext);
        }
    }

    if (! IsSemanticAccessibilityControl(current))
    {
        return;
    }

    if (const auto* textField = dynamic_cast<const TextField*>(current); textField && textField->IsPasswordRevealButtonVisibleForAccessibility())
    {
        AppendTransformedAccessibilityPointHit(
            snapshot, AccessibilityFragmentKind::TextFieldPasswordRevealButton, basePath, textField->GetPasswordRevealButtonAccessibilityRect(), context);
    }
    if (const auto* tree = dynamic_cast<const Tree*>(current))
    {
        AppendTreeAccessibilityPointHits(*tree, basePath, snapshot, context);
    }
    else if (const auto* grid = dynamic_cast<const Grid*>(current))
    {
        AppendGridAccessibilityPointHits(*grid, basePath, snapshot, context);
    }

    AppendTransformedAccessibilityPointHit(snapshot, AccessibilityFragmentKind::Control, basePath, current->GetHitBounds(), context);
}

const AccessibilityPointHitSnapshot* FindSnapshotPointHit(const AccessibilitySnapshot& snapshot, D2D1_POINT_2F pointDip) noexcept
{
    for (const AccessibilityPointHitSnapshot& hit : snapshot.pointHitRecords)
    {
        if (PointInRect(hit.hitRectDip, pointDip))
        {
            return &hit;
        }
    }

    return nullptr;
}

bool SnapshotPointHitMatchesFragment(const AccessibilityPointHitSnapshot& hit,
                                     AccessibilityFragmentKind kind,
                                     const ControlPath& path,
                                     size_t treeVisibleIndex,
                                     uint64_t gridRowId,
                                     size_t gridColumnIndex) noexcept
{
    if (hit.kind != kind || ! AreControlPathsEqual(hit.path, path))
    {
        return false;
    }

    switch (kind)
    {
        case AccessibilityFragmentKind::TreeItem: return hit.treeVisibleIndex == treeVisibleIndex;
        case AccessibilityFragmentKind::GridHeader: return hit.gridColumnIndex == gridColumnIndex;
        case AccessibilityFragmentKind::GridRow: return hit.gridRowId == gridRowId;
        case AccessibilityFragmentKind::GridCell: return hit.gridRowId == gridRowId && hit.gridColumnIndex == gridColumnIndex;
        case AccessibilityFragmentKind::Control:
        case AccessibilityFragmentKind::TextFieldPasswordRevealButton: return true;
        case AccessibilityFragmentKind::Root: return false;
    }

    return false;
}

std::optional<D2D1_RECT_F> FindSnapshotFragmentBounds(const AccessibilitySnapshot& snapshot,
                                                      AccessibilityFragmentKind kind,
                                                      const ControlPath& path,
                                                      size_t treeVisibleIndex,
                                                      uint64_t gridRowId,
                                                      size_t gridColumnIndex) noexcept
{
    for (const AccessibilityPointHitSnapshot& hit : snapshot.pointHitRecords)
    {
        if (SnapshotPointHitMatchesFragment(hit, kind, path, treeVisibleIndex, gridRowId, gridColumnIndex))
        {
            return hit.hitRectDip;
        }
    }

    return std::nullopt;
}

void AppendAccessibilitySnapshotNavigation(
    WindowHost& host, const Control* root, const Control* current, const ControlPath& basePath, AccessibilitySnapshot& snapshot)
{
    if (! current || ! current->IsVisible())
    {
        return;
    }

    if (IsSemanticAccessibilityControl(current))
    {
        snapshot.semanticControlOrder.push_back(basePath);

        AccessibilityControlNavigationSnapshot record{};
        record.path                      = basePath;
        record.controlTypeId             = GetControlTypeId(current);
        record.controlAccessibleName     = std::wstring(GetControlAccessibleName(root, current));
        record.controlAccessibleHelpText = current->GetAccessibleHelpText();
        record.controlVisible            = current->IsVisible();
        record.controlEnabled            = current->IsEnabled();
        record.controlFocusable          = current->IsFocusable();
        record.controlHasFocus           = current->HasFocus();
        if (const auto* textField = dynamic_cast<const TextField*>(current))
        {
            record.controlIsPassword = textField->IsMasked();
        }
        record.controlSupportsInvoke     = SupportsInvokePattern(current);
        record.controlSupportsToggle     = SupportsTogglePattern(current);
        record.controlSupportsValue      = SupportsValuePattern(current);
        record.controlSupportsText       = SupportsTextPattern(current);
        record.controlSupportsRangeValue = SupportsRangeValuePattern(current);
        record.controlValueReadOnly      = IsValueReadOnly(current);
        if (record.controlSupportsValue)
        {
            record.controlAccessibleValue = std::wstring(GetControlAccessibleValue(current));
        }
        if (record.controlSupportsText)
        {
            TextRangeSpan selectionRange{};
            bool hasSelectionRange = false;
            NativeTextInputState nativeState{};
            if (host.TryReadNativeTextInputState(current, nativeState))
            {
                record.controlAccessibleText = nativeState.text;
                selectionRange    = GetTextRangeSpanFromState(nativeState.caretIndex, nativeState.selectionAnchorIndex, record.controlAccessibleText.size());
                hasSelectionRange = true;
                record.controlTextSelectionStart        = selectionRange.start;
                record.controlTextSelectionEnd          = selectionRange.end;
                record.controlTextCompositionStart      = nativeState.compositionStartIndex;
                record.controlTextCompositionEnd        = nativeState.compositionEndIndex;
                record.controlTextConversionTargetStart = nativeState.conversionTargetStartIndex;
                record.controlTextConversionTargetEnd   = nativeState.conversionTargetEndIndex;
            }
            else
            {
                record.controlAccessibleText     = GetControlAccessibleTextRangeText(current);
                selectionRange                   = GetControlAccessibleSelectionRange(&host, current, record.controlAccessibleText.size());
                hasSelectionRange                = true;
                record.controlTextSelectionStart = selectionRange.start;
                record.controlTextSelectionEnd   = selectionRange.end;
            }

            if (hasSelectionRange && selectionRange.start != selectionRange.end)
            {
                if (std::optional<std::vector<D2D1_RECT_F>> bounds =
                        TryResolveTextRangeCaretRects(host, *current, record.controlAccessibleText, selectionRange);
                    bounds.has_value() && ! bounds->empty())
                {
                    record.controlTextSelectionBoundsDip = std::move(bounds.value());
                }
                else
                {
                    record.controlTextSelectionBoundsDip.push_back(ResolveTextPatternViewportRect(current));
                }
            }
        }
        if (const auto* toggle = dynamic_cast<const Toggle*>(current))
        {
            record.controlToggleChecked = toggle->IsChecked();
        }
        if (const auto* slider = dynamic_cast<const Slider*>(current))
        {
            record.controlRangeValue       = slider->GetValue();
            record.controlRangeMinimum     = slider->GetMinimum();
            record.controlRangeMaximum     = slider->GetMaximum();
            record.controlRangeSmallChange = slider->GetStep();
            record.controlRangeLargeChange = slider->GetLargeStep();
        }
        if (const auto* textField = dynamic_cast<const TextField*>(current))
        {
            record.hasPasswordRevealButton = textField->IsPasswordRevealButtonVisibleForAccessibility();
            if (record.hasPasswordRevealButton)
            {
                record.passwordRevealButtonAccessibleName = textField->GetPasswordRevealAccessibleName();
                record.passwordRevealButtonEnabled        = textField->IsEnabled();
            }
        }
        else if (const auto* tree = dynamic_cast<const Tree*>(current))
        {
            record.isTree        = true;
            record.treeIsEnabled = tree->IsEnabled();
            record.treeHasFocus  = tree->HasFocus();
            if (const auto* model = tree->GetModel())
            {
                record.treeVisibleItemCount = model->GetVisibleItemCount();
                record.treeItems.reserve(record.treeVisibleItemCount);
                const std::optional<uint64_t> selectedItemId = tree->GetSelectedItemId();
                for (size_t visibleIndex = 0u; visibleIndex < record.treeVisibleItemCount; ++visibleIndex)
                {
                    TreeItemData item{};
                    model->GetVisibleItem(visibleIndex, item);
                    record.treeItems.push_back(AccessibilityTreeItemSnapshotRecord{.visibleIndex = visibleIndex,
                                                                                   .itemId       = item.id,
                                                                                   .text         = item.text,
                                                                                   .depth        = item.depth,
                                                                                   .hasChildren  = item.hasChildren,
                                                                                   .expanded     = item.expanded});
                    if (selectedItemId && selectedItemId.value() == item.id)
                    {
                        record.selectedTreeVisibleIndex = visibleIndex;
                    }
                }
            }
        }
        else if (const auto* grid = dynamic_cast<const Grid*>(current))
        {
            const bool captureGridSnapshotPerf = Debug::Perf::IsCaptureEnabled();
            const auto gridSnapshotStartedAt   = std::chrono::steady_clock::now();
            size_t materializedOffscreenRowCount = 0u;
            record.isGrid                = true;
            record.gridCanSelectMultiple = grid->GetSelectionMode() != GridSelectionMode::Single;
            record.gridIsEnabled         = grid->IsEnabled();
            record.gridHasFocus          = grid->HasFocus();
            if (const auto* model = grid->GetModel())
            {
                record.gridRowCount    = model->GetRowCount();
                record.gridColumnCount = model->GetColumnCount();

                const size_t visibleColumnCount = grid->GetVisibleColumnCount();
                record.gridVisibleColumns.reserve(visibleColumnCount);
                record.gridHeaders.reserve(visibleColumnCount);
                for (size_t visibleColumnIndex = 0u; visibleColumnIndex < visibleColumnCount; ++visibleColumnIndex)
                {
                    if (const std::optional<size_t> columnIndex = grid->GetVisibleColumnAt(visibleColumnIndex))
                    {
                        record.gridVisibleColumns.push_back(columnIndex.value());
                        const GridColumnDesc columnDesc = model->GetColumn(columnIndex.value());
                        record.gridHeaders.push_back(AccessibilityGridHeaderSnapshotRecord{
                            .columnIndex = columnIndex.value(), .gridHeaderName = std::wstring(GetGridHeaderAccessibleName(columnDesc))});
                    }
                }

                const size_t visibleRowCount = grid->GetVisibleRowCount();
                record.gridVisibleRows.reserve(visibleRowCount);
                record.gridVisibleRowIds.reserve(visibleRowCount);
                record.gridRows.reserve(visibleRowCount);
                for (size_t visibleRowIndex = 0u; visibleRowIndex < visibleRowCount; ++visibleRowIndex)
                {
                    if (const std::optional<size_t> rowIndex = grid->GetVisibleRowAt(visibleRowIndex))
                    {
                        const uint64_t rowId = model->GetStableRowId(rowIndex.value());
                        record.gridVisibleRows.push_back(rowIndex.value());
                        record.gridVisibleRowIds.push_back(rowId);
                        record.gridRows.push_back(AccessibilityGridRowSnapshotRecord{
                            .rowIndex         = rowIndex.value(),
                            .rowId            = rowId,
                            .gridRowOffscreen = ! grid->IsVisible() || ! grid->GetVisibleRowRect(rowIndex.value()).has_value()});
                    }
                }

                record.gridCells.reserve(record.gridVisibleRows.size() * record.gridColumnCount);
                for (size_t rowOrdinal = 0u; rowOrdinal < record.gridVisibleRows.size(); ++rowOrdinal)
                {
                    const size_t rowIndex = record.gridVisibleRows[rowOrdinal];
                    const uint64_t rowId  = record.gridVisibleRowIds[rowOrdinal];
                    for (size_t columnIndex = 0u; columnIndex < record.gridColumnCount; ++columnIndex)
                    {
                        GridCellData cellData{};
                        model->GetCellData(rowIndex, columnIndex, cellData);

                        std::wstring helpText;
                        std::wstring accessibleText = BuildGridCellAccessibleText(cellData);
                        if (! accessibleText.empty() && rowOrdinal < record.gridRows.size())
                        {
                            std::wstring& rowName = record.gridRows[rowOrdinal].gridRowAccessibleName;
                            if (! rowName.empty())
                            {
                                rowName.append(L" | ");
                            }
                            rowName.append(accessibleText);
                        }
                        if (! cellData.tooltipText.empty() && cellData.tooltipText != cellData.text && cellData.tooltipText != accessibleText)
                        {
                            helpText = cellData.tooltipText;
                        }

                        record.gridCells.push_back(AccessibilityGridCellStateSnapshotRecord{
                            .rowIndex                   = rowIndex,
                            .rowId                      = rowId,
                            .columnIndex                = columnIndex,
                            .gridCellControlTypeId      = GetGridCellControlTypeId(cellData),
                            .gridCellAccessibleText     = std::move(accessibleText),
                            .gridCellHelpText           = std::move(helpText),
                            .gridCellEnabled            = cellData.enabled,
                            .gridCellOffscreen          = ! grid->IsVisible() || ! grid->GetVisibleCellRect(rowIndex, columnIndex).has_value(),
                            .gridCellChecked            = cellData.checked,
                            .gridCellSupportsToggle     = GridCellSupportsTogglePattern(cellData),
                            .gridCellSupportsValue      = GridCellSupportsValuePattern(cellData),
                            .gridCellSupportsRangeValue = GridCellSupportsRangeValuePattern(cellData),
                            .gridCellRangeValue         = GetGridCellRangeValue(cellData)});
                    }
                }

                const auto selection = grid->GetSelectionModel().GetOrderedSelection();
                record.selectedGridRowIds.reserve(selection.size());
                std::unordered_set<uint64_t> visibleRowIds(record.gridVisibleRowIds.begin(), record.gridVisibleRowIds.end());
                for (const uint64_t rowId : selection)
                {
                    const std::optional<size_t> selectedRowIndex = model->FindRowByStableId(rowId);
                    if (selectedRowIndex)
                    {
                        record.selectedGridRowIds.push_back(rowId);
                        if (! visibleRowIds.contains(rowId) &&
                            materializedOffscreenRowCount < AccessibilityOffscreenSelectedRowMaterializationLimit())
                        {
                            ++materializedOffscreenRowCount;
                            AccessibilityGridRowSnapshotRecord rowRecord{
                                .rowIndex              = selectedRowIndex.value(),
                                .rowId                 = rowId,
                                .gridRowOffscreen      = true};
                            for (size_t columnIndex = 0u; columnIndex < record.gridColumnCount; ++columnIndex)
                            {
                                GridCellData cellData{};
                                model->GetCellData(selectedRowIndex.value(), columnIndex, cellData);
                                std::wstring accessibleText = BuildGridCellAccessibleText(cellData);
                                if (! accessibleText.empty())
                                {
                                    if (! rowRecord.gridRowAccessibleName.empty())
                                    {
                                        rowRecord.gridRowAccessibleName.append(L" | ");
                                    }
                                    rowRecord.gridRowAccessibleName.append(accessibleText);
                                }
                            }
                            record.gridRows.push_back(std::move(rowRecord));
                        }
                    }
                }
            }
            if (captureGridSnapshotPerf)
            {
                Debug::Perf::Emit(L"dxui.accessibility.grid_snapshot_rebuild_us",
                                  AccessibilityGridSnapshotPerfDetail(),
                                  Debug::Perf::ElapsedUs(gridSnapshotStartedAt),
                                  record.selectedGridRowIds.size(),
                                  materializedOffscreenRowCount,
                                  S_OK);
            }
        }

        record.controlSupportsSelection = record.isTree || record.isGrid;
        record.controlSupportsTable     = record.isGrid && record.gridColumnCount > 0u;
        snapshot.controlNavigationRecords.push_back(std::move(record));
    }

    if (const auto* panel = dynamic_cast<const Panel*>(current))
    {
        const auto children = panel->GetChildren();
        for (size_t index = 0u; index < children.size(); ++index)
        {
            if (! children[index])
            {
                continue;
            }

            ControlPath childPath{};
            if (! TryAppendPathIndex(basePath, index, childPath))
            {
                continue;
            }

            AppendAccessibilitySnapshotNavigation(host, root, children[index].get(), childPath, snapshot);
        }
    }
}

const AccessibilityControlNavigationSnapshot* FindControlNavigationRecord(const AccessibilitySnapshot& snapshot, const ControlPath& path) noexcept
{
    const auto it = std::ranges::find_if(snapshot.controlNavigationRecords, [&](const AccessibilityControlNavigationSnapshot& record) noexcept {
        return AreControlPathsEqual(record.path, path);
    });
    return it == snapshot.controlNavigationRecords.end() ? nullptr : &*it;
}

const AccessibilityControlNavigationSnapshot* ResolveSnapshotControlRecord(const AccessibilitySnapshot& snapshot,
                                                                           AccessibilityFragmentKind kind,
                                                                           const ControlPath& path) noexcept
{
    if (kind == AccessibilityFragmentKind::Control)
    {
        return FindControlNavigationRecord(snapshot, path);
    }
    if (kind == AccessibilityFragmentKind::Root && SnapshotHasCollapsedSemanticRoot(snapshot))
    {
        return FindControlNavigationRecord(snapshot, snapshot.semanticControlOrder.front());
    }
    return nullptr;
}

std::optional<size_t> FindSemanticControlOrderIndex(const AccessibilitySnapshot& snapshot, const ControlPath& path) noexcept
{
    for (size_t index = 0u; index < snapshot.semanticControlOrder.size(); ++index)
    {
        if (AreControlPathsEqual(snapshot.semanticControlOrder[index], path))
        {
            return index;
        }
    }

    return std::nullopt;
}

bool SnapshotHasCollapsedSemanticRoot(const AccessibilitySnapshot& snapshot) noexcept
{
    return snapshot.hasCollapsedSemanticRoot;
}

bool SnapshotPathIsCollapsedSemanticRoot(const AccessibilitySnapshot& snapshot, const ControlPath& path) noexcept
{
    return SnapshotHasCollapsedSemanticRoot(snapshot) && AreControlPathsEqual(snapshot.semanticControlOrder.front(), path);
}

std::optional<size_t> FindSizeValueIndex(std::span<const size_t> values, size_t value) noexcept
{
    for (size_t index = 0u; index < values.size(); ++index)
    {
        if (values[index] == value)
        {
            return index;
        }
    }

    return std::nullopt;
}

std::optional<size_t> FindU64ValueIndex(std::span<const uint64_t> values, uint64_t value) noexcept
{
    for (size_t index = 0u; index < values.size(); ++index)
    {
        if (values[index] == value)
        {
            return index;
        }
    }

    return std::nullopt;
}

bool SnapshotContainsGridRow(const AccessibilityControlNavigationSnapshot& record, uint64_t rowId) noexcept
{
    return FindU64ValueIndex(record.gridVisibleRowIds, rowId).has_value();
}

const AccessibilityGridRowSnapshotRecord* FindSnapshotGridRowRecord(const AccessibilityControlNavigationSnapshot& record, uint64_t rowId) noexcept
{
    if (! record.isGrid)
    {
        return nullptr;
    }

    for (const AccessibilityGridRowSnapshotRecord& row : record.gridRows)
    {
        if (row.rowId == rowId)
        {
            return &row;
        }
    }

    return nullptr;
}

const AccessibilityTreeItemSnapshotRecord* FindSnapshotTreeItemRecord(const AccessibilityControlNavigationSnapshot& record, size_t treeVisibleIndex) noexcept
{
    if (! record.isTree || treeVisibleIndex >= record.treeItems.size())
    {
        return nullptr;
    }

    const AccessibilityTreeItemSnapshotRecord& item = record.treeItems[treeVisibleIndex];
    return item.visibleIndex == treeVisibleIndex ? &item : nullptr;
}

const AccessibilityGridHeaderSnapshotRecord* FindSnapshotGridHeaderRecord(const AccessibilityControlNavigationSnapshot& record, size_t gridColumnIndex) noexcept
{
    if (! record.isGrid)
    {
        return nullptr;
    }

    for (const AccessibilityGridHeaderSnapshotRecord& header : record.gridHeaders)
    {
        if (header.columnIndex == gridColumnIndex)
        {
            return &header;
        }
    }

    return nullptr;
}

bool SnapshotContainsTreeItem(const AccessibilityControlNavigationSnapshot& record, size_t treeVisibleIndex) noexcept
{
    return FindSnapshotTreeItemRecord(record, treeVisibleIndex) != nullptr;
}

bool SnapshotSupportsSelectionProvider(const AccessibilityControlNavigationSnapshot& record) noexcept
{
    return record.isTree || record.isGrid;
}

bool SnapshotTreeItemIsSelected(const AccessibilityControlNavigationSnapshot& record, size_t treeVisibleIndex) noexcept
{
    return record.selectedTreeVisibleIndex && record.selectedTreeVisibleIndex.value() == treeVisibleIndex;
}

bool SnapshotGridRowIsSelected(const AccessibilityControlNavigationSnapshot& record, uint64_t rowId) noexcept
{
    return FindU64ValueIndex(record.selectedGridRowIds, rowId).has_value();
}

std::optional<AccessibilityGridCellSnapshotRecord> FindSnapshotGridCellRecord(const AccessibilitySnapshot& snapshot,
                                                                              const ControlPath& path,
                                                                              uint64_t rowId,
                                                                              size_t columnIndex) noexcept
{
    const AccessibilityControlNavigationSnapshot* record = FindControlNavigationRecord(snapshot, path);
    if (! record || ! record->isGrid || columnIndex >= record->gridColumnCount)
    {
        return std::nullopt;
    }

    const auto it = std::ranges::find_if(record->gridCells, [&](const AccessibilityGridCellStateSnapshotRecord& cell) noexcept {
        return cell.rowId == rowId && cell.columnIndex == columnIndex;
    });
    if (it == record->gridCells.end())
    {
        return std::nullopt;
    }

    return AccessibilityGridCellSnapshotRecord{.controlRecord = record, .cellRecord = &*it, .rowIndex = it->rowIndex, .columnIndex = it->columnIndex};
}

bool SnapshotGridCellSupportsTogglePattern(const AccessibilityGridCellSnapshotRecord& cell) noexcept
{
    return cell.cellRecord && cell.cellRecord->gridCellSupportsToggle;
}

bool SnapshotGridCellSupportsValuePattern(const AccessibilityGridCellSnapshotRecord& cell) noexcept
{
    return cell.cellRecord && cell.cellRecord->gridCellSupportsValue;
}

bool SnapshotGridCellSupportsRangeValuePattern(const AccessibilityGridCellSnapshotRecord& cell) noexcept
{
    return cell.cellRecord && cell.cellRecord->gridCellSupportsRangeValue;
}

enum class AccessibilityPatternKind
{
    Invoke,
    Toggle,
    Value,
    Text,
    TextEdit,
    RangeValue,
    Selection,
    Table,
    SelectionItem,
    ExpandCollapse,
    GridItem,
    TableItem,
};

struct AccessibilityPatternQueryResult
{
    void* queryInterface      = nullptr;
    IUnknown* patternProvider = nullptr;
};

std::optional<AccessibilityPatternKind> PatternKindFromInterfaceId(REFIID riid) noexcept
{
    if (riid == __uuidof(IInvokeProvider))
    {
        return AccessibilityPatternKind::Invoke;
    }
    if (riid == __uuidof(IToggleProvider))
    {
        return AccessibilityPatternKind::Toggle;
    }
    if (riid == __uuidof(IValueProvider))
    {
        return AccessibilityPatternKind::Value;
    }
    if (riid == __uuidof(ITextProvider))
    {
        return AccessibilityPatternKind::Text;
    }
    if (riid == __uuidof(ITextEditProvider))
    {
        return AccessibilityPatternKind::TextEdit;
    }
    if (riid == __uuidof(IRangeValueProvider))
    {
        return AccessibilityPatternKind::RangeValue;
    }
    if (riid == __uuidof(ISelectionProvider))
    {
        return AccessibilityPatternKind::Selection;
    }
    if (riid == __uuidof(ITableProvider))
    {
        return AccessibilityPatternKind::Table;
    }
    if (riid == __uuidof(ISelectionItemProvider))
    {
        return AccessibilityPatternKind::SelectionItem;
    }
    if (riid == __uuidof(IExpandCollapseProvider))
    {
        return AccessibilityPatternKind::ExpandCollapse;
    }
    if (riid == __uuidof(IGridItemProvider))
    {
        return AccessibilityPatternKind::GridItem;
    }
    if (riid == __uuidof(ITableItemProvider))
    {
        return AccessibilityPatternKind::TableItem;
    }
    return std::nullopt;
}

std::optional<AccessibilityPatternKind> PatternKindFromPatternId(PATTERNID patternId) noexcept
{
    switch (patternId)
    {
        case UIA_InvokePatternId: return AccessibilityPatternKind::Invoke;
        case UIA_TogglePatternId: return AccessibilityPatternKind::Toggle;
        case UIA_ValuePatternId: return AccessibilityPatternKind::Value;
        case UIA_TextPatternId: return AccessibilityPatternKind::Text;
        case UIA_TextEditPatternId: return AccessibilityPatternKind::TextEdit;
        case UIA_RangeValuePatternId: return AccessibilityPatternKind::RangeValue;
        case UIA_SelectionPatternId: return AccessibilityPatternKind::Selection;
        case UIA_TablePatternId: return AccessibilityPatternKind::Table;
        case UIA_SelectionItemPatternId: return AccessibilityPatternKind::SelectionItem;
        case UIA_ExpandCollapsePatternId: return AccessibilityPatternKind::ExpandCollapse;
        case UIA_GridItemPatternId: return AccessibilityPatternKind::GridItem;
        case UIA_TableItemPatternId: return AccessibilityPatternKind::TableItem;
        default: return std::nullopt;
    }
}

AccessibilityNavigationTarget MakeControlNavigationTarget(const ControlPath& path) noexcept
{
    AccessibilityNavigationTarget target{};
    target.kind = AccessibilityFragmentKind::Control;
    target.path = path;
    return target;
}

AccessibilityNavigationTarget MakeNavigationTarget(
    AccessibilityFragmentKind kind, const ControlPath& path, size_t treeVisibleIndex = 0u, uint64_t gridRowId = 0u, size_t gridColumnIndex = 0u) noexcept
{
    AccessibilityNavigationTarget target{};
    target.kind             = kind;
    target.path             = path;
    target.treeVisibleIndex = treeVisibleIndex;
    target.gridRowId        = gridRowId;
    target.gridColumnIndex  = gridColumnIndex;
    return target;
}

std::optional<AccessibilityNavigationTarget> ResolveSnapshotNavigationTarget(const AccessibilitySnapshot& snapshot,
                                                                             AccessibilityFragmentKind kind,
                                                                             const ControlPath& path,
                                                                             size_t treeVisibleIndex,
                                                                             uint64_t gridRowId,
                                                                             size_t gridColumnIndex,
                                                                             NavigateDirection direction) noexcept
{
    const AccessibilityControlNavigationSnapshot* controlRecord = FindControlNavigationRecord(snapshot, path);
    switch (direction)
    {
        case NavigateDirection_FirstChild:
            if (kind == AccessibilityFragmentKind::Root)
            {
                if (SnapshotHasCollapsedSemanticRoot(snapshot))
                {
                    return std::nullopt;
                }
                if (! snapshot.semanticControlOrder.empty())
                {
                    return MakeControlNavigationTarget(snapshot.semanticControlOrder.front());
                }
                return std::nullopt;
            }
            if (kind == AccessibilityFragmentKind::Control && controlRecord)
            {
                if (controlRecord->treeVisibleItemCount > 0u)
                {
                    return MakeNavigationTarget(AccessibilityFragmentKind::TreeItem, path, 0u);
                }
                if (! controlRecord->gridVisibleColumns.empty())
                {
                    return MakeNavigationTarget(AccessibilityFragmentKind::GridHeader, path, 0u, 0u, controlRecord->gridVisibleColumns.front());
                }
                if (! controlRecord->gridVisibleRowIds.empty())
                {
                    return MakeNavigationTarget(AccessibilityFragmentKind::GridRow, path, 0u, controlRecord->gridVisibleRowIds.front());
                }
                if (controlRecord->hasPasswordRevealButton)
                {
                    return MakeNavigationTarget(AccessibilityFragmentKind::TextFieldPasswordRevealButton, path);
                }
            }
            if (kind == AccessibilityFragmentKind::GridRow && controlRecord && SnapshotContainsGridRow(*controlRecord, gridRowId) &&
                controlRecord->gridColumnCount > 0u)
            {
                return MakeNavigationTarget(AccessibilityFragmentKind::GridCell, path, 0u, gridRowId, 0u);
            }
            return std::nullopt;
        case NavigateDirection_LastChild:
            if (kind == AccessibilityFragmentKind::Root)
            {
                if (SnapshotHasCollapsedSemanticRoot(snapshot))
                {
                    return std::nullopt;
                }
                if (! snapshot.semanticControlOrder.empty())
                {
                    return MakeControlNavigationTarget(snapshot.semanticControlOrder.back());
                }
                return std::nullopt;
            }
            if (kind == AccessibilityFragmentKind::Control && controlRecord)
            {
                if (controlRecord->treeVisibleItemCount > 0u)
                {
                    return MakeNavigationTarget(AccessibilityFragmentKind::TreeItem, path, controlRecord->treeVisibleItemCount - 1u);
                }
                if (! controlRecord->gridVisibleRowIds.empty())
                {
                    return MakeNavigationTarget(AccessibilityFragmentKind::GridRow, path, 0u, controlRecord->gridVisibleRowIds.back());
                }
                if (! controlRecord->gridVisibleColumns.empty())
                {
                    return MakeNavigationTarget(AccessibilityFragmentKind::GridHeader, path, 0u, 0u, controlRecord->gridVisibleColumns.back());
                }
                if (controlRecord->hasPasswordRevealButton)
                {
                    return MakeNavigationTarget(AccessibilityFragmentKind::TextFieldPasswordRevealButton, path);
                }
            }
            if (kind == AccessibilityFragmentKind::GridRow && controlRecord && SnapshotContainsGridRow(*controlRecord, gridRowId) &&
                controlRecord->gridColumnCount > 0u)
            {
                return MakeNavigationTarget(AccessibilityFragmentKind::GridCell, path, 0u, gridRowId, controlRecord->gridColumnCount - 1u);
            }
            return std::nullopt;
        case NavigateDirection_Parent:
            if (kind == AccessibilityFragmentKind::Control)
            {
                if (SnapshotPathIsCollapsedSemanticRoot(snapshot, path))
                {
                    return std::nullopt;
                }
                return MakeNavigationTarget(AccessibilityFragmentKind::Root, ControlPath{});
            }
            if (kind == AccessibilityFragmentKind::GridCell)
            {
                return MakeNavigationTarget(AccessibilityFragmentKind::GridRow, path, 0u, gridRowId);
            }
            if (kind == AccessibilityFragmentKind::TreeItem || kind == AccessibilityFragmentKind::GridHeader || kind == AccessibilityFragmentKind::GridRow ||
                kind == AccessibilityFragmentKind::TextFieldPasswordRevealButton)
            {
                return MakeControlNavigationTarget(path);
            }
            return std::nullopt;
        case NavigateDirection_NextSibling:
            if (kind == AccessibilityFragmentKind::Control)
            {
                if (SnapshotPathIsCollapsedSemanticRoot(snapshot, path))
                {
                    return std::nullopt;
                }
                const std::optional<size_t> ordinal = FindSemanticControlOrderIndex(snapshot, path);
                if (ordinal && (ordinal.value() + 1u) < snapshot.semanticControlOrder.size())
                {
                    return MakeControlNavigationTarget(snapshot.semanticControlOrder[ordinal.value() + 1u]);
                }
            }
            else if (kind == AccessibilityFragmentKind::TreeItem && controlRecord && (treeVisibleIndex + 1u) < controlRecord->treeVisibleItemCount)
            {
                return MakeNavigationTarget(AccessibilityFragmentKind::TreeItem, path, treeVisibleIndex + 1u);
            }
            else if (kind == AccessibilityFragmentKind::GridHeader && controlRecord)
            {
                const std::optional<size_t> ordinal = FindSizeValueIndex(controlRecord->gridVisibleColumns, gridColumnIndex);
                if (ordinal && (ordinal.value() + 1u) < controlRecord->gridVisibleColumns.size())
                {
                    return MakeNavigationTarget(AccessibilityFragmentKind::GridHeader, path, 0u, 0u, controlRecord->gridVisibleColumns[ordinal.value() + 1u]);
                }
                if (ordinal && ! controlRecord->gridVisibleRowIds.empty())
                {
                    return MakeNavigationTarget(AccessibilityFragmentKind::GridRow, path, 0u, controlRecord->gridVisibleRowIds.front());
                }
            }
            else if (kind == AccessibilityFragmentKind::GridRow && controlRecord)
            {
                const std::optional<size_t> ordinal = FindU64ValueIndex(controlRecord->gridVisibleRowIds, gridRowId);
                if (ordinal && (ordinal.value() + 1u) < controlRecord->gridVisibleRowIds.size())
                {
                    return MakeNavigationTarget(AccessibilityFragmentKind::GridRow, path, 0u, controlRecord->gridVisibleRowIds[ordinal.value() + 1u]);
                }
            }
            else if (kind == AccessibilityFragmentKind::GridCell && controlRecord && SnapshotContainsGridRow(*controlRecord, gridRowId))
            {
                if ((gridColumnIndex + 1u) < controlRecord->gridColumnCount)
                {
                    return MakeNavigationTarget(AccessibilityFragmentKind::GridCell, path, 0u, gridRowId, gridColumnIndex + 1u);
                }
            }
            return std::nullopt;
        case NavigateDirection_PreviousSibling:
            if (kind == AccessibilityFragmentKind::Control)
            {
                if (SnapshotPathIsCollapsedSemanticRoot(snapshot, path))
                {
                    return std::nullopt;
                }
                const std::optional<size_t> ordinal = FindSemanticControlOrderIndex(snapshot, path);
                if (ordinal && ordinal.value() > 0u)
                {
                    return MakeControlNavigationTarget(snapshot.semanticControlOrder[ordinal.value() - 1u]);
                }
            }
            else if (kind == AccessibilityFragmentKind::TreeItem && controlRecord && treeVisibleIndex > 0u &&
                     treeVisibleIndex < controlRecord->treeVisibleItemCount)
            {
                return MakeNavigationTarget(AccessibilityFragmentKind::TreeItem, path, treeVisibleIndex - 1u);
            }
            else if (kind == AccessibilityFragmentKind::GridHeader && controlRecord)
            {
                const std::optional<size_t> ordinal = FindSizeValueIndex(controlRecord->gridVisibleColumns, gridColumnIndex);
                if (ordinal && ordinal.value() > 0u)
                {
                    return MakeNavigationTarget(AccessibilityFragmentKind::GridHeader, path, 0u, 0u, controlRecord->gridVisibleColumns[ordinal.value() - 1u]);
                }
            }
            else if (kind == AccessibilityFragmentKind::GridRow && controlRecord)
            {
                const std::optional<size_t> ordinal = FindU64ValueIndex(controlRecord->gridVisibleRowIds, gridRowId);
                if (ordinal && ordinal.value() > 0u)
                {
                    return MakeNavigationTarget(AccessibilityFragmentKind::GridRow, path, 0u, controlRecord->gridVisibleRowIds[ordinal.value() - 1u]);
                }
                if (ordinal && ! controlRecord->gridVisibleColumns.empty())
                {
                    return MakeNavigationTarget(AccessibilityFragmentKind::GridHeader, path, 0u, 0u, controlRecord->gridVisibleColumns.back());
                }
            }
            else if (kind == AccessibilityFragmentKind::GridCell && controlRecord && SnapshotContainsGridRow(*controlRecord, gridRowId))
            {
                if (gridColumnIndex > 0u && gridColumnIndex < controlRecord->gridColumnCount)
                {
                    return MakeNavigationTarget(AccessibilityFragmentKind::GridCell, path, 0u, gridRowId, gridColumnIndex - 1u);
                }
            }
            return std::nullopt;
        default: return std::nullopt;
    }
}

[[nodiscard]] std::wstring_view FindAssociatedLabelText(const Control* root, const Control* target, uint32_t depth = 0u) noexcept
{
    if (! root || ! target || depth >= kAccessibilityMaxDepth)
    {
        return {};
    }

    if (const auto* label = dynamic_cast<const Label*>(root))
    {
        if (label->GetMnemonicTarget() == target)
        {
            return label->GetText();
        }
    }

    if (const auto* panel = dynamic_cast<const Panel*>(root))
    {
        const auto children = panel->GetChildren();
        for (const auto& child : children)
        {
            if (! child)
            {
                continue;
            }

            if (const std::wstring_view text = FindAssociatedLabelText(child.get(), target, depth + 1u); ! text.empty())
            {
                return text;
            }
        }
    }

    return {};
}

[[nodiscard]] std::wstring_view GetControlAccessibleName(const Control* root, const Control* control) noexcept
{
    if (! control)
    {
        return {};
    }

    if (const std::wstring_view explicitName = control->GetAccessibleName(); ! explicitName.empty())
    {
        return explicitName;
    }

    if (const auto* toggle = dynamic_cast<const Toggle*>(control))
    {
        if (const std::wstring_view labelText = FindAssociatedLabelText(root, toggle); ! labelText.empty())
        {
            return labelText;
        }
        return toggle->GetDisplayedText();
    }
    if (const auto* button = dynamic_cast<const Button*>(control))
    {
        return button->GetText();
    }
    if (const auto* label = dynamic_cast<const Label*>(control))
    {
        return label->GetText();
    }
    if (const auto* textField = dynamic_cast<const TextField*>(control))
    {
        if (const std::wstring_view labelText = FindAssociatedLabelText(root, textField); ! labelText.empty())
        {
            return labelText;
        }
        if (textField->IsMasked())
        {
            return {};
        }
        return textField->GetText();
    }
    if (const auto* comboBox = dynamic_cast<const ComboBox*>(control))
    {
        if (const std::wstring_view labelText = FindAssociatedLabelText(root, comboBox); ! labelText.empty())
        {
            return labelText;
        }
        return comboBox->GetDisplayedText();
    }
    if (const auto* slider = dynamic_cast<const Slider*>(control))
    {
        if (const std::wstring_view labelText = FindAssociatedLabelText(root, slider); ! labelText.empty())
        {
            return labelText;
        }
    }

    return FindAssociatedLabelText(root, control);
}

[[nodiscard]] std::wstring_view GetControlAccessibleValue(const Control* control) noexcept
{
    if (const auto* textField = dynamic_cast<const TextField*>(control))
    {
        if (textField->IsMasked())
        {
            return {};
        }
        return textField->GetText();
    }
    if (const auto* comboBox = dynamic_cast<const ComboBox*>(control))
    {
        return comboBox->GetDisplayedText();
    }
    return {};
}

[[nodiscard]] bool SupportsTextPattern(const Control* control) noexcept
{
    if (const auto* textField = dynamic_cast<const TextField*>(control))
    {
        return ! textField->IsMasked();
    }
    if (const auto* comboBox = dynamic_cast<const ComboBox*>(control))
    {
        return comboBox->IsEditable();
    }
    return false;
}

[[nodiscard]] std::wstring GetControlAccessibleTextRangeText(const Control* control)
{
    if (const auto* textField = dynamic_cast<const TextField*>(control))
    {
        if (textField->IsMasked())
        {
            return {};
        }
        return std::wstring(textField->GetText());
    }
    if (const auto* comboBox = dynamic_cast<const ComboBox*>(control))
    {
        if (! comboBox->IsEditable())
        {
            return {};
        }
        return std::wstring(comboBox->GetText());
    }
    return {};
}

[[nodiscard]] D2D1_RECT_F ResolveTextPatternViewportRect(const Control* control) noexcept
{
    if (! control)
    {
        return D2D1::RectF();
    }
    if (const std::optional<D2D1_RECT_F> viewport = control->TryGetTextInputViewportRect(); viewport.has_value())
    {
        return viewport.value();
    }
    return control->GetHitBounds();
}

[[nodiscard]] bool IsSimpleTextRangeCaretGeometrySupported(const Control& control, std::wstring_view text) noexcept
{
    const DWRITE_READING_DIRECTION readingDirection = ResolveReadingDirection(control.GetFlowDirection());
    return ! ShouldEmitSingleLineBiDiTextMetric(text, readingDirection);
}

[[nodiscard]] bool TextRangeContainsLineBreak(std::wstring_view text, const TextRangeSpan& range) noexcept
{
    const size_t start = (std::min)(range.start, text.size());
    const size_t end   = (std::min)(range.end, text.size());
    return text.substr(start, end - start).find_first_of(L"\r\n") != std::wstring_view::npos;
}

[[nodiscard]] size_t FindNextLineBreak(std::wstring_view text, size_t start, size_t end) noexcept
{
    const size_t clampedStart = (std::min)(start, text.size());
    const size_t clampedEnd   = (std::min)(end, text.size());
    const size_t breakIndex   = text.substr(clampedStart, clampedEnd - clampedStart).find_first_of(L"\r\n");
    return breakIndex == std::wstring_view::npos ? clampedEnd : clampedStart + breakIndex;
}

[[nodiscard]] size_t AdvancePastLineBreak(std::wstring_view text, size_t index, size_t end) noexcept
{
    const size_t clampedIndex = (std::min)(index, text.size());
    const size_t clampedEnd   = (std::min)(end, text.size());
    if (clampedIndex >= clampedEnd)
    {
        return clampedIndex;
    }
    if (text[clampedIndex] == L'\r' && clampedIndex + 1u < clampedEnd && text[clampedIndex + 1u] == L'\n')
    {
        return clampedIndex + 2u;
    }
    return clampedIndex + 1u;
}

[[nodiscard]] bool CaretRectsShareVisualLine(const D2D1_RECT_F& first, const D2D1_RECT_F& second) noexcept
{
    constexpr float kLineToleranceDip = 1.0f;
    return std::fabs(first.top - second.top) <= kLineToleranceDip && std::fabs(first.bottom - second.bottom) <= kLineToleranceDip;
}

[[nodiscard]] D2D1_RECT_F ClipRectToBounds(D2D1_RECT_F rect, const D2D1_RECT_F& bounds) noexcept
{
    rect.left   = std::clamp(rect.left, bounds.left, bounds.right);
    rect.right  = std::clamp(rect.right, rect.left, bounds.right);
    rect.top    = std::clamp(rect.top, bounds.top, bounds.bottom);
    rect.bottom = std::clamp(rect.bottom, rect.top, bounds.bottom);
    if (rect.right <= rect.left)
    {
        rect.right = (std::min)(bounds.right, rect.left + 1.0f);
    }
    if (rect.bottom <= rect.top)
    {
        rect.bottom = (std::min)(bounds.bottom, rect.top + 1.0f);
    }
    return rect;
}

[[nodiscard]] std::optional<D2D1_RECT_F> TryResolveSimpleTextRangeCaretRect(const WindowHost& host,
                                                                            const Control& control,
                                                                            std::wstring_view text,
                                                                            const TextRangeSpan& range) noexcept
{
    if (range.start == range.end || ! IsSimpleTextRangeCaretGeometrySupported(control, text) || TextRangeContainsLineBreak(text, range))
    {
        return std::nullopt;
    }

    const std::optional<D2D1_RECT_F> startRect = control.TryGetTextInputCaretRect(host, range.start);
    const std::optional<D2D1_RECT_F> endRect   = control.TryGetTextInputCaretRect(host, range.end);
    if (! startRect.has_value() || ! endRect.has_value())
    {
        return std::nullopt;
    }
    if (! CaretRectsShareVisualLine(startRect.value(), endRect.value()))
    {
        return std::nullopt;
    }

    const D2D1_RECT_F viewport = ResolveTextPatternViewportRect(&control);
    const D2D1_RECT_F rect     = D2D1::RectF((std::min)(startRect->left, endRect->left),
                                             (std::min)(startRect->top, endRect->top),
                                             (std::max)(startRect->right, endRect->right),
                                             (std::max)(startRect->bottom, endRect->bottom));
    return ClipRectToBounds(rect, viewport);
}

[[nodiscard]] std::optional<std::vector<D2D1_RECT_F>> TryResolveTextRangeCaretRects(const WindowHost& host,
                                                                                    const Control& control,
                                                                                    std::wstring_view text,
                                                                                    const TextRangeSpan& range)
{
    if (range.start == range.end)
    {
        return std::nullopt;
    }

    if (! IsSimpleTextRangeCaretGeometrySupported(control, text))
    {
        if (std::optional<std::vector<D2D1_RECT_F>> rects = control.TryGetTextInputRangeRects(host, range.start, range.end);
            rects.has_value() && ! rects->empty())
        {
            return rects;
        }
        return std::nullopt;
    }

    if (! TextRangeContainsLineBreak(text, range))
    {
        if (std::optional<D2D1_RECT_F> rect = TryResolveSimpleTextRangeCaretRect(host, control, text, range); rect.has_value())
        {
            return std::vector<D2D1_RECT_F>{rect.value()};
        }
        if (std::optional<std::vector<D2D1_RECT_F>> rects = control.TryGetTextInputRangeRects(host, range.start, range.end);
            rects.has_value() && ! rects->empty())
        {
            return rects;
        }
        return std::nullopt;
    }

    const D2D1_RECT_F viewport = ResolveTextPatternViewportRect(&control);
    std::vector<D2D1_RECT_F> rects;
    size_t segmentStart   = (std::min)(range.start, text.size());
    const size_t rangeEnd = (std::min)(range.end, text.size());
    while (segmentStart < rangeEnd)
    {
        const size_t segmentEnd = FindNextLineBreak(text, segmentStart, rangeEnd);
        if (segmentStart < segmentEnd)
        {
            const std::optional<D2D1_RECT_F> startRect = control.TryGetTextInputCaretRect(host, segmentStart);
            const std::optional<D2D1_RECT_F> endRect   = control.TryGetTextInputCaretRect(host, segmentEnd);
            if (! startRect.has_value() || ! endRect.has_value())
            {
                return std::nullopt;
            }

            if (CaretRectsShareVisualLine(startRect.value(), endRect.value()))
            {
                const D2D1_RECT_F rect = D2D1::RectF((std::min)(startRect->left, endRect->left),
                                                     (std::min)(startRect->top, endRect->top),
                                                     (std::max)(startRect->right, endRect->right),
                                                     (std::max)(startRect->bottom, endRect->bottom));
                rects.push_back(ClipRectToBounds(rect, viewport));
            }
            else if (std::optional<std::vector<D2D1_RECT_F>> segmentRects = control.TryGetTextInputRangeRects(host, segmentStart, segmentEnd);
                     segmentRects.has_value() && ! segmentRects->empty())
            {
                rects.insert(rects.end(), segmentRects->begin(), segmentRects->end());
            }
            else
            {
                return std::nullopt;
            }
        }

        if (segmentEnd >= rangeEnd)
        {
            break;
        }
        segmentStart = AdvancePastLineBreak(text, segmentEnd, rangeEnd);
    }

    if (rects.empty())
    {
        return std::nullopt;
    }
    return rects;
}

[[nodiscard]] size_t GetTextRangeEndpointPosition(const TextRangeSpan& range, TextPatternRangeEndpoint endpoint) noexcept
{
    return endpoint == TextPatternRangeEndpoint_Start ? range.start : range.end;
}

struct TextRangeUnitMoveResult
{
    size_t position = 0u;
    int moved       = 0;
};

[[nodiscard]] TextRangeUnitMoveResult MoveTextRangePositionByUnit(std::wstring_view text, size_t position, TextUnit unit, int count) noexcept
{
    const size_t clampedPosition = std::min(position, text.size());
    if (count == 0)
    {
        return TextRangeUnitMoveResult{clampedPosition, 0};
    }

    if (unit == TextUnit_Character)
    {
        size_t cursor = SnapCaretIndexToTextElementBoundary(text, clampedPosition);
        int moved     = 0;
        if (count > 0)
        {
            while (moved < count)
            {
                const size_t next = StepToNextTextElement(text, cursor);
                if (next == cursor)
                {
                    break;
                }

                cursor = next;
                ++moved;
            }
        }
        else
        {
            while (moved > count)
            {
                const size_t previous = StepToPreviousTextElement(text, cursor);
                if (previous == cursor)
                {
                    break;
                }

                cursor = previous;
                --moved;
            }
        }

        return TextRangeUnitMoveResult{cursor, moved};
    }

    if (unit != TextUnit_Word)
    {
        return TextRangeUnitMoveResult{clampedPosition, 0};
    }

    size_t cursor = clampedPosition;
    int moved     = 0;
    if (count > 0)
    {
        while (moved < count)
        {
            const size_t next = FindNextWordBoundary(text, cursor);
            if (next == cursor)
            {
                break;
            }

            cursor = next;
            ++moved;
        }
    }
    else
    {
        while (moved > count)
        {
            const size_t previous = FindPreviousWordBoundary(text, cursor);
            if (previous == cursor)
            {
                break;
            }

            cursor = previous;
            --moved;
        }
    }

    return TextRangeUnitMoveResult{cursor, moved};
}

[[nodiscard]] TextRangeSpan ClampTextRangeSpan(size_t start, size_t end, size_t textLength) noexcept
{
    start = (std::min)(start, textLength);
    end   = (std::min)(end, textLength);
    if (start > end)
    {
        std::swap(start, end);
    }
    return TextRangeSpan{start, end};
}

[[nodiscard]] bool IsAccessibilityTextRangeWhitespace(wchar_t value) noexcept
{
    WORD charType = 0u;
    return GetStringTypeW(CT_CTYPE1, &value, 1, &charType) && ((charType & C1_SPACE) != 0u);
}

[[nodiscard]] size_t TrimTrailingTextRangeWhitespace(std::wstring_view text, size_t start, size_t end) noexcept
{
    start = (std::min)(start, text.size());
    end   = (std::min)(end, text.size());
    while (end > start)
    {
        const size_t previous = StepToPreviousCodePoint(text, end);
        if (! IsAccessibilityTextRangeWhitespace(text[previous]))
        {
            break;
        }

        end = previous;
    }

    return end;
}

[[nodiscard]] TextRangeSpan GetTextRangeWordSpanAtPosition(std::wstring_view text, size_t position) noexcept
{
    const size_t start = (std::min)(position, text.size());
    if (start >= text.size())
    {
        return TextRangeSpan{text.size(), text.size()};
    }

    const TextRangeUnitMoveResult endBoundary = MoveTextRangePositionByUnit(text, start, TextUnit_Word, 1);
    const size_t end                          = TrimTrailingTextRangeWhitespace(text, start, endBoundary.position);
    return TextRangeSpan{start, end};
}

struct TextRangeSpanMoveResult
{
    TextRangeSpan range{};
    int moved = 0;
};

[[nodiscard]] TextRangeSpanMoveResult MoveTextRangeSpanByWord(std::wstring_view text, const TextRangeSpan& range, int count) noexcept
{
    const TextRangeSpan clampedRange = ClampTextRangeSpan(range.start, range.end, text.size());
    if (count == 0)
    {
        return TextRangeSpanMoveResult{clampedRange, 0};
    }

    size_t cursor = clampedRange.start;
    int moved     = 0;
    if (count > 0)
    {
        while (moved < count)
        {
            const TextRangeUnitMoveResult moveResult = MoveTextRangePositionByUnit(text, cursor, TextUnit_Word, 1);
            if (moveResult.position == cursor || moveResult.moved == 0)
            {
                break;
            }

            cursor = moveResult.position;
            ++moved;
        }
    }
    else
    {
        while (moved > count)
        {
            const TextRangeUnitMoveResult moveResult = MoveTextRangePositionByUnit(text, cursor, TextUnit_Word, -1);
            if (moveResult.position == cursor || moveResult.moved == 0)
            {
                break;
            }

            cursor = moveResult.position;
            --moved;
        }
    }

    if (moved == 0)
    {
        return TextRangeSpanMoveResult{clampedRange, 0};
    }

    return TextRangeSpanMoveResult{GetTextRangeWordSpanAtPosition(text, cursor), moved};
}

[[nodiscard]] bool IsAccessibilityTextRangeLineBreak(wchar_t value) noexcept
{
    return value == L'\r' || value == L'\n';
}

[[nodiscard]] size_t GetTextRangeLineStart(std::wstring_view text, size_t position) noexcept
{
    size_t cursor = (std::min)(position, text.size());
    while (cursor > 0u)
    {
        const size_t previous = StepToPreviousCodePoint(text, cursor);
        if (IsAccessibilityTextRangeLineBreak(text[previous]))
        {
            return cursor;
        }

        cursor = previous;
    }

    return 0u;
}

[[nodiscard]] size_t GetTextRangeLineEnd(std::wstring_view text, size_t lineStart) noexcept
{
    size_t cursor = (std::min)(lineStart, text.size());
    while (cursor < text.size() && ! IsAccessibilityTextRangeLineBreak(text[cursor]))
    {
        cursor = StepToNextCodePoint(text, cursor);
    }

    return cursor;
}

[[nodiscard]] size_t AdvancePastTextRangeLineBreak(std::wstring_view text, size_t position) noexcept
{
    position = (std::min)(position, text.size());
    if (position >= text.size())
    {
        return text.size();
    }

    if (text[position] == L'\r')
    {
        const size_t afterCarriageReturn = StepToNextCodePoint(text, position);
        if (afterCarriageReturn < text.size() && text[afterCarriageReturn] == L'\n')
        {
            return StepToNextCodePoint(text, afterCarriageReturn);
        }

        return afterCarriageReturn;
    }

    if (text[position] == L'\n')
    {
        return StepToNextCodePoint(text, position);
    }

    return position;
}

[[nodiscard]] TextRangeSpan GetTextRangeLineSpanAtPosition(std::wstring_view text, size_t position) noexcept
{
    const size_t lineStart = GetTextRangeLineStart(text, position);
    return TextRangeSpan{lineStart, GetTextRangeLineEnd(text, lineStart)};
}

[[nodiscard]] size_t GetPreviousTextRangeLineStart(std::wstring_view text, size_t lineStart) noexcept
{
    lineStart = (std::min)(lineStart, text.size());
    if (lineStart == 0u)
    {
        return 0u;
    }

    size_t previousLineEnd = StepToPreviousCodePoint(text, lineStart);
    if (text[previousLineEnd] == L'\n' && previousLineEnd > 0u)
    {
        const size_t beforeLineFeed = StepToPreviousCodePoint(text, previousLineEnd);
        if (text[beforeLineFeed] == L'\r')
        {
            previousLineEnd = beforeLineFeed;
        }
    }

    return GetTextRangeLineStart(text, previousLineEnd);
}

[[nodiscard]] TextRangeSpanMoveResult MoveTextRangeSpanByLine(std::wstring_view text, const TextRangeSpan& range, int count) noexcept
{
    const TextRangeSpan clampedRange = ClampTextRangeSpan(range.start, range.end, text.size());
    if (count == 0)
    {
        return TextRangeSpanMoveResult{clampedRange, 0};
    }

    size_t cursor = GetTextRangeLineStart(text, clampedRange.start);
    int moved     = 0;
    if (count > 0)
    {
        while (moved < count)
        {
            const size_t lineEnd = GetTextRangeLineEnd(text, cursor);
            if (lineEnd >= text.size())
            {
                break;
            }

            const size_t nextLineStart = AdvancePastTextRangeLineBreak(text, lineEnd);
            if (nextLineStart == cursor)
            {
                break;
            }

            cursor = nextLineStart;
            ++moved;
        }
    }
    else
    {
        while (moved > count)
        {
            const size_t previousLineStart = GetPreviousTextRangeLineStart(text, cursor);
            if (previousLineStart == cursor)
            {
                break;
            }

            cursor = previousLineStart;
            --moved;
        }
    }

    if (moved == 0)
    {
        return TextRangeSpanMoveResult{clampedRange, 0};
    }

    return TextRangeSpanMoveResult{GetTextRangeLineSpanAtPosition(text, cursor), moved};
}

[[nodiscard]] std::optional<std::vector<TextRangeSpan>> TryResolveTextRangeVisualLineSpans(const WindowHost& host,
                                                                                           const Control& control,
                                                                                           std::wstring_view text) noexcept
{
    if (text.empty() || ! IsSimpleTextRangeCaretGeometrySupported(control, text))
    {
        return std::nullopt;
    }

    std::vector<TextRangeSpan> spans;
    size_t logicalLineCount = 0u;
    size_t logicalLineStart = 0u;
    while (logicalLineStart <= text.size())
    {
        ++logicalLineCount;
        const size_t logicalLineEnd = GetTextRangeLineEnd(text, logicalLineStart);
        if (logicalLineStart == logicalLineEnd)
        {
            spans.push_back(TextRangeSpan{logicalLineStart, logicalLineEnd});
        }
        else
        {
            size_t visualLineStart                    = logicalLineStart;
            std::optional<D2D1_RECT_F> visualLineRect = control.TryGetTextInputCaretRect(host, visualLineStart);
            if (! visualLineRect.has_value())
            {
                return std::nullopt;
            }

            size_t cursor = logicalLineStart;
            while (cursor < logicalLineEnd)
            {
                const size_t next                         = StepToNextCodePoint(text, cursor);
                const std::optional<D2D1_RECT_F> nextRect = control.TryGetTextInputCaretRect(host, next);
                if (! nextRect.has_value())
                {
                    return std::nullopt;
                }

                if (! CaretRectsShareVisualLine(visualLineRect.value(), nextRect.value()))
                {
                    if (visualLineStart < next)
                    {
                        spans.push_back(TextRangeSpan{visualLineStart, next});
                    }
                    visualLineStart = next;
                    visualLineRect  = nextRect;
                }

                cursor = next;
            }

            if (visualLineStart < logicalLineEnd)
            {
                spans.push_back(TextRangeSpan{visualLineStart, logicalLineEnd});
            }
        }

        if (logicalLineEnd >= text.size())
        {
            break;
        }

        const size_t nextLogicalLineStart = AdvancePastTextRangeLineBreak(text, logicalLineEnd);
        if (nextLogicalLineStart <= logicalLineStart)
        {
            break;
        }
        logicalLineStart = nextLogicalLineStart;
    }

    if (spans.size() <= logicalLineCount)
    {
        return std::nullopt;
    }
    return spans;
}

[[nodiscard]] size_t FindTextRangeVisualLineIndex(std::span<const TextRangeSpan> spans, size_t position) noexcept
{
    if (spans.empty())
    {
        return 0u;
    }

    for (size_t index = 0u; index < spans.size(); ++index)
    {
        const TextRangeSpan& span = spans[index];
        if (position == span.start || (position > span.start && position < span.end))
        {
            return index;
        }
        if (position < span.start)
        {
            return index == 0u ? 0u : index - 1u;
        }
    }

    return spans.size() - 1u;
}

[[nodiscard]] std::optional<TextRangeSpanMoveResult> TryMoveTextRangeSpanByVisualLine(
    const WindowHost& host, const Control& control, std::wstring_view text, const TextRangeSpan& range, int count) noexcept
{
    const TextRangeSpan clampedRange = ClampTextRangeSpan(range.start, range.end, text.size());
    if (count == 0)
    {
        return TextRangeSpanMoveResult{clampedRange, 0};
    }

    const std::optional<std::vector<TextRangeSpan>> spans = TryResolveTextRangeVisualLineSpans(host, control, text);
    if (! spans.has_value() || spans->empty())
    {
        return std::nullopt;
    }

    const size_t currentIndex = FindTextRangeVisualLineIndex(*spans, clampedRange.start);
    const int requestedIndex  = static_cast<int>(currentIndex) + count;
    const int targetIndex     = std::clamp(requestedIndex, 0, static_cast<int>(spans->size() - 1u));
    const int moved           = targetIndex - static_cast<int>(currentIndex);
    if (moved == 0)
    {
        return TextRangeSpanMoveResult{clampedRange, 0};
    }

    return TextRangeSpanMoveResult{spans->at(static_cast<size_t>(targetIndex)), moved};
}

[[nodiscard]] TextRangeUnitMoveResult MoveTextRangePositionByLine(std::wstring_view text, size_t position, int count) noexcept
{
    const size_t clampedPosition = (std::min)(position, text.size());
    if (count == 0)
    {
        return TextRangeUnitMoveResult{clampedPosition, 0};
    }

    const TextRangeSpan lineRange            = GetTextRangeLineSpanAtPosition(text, clampedPosition);
    const TextRangeSpanMoveResult moveResult = MoveTextRangeSpanByLine(text, lineRange, count);
    if (moveResult.moved == 0)
    {
        return TextRangeUnitMoveResult{clampedPosition, 0};
    }

    return TextRangeUnitMoveResult{moveResult.range.start, moveResult.moved};
}

[[nodiscard]] std::optional<TextRangeUnitMoveResult> TryMoveTextRangePositionByVisualLine(
    const WindowHost& host, const Control& control, std::wstring_view text, size_t position, int count) noexcept
{
    const size_t clampedPosition = (std::min)(position, text.size());
    if (count == 0)
    {
        return TextRangeUnitMoveResult{clampedPosition, 0};
    }

    const std::optional<TextRangeSpanMoveResult> moveResult =
        TryMoveTextRangeSpanByVisualLine(host, control, text, TextRangeSpan{clampedPosition, clampedPosition}, count);
    if (! moveResult.has_value())
    {
        return std::nullopt;
    }

    return TextRangeUnitMoveResult{moveResult->range.start, moveResult->moved};
}

[[nodiscard]] TextRangeSpan GetTextRangeSpanFromState(size_t caretIndex, std::optional<size_t> selectionAnchorIndex, size_t textLength) noexcept
{
    const size_t clampedCaret = (std::min)(caretIndex, textLength);
    if (! selectionAnchorIndex)
    {
        return TextRangeSpan{clampedCaret, clampedCaret};
    }

    return ClampTextRangeSpan(selectionAnchorIndex.value(), clampedCaret, textLength);
}

[[nodiscard]] TextRangeSpan GetControlAccessibleSelectionRange(const WindowHost* host, const Control* control, size_t textLength) noexcept
{
    if (host && control)
    {
        NativeTextInputState nativeState{};
        if (host->TryReadNativeTextInputState(control, nativeState))
        {
            return GetTextRangeSpanFromState(nativeState.caretIndex, nativeState.selectionAnchorIndex, textLength);
        }

        TextInputState textInputState{};
        if (host->TryReadTextInputState(control, textInputState))
        {
            return GetTextRangeSpanFromState(textInputState.caretIndex, textInputState.selectionAnchorIndex, textLength);
        }
    }

    if (const auto* textField = dynamic_cast<const TextField*>(control))
    {
        if (const std::optional<std::pair<size_t, size_t>> selectionRange = textField->GetSelectionRange())
        {
            return ClampTextRangeSpan(selectionRange->first, selectionRange->second, textLength);
        }
    }
    if (const auto* comboBox = dynamic_cast<const ComboBox*>(control); comboBox && comboBox->IsEditable())
    {
        if (const std::optional<std::pair<size_t, size_t>> selectionRange = comboBox->GetEditableSelectionRange())
        {
            return ClampTextRangeSpan(selectionRange->first, selectionRange->second, textLength);
        }
    }

    return TextRangeSpan{textLength, textLength};
}

[[nodiscard]] std::wstring BuildGridCellAccessibleText(const GridCellData& cellData)
{
    std::wstring text;
    if (cellData.kind == GridCellKind::Checkbox)
    {
        text.assign(cellData.checked ? L"[x]" : L"[ ]");
        if (! cellData.text.empty())
        {
            text.push_back(L' ');
        }
    }

    if (! cellData.text.empty())
    {
        text.append(cellData.text);
    }
    else if (cellData.kind == GridCellKind::ColorSwatch && cellData.hasSwatchValue)
    {
        text.append(std::format(L"#{:08X}", cellData.swatchArgb));
    }
    else if (cellData.kind == GridCellKind::IconText && ! cellData.iconText.empty())
    {
        text.append(cellData.iconText);
    }

    if (! cellData.badgeText.empty())
    {
        if (! text.empty())
        {
            text.push_back(L' ');
        }
        text.push_back(L'[');
        text.append(cellData.badgeText);
        text.push_back(L']');
    }

    return text;
}

[[nodiscard]] std::wstring_view GetGridHeaderAccessibleName(const GridColumnDesc& columnDesc) noexcept
{
    if (! columnDesc.title.empty())
    {
        return columnDesc.title;
    }

    return columnDesc.id;
}

[[nodiscard]] CONTROLTYPEID GetGridCellControlTypeId(const GridCellData& cellData) noexcept
{
    switch (cellData.kind)
    {
        case GridCellKind::Checkbox: return UIA_CheckBoxControlTypeId;
        case GridCellKind::ColorSwatch: return UIA_ImageControlTypeId;
        case GridCellKind::Spinner:
        case GridCellKind::Marquee: return UIA_ProgressBarControlTypeId;
        case GridCellKind::IconText: return (! cellData.iconText.empty() && cellData.text.empty()) ? UIA_ImageControlTypeId : UIA_TextControlTypeId;
        case GridCellKind::Text:
        default: return UIA_TextControlTypeId;
    }
}

[[nodiscard]] bool GridCellSupportsTogglePattern(const GridCellData& cellData) noexcept
{
    return cellData.kind == GridCellKind::Checkbox && cellData.enabled;
}

[[nodiscard]] bool GridCellSupportsValuePattern(const GridCellData& cellData) noexcept
{
    switch (cellData.kind)
    {
        case GridCellKind::Text:
        case GridCellKind::ColorSwatch: return true;
        case GridCellKind::IconText: return ! cellData.text.empty();
        case GridCellKind::Checkbox:
        case GridCellKind::Spinner:
        case GridCellKind::Marquee:
        default: return false;
    }
}

[[nodiscard]] bool GridCellSupportsRangeValuePattern(const GridCellData& cellData) noexcept
{
    return cellData.kind == GridCellKind::Marquee && std::isfinite(cellData.progress) && cellData.progress > 0.0f;
}

[[nodiscard]] double GetGridCellRangeValue(const GridCellData& cellData) noexcept
{
    return std::clamp(static_cast<double>(cellData.progress), 0.0, 1.0);
}

[[nodiscard]] CONTROLTYPEID GetControlTypeId(const Control* control) noexcept
{
    if (dynamic_cast<const Checkbox*>(control) != nullptr)
    {
        return UIA_CheckBoxControlTypeId;
    }
    if (dynamic_cast<const Toggle*>(control) != nullptr || dynamic_cast<const Button*>(control) != nullptr)
    {
        return UIA_ButtonControlTypeId;
    }
    if (dynamic_cast<const TextField*>(control) != nullptr)
    {
        return UIA_EditControlTypeId;
    }
    if (dynamic_cast<const ComboBox*>(control) != nullptr)
    {
        return UIA_ComboBoxControlTypeId;
    }
    if (dynamic_cast<const Slider*>(control) != nullptr)
    {
        return UIA_SliderControlTypeId;
    }
    if (dynamic_cast<const Tree*>(control) != nullptr)
    {
        return UIA_TreeControlTypeId;
    }
    if (dynamic_cast<const Grid*>(control) != nullptr)
    {
        return UIA_DataGridControlTypeId;
    }
    if (dynamic_cast<const ColorSwatch*>(control) != nullptr)
    {
        return UIA_ImageControlTypeId;
    }
    return UIA_TextControlTypeId;
}

[[nodiscard]] bool SupportsInvokePattern(const Control* control) noexcept
{
    return dynamic_cast<const Button*>(control) != nullptr && dynamic_cast<const Toggle*>(control) == nullptr;
}

[[nodiscard]] bool SupportsTogglePattern(const Control* control) noexcept
{
    return dynamic_cast<const Toggle*>(control) != nullptr;
}

[[nodiscard]] bool SupportsValuePattern(const Control* control) noexcept
{
    if (const auto* textField = dynamic_cast<const TextField*>(control))
    {
        return ! textField->IsMultiline();
    }
    return dynamic_cast<const ComboBox*>(control) != nullptr;
}

[[nodiscard]] bool SupportsRangeValuePattern(const Control* control) noexcept
{
    return dynamic_cast<const Slider*>(control) != nullptr;
}

[[nodiscard]] bool IsValueReadOnly(const Control* control) noexcept
{
    if (const auto* textField = dynamic_cast<const TextField*>(control))
    {
        return textField->IsReadOnly();
    }
    if (const auto* comboBox = dynamic_cast<const ComboBox*>(control))
    {
        return ! comboBox->IsEditable();
    }
    if (dynamic_cast<const Slider*>(control) != nullptr)
    {
        return false;
    }
    return true;
}

[[nodiscard]] VARIANT VariantFromBool(bool value) noexcept
{
    VARIANT variant{};
    variant.vt      = VT_BOOL;
    variant.boolVal = value ? VARIANT_TRUE : VARIANT_FALSE;
    return variant;
}

[[nodiscard]] VARIANT VariantFromInt(int value) noexcept
{
    VARIANT variant{};
    variant.vt   = VT_I4;
    variant.lVal = value;
    return variant;
}

[[nodiscard]] VARIANT VariantFromDouble(double value) noexcept
{
    VARIANT variant{};
    variant.vt     = VT_R8;
    variant.dblVal = value;
    return variant;
}

HRESULT SetVariantFromString(VARIANT* outVariant, std::wstring_view value) noexcept
{
    if (! outVariant)
    {
        return E_POINTER;
    }

    VariantInit(outVariant);
    BSTR text = SysAllocStringLen(value.data(), static_cast<UINT>(value.size()));
    if (! text && ! value.empty())
    {
        return E_OUTOFMEMORY;
    }

    outVariant->vt      = VT_BSTR;
    outVariant->bstrVal = text;
    return S_OK;
}

using RuntimeIdValueBuffer = std::array<LONG, kAccessibilityMaxRuntimeIdValueCount>;

[[nodiscard]] bool AppendRuntimeIdValue(RuntimeIdValueBuffer& values, size_t& valueCount, LONG value) noexcept
{
    if (valueCount >= values.size())
    {
        return false;
    }

    values[valueCount] = value;
    ++valueCount;
    return true;
}

[[nodiscard]] bool AppendControlPathRuntimeIdPrefix(RuntimeIdValueBuffer& values, size_t& valueCount, const ControlPath& path) noexcept
{
    for (uint32_t depth = 0u; depth < path.depth; ++depth)
    {
        if (! AppendRuntimeIdValue(values, valueCount, static_cast<LONG>(path.indices[depth])))
        {
            return false;
        }
    }

    return AppendRuntimeIdValue(values, valueCount, static_cast<LONG>(path.depth));
}

[[nodiscard]] bool AppendWindowRuntimeIdPrefix(RuntimeIdValueBuffer& values, size_t& valueCount, HWND hwnd) noexcept
{
    return AppendRuntimeIdValue(values, valueCount, UiaAppendRuntimeId) && AppendRuntimeIdValue(values, valueCount, HandleToLong(hwnd)) &&
           AppendRuntimeIdValue(values, valueCount, static_cast<LONG>((static_cast<ULONG_PTR>(reinterpret_cast<ULONG_PTR>(hwnd)) >> 32) & 0x7FFFFFFF));
}

[[nodiscard]] bool AppendFragmentRuntimeIdPrefix(RuntimeIdValueBuffer& values, size_t& valueCount, const ControlPath& path) noexcept
{
    return AppendRuntimeIdValue(values, valueCount, UiaAppendRuntimeId) && AppendControlPathRuntimeIdPrefix(values, valueCount, path);
}

HRESULT BuildRuntimeId(SAFEARRAY** outArray, std::span<const LONG> values) noexcept
{
    if (! outArray)
    {
        return E_POINTER;
    }

    *outArray = nullptr;
    unique_safearray array(SafeArrayCreateVector(VT_I4, 0, static_cast<ULONG>(values.size())));
    if (! array)
    {
        return E_OUTOFMEMORY;
    }

    for (size_t valueIndex = 0u; valueIndex < values.size(); ++valueIndex)
    {
        LONG arrayIndex  = static_cast<LONG>(valueIndex);
        LONG value       = values[valueIndex];
        const HRESULT hr = SafeArrayPutElement(array.get(), &arrayIndex, &value);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    *outArray = array.release();
    return S_OK;
}

HRESULT SetRuntimeId(SAFEARRAY** outArray, HWND hwnd, const ControlPath* path) noexcept
{
    RuntimeIdValueBuffer values{};
    size_t valueCount = 0u;
    if (! AppendWindowRuntimeIdPrefix(values, valueCount, hwnd))
    {
        return E_INVALIDARG;
    }

    if (path)
    {
        if (! AppendControlPathRuntimeIdPrefix(values, valueCount, *path))
        {
            return E_INVALIDARG;
        }
    }

    return BuildRuntimeId(outArray, std::span<const LONG>(values.data(), valueCount));
}

HRESULT SetTextFieldPasswordRevealButtonRuntimeId(SAFEARRAY** outArray, const ControlPath& path) noexcept
{
    RuntimeIdValueBuffer values{};
    size_t valueCount = 0u;
    if (! AppendFragmentRuntimeIdPrefix(values, valueCount, path) || ! AppendRuntimeIdValue(values, valueCount, kAccessibilityRuntimeIdPasswordRevealButton))
    {
        return E_INVALIDARG;
    }

    return BuildRuntimeId(outArray, std::span<const LONG>(values.data(), valueCount));
}

HRESULT SetTreeItemRuntimeId(SAFEARRAY** outArray, const ControlPath& path, size_t visibleIndex) noexcept
{
    if (! outArray)
    {
        return E_POINTER;
    }

    if (visibleIndex > static_cast<size_t>((std::numeric_limits<LONG>::max)()))
    {
        return E_INVALIDARG;
    }

    RuntimeIdValueBuffer values{};
    size_t valueCount = 0u;
    if (! AppendFragmentRuntimeIdPrefix(values, valueCount, path) || ! AppendRuntimeIdValue(values, valueCount, kAccessibilityRuntimeIdTreeItem) ||
        ! AppendRuntimeIdValue(values, valueCount, static_cast<LONG>(visibleIndex)))
    {
        return E_INVALIDARG;
    }

    return BuildRuntimeId(outArray, std::span<const LONG>(values.data(), valueCount));
}

HRESULT SetGridRowRuntimeId(SAFEARRAY** outArray, const ControlPath& path, uint64_t rowId) noexcept
{
    RuntimeIdValueBuffer values{};
    size_t valueCount = 0u;
    if (! AppendFragmentRuntimeIdPrefix(values, valueCount, path) || ! AppendRuntimeIdValue(values, valueCount, kAccessibilityRuntimeIdGridRow) ||
        ! AppendRuntimeIdValue(values, valueCount, static_cast<LONG>(rowId & 0xFFFFFFFFull)) ||
        ! AppendRuntimeIdValue(values, valueCount, static_cast<LONG>((rowId >> 32u) & 0xFFFFFFFFull)))
    {
        return E_INVALIDARG;
    }

    return BuildRuntimeId(outArray, std::span<const LONG>(values.data(), valueCount));
}

HRESULT SetGridHeaderRuntimeId(SAFEARRAY** outArray, const ControlPath& path, size_t columnIndex) noexcept
{
    if (! outArray)
    {
        return E_POINTER;
    }

    if (columnIndex > static_cast<size_t>((std::numeric_limits<LONG>::max)()))
    {
        return E_INVALIDARG;
    }

    RuntimeIdValueBuffer values{};
    size_t valueCount = 0u;
    if (! AppendFragmentRuntimeIdPrefix(values, valueCount, path) || ! AppendRuntimeIdValue(values, valueCount, kAccessibilityRuntimeIdGridHeader) ||
        ! AppendRuntimeIdValue(values, valueCount, static_cast<LONG>(columnIndex)))
    {
        return E_INVALIDARG;
    }

    return BuildRuntimeId(outArray, std::span<const LONG>(values.data(), valueCount));
}

HRESULT SetGridCellRuntimeId(SAFEARRAY** outArray, const ControlPath& path, uint64_t rowId, size_t columnIndex) noexcept
{
    if (! outArray)
    {
        return E_POINTER;
    }

    if (columnIndex > static_cast<size_t>((std::numeric_limits<LONG>::max)()))
    {
        return E_INVALIDARG;
    }

    RuntimeIdValueBuffer values{};
    size_t valueCount = 0u;
    if (! AppendFragmentRuntimeIdPrefix(values, valueCount, path) || ! AppendRuntimeIdValue(values, valueCount, kAccessibilityRuntimeIdGridCell) ||
        ! AppendRuntimeIdValue(values, valueCount, static_cast<LONG>(rowId & 0xFFFFFFFFull)) ||
        ! AppendRuntimeIdValue(values, valueCount, static_cast<LONG>((rowId >> 32u) & 0xFFFFFFFFull)) ||
        ! AppendRuntimeIdValue(values, valueCount, static_cast<LONG>(columnIndex)))
    {
        return E_INVALIDARG;
    }

    return BuildRuntimeId(outArray, std::span<const LONG>(values.data(), valueCount));
}

HRESULT SetProviderArray(SAFEARRAY** outArray, std::span<IRawElementProviderSimple* const> providers) noexcept
{
    if (! outArray)
    {
        return E_POINTER;
    }

    *outArray = nullptr;
    unique_safearray array(SafeArrayCreateVector(VT_UNKNOWN, 0, static_cast<ULONG>(providers.size())));
    if (! array)
    {
        return E_OUTOFMEMORY;
    }

    for (LONG index = 0; index < static_cast<LONG>(providers.size()); ++index)
    {
        IRawElementProviderSimple* provider = providers[static_cast<size_t>(index)];
        const HRESULT hr                    = SafeArrayPutElement(array.get(), &index, provider);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    *outArray = array.release();
    return S_OK;
}

HRESULT SetTextRangeProviderArray(SAFEARRAY** outArray, std::span<ITextRangeProvider* const> providers) noexcept
{
    if (! outArray)
    {
        return E_POINTER;
    }

    *outArray = nullptr;
    unique_safearray array(SafeArrayCreateVector(VT_UNKNOWN, 0, static_cast<ULONG>(providers.size())));
    if (! array)
    {
        return E_OUTOFMEMORY;
    }

    for (LONG index = 0; index < static_cast<LONG>(providers.size()); ++index)
    {
        ITextRangeProvider* provider = providers[static_cast<size_t>(index)];
        const HRESULT hr             = SafeArrayPutElement(array.get(), &index, provider);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    *outArray = array.release();
    return S_OK;
}

HRESULT SetDoubleArray(SAFEARRAY** outArray, std::span<const double> values) noexcept
{
    if (! outArray)
    {
        return E_POINTER;
    }

    *outArray = nullptr;
    unique_safearray array(SafeArrayCreateVector(VT_R8, 0, static_cast<ULONG>(values.size())));
    if (! array)
    {
        return E_OUTOFMEMORY;
    }

    for (LONG index = 0; index < static_cast<LONG>(values.size()); ++index)
    {
        double value     = values[static_cast<size_t>(index)];
        const HRESULT hr = SafeArrayPutElement(array.get(), &index, &value);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    *outArray = array.release();
    return S_OK;
}

HRESULT SetScreenRectDoubleArray(SAFEARRAY** outArray, HWND hwnd, std::span<const D2D1_RECT_F> boundsDip, float dipToPixelScale) noexcept
{
    if (! std::isfinite(dipToPixelScale) || dipToPixelScale <= 0.0f)
    {
        dipToPixelScale = 1.0f;
    }

    std::vector<double> values;
    values.reserve(boundsDip.size() * 4u);
    for (const D2D1_RECT_F& rectDip : boundsDip)
    {
        POINT topLeft{static_cast<LONG>(std::lround(rectDip.left * dipToPixelScale)), static_cast<LONG>(std::lround(rectDip.top * dipToPixelScale))};
        POINT bottomRight{static_cast<LONG>(std::lround(rectDip.right * dipToPixelScale)), static_cast<LONG>(std::lround(rectDip.bottom * dipToPixelScale))};
        ClientToScreen(hwnd, &topLeft);
        ClientToScreen(hwnd, &bottomRight);
        values.push_back(static_cast<double>(topLeft.x));
        values.push_back(static_cast<double>(topLeft.y));
        values.push_back(static_cast<double>(bottomRight.x - topLeft.x));
        values.push_back(static_cast<double>(bottomRight.y - topLeft.y));
    }

    return SetDoubleArray(outArray, values);
}

class AccessibilityTextRangeProvider final : public ITextRangeProvider
{
public:
    AccessibilityTextRangeProvider(WindowHostAccessibilityTarget* target, HWND hwnd, const ControlPath& path, size_t start, size_t end) noexcept
        : _target(target),
          _hwnd(hwnd),
          _snapshot(CaptureAccessibilitySnapshot(target, hwnd)),
          _path(path),
          _rangeStart(start),
          _rangeEnd(end)
    {
    }

    AccessibilityTextRangeProvider(
        WindowHostAccessibilityTarget* target, HWND hwnd, const ControlPath& path, size_t start, size_t end, std::wstring textOverride) noexcept
        : _target(target),
          _hwnd(hwnd),
          _snapshot(CaptureAccessibilitySnapshot(target, hwnd)),
          _textOverride(std::move(textOverride)),
          _path(path),
          _rangeStart(start),
          _rangeEnd(end)
    {
    }

    AccessibilityTextRangeProvider(WindowHostAccessibilityTarget* target,
                                   HWND hwnd,
                                   const ControlPath& path,
                                   size_t start,
                                   size_t end,
                                   std::vector<D2D1_RECT_F> boundsOverrideDip) noexcept
        : _target(target),
          _hwnd(hwnd),
          _snapshot(CaptureAccessibilitySnapshot(target, hwnd)),
          _boundsOverrideDip(std::move(boundsOverrideDip)),
          _path(path),
          _rangeStart(start),
          _rangeEnd(end)
    {
    }

    AccessibilityTextRangeProvider(const AccessibilityTextRangeProvider&)            = delete;
    AccessibilityTextRangeProvider& operator=(const AccessibilityTextRangeProvider&) = delete;
    AccessibilityTextRangeProvider(AccessibilityTextRangeProvider&&)                 = delete;
    AccessibilityTextRangeProvider& operator=(AccessibilityTextRangeProvider&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override;
    ULONG STDMETHODCALLTYPE AddRef() noexcept override;
    ULONG STDMETHODCALLTYPE Release() noexcept override;

    HRESULT STDMETHODCALLTYPE Clone(ITextRangeProvider** outClone) noexcept override;
    HRESULT STDMETHODCALLTYPE Compare(ITextRangeProvider* range, BOOL* outSame) noexcept override;
    HRESULT STDMETHODCALLTYPE CompareEndpoints(TextPatternRangeEndpoint endpoint,
                                               ITextRangeProvider* targetRange,
                                               TextPatternRangeEndpoint targetEndpoint,
                                               int* outResult) noexcept override;
    HRESULT STDMETHODCALLTYPE ExpandToEnclosingUnit(TextUnit unit) noexcept override;
    HRESULT STDMETHODCALLTYPE FindAttribute(TEXTATTRIBUTEID attributeId, VARIANT value, BOOL backward, ITextRangeProvider** outRange) noexcept override;
    HRESULT STDMETHODCALLTYPE FindText(BSTR text, BOOL backward, BOOL ignoreCase, ITextRangeProvider** outRange) noexcept override;
    HRESULT STDMETHODCALLTYPE GetAttributeValue(TEXTATTRIBUTEID attributeId, VARIANT* outValue) noexcept override;
    HRESULT STDMETHODCALLTYPE GetBoundingRectangles(SAFEARRAY** outRectangles) noexcept override;
    HRESULT STDMETHODCALLTYPE GetEnclosingElement(IRawElementProviderSimple** outElement) noexcept override;
    HRESULT STDMETHODCALLTYPE GetText(int maxLength, BSTR* outText) noexcept override;
    HRESULT STDMETHODCALLTYPE Move(TextUnit unit, int count, int* outMoved) noexcept override;
    HRESULT STDMETHODCALLTYPE MoveEndpointByUnit(TextPatternRangeEndpoint endpoint, TextUnit unit, int count, int* outMoved) noexcept override;
    HRESULT STDMETHODCALLTYPE MoveEndpointByRange(TextPatternRangeEndpoint endpoint,
                                                  ITextRangeProvider* targetRange,
                                                  TextPatternRangeEndpoint targetEndpoint) noexcept override;
    HRESULT STDMETHODCALLTYPE Select() noexcept override;
    HRESULT STDMETHODCALLTYPE AddToSelection() noexcept override;
    HRESULT STDMETHODCALLTYPE RemoveFromSelection() noexcept override;
    HRESULT STDMETHODCALLTYPE ScrollIntoView(BOOL alignToTop) noexcept override;
    HRESULT STDMETHODCALLTYPE GetChildren(SAFEARRAY** outChildren) noexcept override;
    HRESULT ExecuteSelectOnWindowThread() noexcept;
    HRESULT ExecuteMoveByVisualLineOnWindowThread(size_t start, size_t end, int count, size_t& outStart, size_t& outEnd, int& outMoved) noexcept;
    HRESULT ExecuteMoveEndpointByVisualLineOnWindowThread(
        size_t start, size_t end, TextPatternRangeEndpoint endpoint, int count, size_t& outStart, size_t& outEnd, int& outMoved) noexcept;
    HRESULT ExecuteResolveBoundsOnWindowThread(size_t start, size_t end, std::vector<D2D1_RECT_F>& outBoundsDip, float& outDipToPixelScale) noexcept;

private:
    [[nodiscard]] bool IsCurrentThreadWindowThread() const noexcept;
    HRESULT DispatchActionToWindowThread(AccessibilityUiActionRequest& request) const noexcept;
    HRESULT DispatchLineMovementToWindowThread(size_t start, size_t end, int count, size_t& outStart, size_t& outEnd, int& outMoved) noexcept;
    HRESULT DispatchEndpointLineMovementToWindowThread(
        size_t start, size_t end, TextPatternRangeEndpoint endpoint, int count, size_t& outStart, size_t& outEnd, int& outMoved) noexcept;
    HRESULT DispatchBoundingRectanglesToWindowThread(size_t start, size_t end, std::vector<D2D1_RECT_F>& outBoundsDip, float& outDipToPixelScale) noexcept;
    [[nodiscard]] WindowHost* ResolveHost() const noexcept;
    [[nodiscard]] const Control* ResolveControl() const noexcept;
    [[nodiscard]] Control* ResolveMutableControl() const noexcept;
    [[nodiscard]] std::wstring ResolveText() const;
    [[nodiscard]] TextRangeSpan ClampCurrentRange(size_t textLength) const noexcept;
    [[nodiscard]] ITextRangeProvider* CreateRange(size_t start, size_t end) const noexcept;

    std::atomic<ULONG> _referenceCount{1u};
    WindowHostAccessibilityTarget* _target = nullptr;
    HWND _hwnd                             = nullptr;
    std::shared_ptr<const AccessibilitySnapshot> _snapshot;
    std::optional<std::wstring> _textOverride;
    std::optional<std::vector<D2D1_RECT_F>> _boundsOverrideDip;
    ControlPath _path{};
    size_t _rangeStart = 0u;
    size_t _rangeEnd   = 0u;
};

class AccessibilityProvider final : public IRawElementProviderSimple,
                                    public IRawElementProviderFragment,
                                    public IRawElementProviderFragmentRoot,
                                    public IInvokeProvider,
                                    public ITextEditProvider,
                                    public ITableProvider,
                                    public IToggleProvider,
                                    public IValueProvider,
                                    public IRangeValueProvider,
                                    public ISelectionProvider,
                                    public ISelectionItemProvider,
                                    public IExpandCollapseProvider,
                                    public IGridItemProvider,
                                    public ITableItemProvider
{
public:
    struct GridHeaderTag
    {
    };

    AccessibilityProvider(WindowHostAccessibilityTarget* target, HWND hwnd) noexcept
        : _target(target),
          _hwnd(hwnd),
          _snapshot(CaptureAccessibilitySnapshot(target, hwnd))
    {
    }

    AccessibilityProvider(WindowHostAccessibilityTarget* target, HWND hwnd, const ControlPath& path) noexcept
        : _target(target),
          _hwnd(hwnd),
          _snapshot(CaptureAccessibilitySnapshot(target, hwnd)),
          _path(path),
          _kind(AccessibilityFragmentKind::Control)
    {
    }

    AccessibilityProvider(WindowHostAccessibilityTarget* target, HWND hwnd, const ControlPath& path, AccessibilityFragmentKind kind) noexcept
        : _target(target),
          _hwnd(hwnd),
          _snapshot(CaptureAccessibilitySnapshot(target, hwnd)),
          _path(path),
          _kind(kind)
    {
    }

    AccessibilityProvider(WindowHostAccessibilityTarget* target, HWND hwnd, const ControlPath& path, size_t treeVisibleIndex) noexcept
        : _target(target),
          _hwnd(hwnd),
          _snapshot(CaptureAccessibilitySnapshot(target, hwnd)),
          _path(path),
          _kind(AccessibilityFragmentKind::TreeItem),
          _treeVisibleIndex(treeVisibleIndex)
    {
    }

    AccessibilityProvider(WindowHostAccessibilityTarget* target, HWND hwnd, const ControlPath& path, size_t gridColumnIndex, GridHeaderTag) noexcept
        : _target(target),
          _hwnd(hwnd),
          _snapshot(CaptureAccessibilitySnapshot(target, hwnd)),
          _path(path),
          _kind(AccessibilityFragmentKind::GridHeader),
          _gridColumnIndex(gridColumnIndex)
    {
    }

    AccessibilityProvider(
        WindowHostAccessibilityTarget* target, HWND hwnd, const ControlPath& path, uint64_t gridRowId, AccessibilityFragmentKind kind) noexcept
        : _target(target),
          _hwnd(hwnd),
          _snapshot(CaptureAccessibilitySnapshot(target, hwnd)),
          _path(path),
          _kind(kind),
          _gridRowId(gridRowId)
    {
    }

    AccessibilityProvider(WindowHostAccessibilityTarget* target, HWND hwnd, const ControlPath& path, uint64_t gridRowId, size_t gridColumnIndex) noexcept
        : _target(target),
          _hwnd(hwnd),
          _snapshot(CaptureAccessibilitySnapshot(target, hwnd)),
          _path(path),
          _kind(AccessibilityFragmentKind::GridCell),
          _gridRowId(gridRowId),
          _gridColumnIndex(gridColumnIndex)
    {
    }

    AccessibilityProvider(const AccessibilityProvider&)            = delete;
    AccessibilityProvider& operator=(const AccessibilityProvider&) = delete;
    AccessibilityProvider(AccessibilityProvider&&)                 = delete;
    AccessibilityProvider& operator=(AccessibilityProvider&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override;
    ULONG STDMETHODCALLTYPE AddRef() noexcept override;
    ULONG STDMETHODCALLTYPE Release() noexcept override;

    HRESULT STDMETHODCALLTYPE get_ProviderOptions(ProviderOptions* outOptions) noexcept override;
    HRESULT STDMETHODCALLTYPE GetPatternProvider(PATTERNID patternId, IUnknown** outProvider) noexcept override;
    HRESULT STDMETHODCALLTYPE GetPropertyValue(PROPERTYID propertyId, VARIANT* outValue) noexcept override;
    HRESULT STDMETHODCALLTYPE get_HostRawElementProvider(IRawElementProviderSimple** outProvider) noexcept override;

    HRESULT STDMETHODCALLTYPE Navigate(NavigateDirection direction, IRawElementProviderFragment** outProvider) noexcept override;
    HRESULT STDMETHODCALLTYPE GetRuntimeId(SAFEARRAY** outRuntimeId) noexcept override;
    HRESULT STDMETHODCALLTYPE get_BoundingRectangle(UiaRect* outRect) noexcept override;
    HRESULT STDMETHODCALLTYPE GetEmbeddedFragmentRoots(SAFEARRAY** outRoots) noexcept override;
    HRESULT STDMETHODCALLTYPE SetFocus() noexcept override;
    HRESULT STDMETHODCALLTYPE get_FragmentRoot(IRawElementProviderFragmentRoot** outRoot) noexcept override;

    HRESULT STDMETHODCALLTYPE ElementProviderFromPoint(double x, double y, IRawElementProviderFragment** outProvider) noexcept override;
    HRESULT STDMETHODCALLTYPE GetFocus(IRawElementProviderFragment** outProvider) noexcept override;

    HRESULT STDMETHODCALLTYPE Invoke() noexcept override;
    HRESULT STDMETHODCALLTYPE Toggle() noexcept override;
    HRESULT STDMETHODCALLTYPE get_ToggleState(ToggleState* outState) noexcept override;
    HRESULT STDMETHODCALLTYPE GetVisibleRanges(SAFEARRAY** outRanges) noexcept override;
    HRESULT STDMETHODCALLTYPE RangeFromChild(IRawElementProviderSimple* childElement, ITextRangeProvider** outRange) noexcept override;
    HRESULT STDMETHODCALLTYPE RangeFromPoint(UiaPoint point, ITextRangeProvider** outRange) noexcept override;
    HRESULT STDMETHODCALLTYPE get_DocumentRange(ITextRangeProvider** outRange) noexcept override;
    HRESULT STDMETHODCALLTYPE get_SupportedTextSelection(SupportedTextSelection* outSupportedSelection) noexcept override;
    HRESULT STDMETHODCALLTYPE GetActiveComposition(ITextRangeProvider** outRange) noexcept override;
    HRESULT STDMETHODCALLTYPE GetConversionTarget(ITextRangeProvider** outRange) noexcept override;
    HRESULT STDMETHODCALLTYPE SetValue(LPCWSTR value) noexcept override;
    HRESULT STDMETHODCALLTYPE SetValue(double value) noexcept override;
    HRESULT STDMETHODCALLTYPE get_Value(BSTR* outValue) noexcept override;
    HRESULT STDMETHODCALLTYPE get_Value(double* outValue) noexcept override;
    HRESULT STDMETHODCALLTYPE get_IsReadOnly(BOOL* outReadOnly) noexcept override;
    HRESULT STDMETHODCALLTYPE get_Maximum(double* outMaximum) noexcept override;
    HRESULT STDMETHODCALLTYPE get_Minimum(double* outMinimum) noexcept override;
    HRESULT STDMETHODCALLTYPE get_LargeChange(double* outLargeChange) noexcept override;
    HRESULT STDMETHODCALLTYPE get_SmallChange(double* outSmallChange) noexcept override;
    HRESULT STDMETHODCALLTYPE GetSelection(SAFEARRAY** outSelection) noexcept override;
    HRESULT STDMETHODCALLTYPE get_CanSelectMultiple(BOOL* outCanSelectMultiple) noexcept override;
    HRESULT STDMETHODCALLTYPE get_IsSelectionRequired(BOOL* outIsSelectionRequired) noexcept override;
    HRESULT STDMETHODCALLTYPE GetRowHeaders(SAFEARRAY** outRowHeaders) noexcept override;
    HRESULT STDMETHODCALLTYPE GetColumnHeaders(SAFEARRAY** outColumnHeaders) noexcept override;
    HRESULT STDMETHODCALLTYPE get_RowOrColumnMajor(RowOrColumnMajor* outRowOrColumnMajor) noexcept override;
    HRESULT STDMETHODCALLTYPE Select() noexcept override;
    HRESULT STDMETHODCALLTYPE AddToSelection() noexcept override;
    HRESULT STDMETHODCALLTYPE RemoveFromSelection() noexcept override;
    HRESULT STDMETHODCALLTYPE get_IsSelected(BOOL* outSelected) noexcept override;
    HRESULT STDMETHODCALLTYPE get_SelectionContainer(IRawElementProviderSimple** outContainer) noexcept override;
    HRESULT STDMETHODCALLTYPE Expand() noexcept override;
    HRESULT STDMETHODCALLTYPE Collapse() noexcept override;
    HRESULT STDMETHODCALLTYPE get_ExpandCollapseState(ExpandCollapseState* outState) noexcept override;
    HRESULT STDMETHODCALLTYPE get_Row(int* outRow) noexcept override;
    HRESULT STDMETHODCALLTYPE get_Column(int* outColumn) noexcept override;
    HRESULT STDMETHODCALLTYPE get_RowSpan(int* outRowSpan) noexcept override;
    HRESULT STDMETHODCALLTYPE get_ColumnSpan(int* outColumnSpan) noexcept override;
    HRESULT STDMETHODCALLTYPE get_ContainingGrid(IRawElementProviderSimple** outContainingGrid) noexcept override;
    HRESULT STDMETHODCALLTYPE GetRowHeaderItems(SAFEARRAY** outRowHeaderItems) noexcept override;
    HRESULT STDMETHODCALLTYPE GetColumnHeaderItems(SAFEARRAY** outColumnHeaderItems) noexcept override;
    HRESULT ExecuteUiThreadAction(AccessibilityUiActionRequest& request) noexcept;

private:
    [[nodiscard]] WindowHost* ResolveHost() const noexcept;
    [[nodiscard]] const Control* ResolveRootControl() const noexcept;
    [[nodiscard]] bool ResolveControlPath(ControlPath& outPath) const noexcept;
    [[nodiscard]] const Control* ResolveControl() const noexcept;
    [[nodiscard]] Control* ResolveMutableControl() const noexcept;
    [[nodiscard]] const Tree* ResolveTreeControl() const noexcept;
    [[nodiscard]] Tree* ResolveMutableTreeControl() const noexcept;
    [[nodiscard]] const Grid* ResolveGridControl() const noexcept;
    [[nodiscard]] Grid* ResolveMutableGridControl() const noexcept;
    [[nodiscard]] bool ResolveTreeItemData(TreeItemData& outItem) const noexcept;
    [[nodiscard]] bool ResolveGridRowIndex(size_t& outRowIndex) const noexcept;
    [[nodiscard]] bool ResolveGridCellData(size_t& outRowIndex, size_t& outColumnIndex, GridCellData& outCellData) const noexcept;
    [[nodiscard]] bool SupportsTreeItemSelectionPattern() const noexcept;
    [[nodiscard]] bool SupportsTreeItemExpandCollapsePattern() const noexcept;
    [[nodiscard]] AccessibilityPatternQueryResult QueryPattern(AccessibilityPatternKind patternKind) noexcept;
    template <typename TInterface, typename TProvider, typename... Args> [[nodiscard]] TInterface* MakeProvider(Args&&... args) const noexcept
    {
        WindowHostAccessibilityTarget* target = AddRefTarget();
        auto* provider                        = new (std::nothrow) TProvider(target, std::forward<Args>(args)...);
        if (! provider && target)
        {
            static_cast<void>(target->Release());
        }
        return provider ? static_cast<TInterface*>(provider) : nullptr;
    }
    [[nodiscard]] IRawElementProviderFragmentRoot* CreateRootProvider() noexcept;
    [[nodiscard]] IRawElementProviderFragment* CreateChildProvider(const ControlPath& path) noexcept;
    [[nodiscard]] IRawElementProviderFragment* CreateTextFieldPasswordRevealButtonProvider(const ControlPath& path) noexcept;
    [[nodiscard]] IRawElementProviderFragment* CreateTreeItemProvider(const ControlPath& path, size_t visibleIndex) noexcept;
    [[nodiscard]] IRawElementProviderFragment* CreateGridHeaderProvider(const ControlPath& path, size_t columnIndex) noexcept;
    [[nodiscard]] IRawElementProviderFragment* CreateGridRowProvider(const ControlPath& path, uint64_t rowId) noexcept;
    [[nodiscard]] IRawElementProviderFragment* CreateGridCellProvider(const ControlPath& path, uint64_t rowId, size_t columnIndex) noexcept;
    [[nodiscard]] IRawElementProviderFragment* CreateProviderFromNavigationTarget(const AccessibilityNavigationTarget& navigationTarget) noexcept;
    [[nodiscard]] ITextRangeProvider* CreateTextRangeProvider(const ControlPath& path, size_t start, size_t end) noexcept;
    [[nodiscard]] ITextRangeProvider* CreateTextRangeProvider(const ControlPath& path,
                                                              size_t start,
                                                              size_t end,
                                                              std::vector<D2D1_RECT_F> boundsOverrideDip) noexcept;
    [[nodiscard]] ITextRangeProvider* CreateTextRangeProvider(const ControlPath& path, size_t start, size_t end, std::wstring textOverride) noexcept;
    [[nodiscard]] ITextRangeProvider* CreateTextDocumentRangeProvider(const AccessibilityControlNavigationSnapshot& record) noexcept;
    [[nodiscard]] WindowHostAccessibilityTarget* AddRefTarget() const noexcept;
    [[nodiscard]] bool IsCurrentThreadWindowThread() const noexcept;
    HRESULT DispatchActionToWindowThread(AccessibilityUiActionRequest& request) noexcept;
    HRESULT ExecuteSetFocusOnWindowThread() noexcept;
    HRESULT ExecuteInvokeOnWindowThread() noexcept;
    HRESULT ExecuteToggleOnWindowThread() noexcept;
    HRESULT ExecuteSetStringValueOnWindowThread(LPCWSTR value) noexcept;
    HRESULT ExecuteSetRangeValueOnWindowThread(double value) noexcept;
    HRESULT ExecuteResolveTextRangeFromPointOnWindowThread(UiaPoint point, size_t& outCaretIndex, size_t& outTextLength) noexcept;
    HRESULT ExecuteSelectOnWindowThread() noexcept;
    HRESULT ExecuteAddToSelectionOnWindowThread() noexcept;
    HRESULT ExecuteRemoveFromSelectionOnWindowThread() noexcept;
    HRESULT ExecuteExpandOnWindowThread(bool expanded) noexcept;

    std::atomic<ULONG> _referenceCount{1u};
    WindowHostAccessibilityTarget* _target = nullptr;
    HWND _hwnd                             = nullptr;
    std::shared_ptr<const AccessibilitySnapshot> _snapshot;
    ControlPath _path{};
    AccessibilityFragmentKind _kind = AccessibilityFragmentKind::Root;
    size_t _treeVisibleIndex        = 0u;
    uint64_t _gridRowId             = 0u;
    size_t _gridColumnIndex         = 0u;
};

HRESULT AccessibilityTextRangeProvider::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    if (! ppvObject)
    {
        return E_POINTER;
    }

    *ppvObject = nullptr;
    if (riid == __uuidof(IUnknown) || riid == __uuidof(ITextRangeProvider))
    {
        *ppvObject = static_cast<ITextRangeProvider*>(this);
        AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}

ULONG AccessibilityTextRangeProvider::AddRef() noexcept
{
    return _referenceCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
}

ULONG AccessibilityTextRangeProvider::Release() noexcept
{
    const ULONG remaining = _referenceCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
    if (remaining == 0u)
    {
        if (_target)
        {
            static_cast<void>(_target->Release());
            _target = nullptr;
        }
        delete this;
    }
    return remaining;
}

HRESULT AccessibilityTextRangeProvider::Clone(ITextRangeProvider** outClone) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outClone)
    {
        return E_POINTER;
    }

    *outClone = CreateRange(_rangeStart, _rangeEnd);
    return *outClone ? S_OK : E_OUTOFMEMORY;
}

HRESULT AccessibilityTextRangeProvider::Compare(ITextRangeProvider* range, BOOL* outSame) noexcept
{
    if (! outSame)
    {
        return E_POINTER;
    }

    const auto* other = dynamic_cast<AccessibilityTextRangeProvider*>(range);
    *outSame = (other && _target == other->_target && _hwnd == other->_hwnd && AreControlPathsEqual(_path, other->_path) && _rangeStart == other->_rangeStart &&
                _rangeEnd == other->_rangeEnd)
                   ? TRUE
                   : FALSE;
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::CompareEndpoints(TextPatternRangeEndpoint endpoint,
                                                         ITextRangeProvider* targetRange,
                                                         TextPatternRangeEndpoint targetEndpoint,
                                                         int* outResult) noexcept
{
    if (! outResult)
    {
        return E_POINTER;
    }
    if (! targetRange)
    {
        return E_INVALIDARG;
    }

    const auto* other = dynamic_cast<AccessibilityTextRangeProvider*>(targetRange);
    if (! other || _target != other->_target || _hwnd != other->_hwnd || ! AreControlPathsEqual(_path, other->_path))
    {
        return E_INVALIDARG;
    }

    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    const std::wstring text        = ResolveText();
    const TextRangeSpan range      = ClampCurrentRange(text.size());
    const TextRangeSpan otherRange = other->ClampCurrentRange(text.size());
    const size_t position          = GetTextRangeEndpointPosition(range, endpoint);
    const size_t targetPosition    = GetTextRangeEndpointPosition(otherRange, targetEndpoint);

    *outResult = position < targetPosition ? -1 : (position > targetPosition ? 1 : 0);
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::ExpandToEnclosingUnit(TextUnit unit) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    const std::wstring text = ResolveText();
    _boundsOverrideDip.reset();
    if (unit == TextUnit_Document)
    {
        _rangeStart = 0u;
        _rangeEnd   = text.size();
    }
    else
    {
        const TextRangeSpan range = ClampCurrentRange(text.size());
        _rangeStart               = range.start;
        _rangeEnd                 = range.end;
    }
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::FindAttribute(TEXTATTRIBUTEID /*attributeId*/,
                                                      VARIANT /*value*/,
                                                      BOOL /*backward*/,
                                                      ITextRangeProvider** outRange) noexcept
{
    if (! outRange)
    {
        return E_POINTER;
    }

    *outRange = nullptr;
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::FindText(BSTR /*text*/, BOOL /*backward*/, BOOL /*ignoreCase*/, ITextRangeProvider** outRange) noexcept
{
    if (! outRange)
    {
        return E_POINTER;
    }

    *outRange = nullptr;
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::GetAttributeValue(TEXTATTRIBUTEID /*attributeId*/, VARIANT* outValue) noexcept
{
    if (! outValue)
    {
        return E_POINTER;
    }

    VariantInit(outValue);
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::GetBoundingRectangles(SAFEARRAY** outRectangles) noexcept
{
    if (! outRectangles)
    {
        return E_POINTER;
    }

    const auto startedAt = std::chrono::steady_clock::now();
    *outRectangles       = nullptr;

    TextRangeSpan range{};
    std::optional<std::vector<D2D1_RECT_F>> boundsOverrideDip;
    size_t textLength = 0u;
    {
        const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
        const std::wstring text = ResolveText();
        textLength              = text.size();
        range                   = ClampCurrentRange(textLength);
        if (_boundsOverrideDip && ! _boundsOverrideDip->empty())
        {
            boundsOverrideDip = _boundsOverrideDip.value();
        }
    }

    if (range.start == range.end)
    {
        const HRESULT hr = SetDoubleArray(outRectangles, {});
        Debug::Perf::Emit(L"dxui.uia.text_range_us", L"bounding-rectangles-empty", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, hr);
        return hr;
    }

    if (boundsOverrideDip && ! boundsOverrideDip->empty())
    {
        const float dipToPixelScale = (_snapshot && _snapshot->pixelsToDipScale > 0.0f) ? 1.0f / _snapshot->pixelsToDipScale : 1.0f;
        const HRESULT hr            = SetScreenRectDoubleArray(outRectangles, _hwnd, boundsOverrideDip.value(), dipToPixelScale);
        Debug::Perf::Emit(L"dxui.uia.text_range_us",
                          L"bounding-rectangles-snapshot",
                          Debug::Perf::ElapsedUs(startedAt),
                          boundsOverrideDip->size(),
                          range.end - range.start,
                          hr);
        return hr;
    }

    if (! IsCurrentThreadWindowThread())
    {
        std::vector<D2D1_RECT_F> boundsDip;
        float dipToPixelScale    = 1.0f;
        const HRESULT dispatchHr = DispatchBoundingRectanglesToWindowThread(range.start, range.end, boundsDip, dipToPixelScale);
        if (FAILED(dispatchHr))
        {
            Debug::Perf::Emit(L"dxui.uia.text_range_us", L"bounding-rectangles-dispatch", Debug::Perf::ElapsedUs(startedAt), 0u, textLength, dispatchHr);
            return dispatchHr;
        }

        const HRESULT hr = SetScreenRectDoubleArray(outRectangles, _hwnd, boundsDip, dipToPixelScale);
        Debug::Perf::Emit(
            L"dxui.uia.text_range_us", L"bounding-rectangles-dispatch", Debug::Perf::ElapsedUs(startedAt), boundsDip.size(), range.end - range.start, hr);
        return hr;
    }

    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    WindowHost* const host       = ResolveHost();
    const Control* const control = ResolveControl();
    if (! host || ! control)
    {
        const HRESULT hr = SetDoubleArray(outRectangles, {});
        Debug::Perf::Emit(L"dxui.uia.text_range_us", L"bounding-rectangles-empty", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, hr);
        return hr;
    }

    const std::wstring text = _textOverride ? _textOverride.value() : GetControlAccessibleTextRangeText(control);
    range                   = ClampCurrentRange(text.size());
    if (range.start == range.end)
    {
        const HRESULT hr = SetDoubleArray(outRectangles, {});
        Debug::Perf::Emit(L"dxui.uia.text_range_us", L"bounding-rectangles-empty", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, hr);
        return hr;
    }

    std::vector<D2D1_RECT_F> boundsDip =
        TryResolveTextRangeCaretRects(*host, *control, text, range).value_or(std::vector<D2D1_RECT_F>{ResolveTextPatternViewportRect(control)});
    const float dipToPixelScale = static_cast<float>(host->GetDpi()) / static_cast<float>(USER_DEFAULT_SCREEN_DPI);
    const HRESULT hr            = SetScreenRectDoubleArray(outRectangles, _hwnd, boundsDip, dipToPixelScale);
    Debug::Perf::Emit(L"dxui.uia.text_range_us", L"bounding-rectangles", Debug::Perf::ElapsedUs(startedAt), boundsDip.size(), range.end - range.start, hr);
    return hr;
}

HRESULT AccessibilityTextRangeProvider::GetEnclosingElement(IRawElementProviderSimple** outElement) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outElement)
    {
        return E_POINTER;
    }

    *outElement = nullptr;
    const AccessibilityControlNavigationSnapshot* record =
        (_snapshot && _snapshot->alive && _snapshot->hasRetainedRoot) ? FindControlNavigationRecord(*_snapshot, _path) : nullptr;
    if (! record || ! record->controlSupportsText)
    {
        return UIA_E_ELEMENTNOTAVAILABLE;
    }

    WindowHostAccessibilityTarget* target = _target;
    if (target)
    {
        static_cast<void>(target->AddRef());
    }

    auto* provider = new (std::nothrow) AccessibilityProvider(target, _hwnd, record->path);
    if (! provider)
    {
        if (target)
        {
            static_cast<void>(target->Release());
        }
        return E_OUTOFMEMORY;
    }

    *outElement = static_cast<IRawElementProviderSimple*>(provider);
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::GetText(int maxLength, BSTR* outText) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outText)
    {
        return E_POINTER;
    }

    *outText                  = nullptr;
    const std::wstring text   = ResolveText();
    const TextRangeSpan range = ClampCurrentRange(text.size());
    std::wstring_view value(text.data() + range.start, range.end - range.start);
    if (maxLength >= 0)
    {
        value = value.substr(0u, (std::min)(value.size(), static_cast<size_t>(maxLength)));
    }

    *outText         = SysAllocStringLen(value.data(), static_cast<UINT>(value.size()));
    const HRESULT hr = (*outText || value.empty()) ? S_OK : E_OUTOFMEMORY;
    Debug::Perf::Emit(L"dxui.uia.text_range_us", L"gettext", Debug::Perf::ElapsedUs(startedAt), value.size(), text.size(), hr);
    return hr;
}

HRESULT AccessibilityTextRangeProvider::Move(TextUnit unit, int count, int* outMoved) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    if (! outMoved)
    {
        return E_POINTER;
    }

    *outMoved = 0;
    if ((unit != TextUnit_Character && unit != TextUnit_Word && unit != TextUnit_Line) || count == 0)
    {
        const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
        const std::wstring text = ResolveText();
        Debug::Perf::Emit(L"dxui.uia.text_range_us", L"move-unsupported", Debug::Perf::ElapsedUs(startedAt), 0u, text.size(), S_OK);
        return S_OK;
    }

    if (unit == TextUnit_Line && ! IsCurrentThreadWindowThread())
    {
        size_t rangeStart = 0u;
        size_t rangeEnd   = 0u;
        size_t textLength = 0u;
        {
            const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
            const std::wstring text   = ResolveText();
            const TextRangeSpan range = ClampCurrentRange(text.size());
            rangeStart                = range.start;
            rangeEnd                  = range.end;
            textLength                = text.size();
        }

        size_t resultStart = rangeStart;
        size_t resultEnd   = rangeEnd;
        int moved          = 0;
        const HRESULT hr   = DispatchLineMovementToWindowThread(rangeStart, rangeEnd, count, resultStart, resultEnd, moved);
        if (SUCCEEDED(hr))
        {
            const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
            const std::wstring text         = ResolveText();
            const TextRangeSpan resultRange = ClampTextRangeSpan(resultStart, resultEnd, text.size());
            _boundsOverrideDip.reset();
            _rangeStart = resultRange.start;
            _rangeEnd   = resultRange.end;
            *outMoved   = moved;
        }
        Debug::Perf::Emit(
            L"dxui.uia.text_range_us", L"move-line-dispatch", Debug::Perf::ElapsedUs(startedAt), SUCCEEDED(hr) ? resultEnd - resultStart : 0u, textLength, hr);
        return hr;
    }

    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    const std::wstring text   = ResolveText();
    const TextRangeSpan range = ClampCurrentRange(text.size());
    if (unit == TextUnit_Line)
    {
        WindowHost* const host       = ResolveHost();
        const Control* const control = ResolveControl();
        const TextRangeSpanMoveResult moveResult =
            (host && control) ? TryMoveTextRangeSpanByVisualLine(*host, *control, text, range, count).value_or(MoveTextRangeSpanByLine(text, range, count))
                              : MoveTextRangeSpanByLine(text, range, count);
        _boundsOverrideDip.reset();
        _rangeStart = moveResult.range.start;
        _rangeEnd   = moveResult.range.end;
        *outMoved   = moveResult.moved;
        Debug::Perf::Emit(L"dxui.uia.text_range_us", L"move-line", Debug::Perf::ElapsedUs(startedAt), _rangeEnd - _rangeStart, text.size(), S_OK);
        return S_OK;
    }

    if (unit == TextUnit_Word)
    {
        if (range.start != range.end)
        {
            const TextRangeSpanMoveResult moveResult = MoveTextRangeSpanByWord(text, range, count);
            _boundsOverrideDip.reset();
            _rangeStart = moveResult.range.start;
            _rangeEnd   = moveResult.range.end;
            *outMoved   = moveResult.moved;
            Debug::Perf::Emit(L"dxui.uia.text_range_us", L"move-word", Debug::Perf::ElapsedUs(startedAt), _rangeEnd - _rangeStart, text.size(), S_OK);
            return S_OK;
        }

        const TextRangeUnitMoveResult moveResult = MoveTextRangePositionByUnit(text, range.start, unit, count);
        _boundsOverrideDip.reset();
        _rangeStart = moveResult.position;
        _rangeEnd   = moveResult.position;
        *outMoved   = moveResult.moved;
        Debug::Perf::Emit(L"dxui.uia.text_range_us", L"move-word", Debug::Perf::ElapsedUs(startedAt), 0u, text.size(), S_OK);
        return S_OK;
    }

    const int64_t rangeStart = static_cast<int64_t>(range.start);
    const int64_t rangeEnd   = static_cast<int64_t>(range.end);
    const int64_t textLength = static_cast<int64_t>(text.size());
    const int64_t minimum    = -rangeStart;
    const int64_t maximum    = textLength - rangeEnd;
    const int64_t requested  = static_cast<int64_t>(count);
    const int64_t moved      = (std::min)((std::max)(requested, minimum), maximum);
    _boundsOverrideDip.reset();
    _rangeStart = static_cast<size_t>(rangeStart + moved);
    _rangeEnd   = static_cast<size_t>(rangeEnd + moved);
    *outMoved   = static_cast<int>(moved);
    Debug::Perf::Emit(L"dxui.uia.text_range_us", L"move-character", Debug::Perf::ElapsedUs(startedAt), _rangeEnd - _rangeStart, text.size(), S_OK);
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::MoveEndpointByUnit(TextPatternRangeEndpoint endpoint, TextUnit unit, int count, int* outMoved) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    if (! outMoved)
    {
        return E_POINTER;
    }

    *outMoved = 0;
    if ((unit != TextUnit_Character && unit != TextUnit_Word && unit != TextUnit_Line) || count == 0)
    {
        const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
        const std::wstring text = ResolveText();
        Debug::Perf::Emit(L"dxui.uia.text_range_us", L"move-endpoint-unsupported", Debug::Perf::ElapsedUs(startedAt), 0u, text.size(), S_OK);
        return S_OK;
    }

    if (unit == TextUnit_Line && ! IsCurrentThreadWindowThread())
    {
        size_t rangeStart = 0u;
        size_t rangeEnd   = 0u;
        size_t textLength = 0u;
        {
            const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
            const std::wstring text   = ResolveText();
            const TextRangeSpan range = ClampCurrentRange(text.size());
            rangeStart                = range.start;
            rangeEnd                  = range.end;
            textLength                = text.size();
        }

        size_t resultStart = rangeStart;
        size_t resultEnd   = rangeEnd;
        int moved          = 0;
        const HRESULT hr   = DispatchEndpointLineMovementToWindowThread(rangeStart, rangeEnd, endpoint, count, resultStart, resultEnd, moved);
        if (SUCCEEDED(hr))
        {
            const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
            const std::wstring text         = ResolveText();
            const TextRangeSpan resultRange = ClampTextRangeSpan(resultStart, resultEnd, text.size());
            _boundsOverrideDip.reset();
            _rangeStart = resultRange.start;
            _rangeEnd   = resultRange.end;
            *outMoved   = moved;
        }
        Debug::Perf::Emit(L"dxui.uia.text_range_us",
                          L"move-endpoint-line-dispatch",
                          Debug::Perf::ElapsedUs(startedAt),
                          SUCCEEDED(hr) ? resultEnd - resultStart : 0u,
                          textLength,
                          hr);
        return hr;
    }

    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    const std::wstring text       = ResolveText();
    const TextRangeSpan range     = ClampCurrentRange(text.size());
    const size_t endpointPosition = GetTextRangeEndpointPosition(range, endpoint);
    TextRangeUnitMoveResult moveResult{};
    if (unit == TextUnit_Line)
    {
        WindowHost* const host       = ResolveHost();
        const Control* const control = ResolveControl();
        moveResult                   = (host && control) ? TryMoveTextRangePositionByVisualLine(*host, *control, text, endpointPosition, count)
                                                               .value_or(MoveTextRangePositionByLine(text, endpointPosition, count))
                                                         : MoveTextRangePositionByLine(text, endpointPosition, count);
    }
    else
    {
        moveResult = MoveTextRangePositionByUnit(text, endpointPosition, unit, count);
    }
    const size_t movedTo = moveResult.position;
    *outMoved            = moveResult.moved;
    _boundsOverrideDip.reset();
    if (endpoint == TextPatternRangeEndpoint_Start)
    {
        _rangeStart = movedTo;
        _rangeEnd   = range.end < _rangeStart ? _rangeStart : range.end;
    }
    else
    {
        _rangeEnd   = movedTo;
        _rangeStart = _rangeEnd < range.start ? _rangeEnd : range.start;
    }
    const wchar_t* detail = unit == TextUnit_Line ? L"move-endpoint-line" : unit == TextUnit_Word ? L"move-endpoint-word" : L"move-endpoint-character";
    Debug::Perf::Emit(L"dxui.uia.text_range_us", detail, Debug::Perf::ElapsedUs(startedAt), _rangeEnd - _rangeStart, text.size(), S_OK);
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::MoveEndpointByRange(TextPatternRangeEndpoint /*endpoint*/,
                                                            ITextRangeProvider* /*targetRange*/,
                                                            TextPatternRangeEndpoint /*targetEndpoint*/) noexcept
{
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::Select() noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.textRangeProvider = this;
        request.kind              = AccessibilityUiActionKind::Select;
        return DispatchActionToWindowThread(request);
    }

    return ExecuteSelectOnWindowThread();
}

HRESULT AccessibilityTextRangeProvider::ExecuteSelectOnWindowThread() noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());

    WindowHost* host = ResolveHost();
    Control* control = ResolveMutableControl();
    if (! host || ! control || ! SupportsTextPattern(control))
    {
        return UIA_E_ELEMENTNOTAVAILABLE;
    }

    const std::wstring text   = GetControlAccessibleTextRangeText(control);
    const TextRangeSpan range = ClampCurrentRange(text.size());
    if (auto* textField = dynamic_cast<TextField*>(control))
    {
        textField->SetSelectionRange(range.start, range.end);
        host->SyncTextInput(textField);
        host->Invalidate();
        Debug::Perf::Emit(L"dxui.uia.text_range_us", L"select", Debug::Perf::ElapsedUs(startedAt), range.end - range.start, text.size(), S_OK);
        return S_OK;
    }
    if (auto* comboBox = dynamic_cast<ComboBox*>(control); comboBox && comboBox->IsEditable())
    {
        comboBox->SetEditableSelectionRange(range.start, range.end);
        host->SyncTextInput(comboBox);
        host->Invalidate();
        Debug::Perf::Emit(L"dxui.uia.text_range_us", L"select", Debug::Perf::ElapsedUs(startedAt), range.end - range.start, text.size(), S_OK);
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityTextRangeProvider::ExecuteMoveByVisualLineOnWindowThread(
    size_t start, size_t end, int count, size_t& outStart, size_t& outEnd, int& outMoved) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());

    WindowHost* const host       = ResolveHost();
    const Control* const control = ResolveControl();
    const std::wstring text      = control ? (_textOverride ? _textOverride.value() : GetControlAccessibleTextRangeText(control)) : ResolveText();
    const TextRangeSpan range    = ClampTextRangeSpan(start, end, text.size());
    const TextRangeSpanMoveResult moveResult =
        (host && control) ? TryMoveTextRangeSpanByVisualLine(*host, *control, text, range, count).value_or(MoveTextRangeSpanByLine(text, range, count))
                          : MoveTextRangeSpanByLine(text, range, count);

    outStart = moveResult.range.start;
    outEnd   = moveResult.range.end;
    outMoved = moveResult.moved;
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::ExecuteMoveEndpointByVisualLineOnWindowThread(
    size_t start, size_t end, TextPatternRangeEndpoint endpoint, int count, size_t& outStart, size_t& outEnd, int& outMoved) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());

    WindowHost* const host        = ResolveHost();
    const Control* const control  = ResolveControl();
    const std::wstring text       = control ? (_textOverride ? _textOverride.value() : GetControlAccessibleTextRangeText(control)) : ResolveText();
    const TextRangeSpan range     = ClampTextRangeSpan(start, end, text.size());
    const size_t endpointPosition = GetTextRangeEndpointPosition(range, endpoint);

    const TextRangeUnitMoveResult moveResult = (host && control) ? TryMoveTextRangePositionByVisualLine(*host, *control, text, endpointPosition, count)
                                                                       .value_or(MoveTextRangePositionByLine(text, endpointPosition, count))
                                                                 : MoveTextRangePositionByLine(text, endpointPosition, count);
    outMoved                                 = moveResult.moved;
    const size_t movedTo                     = moveResult.position;
    if (endpoint == TextPatternRangeEndpoint_Start)
    {
        outStart = movedTo;
        outEnd   = range.end < outStart ? outStart : range.end;
    }
    else
    {
        outEnd   = movedTo;
        outStart = outEnd < range.start ? outEnd : range.start;
    }

    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::ExecuteResolveBoundsOnWindowThread(size_t start,
                                                                           size_t end,
                                                                           std::vector<D2D1_RECT_F>& outBoundsDip,
                                                                           float& outDipToPixelScale) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());

    outBoundsDip.clear();
    outDipToPixelScale = 1.0f;

    WindowHost* const host       = ResolveHost();
    const Control* const control = ResolveControl();
    if (! host || ! control)
    {
        return S_OK;
    }

    outDipToPixelScale        = static_cast<float>(host->GetDpi()) / static_cast<float>(USER_DEFAULT_SCREEN_DPI);
    const std::wstring text   = _textOverride ? _textOverride.value() : GetControlAccessibleTextRangeText(control);
    const TextRangeSpan range = ClampTextRangeSpan(start, end, text.size());
    if (range.start == range.end)
    {
        return S_OK;
    }

    outBoundsDip = TryResolveTextRangeCaretRects(*host, *control, text, range).value_or(std::vector<D2D1_RECT_F>{ResolveTextPatternViewportRect(control)});
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::AddToSelection() noexcept
{
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::RemoveFromSelection() noexcept
{
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::ScrollIntoView(BOOL /*alignToTop*/) noexcept
{
    return S_OK;
}

HRESULT AccessibilityTextRangeProvider::GetChildren(SAFEARRAY** outChildren) noexcept
{
    if (! outChildren)
    {
        return E_POINTER;
    }

    return SetProviderArray(outChildren, {});
}

bool AccessibilityTextRangeProvider::IsCurrentThreadWindowThread() const noexcept
{
    return _hwnd && GetWindowThreadProcessId(_hwnd, nullptr) == GetCurrentThreadId();
}

HRESULT DispatchAccessibilityUiActionToWindowThread(HWND hwnd, AccessibilityUiActionRequest& request) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return UIA_E_ELEMENTNOTAVAILABLE;
    }

    auto dispatch = std::shared_ptr<AccessibilityUiActionDispatch>(new (std::nothrow) AccessibilityUiActionDispatch());
    if (! dispatch)
    {
        return E_OUTOFMEMORY;
    }

    dispatch->request = std::move(request);
    if (dispatch->request.provider)
    {
        dispatch->providerKeepAlive = static_cast<IRawElementProviderSimple*>(dispatch->request.provider);
    }
    if (dispatch->request.textRangeProvider)
    {
        dispatch->textRangeProviderKeepAlive = static_cast<ITextRangeProvider*>(dispatch->request.textRangeProvider);
    }

    dispatch->completedEvent.reset(::CreateEventW(nullptr, TRUE, FALSE, nullptr));
    if (! dispatch->completedEvent)
    {
        const DWORD lastError = ::GetLastError();
        return lastError != 0u ? HRESULT_FROM_WIN32(lastError) : E_OUTOFMEMORY;
    }

    auto payload = std::unique_ptr<AccessibilityUiActionPayload>(new (std::nothrow) AccessibilityUiActionPayload());
    if (! payload)
    {
        return E_OUTOFMEMORY;
    }
    payload->dispatch = dispatch;

    ::SetLastError(ERROR_SUCCESS);
    if (! PostMessagePayload(hwnd, kWindowHostAccessibilityActionMessage, 0, std::move(payload)))
    {
        const DWORD lastError = ::GetLastError();
        return lastError != 0u ? HRESULT_FROM_WIN32(lastError) : static_cast<HRESULT>(UIA_E_ELEMENTNOTAVAILABLE);
    }

#if defined(ENABLE_TESTS)
    if (const HANDLE postedEvent = g_accessibilityUiActionPostedEvent.load(std::memory_order_acquire); postedEvent)
    {
        static_cast<void>(::SetEvent(postedEvent));
    }
#endif

    const DWORD waitResult = ::WaitForSingleObject(dispatch->completedEvent.get(), AccessibilityUiActionDispatchTimeoutMs());
    if (waitResult != WAIT_OBJECT_0)
    {
        if (waitResult == WAIT_TIMEOUT)
        {
            AccessibilityUiActionDispatch::State expected = AccessibilityUiActionDispatch::State::Pending;
            if (dispatch->state.compare_exchange_strong(
                    expected, AccessibilityUiActionDispatch::State::Abandoned, std::memory_order_acq_rel, std::memory_order_acquire))
            {
                return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
            }

            if (expected == AccessibilityUiActionDispatch::State::Taken && ::WaitForSingleObject(dispatch->completedEvent.get(), 0u) == WAIT_OBJECT_0)
            {
                request = std::move(dispatch->request);
                return request.result;
            }
            return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
        }

        const DWORD lastError = ::GetLastError();
        return lastError != 0u ? HRESULT_FROM_WIN32(lastError) : E_FAIL;
    }

    request = std::move(dispatch->request);
    return request.result;
}

HRESULT AccessibilityTextRangeProvider::DispatchActionToWindowThread(AccessibilityUiActionRequest& request) const noexcept
{
    return DispatchAccessibilityUiActionToWindowThread(_hwnd, request);
}

HRESULT AccessibilityTextRangeProvider::DispatchLineMovementToWindowThread(
    size_t start, size_t end, int count, size_t& outStart, size_t& outEnd, int& outMoved) noexcept
{
    AccessibilityUiActionRequest request{};
    request.textRangeProvider    = this;
    request.kind                 = AccessibilityUiActionKind::MoveTextRangeByVisualLine;
    request.textRangeStart       = start;
    request.textRangeEnd         = end;
    request.textRangeMoveCount   = count;
    request.textRangeMoved       = 0;
    request.textRangeResultStart = start;
    request.textRangeResultEnd   = end;

    const HRESULT hr = DispatchActionToWindowThread(request);
    outStart         = request.textRangeResultStart;
    outEnd           = request.textRangeResultEnd;
    outMoved         = request.textRangeMoved;
    return hr;
}

HRESULT AccessibilityTextRangeProvider::DispatchEndpointLineMovementToWindowThread(
    size_t start, size_t end, TextPatternRangeEndpoint endpoint, int count, size_t& outStart, size_t& outEnd, int& outMoved) noexcept
{
    AccessibilityUiActionRequest request{};
    request.textRangeProvider    = this;
    request.kind                 = AccessibilityUiActionKind::MoveTextRangeEndpointByVisualLine;
    request.textRangeStart       = start;
    request.textRangeEnd         = end;
    request.textRangeEndpoint    = endpoint;
    request.textRangeMoveCount   = count;
    request.textRangeMoved       = 0;
    request.textRangeResultStart = start;
    request.textRangeResultEnd   = end;

    const HRESULT hr = DispatchActionToWindowThread(request);
    outStart         = request.textRangeResultStart;
    outEnd           = request.textRangeResultEnd;
    outMoved         = request.textRangeMoved;
    return hr;
}

HRESULT AccessibilityTextRangeProvider::DispatchBoundingRectanglesToWindowThread(size_t start,
                                                                                 size_t end,
                                                                                 std::vector<D2D1_RECT_F>& outBoundsDip,
                                                                                 float& outDipToPixelScale) noexcept
{
    AccessibilityUiActionRequest request{};
    request.textRangeProvider        = this;
    request.kind                     = AccessibilityUiActionKind::ResolveTextRangeBounds;
    request.textRangeStart           = start;
    request.textRangeEnd             = end;
    request.textRangeDipToPixelScale = outDipToPixelScale;

    const HRESULT hr   = DispatchActionToWindowThread(request);
    outBoundsDip       = std::move(request.textRangeBoundsDip);
    outDipToPixelScale = request.textRangeDipToPixelScale;
    return hr;
}

WindowHost* AccessibilityTextRangeProvider::ResolveHost() const noexcept
{
    return (_target && _target->hwnd == _hwnd) ? _target->ResolveHost() : nullptr;
}

const Control* AccessibilityTextRangeProvider::ResolveControl() const noexcept
{
    WindowHost* host = ResolveHost();
    if (! host || ! IsControlPathVisible(host->GetRoot(), _path))
    {
        return nullptr;
    }

    return ResolveControlAtPath(host->GetRoot(), _path);
}

Control* AccessibilityTextRangeProvider::ResolveMutableControl() const noexcept
{
    WindowHost* host = ResolveHost();
    if (! host || ! IsControlPathVisible(host->GetRoot(), _path))
    {
        return nullptr;
    }

    return ResolveControlAtPath(host->GetRoot(), _path);
}

std::wstring AccessibilityTextRangeProvider::ResolveText() const
{
    if (_textOverride)
    {
        return _textOverride.value();
    }

    const AccessibilityControlNavigationSnapshot* record =
        (_snapshot && _snapshot->alive && _snapshot->hasRetainedRoot) ? FindControlNavigationRecord(*_snapshot, _path) : nullptr;
    if (! record || ! record->controlSupportsText)
    {
        return {};
    }

    return record->controlAccessibleText;
}

TextRangeSpan AccessibilityTextRangeProvider::ClampCurrentRange(size_t textLength) const noexcept
{
    return ClampTextRangeSpan(_rangeStart, _rangeEnd, textLength);
}

ITextRangeProvider* AccessibilityTextRangeProvider::CreateRange(size_t start, size_t end) const noexcept
{
    WindowHostAccessibilityTarget* target = _target;
    if (target)
    {
        static_cast<void>(target->AddRef());
    }

    AccessibilityTextRangeProvider* provider = nullptr;
    if (_textOverride)
    {
        provider = new (std::nothrow) AccessibilityTextRangeProvider(target, _hwnd, _path, start, end, _textOverride.value());
    }
    else if (_boundsOverrideDip && start == _rangeStart && end == _rangeEnd)
    {
        provider = new (std::nothrow) AccessibilityTextRangeProvider(target, _hwnd, _path, start, end, _boundsOverrideDip.value());
    }
    else
    {
        provider = new (std::nothrow) AccessibilityTextRangeProvider(target, _hwnd, _path, start, end);
    }

    if (! provider && target)
    {
        static_cast<void>(target->Release());
    }
    return provider ? static_cast<ITextRangeProvider*>(provider) : nullptr;
}

AccessibilityPatternQueryResult AccessibilityProvider::QueryPattern(AccessibilityPatternKind patternKind) noexcept
{
    const auto makeResult = []<typename TInterface>(TInterface* provider) noexcept -> AccessibilityPatternQueryResult
    { return AccessibilityPatternQueryResult{.queryInterface = static_cast<void*>(provider), .patternProvider = static_cast<IUnknown*>(provider)}; };

    if (_kind == AccessibilityFragmentKind::TreeItem && patternKind == AccessibilityPatternKind::SelectionItem)
    {
        const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
        const AccessibilityControlNavigationSnapshot* record =
            (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? FindControlNavigationRecord(*snapshot, _path) : nullptr;
        return (record && SnapshotContainsTreeItem(*record, _treeVisibleIndex)) ? makeResult(static_cast<ISelectionItemProvider*>(this))
                                                                                : AccessibilityPatternQueryResult{};
    }

    if (_kind == AccessibilityFragmentKind::TreeItem && patternKind == AccessibilityPatternKind::ExpandCollapse)
    {
        const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
        const AccessibilityControlNavigationSnapshot* record =
            (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? FindControlNavigationRecord(*snapshot, _path) : nullptr;
        const AccessibilityTreeItemSnapshotRecord* item = record ? FindSnapshotTreeItemRecord(*record, _treeVisibleIndex) : nullptr;
        return (item && item->hasChildren) ? makeResult(static_cast<IExpandCollapseProvider*>(this)) : AccessibilityPatternQueryResult{};
    }

    if (_kind == AccessibilityFragmentKind::GridRow)
    {
        if (patternKind != AccessibilityPatternKind::SelectionItem)
        {
            return {};
        }

        const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
        const AccessibilityControlNavigationSnapshot* record =
            (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? FindControlNavigationRecord(*snapshot, _path) : nullptr;
        return (record && (SnapshotContainsGridRow(*record, _gridRowId) || SnapshotGridRowIsSelected(*record, _gridRowId)))
                   ? makeResult(static_cast<ISelectionItemProvider*>(this))
                   : AccessibilityPatternQueryResult{};
    }

    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        const std::shared_ptr<const AccessibilitySnapshot> snapshot   = CaptureAccessibilitySnapshot(_target, _hwnd);
        const std::optional<AccessibilityGridCellSnapshotRecord> cell = (snapshot && snapshot->alive && snapshot->hasRetainedRoot)
                                                                            ? FindSnapshotGridCellRecord(*snapshot, _path, _gridRowId, _gridColumnIndex)
                                                                            : std::nullopt;
        if (! cell)
        {
            return {};
        }

        switch (patternKind)
        {
            case AccessibilityPatternKind::Invoke:
            case AccessibilityPatternKind::Text:
            case AccessibilityPatternKind::TextEdit:
            case AccessibilityPatternKind::Selection:
            case AccessibilityPatternKind::Table:
            case AccessibilityPatternKind::SelectionItem:
            case AccessibilityPatternKind::ExpandCollapse: return {};
            case AccessibilityPatternKind::Toggle:
                return SnapshotGridCellSupportsTogglePattern(cell.value()) ? makeResult(static_cast<IToggleProvider*>(this))
                                                                           : AccessibilityPatternQueryResult{};
            case AccessibilityPatternKind::Value:
                return SnapshotGridCellSupportsValuePattern(cell.value()) ? makeResult(static_cast<IValueProvider*>(this)) : AccessibilityPatternQueryResult{};
            case AccessibilityPatternKind::RangeValue:
                return SnapshotGridCellSupportsRangeValuePattern(cell.value()) ? makeResult(static_cast<IRangeValueProvider*>(this))
                                                                               : AccessibilityPatternQueryResult{};
            case AccessibilityPatternKind::GridItem: return makeResult(static_cast<IGridItemProvider*>(this));
            case AccessibilityPatternKind::TableItem:
                return (cell->controlRecord && cell->controlRecord->gridColumnCount > 0u) ? makeResult(static_cast<ITableItemProvider*>(this))
                                                                                          : AccessibilityPatternQueryResult{};
        }
        return {};
    }

    if (_kind == AccessibilityFragmentKind::TextFieldPasswordRevealButton)
    {
        if (patternKind != AccessibilityPatternKind::Invoke)
        {
            return {};
        }

        const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
        const AccessibilityControlNavigationSnapshot* record =
            (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? FindControlNavigationRecord(*snapshot, _path) : nullptr;
        return (record && record->hasPasswordRevealButton) ? makeResult(static_cast<IInvokeProvider*>(this)) : AccessibilityPatternQueryResult{};
    }

    if (_kind != AccessibilityFragmentKind::Control && _kind != AccessibilityFragmentKind::Root)
    {
        return {};
    }

    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (! record)
    {
        return {};
    }

    switch (patternKind)
    {
        case AccessibilityPatternKind::Invoke:
            return record->controlSupportsInvoke ? makeResult(static_cast<IInvokeProvider*>(this)) : AccessibilityPatternQueryResult{};
        case AccessibilityPatternKind::Toggle:
            return record->controlSupportsToggle ? makeResult(static_cast<IToggleProvider*>(this)) : AccessibilityPatternQueryResult{};
        case AccessibilityPatternKind::Value:
            return record->controlSupportsValue ? makeResult(static_cast<IValueProvider*>(this)) : AccessibilityPatternQueryResult{};
        case AccessibilityPatternKind::Text:
            return record->controlSupportsText ? makeResult(static_cast<ITextProvider*>(static_cast<ITextEditProvider*>(this)))
                                               : AccessibilityPatternQueryResult{};
        case AccessibilityPatternKind::TextEdit:
            return record->controlSupportsText ? makeResult(static_cast<ITextEditProvider*>(this)) : AccessibilityPatternQueryResult{};
        case AccessibilityPatternKind::RangeValue:
            return record->controlSupportsRangeValue ? makeResult(static_cast<IRangeValueProvider*>(this)) : AccessibilityPatternQueryResult{};
        case AccessibilityPatternKind::Selection:
            return record->controlSupportsSelection ? makeResult(static_cast<ISelectionProvider*>(this)) : AccessibilityPatternQueryResult{};
        case AccessibilityPatternKind::Table:
            return record->controlSupportsTable ? makeResult(static_cast<ITableProvider*>(this)) : AccessibilityPatternQueryResult{};
        case AccessibilityPatternKind::SelectionItem:
        case AccessibilityPatternKind::ExpandCollapse:
        case AccessibilityPatternKind::GridItem:
        case AccessibilityPatternKind::TableItem: return {};
    }
    return {};
}

HRESULT AccessibilityProvider::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! ppvObject)
    {
        return E_POINTER;
    }

    *ppvObject = nullptr;
    if (riid == __uuidof(IUnknown) || riid == __uuidof(IRawElementProviderSimple))
    {
        *ppvObject = static_cast<IRawElementProviderSimple*>(this);
        AddRef();
        return S_OK;
    }
    if (riid == __uuidof(IRawElementProviderFragment))
    {
        *ppvObject = static_cast<IRawElementProviderFragment*>(this);
        AddRef();
        return S_OK;
    }
    if (_kind == AccessibilityFragmentKind::Root && riid == __uuidof(IRawElementProviderFragmentRoot))
    {
        *ppvObject = static_cast<IRawElementProviderFragmentRoot*>(this);
        AddRef();
        return S_OK;
    }

    const std::optional<AccessibilityPatternKind> patternKind = PatternKindFromInterfaceId(riid);
    if (! patternKind)
    {
        return E_NOINTERFACE;
    }

    const AccessibilityPatternQueryResult pattern = QueryPattern(patternKind.value());
    if (! pattern.queryInterface)
    {
        return E_NOINTERFACE;
    }

    *ppvObject = pattern.queryInterface;
    AddRef();
    return S_OK;
}

ULONG AccessibilityProvider::AddRef() noexcept
{
    return _referenceCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
}

ULONG AccessibilityProvider::Release() noexcept
{
    const ULONG remaining = _referenceCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
    if (remaining == 0u)
    {
        if (_target)
        {
            static_cast<void>(_target->Release());
            _target = nullptr;
        }
        delete this;
    }
    return remaining;
}

HRESULT AccessibilityProvider::get_ProviderOptions(ProviderOptions* outOptions) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outOptions)
    {
        return E_POINTER;
    }

    *outOptions = ProviderOptions_ServerSideProvider;
    return S_OK;
}

HRESULT AccessibilityProvider::GetPatternProvider(PATTERNID patternId, IUnknown** outProvider) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outProvider)
    {
        return E_POINTER;
    }

    *outProvider                                              = nullptr;
    const std::optional<AccessibilityPatternKind> patternKind = PatternKindFromPatternId(patternId);
    if (! patternKind)
    {
        return S_OK;
    }

    const AccessibilityPatternQueryResult pattern = QueryPattern(patternKind.value());
    if (pattern.patternProvider)
    {
        *outProvider = pattern.patternProvider;
        AddRef();
    }

    return S_OK;
}

HRESULT AccessibilityProvider::GetPropertyValue(PROPERTYID propertyId, VARIANT* outValue) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outValue)
    {
        return E_POINTER;
    }

    VariantInit(outValue);
    if (_kind == AccessibilityFragmentKind::TextFieldPasswordRevealButton)
    {
        const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
        const AccessibilityControlNavigationSnapshot* record =
            (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? FindControlNavigationRecord(*snapshot, _path) : nullptr;
        const bool buttonVisible = record && record->hasPasswordRevealButton;

        switch (propertyId)
        {
            case UIA_ControlTypePropertyId: *outValue = VariantFromInt(UIA_ButtonControlTypeId); return S_OK;
            case UIA_NamePropertyId:
            {
                const std::wstring_view name = buttonVisible ? std::wstring_view(record->passwordRevealButtonAccessibleName) : std::wstring_view{};
                return SetVariantFromString(outValue, name);
            }
            case UIA_IsControlElementPropertyId:
            case UIA_IsContentElementPropertyId: *outValue = VariantFromBool(buttonVisible); return S_OK;
            case UIA_IsEnabledPropertyId: *outValue = VariantFromBool(buttonVisible && record->passwordRevealButtonEnabled); return S_OK;
            case UIA_IsKeyboardFocusablePropertyId: *outValue = VariantFromBool(buttonVisible); return S_OK;
            case UIA_HasKeyboardFocusPropertyId: *outValue = VariantFromBool(false); return S_OK;
            case UIA_IsOffscreenPropertyId: *outValue = VariantFromBool(! buttonVisible); return S_OK;
            default: return S_OK;
        }
    }

    if (propertyId == UIA_SelectionItemIsSelectedPropertyId && (_kind == AccessibilityFragmentKind::TreeItem || _kind == AccessibilityFragmentKind::GridRow))
    {
        const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
        const AccessibilityControlNavigationSnapshot* record =
            (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? FindControlNavigationRecord(*snapshot, _path) : nullptr;
        if (_kind == AccessibilityFragmentKind::TreeItem)
        {
            if (record && SnapshotContainsTreeItem(*record, _treeVisibleIndex))
            {
                *outValue = VariantFromBool(SnapshotTreeItemIsSelected(*record, _treeVisibleIndex));
            }
            return S_OK;
        }

        if (record && (SnapshotContainsGridRow(*record, _gridRowId) || SnapshotGridRowIsSelected(*record, _gridRowId)))
        {
            *outValue = VariantFromBool(SnapshotGridRowIsSelected(*record, _gridRowId));
        }
        return S_OK;
    }
    if (_kind == AccessibilityFragmentKind::GridRow)
    {
        const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
        const AccessibilityControlNavigationSnapshot* record =
            (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? FindControlNavigationRecord(*snapshot, _path) : nullptr;
        const AccessibilityGridRowSnapshotRecord* row = record ? FindSnapshotGridRowRecord(*record, _gridRowId) : nullptr;
        if (! record || ! row)
        {
            return S_OK;
        }

        switch (propertyId)
        {
            case UIA_ControlTypePropertyId: *outValue = VariantFromInt(UIA_DataItemControlTypeId); return S_OK;
            case UIA_NamePropertyId: return SetVariantFromString(outValue, row->gridRowAccessibleName);
            case UIA_IsControlElementPropertyId:
            case UIA_IsContentElementPropertyId:
            case UIA_IsKeyboardFocusablePropertyId: *outValue = VariantFromBool(true); return S_OK;
            case UIA_IsEnabledPropertyId: *outValue = VariantFromBool(record->gridIsEnabled); return S_OK;
            case UIA_HasKeyboardFocusPropertyId:
                *outValue = VariantFromBool(record->gridHasFocus && SnapshotGridRowIsSelected(*record, _gridRowId));
                return S_OK;
            case UIA_IsOffscreenPropertyId: *outValue = VariantFromBool(row->gridRowOffscreen); return S_OK;
            case UIA_SelectionItemIsSelectedPropertyId: *outValue = VariantFromBool(SnapshotGridRowIsSelected(*record, _gridRowId)); return S_OK;
            default: return S_OK;
        }
    }
    if (_kind == AccessibilityFragmentKind::TreeItem)
    {
        const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
        const AccessibilityControlNavigationSnapshot* record =
            (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? FindControlNavigationRecord(*snapshot, _path) : nullptr;
        const AccessibilityTreeItemSnapshotRecord* item = record ? FindSnapshotTreeItemRecord(*record, _treeVisibleIndex) : nullptr;
        if (! record || ! item)
        {
            return S_OK;
        }

        switch (propertyId)
        {
            case UIA_ControlTypePropertyId: *outValue = VariantFromInt(UIA_TreeItemControlTypeId); return S_OK;
            case UIA_NamePropertyId: return SetVariantFromString(outValue, item->text);
            case UIA_IsControlElementPropertyId:
            case UIA_IsContentElementPropertyId:
            case UIA_IsKeyboardFocusablePropertyId: *outValue = VariantFromBool(true); return S_OK;
            case UIA_IsEnabledPropertyId: *outValue = VariantFromBool(record->treeIsEnabled); return S_OK;
            case UIA_HasKeyboardFocusPropertyId:
                *outValue = VariantFromBool(record->treeHasFocus && SnapshotTreeItemIsSelected(*record, _treeVisibleIndex));
                return S_OK;
            case UIA_IsOffscreenPropertyId: *outValue = VariantFromBool(false); return S_OK;
            case UIA_LevelPropertyId:
                *outValue = VariantFromInt(static_cast<LONG>((std::min)(item->depth + 1u, static_cast<size_t>((std::numeric_limits<LONG>::max)()))));
                return S_OK;
            case UIA_ExpandCollapseExpandCollapseStatePropertyId:
                if (item->hasChildren)
                {
                    *outValue = VariantFromInt(item->expanded ? ExpandCollapseState_Expanded : ExpandCollapseState_Collapsed);
                }
                return S_OK;
            default: return S_OK;
        }
    }
    if (_kind == AccessibilityFragmentKind::Control && (propertyId == UIA_GridRowCountPropertyId || propertyId == UIA_GridColumnCountPropertyId))
    {
        const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
        const AccessibilityControlNavigationSnapshot* record =
            (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? FindControlNavigationRecord(*snapshot, _path) : nullptr;
        if (! record || ! record->isGrid)
        {
            return S_OK;
        }

        const size_t count = propertyId == UIA_GridRowCountPropertyId ? record->gridRowCount : record->gridColumnCount;
        *outValue          = VariantFromInt(static_cast<LONG>((std::min)(count, static_cast<size_t>((std::numeric_limits<LONG>::max)()))));
        return S_OK;
    }
    if (_kind == AccessibilityFragmentKind::GridHeader)
    {
        const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
        const AccessibilityControlNavigationSnapshot* record =
            (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? FindControlNavigationRecord(*snapshot, _path) : nullptr;
        const AccessibilityGridHeaderSnapshotRecord* header = record ? FindSnapshotGridHeaderRecord(*record, _gridColumnIndex) : nullptr;
        if (! record || ! header)
        {
            return S_OK;
        }

        switch (propertyId)
        {
            case UIA_ControlTypePropertyId: *outValue = VariantFromInt(UIA_HeaderItemControlTypeId); return S_OK;
            case UIA_NamePropertyId: return SetVariantFromString(outValue, header->gridHeaderName);
            case UIA_IsControlElementPropertyId:
            case UIA_IsContentElementPropertyId: *outValue = VariantFromBool(true); return S_OK;
            case UIA_IsEnabledPropertyId: *outValue = VariantFromBool(record->gridIsEnabled); return S_OK;
            case UIA_IsKeyboardFocusablePropertyId:
            case UIA_HasKeyboardFocusPropertyId: *outValue = VariantFromBool(false); return S_OK;
            case UIA_IsOffscreenPropertyId: *outValue = VariantFromBool(false); return S_OK;
            default: return S_OK;
        }
    }
    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        const std::shared_ptr<const AccessibilitySnapshot> snapshot   = CaptureAccessibilitySnapshot(_target, _hwnd);
        const std::optional<AccessibilityGridCellSnapshotRecord> cell = (snapshot && snapshot->alive && snapshot->hasRetainedRoot)
                                                                            ? FindSnapshotGridCellRecord(*snapshot, _path, _gridRowId, _gridColumnIndex)
                                                                            : std::nullopt;
        const AccessibilityGridCellStateSnapshotRecord* cellRecord    = cell ? cell->cellRecord : nullptr;
        if (! cell || ! cellRecord)
        {
            return S_OK;
        }

        switch (propertyId)
        {
            case UIA_ControlTypePropertyId: *outValue = VariantFromInt(cellRecord->gridCellControlTypeId); return S_OK;
            case UIA_NamePropertyId: return SetVariantFromString(outValue, cellRecord->gridCellAccessibleText);
            case UIA_HelpTextPropertyId:
                if (! cellRecord->gridCellHelpText.empty())
                {
                    return SetVariantFromString(outValue, cellRecord->gridCellHelpText);
                }
                return S_OK;
            case UIA_IsControlElementPropertyId:
            case UIA_IsContentElementPropertyId: *outValue = VariantFromBool(true); return S_OK;
            case UIA_IsEnabledPropertyId:
                *outValue = VariantFromBool(cell->controlRecord && cell->controlRecord->gridIsEnabled && cellRecord->gridCellEnabled);
                return S_OK;
            case UIA_IsKeyboardFocusablePropertyId:
            case UIA_HasKeyboardFocusPropertyId: *outValue = VariantFromBool(false); return S_OK;
            case UIA_IsOffscreenPropertyId: *outValue = VariantFromBool(cellRecord->gridCellOffscreen); return S_OK;
            case UIA_ToggleToggleStatePropertyId:
                if (SnapshotGridCellSupportsTogglePattern(cell.value()))
                {
                    *outValue = VariantFromInt(cellRecord->gridCellChecked ? ToggleState_On : ToggleState_Off);
                }
                return S_OK;
            case UIA_ValueValuePropertyId:
                if (SnapshotGridCellSupportsValuePattern(cell.value()))
                {
                    return SetVariantFromString(outValue, cellRecord->gridCellAccessibleText);
                }
                return S_OK;
            case UIA_ValueIsReadOnlyPropertyId:
                if (SnapshotGridCellSupportsValuePattern(cell.value()) || SnapshotGridCellSupportsRangeValuePattern(cell.value()))
                {
                    *outValue = VariantFromBool(true);
                }
                return S_OK;
            case UIA_RangeValueValuePropertyId:
                if (SnapshotGridCellSupportsRangeValuePattern(cell.value()))
                {
                    *outValue = VariantFromDouble(cellRecord->gridCellRangeValue);
                }
                return S_OK;
            case UIA_RangeValueMinimumPropertyId:
                if (SnapshotGridCellSupportsRangeValuePattern(cell.value()))
                {
                    *outValue = VariantFromDouble(0.0);
                }
                return S_OK;
            case UIA_RangeValueMaximumPropertyId:
                if (SnapshotGridCellSupportsRangeValuePattern(cell.value()))
                {
                    *outValue = VariantFromDouble(1.0);
                }
                return S_OK;
            case UIA_RangeValueLargeChangePropertyId:
            case UIA_RangeValueSmallChangePropertyId:
                if (SnapshotGridCellSupportsRangeValuePattern(cell.value()))
                {
                    *outValue = VariantFromDouble(0.0);
                }
                return S_OK;
            default: return S_OK;
        }
    }

    if (_kind != AccessibilityFragmentKind::Control && _kind != AccessibilityFragmentKind::Root)
    {
        return S_OK;
    }

    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (! record)
    {
        if (_kind == AccessibilityFragmentKind::Root)
        {
            switch (propertyId)
            {
                case UIA_ControlTypePropertyId: *outValue = VariantFromInt(UIA_PaneControlTypeId); return S_OK;
                case UIA_IsControlElementPropertyId:
                case UIA_IsContentElementPropertyId:
                case UIA_IsEnabledPropertyId:
                case UIA_IsKeyboardFocusablePropertyId: *outValue = VariantFromBool(true); return S_OK;
                case UIA_HasKeyboardFocusPropertyId: *outValue = VariantFromBool(false); return S_OK;
                case UIA_NamePropertyId:
                    if (snapshot)
                    {
                        return SetVariantFromString(outValue, snapshot->windowName);
                    }
                    return S_OK;
                default: return S_OK;
            }
        }
        return S_OK;
    }

    switch (propertyId)
    {
        case UIA_ControlTypePropertyId: *outValue = VariantFromInt(record->controlTypeId); return S_OK;
        case UIA_NamePropertyId: return SetVariantFromString(outValue, record->controlAccessibleName);
        case UIA_HelpTextPropertyId: return SetVariantFromString(outValue, record->controlAccessibleHelpText);
        case UIA_IsControlElementPropertyId:
        case UIA_IsContentElementPropertyId: *outValue = VariantFromBool(record->controlVisible); return S_OK;
        case UIA_IsEnabledPropertyId: *outValue = VariantFromBool(record->controlVisible && record->controlEnabled); return S_OK;
        case UIA_IsKeyboardFocusablePropertyId: *outValue = VariantFromBool(record->controlVisible && record->controlFocusable); return S_OK;
        case UIA_HasKeyboardFocusPropertyId: *outValue = VariantFromBool(record->controlVisible && record->controlHasFocus); return S_OK;
        case UIA_IsOffscreenPropertyId: *outValue = VariantFromBool(! record->controlVisible); return S_OK;
        case UIA_IsPasswordPropertyId: *outValue = VariantFromBool(record->controlVisible && record->controlIsPassword); return S_OK;
        case UIA_ValueValuePropertyId:
            if (record->controlVisible && record->controlSupportsValue)
            {
                return SetVariantFromString(outValue, record->controlAccessibleValue);
            }
            return S_OK;
        case UIA_ValueIsReadOnlyPropertyId:
            if (record->controlVisible && (record->controlSupportsValue || record->controlSupportsRangeValue))
            {
                *outValue = VariantFromBool(record->controlValueReadOnly);
            }
            return S_OK;
        case UIA_RangeValueValuePropertyId:
            if (record->controlVisible && record->controlSupportsRangeValue)
            {
                *outValue = VariantFromDouble(record->controlRangeValue);
            }
            return S_OK;
        case UIA_RangeValueMinimumPropertyId:
            if (record->controlVisible && record->controlSupportsRangeValue)
            {
                *outValue = VariantFromDouble(record->controlRangeMinimum);
            }
            return S_OK;
        case UIA_RangeValueMaximumPropertyId:
            if (record->controlVisible && record->controlSupportsRangeValue)
            {
                *outValue = VariantFromDouble(record->controlRangeMaximum);
            }
            return S_OK;
        case UIA_RangeValueSmallChangePropertyId:
            if (record->controlVisible && record->controlSupportsRangeValue)
            {
                *outValue = VariantFromDouble(record->controlRangeSmallChange);
            }
            return S_OK;
        case UIA_RangeValueLargeChangePropertyId:
            if (record->controlVisible && record->controlSupportsRangeValue)
            {
                *outValue = VariantFromDouble(record->controlRangeLargeChange);
            }
            return S_OK;
        default: return S_OK;
    }
}

HRESULT AccessibilityProvider::get_HostRawElementProvider(IRawElementProviderSimple** outProvider) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outProvider)
    {
        return E_POINTER;
    }

    *outProvider = nullptr;
    if (_kind != AccessibilityFragmentKind::Root)
    {
        return S_OK;
    }
    return UiaHostProviderFromHwnd(_hwnd, outProvider);
}

HRESULT AccessibilityProvider::Navigate(NavigateDirection direction, IRawElementProviderFragment** outProvider) noexcept
{
    if (! outProvider)
    {
        return E_POINTER;
    }

    *outProvider                                                = nullptr;
    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    if (! snapshot || ! snapshot->alive || ! snapshot->hasRetainedRoot)
    {
        return S_OK;
    }

    const std::optional<AccessibilityNavigationTarget> navigationTarget =
        ResolveSnapshotNavigationTarget(*snapshot, _kind, _path, _treeVisibleIndex, _gridRowId, _gridColumnIndex, direction);
    if (! navigationTarget)
    {
        return S_OK;
    }

    *outProvider = CreateProviderFromNavigationTarget(navigationTarget.value());
    return S_OK;
}

HRESULT AccessibilityProvider::GetRuntimeId(SAFEARRAY** outRuntimeId) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (_kind == AccessibilityFragmentKind::Root)
    {
        return SetRuntimeId(outRuntimeId, _hwnd, nullptr);
    }
    if (_kind == AccessibilityFragmentKind::TreeItem)
    {
        return SetTreeItemRuntimeId(outRuntimeId, _path, _treeVisibleIndex);
    }
    if (_kind == AccessibilityFragmentKind::GridHeader)
    {
        return SetGridHeaderRuntimeId(outRuntimeId, _path, _gridColumnIndex);
    }
    if (_kind == AccessibilityFragmentKind::GridRow)
    {
        return SetGridRowRuntimeId(outRuntimeId, _path, _gridRowId);
    }
    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        return SetGridCellRuntimeId(outRuntimeId, _path, _gridRowId, _gridColumnIndex);
    }
    if (_kind == AccessibilityFragmentKind::TextFieldPasswordRevealButton)
    {
        return SetTextFieldPasswordRevealButtonRuntimeId(outRuntimeId, _path);
    }
    return SetRuntimeId(outRuntimeId, _hwnd, &_path);
}

HRESULT AccessibilityProvider::get_BoundingRectangle(UiaRect* outRect) noexcept
{
    if (! outRect)
    {
        return E_POINTER;
    }

    *outRect                                                    = UiaRect{};
    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    if (! snapshot || ! snapshot->alive)
    {
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::Root)
    {
        if (SnapshotHasCollapsedSemanticRoot(*snapshot) && snapshot->hwnd && snapshot->pixelsToDipScale > 0.0f)
        {
            const std::optional<D2D1_RECT_F> boundsDip =
                FindSnapshotFragmentBounds(*snapshot, AccessibilityFragmentKind::Control, snapshot->semanticControlOrder.front(), 0u, 0u, 0u);
            if (boundsDip)
            {
                const float dipToPixelsScale = 1.0f / snapshot->pixelsToDipScale;
                POINT topLeft{static_cast<LONG>(std::lround(boundsDip->left * dipToPixelsScale)),
                              static_cast<LONG>(std::lround(boundsDip->top * dipToPixelsScale))};
                POINT bottomRight{static_cast<LONG>(std::lround(boundsDip->right * dipToPixelsScale)),
                                  static_cast<LONG>(std::lround(boundsDip->bottom * dipToPixelsScale))};
                if (ClientToScreen(snapshot->hwnd, &topLeft) != FALSE && ClientToScreen(snapshot->hwnd, &bottomRight) != FALSE)
                {
                    outRect->left   = static_cast<double>(topLeft.x);
                    outRect->top    = static_cast<double>(topLeft.y);
                    outRect->width  = static_cast<double>(bottomRight.x - topLeft.x);
                    outRect->height = static_cast<double>(bottomRight.y - topLeft.y);
                    return S_OK;
                }
            }
        }

        RECT windowRect{};
        if (! snapshot->hwnd || GetWindowRect(snapshot->hwnd, &windowRect) == FALSE)
        {
            return S_OK;
        }

        outRect->left   = static_cast<double>(windowRect.left);
        outRect->top    = static_cast<double>(windowRect.top);
        outRect->width  = static_cast<double>(windowRect.right - windowRect.left);
        outRect->height = static_cast<double>(windowRect.bottom - windowRect.top);
        return S_OK;
    }

    if (! snapshot->hasRetainedRoot || ! snapshot->hwnd || snapshot->pixelsToDipScale <= 0.0f)
    {
        return S_OK;
    }

    const std::optional<D2D1_RECT_F> boundsDip = FindSnapshotFragmentBounds(*snapshot, _kind, _path, _treeVisibleIndex, _gridRowId, _gridColumnIndex);
    if (! boundsDip)
    {
        return S_OK;
    }

    const float dipToPixelsScale = 1.0f / snapshot->pixelsToDipScale;
    POINT topLeft{static_cast<LONG>(std::lround(boundsDip->left * dipToPixelsScale)), static_cast<LONG>(std::lround(boundsDip->top * dipToPixelsScale))};
    POINT bottomRight{static_cast<LONG>(std::lround(boundsDip->right * dipToPixelsScale)),
                      static_cast<LONG>(std::lround(boundsDip->bottom * dipToPixelsScale))};
    if (ClientToScreen(snapshot->hwnd, &topLeft) == FALSE || ClientToScreen(snapshot->hwnd, &bottomRight) == FALSE)
    {
        return S_OK;
    }

    outRect->left   = static_cast<double>(topLeft.x);
    outRect->top    = static_cast<double>(topLeft.y);
    outRect->width  = static_cast<double>(bottomRight.x - topLeft.x);
    outRect->height = static_cast<double>(bottomRight.y - topLeft.y);
    return S_OK;
}

HRESULT AccessibilityProvider::GetEmbeddedFragmentRoots(SAFEARRAY** outRoots) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outRoots)
    {
        return E_POINTER;
    }

    *outRoots = nullptr;
    return S_OK;
}

HRESULT AccessibilityProvider::SetFocus() noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.provider = this;
        request.kind     = AccessibilityUiActionKind::SetFocus;
        return DispatchActionToWindowThread(request);
    }

    return ExecuteSetFocusOnWindowThread();
}

HRESULT AccessibilityProvider::get_FragmentRoot(IRawElementProviderFragmentRoot** outRoot) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outRoot)
    {
        return E_POINTER;
    }

    *outRoot = CreateRootProvider();
    return S_OK;
}

HRESULT AccessibilityProvider::ElementProviderFromPoint(double x, double y, IRawElementProviderFragment** outProvider) noexcept
{
    if (! outProvider)
    {
        return E_POINTER;
    }

    *outProvider                                                = nullptr;
    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    if (! snapshot || ! snapshot->alive || ! snapshot->hasRetainedRoot)
    {
        return S_OK;
    }

    POINT pointPx{static_cast<LONG>(std::lround(x)), static_cast<LONG>(std::lround(y))};
    if (! snapshot->hwnd || ScreenToClient(snapshot->hwnd, &pointPx) == FALSE)
    {
        return S_OK;
    }

    const D2D1_POINT_2F pointDip =
        D2D1::Point2F(static_cast<float>(pointPx.x) * snapshot->pixelsToDipScale, static_cast<float>(pointPx.y) * snapshot->pixelsToDipScale);
    const AccessibilityPointHitSnapshot* hit = FindSnapshotPointHit(*snapshot, pointDip);
    if (! hit)
    {
        return S_OK;
    }

    switch (hit->kind)
    {
        case AccessibilityFragmentKind::TextFieldPasswordRevealButton: *outProvider = CreateTextFieldPasswordRevealButtonProvider(hit->path); return S_OK;
        case AccessibilityFragmentKind::TreeItem: *outProvider = CreateTreeItemProvider(hit->path, hit->treeVisibleIndex); return S_OK;
        case AccessibilityFragmentKind::GridHeader: *outProvider = CreateGridHeaderProvider(hit->path, hit->gridColumnIndex); return S_OK;
        case AccessibilityFragmentKind::GridRow: *outProvider = CreateGridRowProvider(hit->path, hit->gridRowId); return S_OK;
        case AccessibilityFragmentKind::GridCell: *outProvider = CreateGridCellProvider(hit->path, hit->gridRowId, hit->gridColumnIndex); return S_OK;
        case AccessibilityFragmentKind::Control:
            *outProvider = SnapshotPathIsCollapsedSemanticRoot(*snapshot, hit->path) ? MakeProvider<IRawElementProviderFragment, AccessibilityProvider>(_hwnd)
                                                                                     : CreateChildProvider(hit->path);
            return S_OK;
        case AccessibilityFragmentKind::Root: *outProvider = MakeProvider<IRawElementProviderFragment, AccessibilityProvider>(_hwnd); return S_OK;
    }

    return S_OK;
}

HRESULT AccessibilityProvider::GetFocus(IRawElementProviderFragment** outProvider) noexcept
{
    if (! outProvider)
    {
        return E_POINTER;
    }

    *outProvider                                                = nullptr;
    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    if (! snapshot || ! snapshot->alive || ! snapshot->hasRetainedRoot || ! snapshot->focusedFragment.has_value())
    {
        return S_OK;
    }

    const AccessibilityFocusedFragmentSnapshot& focusedFragment = snapshot->focusedFragment.value();
    switch (focusedFragment.kind)
    {
        case AccessibilityFragmentKind::TreeItem: *outProvider = CreateTreeItemProvider(focusedFragment.path, focusedFragment.treeVisibleIndex); return S_OK;
        case AccessibilityFragmentKind::GridRow: *outProvider = CreateGridRowProvider(focusedFragment.path, focusedFragment.gridRowId); return S_OK;
        case AccessibilityFragmentKind::Control:
            *outProvider = SnapshotPathIsCollapsedSemanticRoot(*snapshot, focusedFragment.path)
                               ? MakeProvider<IRawElementProviderFragment, AccessibilityProvider>(_hwnd)
                               : CreateChildProvider(focusedFragment.path);
            return S_OK;
        case AccessibilityFragmentKind::Root:
        case AccessibilityFragmentKind::GridHeader:
        case AccessibilityFragmentKind::GridCell:
        case AccessibilityFragmentKind::TextFieldPasswordRevealButton: return S_OK;
    }

    return S_OK;
}

HRESULT AccessibilityProvider::Invoke() noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.provider = this;
        request.kind     = AccessibilityUiActionKind::Invoke;
        return DispatchActionToWindowThread(request);
    }

    return ExecuteInvokeOnWindowThread();
}

HRESULT AccessibilityProvider::Toggle() noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.provider = this;
        request.kind     = AccessibilityUiActionKind::Toggle;
        return DispatchActionToWindowThread(request);
    }

    return ExecuteToggleOnWindowThread();
}

HRESULT AccessibilityProvider::get_ToggleState(ToggleState* outState) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outState)
    {
        return E_POINTER;
    }

    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        const std::shared_ptr<const AccessibilitySnapshot> snapshot   = CaptureAccessibilitySnapshot(_target, _hwnd);
        const std::optional<AccessibilityGridCellSnapshotRecord> cell = (snapshot && snapshot->alive && snapshot->hasRetainedRoot)
                                                                            ? FindSnapshotGridCellRecord(*snapshot, _path, _gridRowId, _gridColumnIndex)
                                                                            : std::nullopt;
        if (! cell || ! SnapshotGridCellSupportsTogglePattern(cell.value()) || ! cell->cellRecord)
        {
            return UIA_E_NOTSUPPORTED;
        }

        *outState = cell->cellRecord->gridCellChecked ? ToggleState_On : ToggleState_Off;
        return S_OK;
    }

    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (! record || ! record->controlSupportsToggle)
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outState = record->controlToggleChecked ? ToggleState_On : ToggleState_Off;
    return S_OK;
}

HRESULT AccessibilityProvider::GetVisibleRanges(SAFEARRAY** outRanges) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outRanges)
    {
        return E_POINTER;
    }

    *outRanges                                                  = nullptr;
    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (! record || ! record->controlSupportsText)
    {
        return UIA_E_NOTSUPPORTED;
    }

    wil::com_ptr_nothrow<ITextRangeProvider> range;
    range.attach(CreateTextDocumentRangeProvider(*record));
    if (! range)
    {
        return E_OUTOFMEMORY;
    }

    ITextRangeProvider* rawRanges[] = {range.get()};
    const HRESULT hr                = SetTextRangeProviderArray(outRanges, rawRanges);
    Debug::Perf::Emit(L"dxui.uia.text_range_us", L"visible-ranges", Debug::Perf::ElapsedUs(startedAt), 1u, record->controlAccessibleText.size(), hr);
    return hr;
}

HRESULT AccessibilityProvider::RangeFromChild(IRawElementProviderSimple* /*childElement*/, ITextRangeProvider** outRange) noexcept
{
    return get_DocumentRange(outRange);
}

HRESULT AccessibilityProvider::RangeFromPoint(UiaPoint point, ITextRangeProvider** outRange) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    if (! outRange)
    {
        return E_POINTER;
    }

    *outRange                                                   = nullptr;
    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (! record || ! record->controlSupportsText)
    {
        return UIA_E_NOTSUPPORTED;
    }

    AccessibilityUiActionRequest request{};
    request.provider       = this;
    request.kind           = AccessibilityUiActionKind::ResolveTextRangeFromPoint;
    request.textRangePoint = point;

    const HRESULT resolveHr = IsCurrentThreadWindowThread()
                                  ? ExecuteResolveTextRangeFromPointOnWindowThread(point, request.textRangeResultStart, request.textRangeTextLength)
                                  : DispatchActionToWindowThread(request);
    if (FAILED(resolveHr) || resolveHr != S_OK)
    {
        return resolveHr;
    }

    *outRange        = CreateTextRangeProvider(record->path, request.textRangeResultStart, request.textRangeResultStart);
    const HRESULT hr = *outRange ? S_OK : E_OUTOFMEMORY;
    Debug::Perf::Emit(
        L"dxui.uia.text_range_us", L"range-from-point", Debug::Perf::ElapsedUs(startedAt), request.textRangeResultStart, request.textRangeTextLength, hr);
    return hr;
}

HRESULT AccessibilityProvider::ExecuteResolveTextRangeFromPointOnWindowThread(UiaPoint point, size_t& outCaretIndex, size_t& outTextLength) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    outCaretIndex = 0u;
    outTextLength = 0u;

    WindowHost* host       = ResolveHost();
    const Control* control = ResolveControl();
    if (! host || ! control || ! SupportsTextPattern(control))
    {
        return UIA_E_NOTSUPPORTED;
    }

    const std::wstring text = GetControlAccessibleTextRangeText(control);
    outTextLength           = text.size();
    const POINT pointScreen{static_cast<LONG>(std::lround(point.x)), static_cast<LONG>(std::lround(point.y))};
    const std::optional<PointDip> pointDip = host->ScreenPointToDipPoint(pointScreen);
    if (! pointDip)
    {
        return E_INVALIDARG;
    }

    const D2D1_POINT_2F hitPoint = D2D1::Point2F(pointDip->x, pointDip->y);
    size_t caretIndex            = 0u;
    if (const std::optional<size_t> hitIndex = control->TryHitTestTextInputPoint(*host, hitPoint); hitIndex.has_value())
    {
        caretIndex = (std::min)(hitIndex.value(), text.size());
    }
    else if (const auto* comboBox = dynamic_cast<const ComboBox*>(control); comboBox && comboBox->IsEditable())
    {
        caretIndex = comboBox->HitTestEditableCaretIndex(*host, hitPoint);
    }
    else
    {
        caretIndex = HitTestCaretIndexDip(
            host, text, FontRole::Body, ResolveTextPatternViewportRect(control), 0.0f, hitPoint, ResolveReadingDirection(control->GetFlowDirection()));
    }

    outCaretIndex = caretIndex;
    return S_OK;
}

HRESULT AccessibilityProvider::get_DocumentRange(ITextRangeProvider** outRange) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outRange)
    {
        return E_POINTER;
    }

    *outRange                                                   = nullptr;
    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (! record || ! record->controlSupportsText)
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outRange        = CreateTextDocumentRangeProvider(*record);
    const HRESULT hr = *outRange ? S_OK : E_OUTOFMEMORY;
    Debug::Perf::Emit(L"dxui.uia.text_range_us", L"document-range", Debug::Perf::ElapsedUs(startedAt), record->controlAccessibleText.size(), 0u, hr);
    return hr;
}

HRESULT AccessibilityProvider::get_SupportedTextSelection(SupportedTextSelection* outSupportedSelection) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outSupportedSelection)
    {
        return E_POINTER;
    }

    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (! record || ! record->controlSupportsText)
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outSupportedSelection = SupportedTextSelection_Single;
    return S_OK;
}

HRESULT AccessibilityProvider::GetActiveComposition(ITextRangeProvider** outRange) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outRange)
    {
        return E_POINTER;
    }

    *outRange                                                   = nullptr;
    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (! record || ! record->controlSupportsText)
    {
        return UIA_E_NOTSUPPORTED;
    }

    if (! record->controlTextCompositionStart || ! record->controlTextCompositionEnd)
    {
        Debug::Perf::Emit(L"dxui.uia.text_range_us", L"active-composition", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, S_OK);
        return S_OK;
    }

    const TextRangeSpan range =
        ClampTextRangeSpan(record->controlTextCompositionStart.value(), record->controlTextCompositionEnd.value(), record->controlAccessibleText.size());
    *outRange        = CreateTextRangeProvider(record->path, range.start, range.end, record->controlAccessibleText);
    const HRESULT hr = *outRange ? S_OK : E_OUTOFMEMORY;
    Debug::Perf::Emit(
        L"dxui.uia.text_range_us", L"active-composition", Debug::Perf::ElapsedUs(startedAt), range.end - range.start, record->controlAccessibleText.size(), hr);
    return hr;
}

HRESULT AccessibilityProvider::GetConversionTarget(ITextRangeProvider** outRange) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outRange)
    {
        return E_POINTER;
    }

    *outRange                                                   = nullptr;
    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (! record || ! record->controlSupportsText)
    {
        return UIA_E_NOTSUPPORTED;
    }

    if (! record->controlTextConversionTargetStart || ! record->controlTextConversionTargetEnd)
    {
        Debug::Perf::Emit(L"dxui.uia.text_range_us", L"conversion-target", Debug::Perf::ElapsedUs(startedAt), 0u, 0u, S_OK);
        return S_OK;
    }

    const TextRangeSpan range = ClampTextRangeSpan(
        record->controlTextConversionTargetStart.value(), record->controlTextConversionTargetEnd.value(), record->controlAccessibleText.size());
    *outRange        = CreateTextRangeProvider(record->path, range.start, range.end, record->controlAccessibleText);
    const HRESULT hr = *outRange ? S_OK : E_OUTOFMEMORY;
    Debug::Perf::Emit(
        L"dxui.uia.text_range_us", L"conversion-target", Debug::Perf::ElapsedUs(startedAt), range.end - range.start, record->controlAccessibleText.size(), hr);
    return hr;
}

HRESULT AccessibilityProvider::SetValue(LPCWSTR value) noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.provider    = this;
        request.kind        = AccessibilityUiActionKind::SetStringValue;
        request.stringValue = value ? value : L"";
        return DispatchActionToWindowThread(request);
    }

    return ExecuteSetStringValueOnWindowThread(value);
}

HRESULT AccessibilityProvider::SetValue(double value) noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.provider    = this;
        request.kind        = AccessibilityUiActionKind::SetRangeValue;
        request.numberValue = value;
        return DispatchActionToWindowThread(request);
    }

    return ExecuteSetRangeValueOnWindowThread(value);
}

HRESULT AccessibilityProvider::get_Value(BSTR* outValue) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outValue)
    {
        return E_POINTER;
    }

    *outValue = nullptr;
    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        const std::shared_ptr<const AccessibilitySnapshot> snapshot   = CaptureAccessibilitySnapshot(_target, _hwnd);
        const std::optional<AccessibilityGridCellSnapshotRecord> cell = (snapshot && snapshot->alive && snapshot->hasRetainedRoot)
                                                                            ? FindSnapshotGridCellRecord(*snapshot, _path, _gridRowId, _gridColumnIndex)
                                                                            : std::nullopt;
        if (! cell || ! SnapshotGridCellSupportsValuePattern(cell.value()) || ! cell->cellRecord)
        {
            return UIA_E_NOTSUPPORTED;
        }

        const std::wstring& value = cell->cellRecord->gridCellAccessibleText;
        *outValue                 = SysAllocStringLen(value.data(), static_cast<UINT>(value.size()));
        return (*outValue || value.empty()) ? S_OK : E_OUTOFMEMORY;
    }

    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (! record || ! record->controlSupportsValue)
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outValue = SysAllocStringLen(record->controlAccessibleValue.data(), static_cast<UINT>(record->controlAccessibleValue.size()));
    return (*outValue || record->controlAccessibleValue.empty()) ? S_OK : E_OUTOFMEMORY;
}

HRESULT AccessibilityProvider::get_Value(double* outValue) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outValue)
    {
        return E_POINTER;
    }

    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        const std::shared_ptr<const AccessibilitySnapshot> snapshot   = CaptureAccessibilitySnapshot(_target, _hwnd);
        const std::optional<AccessibilityGridCellSnapshotRecord> cell = (snapshot && snapshot->alive && snapshot->hasRetainedRoot)
                                                                            ? FindSnapshotGridCellRecord(*snapshot, _path, _gridRowId, _gridColumnIndex)
                                                                            : std::nullopt;
        if (! cell || ! SnapshotGridCellSupportsRangeValuePattern(cell.value()) || ! cell->cellRecord)
        {
            return UIA_E_NOTSUPPORTED;
        }

        *outValue = cell->cellRecord->gridCellRangeValue;
        return S_OK;
    }

    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (record && record->controlSupportsRangeValue)
    {
        *outValue = record->controlRangeValue;
        return S_OK;
    }
    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::get_IsReadOnly(BOOL* outReadOnly) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outReadOnly)
    {
        return E_POINTER;
    }

    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        const std::shared_ptr<const AccessibilitySnapshot> snapshot   = CaptureAccessibilitySnapshot(_target, _hwnd);
        const std::optional<AccessibilityGridCellSnapshotRecord> cell = (snapshot && snapshot->alive && snapshot->hasRetainedRoot)
                                                                            ? FindSnapshotGridCellRecord(*snapshot, _path, _gridRowId, _gridColumnIndex)
                                                                            : std::nullopt;
        if (! cell || (! SnapshotGridCellSupportsValuePattern(cell.value()) && ! SnapshotGridCellSupportsRangeValuePattern(cell.value())))
        {
            return UIA_E_NOTSUPPORTED;
        }

        *outReadOnly = TRUE;
        return S_OK;
    }

    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (! record || (! record->controlSupportsValue && ! record->controlSupportsRangeValue))
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outReadOnly = record->controlValueReadOnly ? TRUE : FALSE;
    return S_OK;
}

HRESULT AccessibilityProvider::get_Maximum(double* outMaximum) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outMaximum)
    {
        return E_POINTER;
    }

    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        const std::shared_ptr<const AccessibilitySnapshot> snapshot   = CaptureAccessibilitySnapshot(_target, _hwnd);
        const std::optional<AccessibilityGridCellSnapshotRecord> cell = (snapshot && snapshot->alive && snapshot->hasRetainedRoot)
                                                                            ? FindSnapshotGridCellRecord(*snapshot, _path, _gridRowId, _gridColumnIndex)
                                                                            : std::nullopt;
        if (! cell || ! SnapshotGridCellSupportsRangeValuePattern(cell.value()))
        {
            return UIA_E_NOTSUPPORTED;
        }

        *outMaximum = 1.0;
        return S_OK;
    }

    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (record && record->controlSupportsRangeValue)
    {
        *outMaximum = record->controlRangeMaximum;
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::get_Minimum(double* outMinimum) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outMinimum)
    {
        return E_POINTER;
    }

    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        const std::shared_ptr<const AccessibilitySnapshot> snapshot   = CaptureAccessibilitySnapshot(_target, _hwnd);
        const std::optional<AccessibilityGridCellSnapshotRecord> cell = (snapshot && snapshot->alive && snapshot->hasRetainedRoot)
                                                                            ? FindSnapshotGridCellRecord(*snapshot, _path, _gridRowId, _gridColumnIndex)
                                                                            : std::nullopt;
        if (! cell || ! SnapshotGridCellSupportsRangeValuePattern(cell.value()))
        {
            return UIA_E_NOTSUPPORTED;
        }

        *outMinimum = 0.0;
        return S_OK;
    }

    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (record && record->controlSupportsRangeValue)
    {
        *outMinimum = record->controlRangeMinimum;
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::get_LargeChange(double* outLargeChange) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outLargeChange)
    {
        return E_POINTER;
    }

    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        const std::shared_ptr<const AccessibilitySnapshot> snapshot   = CaptureAccessibilitySnapshot(_target, _hwnd);
        const std::optional<AccessibilityGridCellSnapshotRecord> cell = (snapshot && snapshot->alive && snapshot->hasRetainedRoot)
                                                                            ? FindSnapshotGridCellRecord(*snapshot, _path, _gridRowId, _gridColumnIndex)
                                                                            : std::nullopt;
        if (! cell || ! SnapshotGridCellSupportsRangeValuePattern(cell.value()))
        {
            return UIA_E_NOTSUPPORTED;
        }

        *outLargeChange = 0.0;
        return S_OK;
    }

    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (record && record->controlSupportsRangeValue)
    {
        *outLargeChange = record->controlRangeLargeChange;
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::get_SmallChange(double* outSmallChange) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outSmallChange)
    {
        return E_POINTER;
    }

    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        const std::shared_ptr<const AccessibilitySnapshot> snapshot   = CaptureAccessibilitySnapshot(_target, _hwnd);
        const std::optional<AccessibilityGridCellSnapshotRecord> cell = (snapshot && snapshot->alive && snapshot->hasRetainedRoot)
                                                                            ? FindSnapshotGridCellRecord(*snapshot, _path, _gridRowId, _gridColumnIndex)
                                                                            : std::nullopt;
        if (! cell || ! SnapshotGridCellSupportsRangeValuePattern(cell.value()))
        {
            return UIA_E_NOTSUPPORTED;
        }

        *outSmallChange = 0.0;
        return S_OK;
    }

    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (record && record->controlSupportsRangeValue)
    {
        *outSmallChange = record->controlRangeSmallChange;
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::GetSelection(SAFEARRAY** outSelection) noexcept
{
    if (! outSelection)
    {
        return E_POINTER;
    }

    *outSelection                                               = nullptr;
    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (record && SnapshotSupportsSelectionProvider(*record))
    {
        std::vector<wil::com_ptr_nothrow<IRawElementProviderSimple>> selectionProviders;
        if (record->isTree)
        {
            if (record->selectedTreeVisibleIndex)
            {
                wil::com_ptr_nothrow<IRawElementProviderFragment> fragment;
                fragment.attach(CreateTreeItemProvider(record->path, record->selectedTreeVisibleIndex.value()));
                wil::com_ptr_nothrow<IRawElementProviderSimple> simple;
                if (! fragment || FAILED(fragment.query_to(simple.put())))
                {
                    return E_OUTOFMEMORY;
                }
                selectionProviders.push_back(std::move(simple));
            }
        }
        else if (record->isGrid)
        {
            selectionProviders.reserve(record->selectedGridRowIds.size());
            for (const uint64_t rowId : record->selectedGridRowIds)
            {
                wil::com_ptr_nothrow<IRawElementProviderFragment> fragment;
                fragment.attach(CreateGridRowProvider(record->path, rowId));
                wil::com_ptr_nothrow<IRawElementProviderSimple> simple;
                if (! fragment || FAILED(fragment.query_to(simple.put())))
                {
                    return E_OUTOFMEMORY;
                }
                selectionProviders.push_back(std::move(simple));
            }
        }

        std::vector<IRawElementProviderSimple*> rawProviders;
        rawProviders.reserve(selectionProviders.size());
        for (const auto& provider : selectionProviders)
        {
            rawProviders.push_back(provider.get());
        }

        return SetProviderArray(outSelection, rawProviders);
    }

    const AccessibilityControlNavigationSnapshot* textRecord =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (textRecord && textRecord->controlSupportsText)
    {
        wil::com_ptr_nothrow<ITextRangeProvider> range;
        if (! textRecord->controlTextSelectionBoundsDip.empty())
        {
            range.attach(CreateTextRangeProvider(
                textRecord->path, textRecord->controlTextSelectionStart, textRecord->controlTextSelectionEnd, textRecord->controlTextSelectionBoundsDip));
        }
        else
        {
            range.attach(CreateTextRangeProvider(textRecord->path, textRecord->controlTextSelectionStart, textRecord->controlTextSelectionEnd));
        }
        if (! range)
        {
            return E_OUTOFMEMORY;
        }

        ITextRangeProvider* rawRanges[] = {range.get()};
        return SetTextRangeProviderArray(outSelection, rawRanges);
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::get_CanSelectMultiple(BOOL* outCanSelectMultiple) noexcept
{
    if (! outCanSelectMultiple)
    {
        return E_POINTER;
    }

    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (! record || ! SnapshotSupportsSelectionProvider(*record))
    {
        return UIA_E_NOTSUPPORTED;
    }

    if (record->isTree)
    {
        *outCanSelectMultiple = FALSE;
        return S_OK;
    }
    if (record->isGrid)
    {
        *outCanSelectMultiple = record->gridCanSelectMultiple ? TRUE : FALSE;
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::get_IsSelectionRequired(BOOL* outIsSelectionRequired) noexcept
{
    if (! outIsSelectionRequired)
    {
        return E_POINTER;
    }

    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (! record || ! SnapshotSupportsSelectionProvider(*record))
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outIsSelectionRequired = FALSE;
    return S_OK;
}

HRESULT AccessibilityProvider::GetRowHeaders(SAFEARRAY** outRowHeaders) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (! outRowHeaders)
    {
        return E_POINTER;
    }

    *outRowHeaders                                              = nullptr;
    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (! record || ! record->isGrid)
    {
        return UIA_E_NOTSUPPORTED;
    }

    return SetProviderArray(outRowHeaders, {});
}

HRESULT AccessibilityProvider::GetColumnHeaders(SAFEARRAY** outColumnHeaders) noexcept
{
    if (! outColumnHeaders)
    {
        return E_POINTER;
    }

    *outColumnHeaders                                           = nullptr;
    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (! record || ! record->isGrid)
    {
        return UIA_E_NOTSUPPORTED;
    }

    std::vector<wil::com_ptr_nothrow<IRawElementProviderSimple>> headerProviders;
    headerProviders.reserve(record->gridVisibleColumns.size());
    for (const size_t columnIndex : record->gridVisibleColumns)
    {
        wil::com_ptr_nothrow<IRawElementProviderFragment> fragment;
        fragment.attach(CreateGridHeaderProvider(record->path, columnIndex));
        wil::com_ptr_nothrow<IRawElementProviderSimple> simple;
        if (! fragment || FAILED(fragment.query_to(simple.put())))
        {
            return E_OUTOFMEMORY;
        }

        headerProviders.push_back(std::move(simple));
    }

    std::vector<IRawElementProviderSimple*> rawProviders;
    rawProviders.reserve(headerProviders.size());
    for (const auto& provider : headerProviders)
    {
        rawProviders.push_back(provider.get());
    }

    return SetProviderArray(outColumnHeaders, rawProviders);
}

HRESULT AccessibilityProvider::get_RowOrColumnMajor(RowOrColumnMajor* outRowOrColumnMajor) noexcept
{
    if (! outRowOrColumnMajor)
    {
        return E_POINTER;
    }

    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? ResolveSnapshotControlRecord(*snapshot, _kind, _path) : nullptr;
    if (! record || ! record->isGrid)
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outRowOrColumnMajor = RowOrColumnMajor_Indeterminate;
    return S_OK;
}

HRESULT AccessibilityProvider::Select() noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.provider = this;
        request.kind     = AccessibilityUiActionKind::Select;
        return DispatchActionToWindowThread(request);
    }

    return ExecuteSelectOnWindowThread();
}

HRESULT AccessibilityProvider::AddToSelection() noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.provider = this;
        request.kind     = AccessibilityUiActionKind::AddToSelection;
        return DispatchActionToWindowThread(request);
    }

    return ExecuteAddToSelectionOnWindowThread();
}

HRESULT AccessibilityProvider::RemoveFromSelection() noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.provider = this;
        request.kind     = AccessibilityUiActionKind::RemoveFromSelection;
        return DispatchActionToWindowThread(request);
    }

    return ExecuteRemoveFromSelectionOnWindowThread();
}

HRESULT AccessibilityProvider::get_IsSelected(BOOL* outSelected) noexcept
{
    if (! outSelected)
    {
        return E_POINTER;
    }

    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? FindControlNavigationRecord(*snapshot, _path) : nullptr;
    if (_kind == AccessibilityFragmentKind::TreeItem)
    {
        if (! record || ! SnapshotContainsTreeItem(*record, _treeVisibleIndex))
        {
            return UIA_E_NOTSUPPORTED;
        }

        *outSelected = SnapshotTreeItemIsSelected(*record, _treeVisibleIndex) ? TRUE : FALSE;
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::GridRow)
    {
        if (! record || ! record->isGrid || ! SnapshotContainsGridRow(*record, _gridRowId))
        {
            return UIA_E_NOTSUPPORTED;
        }

        *outSelected = SnapshotGridRowIsSelected(*record, _gridRowId) ? TRUE : FALSE;
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::get_SelectionContainer(IRawElementProviderSimple** outContainer) noexcept
{
    if (! outContainer)
    {
        return E_POINTER;
    }

    *outContainer                                               = nullptr;
    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? FindControlNavigationRecord(*snapshot, _path) : nullptr;
    const bool treeItemSupported = _kind == AccessibilityFragmentKind::TreeItem && record && SnapshotContainsTreeItem(*record, _treeVisibleIndex);
    const bool gridRowSupported  = _kind == AccessibilityFragmentKind::GridRow && record && record->isGrid && SnapshotContainsGridRow(*record, _gridRowId);
    if (! treeItemSupported && ! gridRowSupported)
    {
        return UIA_E_NOTSUPPORTED;
    }

    WindowHostAccessibilityTarget* target = AddRefTarget();
    auto* provider                        = new (std::nothrow) AccessibilityProvider(target, _hwnd, _path);
    if (! provider)
    {
        if (target)
        {
            static_cast<void>(target->Release());
        }
        return E_OUTOFMEMORY;
    }

    *outContainer = static_cast<IRawElementProviderSimple*>(provider);
    return S_OK;
}

HRESULT AccessibilityProvider::Expand() noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.provider = this;
        request.kind     = AccessibilityUiActionKind::Expand;
        return DispatchActionToWindowThread(request);
    }

    return ExecuteExpandOnWindowThread(true);
}

HRESULT AccessibilityProvider::Collapse() noexcept
{
    if (! IsCurrentThreadWindowThread())
    {
        AccessibilityUiActionRequest request{};
        request.provider = this;
        request.kind     = AccessibilityUiActionKind::Collapse;
        return DispatchActionToWindowThread(request);
    }

    return ExecuteExpandOnWindowThread(false);
}

HRESULT AccessibilityProvider::get_ExpandCollapseState(ExpandCollapseState* outState) noexcept
{
    if (! outState)
    {
        return E_POINTER;
    }

    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const AccessibilityControlNavigationSnapshot* record =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? FindControlNavigationRecord(*snapshot, _path) : nullptr;
    const AccessibilityTreeItemSnapshotRecord* item = record ? FindSnapshotTreeItemRecord(*record, _treeVisibleIndex) : nullptr;
    if (! item || ! item->hasChildren)
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outState = item->expanded ? ExpandCollapseState_Expanded : ExpandCollapseState_Collapsed;
    return S_OK;
}

HRESULT AccessibilityProvider::get_Row(int* outRow) noexcept
{
    if (! outRow)
    {
        return E_POINTER;
    }

    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const std::optional<AccessibilityGridCellSnapshotRecord> cell =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? FindSnapshotGridCellRecord(*snapshot, _path, _gridRowId, _gridColumnIndex) : std::nullopt;
    if (! cell || cell->rowIndex > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outRow = static_cast<int>(cell->rowIndex);
    return S_OK;
}

HRESULT AccessibilityProvider::get_Column(int* outColumn) noexcept
{
    if (! outColumn)
    {
        return E_POINTER;
    }

    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const std::optional<AccessibilityGridCellSnapshotRecord> cell =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? FindSnapshotGridCellRecord(*snapshot, _path, _gridRowId, _gridColumnIndex) : std::nullopt;
    if (! cell || cell->columnIndex > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outColumn = static_cast<int>(cell->columnIndex);
    return S_OK;
}

HRESULT AccessibilityProvider::get_RowSpan(int* outRowSpan) noexcept
{
    if (! outRowSpan)
    {
        return E_POINTER;
    }

    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const std::optional<AccessibilityGridCellSnapshotRecord> cell =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? FindSnapshotGridCellRecord(*snapshot, _path, _gridRowId, _gridColumnIndex) : std::nullopt;
    if (! cell)
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outRowSpan = 1;
    return S_OK;
}

HRESULT AccessibilityProvider::get_ColumnSpan(int* outColumnSpan) noexcept
{
    if (! outColumnSpan)
    {
        return E_POINTER;
    }

    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const std::optional<AccessibilityGridCellSnapshotRecord> cell =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? FindSnapshotGridCellRecord(*snapshot, _path, _gridRowId, _gridColumnIndex) : std::nullopt;
    if (! cell)
    {
        return UIA_E_NOTSUPPORTED;
    }

    *outColumnSpan = 1;
    return S_OK;
}

HRESULT AccessibilityProvider::get_ContainingGrid(IRawElementProviderSimple** outContainingGrid) noexcept
{
    if (! outContainingGrid)
    {
        return E_POINTER;
    }

    *outContainingGrid                                          = nullptr;
    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const std::optional<AccessibilityGridCellSnapshotRecord> cell =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? FindSnapshotGridCellRecord(*snapshot, _path, _gridRowId, _gridColumnIndex) : std::nullopt;
    if (! cell)
    {
        return UIA_E_NOTSUPPORTED;
    }

    WindowHostAccessibilityTarget* target = AddRefTarget();
    auto* provider                        = new (std::nothrow) AccessibilityProvider(target, _hwnd, _path);
    if (! provider)
    {
        if (target)
        {
            static_cast<void>(target->Release());
        }
        return E_OUTOFMEMORY;
    }

    *outContainingGrid = static_cast<IRawElementProviderSimple*>(provider);
    return S_OK;
}

HRESULT AccessibilityProvider::GetRowHeaderItems(SAFEARRAY** outRowHeaderItems) noexcept
{
    if (! outRowHeaderItems)
    {
        return E_POINTER;
    }

    *outRowHeaderItems                                          = nullptr;
    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const std::optional<AccessibilityGridCellSnapshotRecord> cell =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? FindSnapshotGridCellRecord(*snapshot, _path, _gridRowId, _gridColumnIndex) : std::nullopt;
    if (! cell || cell->controlRecord->gridVisibleColumns.empty())
    {
        return UIA_E_NOTSUPPORTED;
    }

    return SetProviderArray(outRowHeaderItems, {});
}

HRESULT AccessibilityProvider::GetColumnHeaderItems(SAFEARRAY** outColumnHeaderItems) noexcept
{
    if (! outColumnHeaderItems)
    {
        return E_POINTER;
    }

    *outColumnHeaderItems                                       = nullptr;
    const std::shared_ptr<const AccessibilitySnapshot> snapshot = CaptureAccessibilitySnapshot(_target, _hwnd);
    const std::optional<AccessibilityGridCellSnapshotRecord> cell =
        (snapshot && snapshot->alive && snapshot->hasRetainedRoot) ? FindSnapshotGridCellRecord(*snapshot, _path, _gridRowId, _gridColumnIndex) : std::nullopt;
    if (! cell || cell->controlRecord->gridVisibleColumns.empty())
    {
        return UIA_E_NOTSUPPORTED;
    }

    wil::com_ptr_nothrow<IRawElementProviderFragment> fragment;
    fragment.attach(CreateGridHeaderProvider(_path, cell->columnIndex));
    wil::com_ptr_nothrow<IRawElementProviderSimple> simple;
    if (! fragment || FAILED(fragment.query_to(simple.put())))
    {
        return E_OUTOFMEMORY;
    }

    IRawElementProviderSimple* provider = simple.get();
    return SetProviderArray(outColumnHeaderItems, std::span<IRawElementProviderSimple* const>(&provider, 1u));
}

WindowHost* AccessibilityProvider::ResolveHost() const noexcept
{
    return _target ? _target->ResolveHost() : nullptr;
}

const Control* AccessibilityProvider::ResolveRootControl() const noexcept
{
    WindowHost* host = ResolveHost();
    return host ? host->GetRoot() : nullptr;
}

bool AccessibilityProvider::ResolveControlPath(ControlPath& outPath) const noexcept
{
    if (_kind == AccessibilityFragmentKind::Control || _kind == AccessibilityFragmentKind::TextFieldPasswordRevealButton)
    {
        outPath = _path;
        return true;
    }

    if (_kind == AccessibilityFragmentKind::Root)
    {
        return TryResolveSingleSemanticRootControlPath(ResolveRootControl(), outPath);
    }

    return false;
}

const Control* AccessibilityProvider::ResolveControl() const noexcept
{
    ControlPath path{};
    if (! ResolveControlPath(path))
    {
        return nullptr;
    }

    WindowHost* host = ResolveHost();
    return host ? ResolveControlAtPath(host->GetRoot(), path) : nullptr;
}

Control* AccessibilityProvider::ResolveMutableControl() const noexcept
{
    ControlPath path{};
    if (! ResolveControlPath(path))
    {
        return nullptr;
    }

    WindowHost* host = ResolveHost();
    return host ? ResolveControlAtPath(host->GetRoot(), path) : nullptr;
}

const Tree* AccessibilityProvider::ResolveTreeControl() const noexcept
{
    ControlPath path{};
    if (_kind == AccessibilityFragmentKind::TreeItem)
    {
        path = _path;
    }
    else if (! ResolveControlPath(path))
    {
        return nullptr;
    }

    WindowHost* host = ResolveHost();
    return host ? dynamic_cast<const Tree*>(ResolveControlAtPath(host->GetRoot(), path)) : nullptr;
}

Tree* AccessibilityProvider::ResolveMutableTreeControl() const noexcept
{
    ControlPath path{};
    if (_kind == AccessibilityFragmentKind::TreeItem)
    {
        path = _path;
    }
    else if (! ResolveControlPath(path))
    {
        return nullptr;
    }

    WindowHost* host = ResolveHost();
    return host ? dynamic_cast<Tree*>(ResolveControlAtPath(host->GetRoot(), path)) : nullptr;
}

const Grid* AccessibilityProvider::ResolveGridControl() const noexcept
{
    ControlPath path{};
    if (_kind == AccessibilityFragmentKind::GridHeader || _kind == AccessibilityFragmentKind::GridRow || _kind == AccessibilityFragmentKind::GridCell)
    {
        path = _path;
    }
    else if (! ResolveControlPath(path))
    {
        return nullptr;
    }

    WindowHost* host = ResolveHost();
    return host ? dynamic_cast<const Grid*>(ResolveControlAtPath(host->GetRoot(), path)) : nullptr;
}

Grid* AccessibilityProvider::ResolveMutableGridControl() const noexcept
{
    ControlPath path{};
    if (_kind == AccessibilityFragmentKind::GridHeader || _kind == AccessibilityFragmentKind::GridRow || _kind == AccessibilityFragmentKind::GridCell)
    {
        path = _path;
    }
    else if (! ResolveControlPath(path))
    {
        return nullptr;
    }

    WindowHost* host = ResolveHost();
    return host ? dynamic_cast<Grid*>(ResolveControlAtPath(host->GetRoot(), path)) : nullptr;
}

bool AccessibilityProvider::ResolveTreeItemData(TreeItemData& outItem) const noexcept
{
    const Tree* tree = ResolveTreeControl();
    if (! tree)
    {
        return false;
    }

    const auto* model = tree->GetModel();
    if (! model || _treeVisibleIndex >= model->GetVisibleItemCount())
    {
        return false;
    }

    model->GetVisibleItem(_treeVisibleIndex, outItem);
    return true;
}

bool AccessibilityProvider::ResolveGridRowIndex(size_t& outRowIndex) const noexcept
{
    const Grid* grid  = ResolveGridControl();
    const auto* model = grid ? grid->GetModel() : nullptr;
    if (! grid || ! model)
    {
        return false;
    }

    const std::optional<size_t> rowIndex = model->FindRowByStableId(_gridRowId);
    if (! rowIndex)
    {
        return false;
    }

    outRowIndex = rowIndex.value();
    return true;
}

bool AccessibilityProvider::ResolveGridCellData(size_t& outRowIndex, size_t& outColumnIndex, GridCellData& outCellData) const noexcept
{
    const Grid* grid  = ResolveGridControl();
    const auto* model = grid ? grid->GetModel() : nullptr;
    if (! grid || ! model || ! ResolveGridRowIndex(outRowIndex) || ! grid->FindVisibleRowOrdinal(outRowIndex) ||
        ! grid->FindVisibleColumnOrdinal(_gridColumnIndex))
    {
        return false;
    }

    outColumnIndex = _gridColumnIndex;
    outCellData    = {};
    model->GetCellData(outRowIndex, outColumnIndex, outCellData);
    return true;
}

bool AccessibilityProvider::SupportsTreeItemSelectionPattern() const noexcept
{
    TreeItemData item;
    return ResolveTreeItemData(item);
}

bool AccessibilityProvider::SupportsTreeItemExpandCollapsePattern() const noexcept
{
    TreeItemData item;
    return ResolveTreeItemData(item) && item.hasChildren;
}

WindowHostAccessibilityTarget* AccessibilityProvider::AddRefTarget() const noexcept
{
    if (_target)
    {
        static_cast<void>(_target->AddRef());
    }

    return _target;
}

bool AccessibilityProvider::IsCurrentThreadWindowThread() const noexcept
{
    return _hwnd && GetWindowThreadProcessId(_hwnd, nullptr) == GetCurrentThreadId();
}

HRESULT AccessibilityProvider::DispatchActionToWindowThread(AccessibilityUiActionRequest& request) noexcept
{
    return DispatchAccessibilityUiActionToWindowThread(_hwnd, request);
}

HRESULT AccessibilityProvider::ExecuteUiThreadAction(AccessibilityUiActionRequest& request) noexcept
{
    switch (request.kind)
    {
        case AccessibilityUiActionKind::SetFocus: return ExecuteSetFocusOnWindowThread();
        case AccessibilityUiActionKind::Invoke: return ExecuteInvokeOnWindowThread();
        case AccessibilityUiActionKind::Toggle: return ExecuteToggleOnWindowThread();
        case AccessibilityUiActionKind::SetStringValue: return ExecuteSetStringValueOnWindowThread(request.stringValue.c_str());
        case AccessibilityUiActionKind::SetRangeValue: return ExecuteSetRangeValueOnWindowThread(request.numberValue);
        case AccessibilityUiActionKind::ResolveTextRangeFromPoint:
            return ExecuteResolveTextRangeFromPointOnWindowThread(request.textRangePoint, request.textRangeResultStart, request.textRangeTextLength);
        case AccessibilityUiActionKind::Select: return ExecuteSelectOnWindowThread();
        case AccessibilityUiActionKind::AddToSelection: return ExecuteAddToSelectionOnWindowThread();
        case AccessibilityUiActionKind::RemoveFromSelection: return ExecuteRemoveFromSelectionOnWindowThread();
        case AccessibilityUiActionKind::Expand: return ExecuteExpandOnWindowThread(true);
        case AccessibilityUiActionKind::Collapse: return ExecuteExpandOnWindowThread(false);
        case AccessibilityUiActionKind::MoveTextRangeByVisualLine: return UIA_E_NOTSUPPORTED;
        case AccessibilityUiActionKind::MoveTextRangeEndpointByVisualLine: return UIA_E_NOTSUPPORTED;
        case AccessibilityUiActionKind::ResolveTextRangeBounds: return UIA_E_NOTSUPPORTED;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::ExecuteSetFocusOnWindowThread() noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    WindowHost* host = ResolveHost();
    if (! host)
    {
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::Root)
    {
        ::SetFocus(_hwnd);
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::TreeItem)
    {
        TreeItemData item{};
        Tree* tree = ResolveMutableTreeControl();
        if (tree && ResolveTreeItemData(item))
        {
            ::SetFocus(_hwnd);
            tree->SetSelectedItemId(item.id);
            host->SetFocusControl(tree);
            RefreshWindowHostAccessibilitySnapshot(_hwnd, host);
            host->Invalidate();
        }
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::GridHeader)
    {
        Grid* grid = ResolveMutableGridControl();
        if (grid)
        {
            ::SetFocus(_hwnd);
            host->SetFocusControl(grid);
            RefreshWindowHostAccessibilitySnapshot(_hwnd, host);
            host->Invalidate();
        }
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::GridRow)
    {
        size_t rowIndex = 0u;
        Grid* grid      = ResolveMutableGridControl();
        if (grid && ResolveGridRowIndex(rowIndex) && grid->RequestSelectRow(rowIndex, 0u))
        {
            ::SetFocus(_hwnd);
            host->SetFocusControl(grid);
            RefreshWindowHostAccessibilitySnapshot(_hwnd, host);
            host->Invalidate();
        }
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        size_t rowIndex    = 0u;
        size_t columnIndex = 0u;
        GridCellData cellData{};
        Grid* grid = ResolveMutableGridControl();
        if (grid && ResolveGridCellData(rowIndex, columnIndex, cellData) && grid->RequestSelectRow(rowIndex, 0u))
        {
            ::SetFocus(_hwnd);
            host->SetFocusControl(grid);
            RefreshWindowHostAccessibilitySnapshot(_hwnd, host);
            host->Invalidate();
        }
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::TextFieldPasswordRevealButton)
    {
        auto* textField = dynamic_cast<TextField*>(ResolveMutableControl());
        if (textField && textField->IsPasswordRevealButtonVisibleForAccessibility())
        {
            ::SetFocus(_hwnd);
            host->SetFocusControl(textField);
            host->Invalidate();
        }
        return S_OK;
    }

    Control* control = ResolveMutableControl();
    if (control && control->IsFocusable())
    {
        ::SetFocus(_hwnd);
        host->SetFocusControl(control);
    }
    return S_OK;
}

HRESULT AccessibilityProvider::ExecuteInvokeOnWindowThread() noexcept
{
    WindowHost* host           = nullptr;
    TextField* revealTextField = nullptr;
    Button* button             = nullptr;

    {
        const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
        host = ResolveHost();
        if (! host)
        {
            return UIA_E_NOTSUPPORTED;
        }

        // Invoke callbacks can rebuild or destroy parts of the DxUi tree and may
        // raise accessibility events. Validate the target while locked, then run
        // the application callback after releasing the accessibility mutex.
        if (_kind == AccessibilityFragmentKind::TextFieldPasswordRevealButton)
        {
            revealTextField = dynamic_cast<TextField*>(ResolveMutableControl());
            if (! revealTextField || ! revealTextField->IsPasswordRevealButtonVisibleForAccessibility())
            {
                return UIA_E_NOTSUPPORTED;
            }
        }
        else
        {
            button = dynamic_cast<Button*>(ResolveMutableControl());
            if (! button || ! SupportsInvokePattern(button))
            {
                return UIA_E_NOTSUPPORTED;
            }
        }
    }

    if (_kind == AccessibilityFragmentKind::TextFieldPasswordRevealButton)
    {
        if (! revealTextField->InvokePasswordRevealButton(*host))
        {
            return UIA_E_NOTSUPPORTED;
        }
        return S_OK;
    }

    return button->Invoke(*host, true) ? S_OK : UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::ExecuteToggleOnWindowThread() noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    WindowHost* host = ResolveHost();
    if (! host)
    {
        return UIA_E_NOTSUPPORTED;
    }

    if (_kind == AccessibilityFragmentKind::GridCell)
    {
        size_t rowIndex    = 0u;
        size_t columnIndex = 0u;
        GridCellData cellData{};
        Grid* grid = ResolveMutableGridControl();
        if (! grid || ! ResolveGridCellData(rowIndex, columnIndex, cellData) || ! GridCellSupportsTogglePattern(cellData))
        {
            return UIA_E_NOTSUPPORTED;
        }

        if (! grid->RequestToggleCheckboxCell(*host, rowIndex, columnIndex))
        {
            return UIA_E_NOTSUPPORTED;
        }

        RefreshWindowHostAccessibilitySnapshot(_hwnd, host);
        return S_OK;
    }

    auto* toggle = dynamic_cast<RedSalamander::DxUi::Toggle*>(ResolveMutableControl());
    if (! toggle)
    {
        return UIA_E_NOTSUPPORTED;
    }

    static_cast<void>(toggle->OnMnemonic(*host));
    RefreshWindowHostAccessibilitySnapshot(_hwnd, host);
    return S_OK;
}

HRESULT AccessibilityProvider::ExecuteSetStringValueOnWindowThread(LPCWSTR value) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    WindowHost* host = ResolveHost();
    Control* control = ResolveMutableControl();
    if (! host || ! control || ! SupportsValuePattern(control) || IsValueReadOnly(control))
    {
        return UIA_E_NOTSUPPORTED;
    }

    if (auto* textField = dynamic_cast<TextField*>(control))
    {
        textField->SetTextAndNotify(value ? value : L"");
        host->SyncTextInput(textField);
        RefreshWindowHostAccessibilitySnapshot(_hwnd, host);
        host->Invalidate();
        return S_OK;
    }
    if (auto* comboBox = dynamic_cast<ComboBox*>(control))
    {
        comboBox->SetTextAndNotify(value ? value : L"");
        host->SyncTextInput(comboBox);
        RefreshWindowHostAccessibilitySnapshot(_hwnd, host);
        host->Invalidate();
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::ExecuteSetRangeValueOnWindowThread(double value) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    WindowHost* host = ResolveHost();
    Control* control = ResolveMutableControl();
    if (! host || ! control || ! SupportsRangeValuePattern(control) || IsValueReadOnly(control))
    {
        return UIA_E_NOTSUPPORTED;
    }

    if (auto* slider = dynamic_cast<Slider*>(control))
    {
        slider->SetValue(value);
        host->Invalidate();
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::ExecuteSelectOnWindowThread() noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    WindowHost* host = ResolveHost();
    if (! host)
    {
        return UIA_E_NOTSUPPORTED;
    }

    if (_kind == AccessibilityFragmentKind::TreeItem)
    {
        Tree* tree = ResolveMutableTreeControl();
        if (! tree || ! SupportsTreeItemSelectionPattern() || ! tree->RequestSelectVisibleItem(_treeVisibleIndex))
        {
            return UIA_E_NOTSUPPORTED;
        }

        host->SetFocusControl(tree);
        RefreshWindowHostAccessibilitySnapshot(_hwnd, host);
        host->Invalidate();
        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::GridRow)
    {
        size_t rowIndex = 0u;
        Grid* grid      = ResolveMutableGridControl();
        if (! grid || ! ResolveGridRowIndex(rowIndex) || ! grid->RequestSelectRow(rowIndex, 0u))
        {
            return UIA_E_NOTSUPPORTED;
        }

        host->SetFocusControl(grid);
        RefreshWindowHostAccessibilitySnapshot(_hwnd, host);
        host->Invalidate();
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::ExecuteAddToSelectionOnWindowThread() noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    if (_kind == AccessibilityFragmentKind::TreeItem)
    {
        return ExecuteSelectOnWindowThread();
    }

    WindowHost* host = ResolveHost();
    size_t rowIndex  = 0u;
    Grid* grid       = ResolveMutableGridControl();
    if (! host || ! grid || _kind != AccessibilityFragmentKind::GridRow || ! ResolveGridRowIndex(rowIndex) || ! grid->RequestSelectRow(rowIndex, MK_CONTROL))
    {
        return UIA_E_NOTSUPPORTED;
    }

    host->SetFocusControl(grid);
    RefreshWindowHostAccessibilitySnapshot(_hwnd, host);
    host->Invalidate();
    return S_OK;
}

HRESULT AccessibilityProvider::ExecuteRemoveFromSelectionOnWindowThread() noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    WindowHost* host = ResolveHost();
    if (! host)
    {
        return UIA_E_NOTSUPPORTED;
    }

    if (_kind == AccessibilityFragmentKind::TreeItem)
    {
        Tree* tree = ResolveMutableTreeControl();
        TreeItemData item;
        if (! tree || ! ResolveTreeItemData(item))
        {
            return UIA_E_NOTSUPPORTED;
        }

        if (tree->GetSelectedItemId() && tree->GetSelectedItemId().value() == item.id)
        {
            tree->SetSelectedItemId(std::nullopt);
            RefreshWindowHostAccessibilitySnapshot(_hwnd, host);
            host->Invalidate();
        }

        return S_OK;
    }

    if (_kind == AccessibilityFragmentKind::GridRow)
    {
        size_t rowIndex = 0u;
        Grid* grid      = ResolveMutableGridControl();
        if (! grid || ! ResolveGridRowIndex(rowIndex) || ! grid->RequestRemoveRowSelection(rowIndex))
        {
            return UIA_E_NOTSUPPORTED;
        }

        RefreshWindowHostAccessibilitySnapshot(_hwnd, host);
        host->Invalidate();
        return S_OK;
    }

    return UIA_E_NOTSUPPORTED;
}

HRESULT AccessibilityProvider::ExecuteExpandOnWindowThread(bool expanded) noexcept
{
    const std::scoped_lock accessibilityLock(GetAccessibilityTargetMutex());
    WindowHost* host = ResolveHost();
    Tree* tree       = ResolveMutableTreeControl();
    if (! host || ! tree || ! SupportsTreeItemExpandCollapsePattern())
    {
        return UIA_E_NOTSUPPORTED;
    }

    TreeItemData item;
    if (! ResolveTreeItemData(item))
    {
        return UIA_E_NOTSUPPORTED;
    }
    const bool stateChanged = item.expanded != expanded;
    if (! tree->RequestExpandedState(_treeVisibleIndex, expanded))
    {
        return UIA_E_NOTSUPPORTED;
    }

    RefreshWindowHostAccessibilitySnapshot(_hwnd, host);
    if (stateChanged && ! host->GetTheme().reducedMotion)
    {
        host->RequestAnimation();
    }
    host->Invalidate();
    return S_OK;
}

IRawElementProviderFragmentRoot* AccessibilityProvider::CreateRootProvider() noexcept
{
    return MakeProvider<IRawElementProviderFragmentRoot, AccessibilityProvider>(_hwnd);
}

IRawElementProviderFragment* AccessibilityProvider::CreateChildProvider(const ControlPath& path) noexcept
{
    return MakeProvider<IRawElementProviderFragment, AccessibilityProvider>(_hwnd, path);
}

IRawElementProviderFragment* AccessibilityProvider::CreateTextFieldPasswordRevealButtonProvider(const ControlPath& path) noexcept
{
    return MakeProvider<IRawElementProviderFragment, AccessibilityProvider>(_hwnd, path, AccessibilityFragmentKind::TextFieldPasswordRevealButton);
}

IRawElementProviderFragment* AccessibilityProvider::CreateTreeItemProvider(const ControlPath& path, size_t visibleIndex) noexcept
{
    return MakeProvider<IRawElementProviderFragment, AccessibilityProvider>(_hwnd, path, visibleIndex);
}

IRawElementProviderFragment* AccessibilityProvider::CreateGridHeaderProvider(const ControlPath& path, size_t columnIndex) noexcept
{
    return MakeProvider<IRawElementProviderFragment, AccessibilityProvider>(_hwnd, path, columnIndex, AccessibilityProvider::GridHeaderTag{});
}

IRawElementProviderFragment* AccessibilityProvider::CreateGridRowProvider(const ControlPath& path, uint64_t rowId) noexcept
{
    return MakeProvider<IRawElementProviderFragment, AccessibilityProvider>(_hwnd, path, rowId, AccessibilityFragmentKind::GridRow);
}

IRawElementProviderFragment* AccessibilityProvider::CreateGridCellProvider(const ControlPath& path, uint64_t rowId, size_t columnIndex) noexcept
{
    return MakeProvider<IRawElementProviderFragment, AccessibilityProvider>(_hwnd, path, rowId, columnIndex);
}

IRawElementProviderFragment* AccessibilityProvider::CreateProviderFromNavigationTarget(const AccessibilityNavigationTarget& navigationTarget) noexcept
{
    switch (navigationTarget.kind)
    {
        case AccessibilityFragmentKind::Root:
        {
            return MakeProvider<IRawElementProviderFragment, AccessibilityProvider>(_hwnd);
        }
        case AccessibilityFragmentKind::Control: return CreateChildProvider(navigationTarget.path);
        case AccessibilityFragmentKind::TextFieldPasswordRevealButton: return CreateTextFieldPasswordRevealButtonProvider(navigationTarget.path);
        case AccessibilityFragmentKind::TreeItem: return CreateTreeItemProvider(navigationTarget.path, navigationTarget.treeVisibleIndex);
        case AccessibilityFragmentKind::GridHeader: return CreateGridHeaderProvider(navigationTarget.path, navigationTarget.gridColumnIndex);
        case AccessibilityFragmentKind::GridRow: return CreateGridRowProvider(navigationTarget.path, navigationTarget.gridRowId);
        case AccessibilityFragmentKind::GridCell:
            return CreateGridCellProvider(navigationTarget.path, navigationTarget.gridRowId, navigationTarget.gridColumnIndex);
        default: return nullptr;
    }
}

ITextRangeProvider* AccessibilityProvider::CreateTextRangeProvider(const ControlPath& path, size_t start, size_t end) noexcept
{
    return MakeProvider<ITextRangeProvider, AccessibilityTextRangeProvider>(_hwnd, path, start, end);
}

ITextRangeProvider* AccessibilityProvider::CreateTextRangeProvider(const ControlPath& path,
                                                                   size_t start,
                                                                   size_t end,
                                                                   std::vector<D2D1_RECT_F> boundsOverrideDip) noexcept
{
    return MakeProvider<ITextRangeProvider, AccessibilityTextRangeProvider>(_hwnd, path, start, end, std::move(boundsOverrideDip));
}

ITextRangeProvider* AccessibilityProvider::CreateTextRangeProvider(const ControlPath& path, size_t start, size_t end, std::wstring textOverride) noexcept
{
    return MakeProvider<ITextRangeProvider, AccessibilityTextRangeProvider>(_hwnd, path, start, end, std::move(textOverride));
}

ITextRangeProvider* AccessibilityProvider::CreateTextDocumentRangeProvider(const AccessibilityControlNavigationSnapshot& record) noexcept
{
    if (! record.controlSupportsText)
    {
        return nullptr;
    }

    return CreateTextRangeProvider(record.path, 0u, record.controlAccessibleText.size());
}

[[nodiscard]] bool FindAccessibilityPathForTarget(const Control* current, const ControlPath& basePath, const Control* target, ControlPath& outPath) noexcept
{
    if (! current || ! current->IsVisible())
    {
        return false;
    }

    if (current == target && IsSemanticAccessibilityControl(current))
    {
        outPath = basePath;
        return true;
    }

    const auto* panel = dynamic_cast<const Panel*>(current);
    if (! panel)
    {
        return false;
    }

    const auto children = panel->GetChildren();
    for (size_t index = 0u; index < children.size(); ++index)
    {
        if (! children[index])
        {
            continue;
        }

        ControlPath childPath{};
        if (! TryAppendPathIndex(basePath, index, childPath))
        {
            continue;
        }

        if (FindAccessibilityPathForTarget(children[index].get(), childPath, target, outPath))
        {
            return true;
        }
    }

    return false;
}

} // namespace

bool RaiseWindowHostTextInputAutomationEvent(HWND hwnd, const Control* control, TextInputAutomationEventKind kind) noexcept
{
    if (! hwnd || ! control)
    {
        return false;
    }

    WindowHostAccessibilityTarget* target = AcquireWindowHostAccessibilityTarget(hwnd);
    if (! target)
    {
        return false;
    }

    WindowHost* const host = target->ResolveHost();
    ControlPath controlPath{};
    if (! host || ! FindAccessibilityPathForTarget(host->GetRoot(), ControlPath{}, control, controlPath))
    {
        static_cast<void>(target->Release());
        return false;
    }

    auto* provider = new (std::nothrow) AccessibilityProvider(target, hwnd, controlPath);
    if (! provider)
    {
        static_cast<void>(target->Release());
        return false;
    }
    const auto releaseProvider = wil::scope_exit([&] { static_cast<void>(provider->Release()); });

    switch (kind)
    {
        case TextInputAutomationEventKind::TextChanged:
            static_cast<void>(UiaRaiseAutomationEvent(static_cast<IRawElementProviderSimple*>(provider), UIA_Text_TextChangedEventId));
            return true;
        case TextInputAutomationEventKind::TextSelectionChanged:
            static_cast<void>(UiaRaiseAutomationEvent(static_cast<IRawElementProviderSimple*>(provider), UIA_Text_TextSelectionChangedEventId));
            return true;
        case TextInputAutomationEventKind::ActiveTextPositionChanged:
        {
            wil::com_ptr_nothrow<ITextRangeProvider> activeRange;
            NativeTextInputState state{};
            if (host->TryReadNativeTextInputState(control, state))
            {
                const size_t caretIndex = std::min(state.caretIndex, state.text.size());
                static_cast<void>(target->AddRef());
                auto* rangeProvider = new (std::nothrow) AccessibilityTextRangeProvider(target, hwnd, controlPath, caretIndex, caretIndex);
                if (rangeProvider)
                {
                    activeRange.attach(static_cast<ITextRangeProvider*>(rangeProvider));
                }
                else
                {
                    static_cast<void>(target->Release());
                }
            }
            static_cast<void>(UiaRaiseActiveTextPositionChangedEvent(static_cast<IRawElementProviderSimple*>(provider), activeRange.get()));
            return true;
        }
        case TextInputAutomationEventKind::TextEditCompositionChanged:
        {
            unique_safearray changedData(SafeArrayCreateVector(VT_BSTR, 0u, 0u));
            static_cast<void>(
                UiaRaiseTextEditTextChangedEvent(static_cast<IRawElementProviderSimple*>(provider), TextEditChangeType_Composition, changedData.get()));
            return true;
        }
        case TextInputAutomationEventKind::TextEditConversionTargetChanged:
            static_cast<void>(UiaRaiseAutomationEvent(static_cast<IRawElementProviderSimple*>(provider), UIA_TextEdit_ConversionTargetChangedEventId));
            return true;
        default: return false;
    }
}

void RegisterWindowHostAccessibilityTarget(HWND hwnd, WindowHost* host) noexcept
{
    if (hwnd && host)
    {
        const std::scoped_lock lock(GetAccessibilityTargetMutex());
        if (auto* target = static_cast<WindowHostAccessibilityTarget*>(GetPropW(hwnd, kWindowHostPropName)))
        {
            target->host.store(host, std::memory_order_release);
            PublishWindowHostAccessibilitySnapshot(*target, *host);
            return;
        }

        auto* target = new (std::nothrow) WindowHostAccessibilityTarget(hwnd, host);
        if (! target)
        {
            return;
        }

        if (SetPropW(hwnd, kWindowHostPropName, target) == 0)
        {
            static_cast<void>(target->Release());
            return;
        }

        PublishWindowHostAccessibilitySnapshot(*target, *host);
    }
}

void UnregisterWindowHostAccessibilityTarget(HWND hwnd, WindowHost* host) noexcept
{
    if (! hwnd)
    {
        return;
    }

    {
        const std::scoped_lock lock(GetAccessibilityTargetMutex());
        auto* target = static_cast<WindowHostAccessibilityTarget*>(GetPropW(hwnd, kWindowHostPropName));
        if (! target)
        {
            return;
        }

        const WindowHost* current = target->host.load(std::memory_order_acquire);
        if (host && current != host)
        {
            return;
        }

        PublishEmptyAccessibilitySnapshot(*target);
        target->host.store(nullptr, std::memory_order_release);
        if (RemovePropW(hwnd, kWindowHostPropName) == target)
        {
            static_cast<void>(target->Release());
        }
    }

    // Retire the provider map only after dropping the target mutex. UI Automation may otherwise
    // retain an earlier provider when Win32 recycles this HWND for a later DxUi host.
    static_cast<void>(UiaReturnRawElementProvider(hwnd, 0, 0, nullptr));
}

void NotifyWindowHostAccessibilityDestroyed(HWND hwnd) noexcept
{
    if (hwnd)
    {
        // UI Automation requires the provider map to be retired while handling window destruction.
        // Unregister repeats this as a fallback for owners that detach before WM_DESTROY.
        static_cast<void>(UiaReturnRawElementProvider(hwnd, 0, 0, nullptr));
    }
}

void RefreshWindowHostAccessibilitySnapshot(HWND hwnd, WindowHost* host) noexcept
{
    if (! hwnd || ! host)
    {
        return;
    }

    const std::scoped_lock lock(GetAccessibilityTargetMutex());
    auto* target = static_cast<WindowHostAccessibilityTarget*>(GetPropW(hwnd, kWindowHostPropName));
    if (! target || target->host.load(std::memory_order_acquire) != host)
    {
        return;
    }

    PublishWindowHostAccessibilitySnapshot(*target, *host);
}

void PublishEmptyWindowHostAccessibilitySnapshot(HWND hwnd, WindowHost* host) noexcept
{
    if (! hwnd)
    {
        return;
    }

    const std::scoped_lock lock(GetAccessibilityTargetMutex());
    auto* target = static_cast<WindowHostAccessibilityTarget*>(GetPropW(hwnd, kWindowHostPropName));
    if (! target)
    {
        return;
    }

    const WindowHost* current = target->host.load(std::memory_order_acquire);
    if (host && current != host)
    {
        return;
    }

    PublishEmptyAccessibilitySnapshot(*target);
}

LRESULT ReturnWindowHostAccessibilityProvider(HWND hwnd, WPARAM wp, LPARAM lp) noexcept
{
    WindowHostAccessibilityTarget* target = AcquireWindowHostAccessibilityTarget(hwnd);
    if (! target || target->ResolveHost() == nullptr)
    {
        if (target)
        {
            static_cast<void>(target->Release());
        }
        return 0;
    }

    auto* provider = new (std::nothrow) AccessibilityProvider(target, hwnd);
    if (! provider)
    {
        static_cast<void>(target->Release());
        return 0;
    }

    const LRESULT result = UiaReturnRawElementProvider(hwnd, wp, lp, static_cast<IRawElementProviderSimple*>(provider));
    static_cast<void>(provider->Release());
    return result;
}

bool TryHandleWindowHostAccessibilityMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, LRESULT& outResult) noexcept
{
    outResult = 0;
    if (msg == WM_GETOBJECT)
    {
        if (lp != static_cast<LPARAM>(UiaRootObjectId))
        {
            return false;
        }

        const LRESULT providerResult = ReturnWindowHostAccessibilityProvider(hwnd, wp, lp);
        if (providerResult == 0)
        {
            return false;
        }

        outResult = providerResult;
        return true;
    }

#if defined(ENABLE_TESTS)
    if (msg == kWindowHostAccessibilityCreateProviderMessage)
    {
        auto** provider = reinterpret_cast<IRawElementProviderFragmentRoot**>(lp);
        if (provider)
        {
            *provider = CreateWindowHostAccessibilityProvider(hwnd);
        }
        return true;
    }
#endif

    if (msg != kWindowHostAccessibilityActionMessage)
    {
        return false;
    }

    auto payload = TakeMessagePayload<AccessibilityUiActionPayload>(lp);
    if (! payload || ! payload->dispatch)
    {
        return true;
    }

    std::shared_ptr<AccessibilityUiActionDispatch> dispatch = payload->dispatch;

#if defined(ENABLE_TESTS)
    MaybeStallAccessibilityUiActionHandlerForTest();
#endif

    AccessibilityUiActionDispatch::State expected = AccessibilityUiActionDispatch::State::Pending;
    if (! dispatch->state.compare_exchange_strong(
            expected, AccessibilityUiActionDispatch::State::Taken, std::memory_order_acq_rel, std::memory_order_acquire))
    {
        return true;
    }

#if defined(ENABLE_TESTS)
    MaybeStallTakenAccessibilityUiActionHandlerForTest();
    g_accessibilityUiActionExecutionCount.fetch_add(1u, std::memory_order_relaxed);
#endif

    auto& request = dispatch->request;

    if (request.provider)
    {
        request.result = request.provider->ExecuteUiThreadAction(request);
        return true;
    }
    if (request.textRangeProvider && request.kind == AccessibilityUiActionKind::Select)
    {
        request.result = request.textRangeProvider->ExecuteSelectOnWindowThread();
        return true;
    }
    if (request.textRangeProvider && request.kind == AccessibilityUiActionKind::MoveTextRangeByVisualLine)
    {
        request.result = request.textRangeProvider->ExecuteMoveByVisualLineOnWindowThread(request.textRangeStart,
                                                                                          request.textRangeEnd,
                                                                                          request.textRangeMoveCount,
                                                                                          request.textRangeResultStart,
                                                                                          request.textRangeResultEnd,
                                                                                          request.textRangeMoved);
        return true;
    }
    if (request.textRangeProvider && request.kind == AccessibilityUiActionKind::MoveTextRangeEndpointByVisualLine)
    {
        request.result = request.textRangeProvider->ExecuteMoveEndpointByVisualLineOnWindowThread(request.textRangeStart,
                                                                                                  request.textRangeEnd,
                                                                                                  request.textRangeEndpoint,
                                                                                                  request.textRangeMoveCount,
                                                                                                  request.textRangeResultStart,
                                                                                                  request.textRangeResultEnd,
                                                                                                  request.textRangeMoved);
        return true;
    }
    if (request.textRangeProvider && request.kind == AccessibilityUiActionKind::ResolveTextRangeBounds)
    {
        request.result = request.textRangeProvider->ExecuteResolveBoundsOnWindowThread(
            request.textRangeStart, request.textRangeEnd, request.textRangeBoundsDip, request.textRangeDipToPixelScale);
        return true;
    }

    request.result = UIA_E_NOTSUPPORTED;
    return true;
}

#if defined(ENABLE_TESTS)
IRawElementProviderFragmentRoot* CreateWindowHostAccessibilityProvider(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return nullptr;
    }

    const DWORD windowThreadId = GetWindowThreadProcessId(hwnd, nullptr);
    if (windowThreadId == 0u)
    {
        return nullptr;
    }
    if (windowThreadId != GetCurrentThreadId())
    {
        // This test-only factory must resolve the live host on its owning thread. SendMessageW is
        // synchronous, so the handler cannot outlive this output slot.
        IRawElementProviderFragmentRoot* provider = nullptr;
        static_cast<void>(SendMessageW(hwnd,
                                       kWindowHostAccessibilityCreateProviderMessage,
                                       0,
                                       reinterpret_cast<LPARAM>(&provider)));
        return provider;
    }

    WindowHostAccessibilityTarget* target = AcquireWindowHostAccessibilityTarget(hwnd);
    if (! target || target->ResolveHost() == nullptr)
    {
        if (target)
        {
            static_cast<void>(target->Release());
        }
        return nullptr;
    }

    auto* provider = new (std::nothrow) AccessibilityProvider(target, hwnd);
    if (! provider)
    {
        static_cast<void>(target->Release());
        return nullptr;
    }
    return provider ? static_cast<IRawElementProviderFragmentRoot*>(provider) : nullptr;
}

void DebugSetAccessibilityUiActionHandlerStallForTest(HANDLE enteredEvent, HANDLE releaseEvent) noexcept
{
    g_accessibilityUiActionHandlerEnteredEvent.store(enteredEvent, std::memory_order_release);
    g_accessibilityUiActionHandlerReleaseEvent.store(releaseEvent, std::memory_order_release);
}

void DebugSetAccessibilityUiActionHandlerTakenStallForTest(HANDLE enteredEvent, HANDLE releaseEvent) noexcept
{
    g_accessibilityUiActionHandlerTakenEnteredEvent.store(enteredEvent, std::memory_order_release);
    g_accessibilityUiActionHandlerTakenReleaseEvent.store(releaseEvent, std::memory_order_release);
}

void DebugSetAccessibilityUiActionPostedEventForTest(HANDLE postedEvent) noexcept
{
    g_accessibilityUiActionPostedEvent.store(postedEvent, std::memory_order_release);
}

void DebugSetAccessibilityUiActionDispatchTimeoutForTest(DWORD timeoutMs) noexcept
{
    g_accessibilityUiActionDispatchTimeoutOverrideMs.store(timeoutMs, std::memory_order_release);
}

void DebugResetAccessibilityUiActionExecutionCountForTest() noexcept
{
    g_accessibilityUiActionExecutionCount.store(0u, std::memory_order_release);
}

uint32_t DebugGetAccessibilityUiActionExecutionCountForTest() noexcept
{
    return g_accessibilityUiActionExecutionCount.load(std::memory_order_acquire);
}

void DebugSetAccessibilityOffscreenSelectedRowMaterializationLimitForTest(size_t limit) noexcept
{
    g_accessibilityOffscreenSelectedRowMaterializationLimitOverride.store(limit, std::memory_order_release);
}
#endif
} // namespace RedSalamander::DxUi
